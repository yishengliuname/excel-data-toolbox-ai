"""Exercise contracts through the existing planner, HTTP worker and XLSX writer."""
import http.client
import io
import json
import threading
import time
from types import SimpleNamespace

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from excel_data_toolbox.business_kernel import numbers
from excel_data_toolbox.hr_report import build_hr_management_report
from excel_data_toolbox.nl_agent import AgentStep, _execute_step
from excel_data_toolbox.session_registry import SessionProxy, SessionRegistry
from excel_data_toolbox.task_contract import TaskScope
from excel_data_toolbox.task_engine import PersistentJobEngine
from excel_data_toolbox.tests.test_inventory_report import _inventory_tables
from excel_data_toolbox.tests.test_hr_report import _hr_tables
from excel_data_toolbox.workbook_fidelity import _write_frame, preserve_workbook_export


def test_actual_http_scope_to_workbook_download_and_versioned_reuse(monkeypatch, tmp_path):
    import excel_data_toolbox.server as server
    import excel_data_toolbox.trusted_automation as trusted
    monkeypatch.setattr(server, 'TASK_REPOSITORY', server.TaskRepository(tmp_path / 'tasks'))
    monkeypatch.setattr(server, '_project_ai_config', lambda: {'configured': False, 'model': 'local'})
    monkeypatch.setattr(server.DeepSeekClient, 'classify_unified_request', lambda *a, **k: pytest.fail('no external API'))
    session = server.AppSession()
    raw = pd.DataFrame({'订单号': list('ABCD'), '日期': ['2026-08-01', '2026-08-02', '2026-09-01', '2026-09-02'],
                        '地区': ['华东', '华西', '华东', '华西'], '销售额': [100, 200, 300, 400]})
    table_id = session.add_table('销售流水', raw.copy(deep=True), source='synthetic.xlsx', original=True)
    registry = SessionRegistry(lambda i: server.AppSession(restore_task_id=i), session)
    monkeypatch.setattr(server, 'SESSION_REGISTRY', registry)
    monkeypatch.setattr(server, 'SESSION', SessionProxy(registry))
    engine = PersistentJobEngine(tmp_path / 'jobs.sqlite', workers=1)
    engine.register('business_autopilot', server._job_business_autopilot)
    monkeypatch.setattr(server, 'JOB_ENGINE', engine)
    engine.start()
    httpd = server.ThreadingHTTPServer(('127.0.0.1', 0), server.ToolboxHandler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    connection = http.client.HTTPConnection('127.0.0.1', httpd.server_address[1], timeout=30)

    def request(path, payload=None):
        connection.request('GET' if payload is None else 'POST', path,
            None if payload is None else json.dumps(payload, ensure_ascii=False).encode(),
            {'Content-Type': 'application/json', 'X-Task-ID': session.task_id})
        response = connection.getresponse()
        body = response.read()
        assert response.status == 200, body.decode(errors='replace')
        return json.loads(body) if 'application/json' in response.headers.get('Content-Type', '') else body

    def finish(job):
        deadline = time.monotonic() + 35
        while job['status'] in {'queued', 'running'} and time.monotonic() < deadline:
            time.sleep(.05)
            job = next(j for j in request('/api/jobs')['jobs'] if j['job_id'] == job['job_id'])
        assert job['status'] == 'completed', job
        return job

    goal = {'prompt': '只分析2026年8月华东地区的销售额、排名和月度趋势，生成Excel报表。', 'table_ids': [table_id]}
    try:
        assert request('/health')['business_kernel'] == 'contracts-v1'
        job = finish(request('/api/ai-run', goal)['job'])
        result = job['result']
        assert result['execution_performed'] and result['download_url'], result
        execution = result['execution']
        assert execution['constraint_checks'] and all(c['status'] == 'passed' for c in execution['constraint_checks'])
        book = load_workbook(io.BytesIO(request(result['download_url'])))
        try:
            metrics = {r[0]: r[1] for r in book['管理层通用总览'].values}
            assert metrics['核心指标：销售额'] == 100
            excluded = list(book['范围排除明细'].values)
            assert len(excluded) >= 4  # header and the three excluded source events
            assert '任务范围与验收' in book.sheetnames
            assert any(ws._charts for ws in book.worksheets)
        finally:
            book.close()
        pd.testing.assert_frame_equal(session.get(table_id).frame, raw)
        repeated = request('/api/ai-run', goal)['job']
        assert repeated['job_id'] == job['job_id']
        # Same source and request must not resurrect an old result after a kernel upgrade.
        old_revision = trusted.engine_revision()
        monkeypatch.setattr(trusted, 'engine_revision', lambda: old_revision + '-changed')
        replacement = finish(request('/api/ai-run', goal)['job'])
        assert replacement['job_id'] != job['job_id']
        pd.testing.assert_frame_equal(session.get(table_id).frame, raw)
        context = SimpleNamespace(task_id=session.task_id, job_id='never-execute', raise_if_cancelled=lambda: None)
        with pytest.raises(server.ApiError, match='内核.*更新'):
            server._job_business_autopilot(context, {'goal': goal, 'signatures': {}, 'engine_revision': old_revision})
        checkpoint = session.task_dir / 'autopilot' / f"{replacement['job_id']}.json"
        checkpoint.write_text('{"phase":"executing"}', encoding='utf-8')
        monkeypatch.setattr(trusted, 'engine_revision', lambda: old_revision + '-changed-again')
        connection.request('POST', '/api/ai-run', json.dumps(goal, ensure_ascii=False).encode(),
                           {'Content-Type': 'application/json', 'X-Task-ID': session.task_id})
        response = connection.getresponse()
        body = response.read().decode()
        assert response.status == 409 and '中断' in body
        assert len(engine.list(session.task_id)) == 2
    finally:
        connection.close()
        httpd.shutdown()
        httpd.server_close()
        thread.join(timeout=5)
        engine.stop(timeout=5)


def test_inventory_period_flows_do_not_discard_prior_stock_movements():
    frames = _inventory_tables()
    july = frames[2].iloc[[0]].copy()
    july.loc[:, '入库日期'], july.loc[:, '入库单号'], july.loc[:, '入库数量'] = '2026/07/25', 'JULY', 7
    frames[2] = pd.concat([frames[2], july], ignore_index=True)
    scope = TaskScope(start='2026-08-01', end_exclusive='2026-09-01')
    step = AgentStep('inventory', 'inventory_management_report', tuple(str(i) for i in range(6)), '库存报告',
                     {'source_names': ['商品资料', '期初库存', '采购入库', '销售出库', '库存调整', '仓库说明'], 'request_scope': scope.to_dict()})
    outputs, report = _execute_step(step, frames)
    detail = outputs['商品库存分析'].set_index('商品编码')
    assert detail.loc['P001', '当前账面库存'] == 11  # opening20 + July7 + Aug10 - sales25 - adjustment1
    assert report['period_start'] == '2026-08-01'
    assert report['period_purchase_amount'] == 100  # excludes July receipt, which remains in closing stock
    assert any(c['status'] == 'needs_review' for c in report['constraint_checks'])  # source stops Aug23


def test_future_unknown_quantity_does_not_poison_current_stock():
    from excel_data_toolbox.inventory_report import build_inventory_management_report
    frames = _inventory_tables()
    future = frames[3].iloc[[0]].copy()
    future.loc[:, '出库日期'], future.loc[:, '出库单号'] = '2026/09/01', 'FUTURE'
    future['数量'] = '待核对'
    frames[3] = pd.concat([frames[3], future], ignore_index=True)
    result = build_inventory_management_report(frames, source_names=['商品资料', '期初库存', '采购入库', '销售出库', '库存调整', '仓库说明'])
    assert result.outputs['商品库存分析'].set_index('商品编码').loc['P001', '当前账面库存'] == 4


def test_invalid_payroll_amount_is_unknown_while_percentage_target_remains_valid():
    frames = _hr_tables()
    frames[3]['金额'] = frames[3]['金额'].astype(object)
    frames[3].loc[0, '金额'] = '10%'
    frames[2]['完成目标比例'] = frames[2]['完成目标比例'].astype(object)
    frames[2].loc[0, '完成目标比例'] = '115%'
    result = build_hr_management_report(frames, source_names=['员工基础信息', '月度考勤记录', '绩效数据', '薪资调整记录', '老板需求说明'])
    detail = result.outputs['员工综合分析'].set_index('员工编号')
    assert detail.loc['E002', '目标完成率'] == 1.15
    assert pd.isna(detail.loc['E001', '预计薪资'])
    assert result.report['expected_active_payroll'] is None
    json.dumps(result.report, allow_nan=False)


@pytest.mark.parametrize('value', ['1,2', '1.234,50'])
def test_ambiguous_locale_numbers_are_not_guessed(value):
    assert numbers(pd.Series([value])).isna().all()


def test_fidelity_source_cannot_be_overwritten_and_formulas_are_not_flattened(tmp_path):
    book = Workbook()
    book.active.append(['金额', '计算结果'])
    book.active.append([100, '=A2*2'])
    source = tmp_path / 'source.xlsx'
    book.save(source)
    original = source.read_bytes()
    with pytest.raises(ValueError, match='覆盖原始'):
        preserve_workbook_export(source, source, {'结果': pd.DataFrame({'金额': [1]})})
    with pytest.raises(ValueError, match='公式'):
        _write_frame(book.active, pd.DataFrame({'金额': [1]}), replace_existing=True)
    assert book.active['B2'].value == '=A2*2'
    assert source.read_bytes() == original
    book.close()


def test_cross_fact_scope_cannot_destroy_refund_attribution_before_join():
    step = AgentStep('report', 'enterprise_diagnosis_report', ('source',), '报表',
                     {'request_scope': TaskScope(start='2026-08-01', end_exclusive='2026-09-01').to_dict()})
    with pytest.raises(ValueError, match='退款追溯'):
        _execute_step(step, [])


def test_engine_revision_can_verify_frozen_bytecode(monkeypatch):
    from pathlib import Path
    from excel_data_toolbox.trusted_automation import engine_revision
    original = Path.is_file
    monkeypatch.setattr(Path, 'is_file', lambda p: False if p.name == 'business_kernel.py' else original(p))
    revision = engine_revision()
    assert len(revision) == 64 and all(c in '0123456789abcdef' for c in revision)


@pytest.mark.parametrize('kind', ['tasks', 'schedules'])
def test_runtime_transactions_close_handles_and_roll_back_failures(monkeypatch, tmp_path, kind):
    import sqlite3
    from excel_data_toolbox.task_store import TaskRepository
    from excel_data_toolbox.scheduler import LocalScheduler
    connections = []
    original = sqlite3.connect

    class TrackedConnection(sqlite3.Connection):
        closed = False
        def close(self):
            self.closed = True
            super().close()

    def connect(*args, **kwargs):
        connection = original(*args, **kwargs, factory=TrackedConnection)
        connections.append(connection)
        return connection

    monkeypatch.setattr(sqlite3, 'connect', connect)
    store = TaskRepository(tmp_path / 'tasks') if kind == 'tasks' else LocalScheduler(tmp_path / 'schedules.sqlite')
    with store._connect() as connection:
        connection.execute('CREATE TABLE probe (value INTEGER)')
    with pytest.raises(RuntimeError, match='rollback'):
        with store._connect() as connection:
            connection.execute('INSERT INTO probe VALUES (1)')
            raise RuntimeError('rollback')
    with store._connect() as connection:
        assert connection.execute('SELECT COUNT(*) FROM probe').fetchone()[0] == 0
    assert all(connection.closed for connection in connections)


def test_explicit_expiry_cleanup_preserves_loaded_and_queued_tasks(tmp_path):
    from datetime import datetime, timedelta
    from excel_data_toolbox.task_store import TaskRepository
    repo = TaskRepository(tmp_path / 'tasks', retention_days=1)
    for task_id in ('ACTIVE-1', 'QUEUED-1', 'EXPIRED-1'):
        repo.create(task_id, 'synthetic')
    removed = repo.purge_expired(now=datetime.now() + timedelta(days=2), exclude={'ACTIVE-1'},
                                is_active=lambda task_id: task_id == 'QUEUED-1')
    assert removed == ['EXPIRED-1']
    assert repo.task_dir('ACTIVE-1').is_dir() and repo.task_dir('QUEUED-1').is_dir()


def test_server_startup_never_implicitly_deletes_customer_history(monkeypatch):
    import excel_data_toolbox.server as server
    def interrupted(**kwargs):
        raise KeyboardInterrupt()
    httpd = SimpleNamespace(serve_forever=interrupted, server_close=lambda: None)
    noop = SimpleNamespace(start=lambda **kwargs: None, stop=lambda: None)
    monkeypatch.setattr(server, 'find_available_port', lambda preferred: preferred)
    monkeypatch.setattr(server, 'ThreadingHTTPServer', lambda *args: httpd)
    monkeypatch.setattr(server, 'SCHEDULER', noop)
    monkeypatch.setattr(server, 'JOB_ENGINE', noop)
    monkeypatch.setattr(server, 'SESSION', SimpleNamespace(close=lambda: None))
    monkeypatch.setattr(server, 'TASK_REPOSITORY', SimpleNamespace(purge_expired=lambda: pytest.fail('startup cannot purge')))
    server.run_server(port=8502, open_browser=False)
