"""Bounded local delivery orchestration, independent of business templates.

Planning never grants approval for unclear rules or external actions. The
orchestrator reuses validated local plans, preserves partial deliveries, and
requires an actual downloadable artifact before claiming workbook delivery.
"""
from __future__ import annotations

from typing import Callable, Mapping, Any


def run_delivery(*, plan: Callable[[], dict], execute: Callable[[str], dict],
                 export: Callable[[dict], dict], progress: Callable[[int, str], None],
                 before_execution: Callable[[], None] = lambda: None,
                 accept: Callable[[dict, dict], dict] | None = None) -> dict[str, Any]:
    progress(15, '正在理解需求、绑定本次输入并检查可执行规则')
    planned = plan_with_recovery(plan, progress)
    mode = planned.get('mode')
    if planned.get('status') != 'ready' or mode not in {'data', 'data_then_chart', 'chart'}:
        return {**planned, 'auto_execute': False, 'delivery_status': 'needs_input',
                'execution_performed': False}
    if mode == 'chart':
        return {**planned, 'delivery_status': 'chart_ready', 'execution_performed': False}
    if planned.get('auto_execute') is not True or not planned.get('plan_token'):
        return {**planned, 'auto_execute': False, 'delivery_status': 'needs_approval',
                'execution_performed': False}
    before_execution()
    progress(45, '正在本机执行受约束计划，保留原始数据与异常证据')
    result = execute(planned['plan_token'])
    progress(85, '正在检查实际交付文件与下载入口')
    if not result.get('download_url'):
        result = {**result, **export(result)}
    if not result.get('download_url'):
        raise ValueError('执行完成但没有可下载交付文件，不能标记交付成功')
    if accept is not None:
        result['task_acceptance'] = accept(planned, result)
        if result['task_acceptance']['status'] != 'passed':
            result['status'] = 'needs_review'
            result['message'] = result.get('message', '') + '；任务级验收未全部通过，已保留可用结果和证据缺口'
    # Tokens are one-shot in-memory approval capabilities, not durable outputs.
    public_plan = {k: v for k, v in planned.items() if k not in {'plan_token', 'dry_run'}}
    return {**public_plan, 'auto_execute': False, 'execution': result,
            'status': result.get('status', 'needs_review'),
            'message': result.get('message', '交付文件已生成'),
            'download_url': result['download_url'], 'execution_performed': True,
            **({'task_acceptance': result['task_acceptance']} if 'task_acceptance' in result else {}),
            'delivery_status': 'partial' if result.get('status') != 'completed' else 'delivered'}


def plan_with_recovery(plan, progress):
    """Only typed transient planning failures get one retry. Never execution."""
    attempts = []
    for number in (1, 2):
        try:
            planned = plan()
            if attempts:
                planned['planning_recovery'] = {'attempts': number, 'events': attempts}
            return planned
        except (TimeoutError, ConnectionError) as exc:
            attempts.append({'attempt': number, 'reason': type(exc).__name__})
            if number == 2:
                raise
            progress(20, '规划发生临时连接故障，重试一次；尚未执行数据处理')


def input_fingerprint(prompt: str, signatures: Mapping[str, Any]) -> str:
    """Stable identity for a goal and exact selected inputs, not a filename."""
    import hashlib
    import json
    body = json.dumps({'prompt': prompt.strip(), 'inputs': signatures},
                      sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(body.encode('utf-8')).hexdigest()
