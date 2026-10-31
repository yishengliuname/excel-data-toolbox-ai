"""Task acceptance is distinct from XLSX serialization acceptance.

This checks the declared task topics and execution evidence, not the truth of
every possible business question. Missing evidence never means a pass.
"""
from __future__ import annotations


def assess_task(result, *, expected_steps, required_topics=()):
    artifact = result.get('acceptance') or {}
    checks = [
        {'item': '原始输入未变', 'status': 'passed' if result.get('inputs_unchanged') is True else 'unverified'},
        {'item': '计划步骤执行', 'status': 'passed' if result.get('steps_completed') == expected_steps and expected_steps > 0 else 'unverified'},
        {'item': '导出文件内容验收', 'status': 'passed' if artifact.get('status') == 'passed' and result.get('download_url') else 'unverified'},
        {'item': '待核验事项', 'status': 'passed' if result.get('pending_reviews') == 0 else 'needs_review'},
    ]
    coverage = result.get('request_coverage') or []
    for topic in sorted(set(required_topics)):
        evidence = [c for c in coverage if c.get('topic') == topic]
        status = 'passed' if evidence and all(c.get('status') == 'available' for c in evidence) else 'needs_review'
        checks.append({'item': topic, 'status': status,
                       'missing': [m for c in evidence for m in c.get('missing', [])] or ([] if evidence else ['缺少需求主题验收证据'])})
    passed = all(c['status'] == 'passed' for c in checks) and result.get('status') == 'completed'
    return {'status': 'passed' if passed else 'needs_review', 'checks': checks,
            'checks_passed': sum(c['status'] == 'passed' for c in checks), 'checks_total': len(checks),
            'boundary': '检查已声明需求主题、输入和导出证据；不等同会计审计或因果证明'}
