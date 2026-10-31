"""Explicit, revocable local-report grants. Customer scopes are not tenant auth.

Old saved profiles remain review-only. No model output can create a grant.
Only selected sources and confirmed rules are used; no filesystem discovery.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
import hashlib
import json
import os
import threading
import uuid

import pandas as pd

from .analysis_compiler import _intent_topics
from .business_contracts import metric_contract
from .business_rules import BusinessRuleError, prepare_business_data
from .business_workflow import BusinessProfileStore


_LOCK = threading.RLock()


def _digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                     allow_nan=False).encode('utf-8')).hexdigest()


def engine_revision():
    from pathlib import Path
    root = Path(__file__).parent
    return _digest({name: hashlib.sha256((root / name).read_bytes()).hexdigest() for name in (
        'business_rules.py', 'business_contracts.py', 'metric_semantics.py',
        'analysis_compiler.py', 'adaptive_report.py', 'trusted_automation.py')})


def validate_limits(raw):
    if not isinstance(raw, dict) or set(raw) != {'valid_days', 'max_rows', 'max_excluded_ratio'}:
        raise BusinessRuleError('自动授权须明确有效天数、最大记录数和最大排除比例')
    days, rows, ratio = (raw[k] for k in ('valid_days', 'max_rows', 'max_excluded_ratio'))
    if type(days) is not int or not 1 <= days <= 365:
        raise BusinessRuleError('授权有效期须为1至365天')
    if type(rows) is not int or not 1 <= rows <= 1_000_000:
        raise BusinessRuleError('授权最大记录数须为1至1000000')
    if type(ratio) not in (int, float) or not 0 <= ratio <= 1:
        raise BusinessRuleError('最大排除比例须在0至1之间')
    return dict(raw)


class TrustedAutomationStore(BusinessProfileStore):
    def _record(self, customer, profile_id):
        # Reuse the same strict scope/id validation without printing identities.
        if not isinstance(profile_id, str) or len(profile_id) != 32 or any(c not in '0123456789abcdef' for c in profile_id):
            raise BusinessRuleError('配置标识无效')
        path = self._scope(customer) / f'{profile_id}.json'
        if not path.is_file():
            raise BusinessRuleError('当前客户没有此配置')
        record = json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(record, dict) or record.get('profile_id') != profile_id or not isinstance(record.get('schemas'), dict):
            raise BusinessRuleError('配置文件结构无效，请重新核对保存')
        return path, record

    @staticmethod
    def _write(path, record):
        temp = path.with_name(f'.{uuid.uuid4().hex}.tmp')
        try:
            temp.write_text(json.dumps(record, ensure_ascii=False, allow_nan=False), encoding='utf-8')
            os.replace(temp, path)
        finally:
            temp.unlink(missing_ok=True)

    @staticmethod
    def _validate_data(frames, names, rules, limits):
        if sum(len(f) for f in frames) > limits['max_rows']:
            raise BusinessRuleError('记录数超过授权上限，请核对数据规模')
        prepared = prepare_business_data(frames, names, rules)
        evidence = []
        for index, (name, frame) in enumerate(zip(names, prepared.frames)):
            raw_count = len(frames[index])
            excluded = raw_count - len(frame)
            if raw_count and excluded / raw_count > limits['max_excluded_ratio']:
                raise BusinessRuleError(f'{name}：排除比例超过授权上限，请核对异常范围')
            context = prepared.contexts[name]
            if not context['status_confirmed']:
                raise BusinessRuleError(f'{name}：存在未授权的状态字段')
            date = context.get('date_field')
            if date and pd.to_datetime(frame[date], errors='coerce', format='mixed').isna().any():
                raise BusinessRuleError(f'{name}：已确认日期包含缺失或无效值')
            for field in context['metrics']:
                contract = metric_contract(index, name, frame, field, context=context)
                if contract.status != 'confirmed':
                    raise BusinessRuleError(f'{name}.{field}：' + ('；'.join(contract.issues) or '指标未形成已确认契约'))
            evidence.append({'source': name, 'input_rows': raw_count, 'included_rows': len(frame),
                             'excluded_rows': excluded, 'confirmed_joins': context.get('confirmed_joins', [])})
        return evidence

    def save_authorized(self, customer, title, rules, frames, names, request, limits):
        limits = validate_limits(limits)
        self._validate_data(frames, names, rules, limits)
        with _LOCK:
            saved = super().save(customer, title, rules, frames, names)
            path, record = self._record(customer, saved['profile_id'])
            now = datetime.now(timezone.utc)
            record['automation'] = {
                'enabled': True, 'revision': uuid.uuid4().hex,
                'operation': 'adaptive_analysis_report', 'approved_at': now.isoformat(),
                'expires_at': (now + timedelta(days=limits['valid_days'])).isoformat(),
                'limits': limits, 'engine_revision': engine_revision(),
                'contract_digest': _digest({'rules': record['rules'], 'schemas': record['schemas']}),
                'topics': sorted(set(_intent_topics(request)) | set(rules.get('requirements', []))),
            }
            self._write(path, record)
            return {**saved, 'automation': record['automation']}

    def list(self, customer):
        with _LOCK:
            profiles = super().list(customer)
            for profile in profiles:
                _, record = self._record(customer, profile['profile_id'])
                grant = record.get('automation', {})
                profile['automation_enabled'] = grant.get('enabled') is True
                profile['expires_at'] = grant.get('expires_at')
            return profiles

    def revoke(self, customer, profile_id):
        with _LOCK:
            path, record = self._record(customer, profile_id)
            if record.get('automation'):
                record['automation']['enabled'] = False
                record['automation']['revision'] = uuid.uuid4().hex
                self._write(path, record)
            return {'profile_id': profile_id, 'automation_enabled': False}

    def authorized_rules(self, customer, profile_id, frames, names, request, revision=None):
        with _LOCK:
            _, record = self._record(customer, profile_id)
            grant = record.get('automation', {})
            if not isinstance(grant, dict):
                raise BusinessRuleError('自动授权结构无效，请重新确认')
            if grant.get('enabled') is not True or grant.get('operation') != 'adaptive_analysis_report':
                raise BusinessRuleError('配置没有自动执行授权或授权已撤销，请重新核对')
            if revision is not None and revision != grant.get('revision'):
                raise BusinessRuleError('授权在排队期间变化，请重新提交')
            try:
                expires = datetime.fromisoformat(grant['expires_at'])
                if expires.tzinfo is None or not isinstance(grant.get('topics'), list) or not all(isinstance(t, str) for t in grant['topics']):
                    raise ValueError('expiry/topics')
            except (KeyError, TypeError, ValueError):
                raise BusinessRuleError('授权有效期或主题结构无效，请重新确认') from None
            if expires <= datetime.now(timezone.utc):
                raise BusinessRuleError('自动执行授权已过期，请重新确认')
            if grant.get('engine_revision') != engine_revision():
                raise BusinessRuleError('执行引擎已更新，请重新核对规则后授权')
            if grant.get('contract_digest') != _digest({'rules': record['rules'], 'schemas': record['schemas']}):
                raise BusinessRuleError('保存的规则或字段结构已变化，请重新核对')
            extra = set(_intent_topics(request)) - set(grant['topics'])
            if extra:
                raise BusinessRuleError('需求超出授权主题：' + '、'.join(sorted(extra)))
            rules = self.load_for_review(customer, profile_id, frames, names)
            evidence = self._validate_data(frames, names, rules, validate_limits(grant['limits']))
            return rules, {'profile_id': profile_id, 'revision': grant['revision'],
                           'expires_at': grant['expires_at'], 'operation': grant['operation'],
                           'source_checks': evidence, 'boundary': '仅本地规则报表；不授权付款、采购、发布或人事处置'}

    @contextmanager
    def execution_guard(self, *args, **kwargs):
        # Revoke cannot race the last check and execution in this local process.
        with _LOCK:
            yield self.authorized_rules(*args, **kwargs)
