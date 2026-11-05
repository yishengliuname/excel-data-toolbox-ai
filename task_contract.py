"""Bound explicit request scope to observed fields; never execute a model's code.

This small compiler supports unambiguous periods and exact entity restrictions.
Unsupported/ambiguous scope stops rather than silently analysing everything.
"""
from __future__ import annotations

from dataclasses import dataclass, asdict
import re

import pandas as pd

from .business_kernel import BusinessInvariantError
from .metric_semantics import classify_metric, classify_sheet_role


_RESTRICT = re.compile(r'只(?:分析|看|统计|保留)|仅(?:分析|看|统计|限|保留)|限定|筛选|范围为|限于|\bonly\b', re.I)
_ENTITY = re.compile(r'地区|区域|门店|店铺|渠道|平台|部门|客户|产品|商品|region|store|channel|department|customer|product', re.I)


@dataclass(frozen=True)
class TaskScope:
    start: str = ''
    end_exclusive: str = ''
    entities: tuple[tuple[str, str], ...] = ()
    top_n: int | None = None

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, value):
        if not isinstance(value, dict) or set(value) != {'start', 'end_exclusive', 'entities', 'top_n'}:
            raise BusinessInvariantError('任务范围契约结构无效')
        if not all(isinstance(value[k], str) for k in ('start', 'end_exclusive')):
            raise BusinessInvariantError('范围日期必须是字符串')
        if bool(value['start']) != bool(value['end_exclusive']):
            raise BusinessInvariantError('范围日期须同时提供起止边界')
        try:
            if value['start']:
                if pd.isna(pd.Timestamp(value['start'])):
                    raise ValueError('无效日期')
            if value['end_exclusive']:
                if pd.isna(pd.Timestamp(value['end_exclusive'])):
                    raise ValueError('无效日期')
            if value['start'] and value['end_exclusive'] and pd.Timestamp(value['start']) >= pd.Timestamp(value['end_exclusive']):
                raise ValueError('日期顺序')
        except (TypeError, ValueError):
            raise BusinessInvariantError('范围日期无效') from None
        entities = value['entities']
        if not isinstance(entities, (tuple, list)) or len(entities) > 5 or any(not isinstance(p, (list, tuple)) or len(p) != 2 or not all(isinstance(v, str) and v for v in p) for p in entities):
            raise BusinessInvariantError('实体范围必须是明确字段和值')
        top = value['top_n']
        if top is not None and (type(top) is not int or not 1 <= top <= 30):
            raise BusinessInvariantError('排名范围无效')
        return cls(value['start'], value['end_exclusive'], tuple(tuple(p) for p in entities), top)


def compile_scope(request: str, frames, names, contexts=None) -> TaskScope:
    text = str(request or '')
    contexts = contexts or {}
    if re.search(r'不(?:统计|分析|包含)|不要统计|排除[^，。；;\n]{0,4}(?:地区|门店|渠道|平台|部门)|\bexclude\b', text, re.I):
        raise BusinessInvariantError('排除范围需要显式确认字段、值及否定条件，不能当作全量分析继续')
    start = end = ''
    # A full date range is stronger than a month. Bounds are half-open so the
    # entire final day, including timestamps, belongs to the selected period.
    day = r'(20\d{2})[-/年](\d{1,2})[-/月](\d{1,2})日?'
    days = list(dict.fromkeys(re.findall(day, text)))
    if days:
        if len(days) == 1 and re.search(r'截至|截止|as of', text, re.I):
            days = [('1900', '01', '01'), days[0]]
        if len(days) != 2 or not re.search(r'至|到|~|—|截止|as of', text, re.I):
            raise BusinessInvariantError('日期范围须明确起止；单一日期请说明当天或截止日')
        try:
            a, b = (pd.Timestamp(year=int(y), month=int(m), day=int(d)) for y, m, d in days)
        except ValueError:
            raise BusinessInvariantError('需求日期无效') from None
        if a > b:
            raise BusinessInvariantError('需求起止日期顺序无效')
        start, end = a.isoformat(), (b + pd.Timedelta(days=1)).isoformat()
    else:
        months = list(dict.fromkeys(re.findall(r'(20\d{2})年(\d{1,2})月|(?<!\d)(20\d{2})-(\d{2})(?![-\d])', text)))
        if not months:
            short = set(re.findall(r'(?<![\d年])(\d{1,2})月(?:份)?', text))
            if short:
                years = set()
                for f in frames:
                    for c in f:
                        if classify_metric(c).kind == 'date':
                            years.update(pd.to_datetime(f[c], errors='coerce', format='mixed').dropna().dt.year.unique())
                if len(short) != 1 or len(years) != 1:
                    raise BusinessInvariantError('需求月份缺少年份或包含多个期间，须确认报告年份和范围')
                months = [(str(next(iter(years))), next(iter(short)), '', '')]
        if months:
            if len(months) != 1:
                raise BusinessInvariantError('多个需求期间须明确比较或合并口径，不能自动选择一个月份')
            y1, m1, y2, m2 = months[0]
            try:
                a = pd.Timestamp(year=int(y1 or y2), month=int(m1 or m2), day=1)
            except ValueError:
                raise BusinessInvariantError('需求月份无效') from None
            start, end = a.isoformat(), (a + pd.offsets.MonthBegin(1)).isoformat()
    entities = []
    clauses = re.split(r'[，。；;\n]', text)
    restricted_clauses = [c for c in clauses if _RESTRICT.search(c)]
    for clause in clauses:
        if clause in restricted_clauses or re.search(r'比较|对比|各地区|所有|全部', clause):
            continue
        for frame in frames:
            for column in frame:
                if _ENTITY.search(str(column)) and classify_metric(column).kind not in {'date', 'additive', 'ratio', 'count'}:
                    for value in frame[column].dropna().astype(str).str.strip().unique():
                        if value and re.search(r'(?:分析|统计|汇总|计算|查看)\s*(?:(?:20\d{2}年)?\d{1,2}月)?\s*' + re.escape(value), clause):
                            restricted_clauses.append(clause)
                            break
    for clause in dict.fromkeys(restricted_clauses):
        candidates = []
        for frame in frames:
            for column in frame:
                if not _ENTITY.search(str(column)) or classify_metric(column).kind in {'date', 'additive', 'ratio', 'count'}:
                    continue
                for value in frame[column].dropna().astype(str).str.strip().unique():
                    if value and value in clause and (len(value) >= 2 or re.search(r'[A-Za-z0-9]', value)):
                        candidates.append((str(column), value))
        # Only one exact entity restriction per clause is currently supported.
        # Two selected stores need an explicit OR-set, not sequential AND filters.
        candidates = list(dict.fromkeys(candidates))
        if len({v for _, v in candidates}) > 1:
            raise BusinessInvariantError('多个对象的筛选须明确并集/交集；请在规则中确认')
        if candidates:
            field, value = candidates[0]
            same_value_fields = {c for c, v in candidates if v == value}
            if len(same_value_fields) > 1:
                raise BusinessInvariantError('筛选对象对应多个字段，请确认筛选字段')
            entities.append((field, value))
        elif re.search(r'地区|门店|渠道|平台|部门|客户|商品|region|store|channel', clause, re.I):
            raise BusinessInvariantError('需求中的筛选对象未能唯一绑定源字段和值，请确认范围')
        elif _RESTRICT.search(clause) and not start:
            raise BusinessInvariantError('限定条件未能绑定明确期间或对象，请确认筛选字段和值')
    top = re.search(r'前\s*(\d+)\s*(?:名|个|项)|\btop\s*(\d+)\b', text, re.I)
    top_n = int(next(g for g in top.groups() if g)) if top else None
    if top_n is not None and not 1 <= top_n <= 30:
        raise BusinessInvariantError('排名范围须为1至30')
    return TaskScope(start, end, tuple(dict.fromkeys(entities)), top_n)


def apply_scope(scope: TaskScope, frames, names, contexts=None):
    contexts = contexts or {}
    selected, excluded, evidence = [], [], []
    for frame, name in zip(frames, names, strict=True):
        role = contexts.get(name, {}).get('role') or classify_sheet_role(name, frame)
        mask = pd.Series(True, index=frame.index)
        if role == 'fact' and (scope.start or scope.entities):
            if scope.start:
                dates = [c for c in frame if classify_metric(c).kind == 'date']
                field = contexts.get(name, {}).get('date_field') or (dates[0] if len(dates) == 1 else '')
                if not field:
                    raise BusinessInvariantError(f'{name}：期间限制缺少唯一已确认业务日期')
                values = pd.to_datetime(frame[field], errors='coerce', format='mixed')
                if values.isna().any():
                    raise BusinessInvariantError(f'{name}：范围筛选日期缺失或无效')
                mask &= values.ge(pd.Timestamp(scope.start)) & values.lt(pd.Timestamp(scope.end_exclusive))
            for field, value in scope.entities:
                if field not in frame:
                    raise BusinessInvariantError(f'{name}：无法将筛选字段{field}绑定此事实域，禁止纳入未筛选总额')
                mask &= frame[field].astype('string').str.strip().eq(value).fillna(False)
            if not mask.any():
                raise BusinessInvariantError(f'{name}：指定范围没有有效记录，不能回退到整表统计')
        part = frame.loc[mask].copy()
        rejected = frame.loc[~mask].copy()
        if not rejected.empty:
            rejected.insert(0, '范围来源表', name)
            rejected.insert(1, '范围排除原因', '不在用户明确指定期间/对象范围内')
            excluded.append(rejected)
        selected.append(part)
        if role == 'fact' and (scope.start or scope.entities):
            evidence.append({'item': f'{name}：显式范围已执行', 'status': 'passed',
                             'input_rows': len(frame), 'included_rows': len(part), 'excluded_rows': len(rejected),
                             'scope': scope.to_dict()})
    if (scope.start or scope.entities) and not evidence:
        raise BusinessInvariantError('指定范围未能绑定任何事实表，不能将未筛选报告交付')
    return selected, excluded, evidence
