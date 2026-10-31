"""Declared-policy, multi-fact execution built from reusable local operations.

No file discovery, customer filename, answer fixture, remote code or eval.
Only a policy sheet explicitly referenced by the request can supply rules.
Unrecognised policy grammar returns to review, rather than guessing approval.
The current grammar covers versioned sales + refunds + dated standard costs,
with optional procurement, stock, services, settlement and receivables facts.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hashlib
import json
from numbers import Real
import re
from typing import Any, Mapping, Sequence

import pandas as pd

from .status_semantics import classify_refund_status, REFUND_CONFIRMED
from .decision_summary import summarise_risks


ZERO = Decimal(0)
CENT = Decimal('0.01')
ALIASES = {
    '销售单号': '订单号', '销售日期': '营业日期', '明细序号': '行号',
    '机构': '分公司名称', '门店': '分公司名称', 'SKU': '商品编码',
    '成交实付': '实付金额', '销售数量': '数量', '优惠金额': '折扣金额',
    '成交单价': '含税单价', '状态': '订单状态',
}
# Signatures bind table roles by fields, not filenames or table positions.
SIGNATURES = {
    'sales': {'订单号', '行号', '营业日期', '商品编码', '数量', '实付金额', '订单状态'},
    'products': {'商品编码', '标准商品名称', '标准类别', '基础单位'},
    'orgs': {'分公司编码', '分公司标准名称', '允许别名'},
    'customers': {'客户编码', '客户名称', '信用额度'},
    'costs': {'商品编码', '生效开始', '生效结束', '标准基础单价'},
    'refunds': {'退款单号', '原订单号', '退款日期', '退款金额', '退款状态'},
    'returns': {'退货单号', '原订单号', '原行号', '验收日期', '验收状态'},
    'purchases': {'入库单号', '采购数量', '采购单位', '采购单价', '入库状态'},
    'opening_stock': {'商品编码', '分公司编码', '期初数量', '库存日期'},
    'adjustments': {'调整单号', '调整数量', '审批状态'},
    'counts': {'商品编码', '盘点日期', '盘点数量'},
    'settlements': {'结算单号', '结算月份', '结算基数', '佣金', '应结实际到账'},
    'opening_ar': {'应收单号', '期初应收金额', '到期日'},
    'receipts': {'回款ID', '回款金额', '到账日期', '到账状态'},
    'applications': {'核销ID', '回款ID', '业务单号', '核销金额', '业务类型'},
    'services': {'工单号', '合同金额', '直接材料费', '直接人工费', '验收或预计验收日'},
    'expenses': {'费用单号', '费用类别', '发生月份', '金额'},
    'targets': {'月份', '分公司编码', '净销售目标'},
    'workforce': {'月份', '分公司编码', '月末人数', '正常工时', '加班工时'},
    'feedback': {'反馈ID', '订单号', '反馈日期', '满意度', '主要问题'},
}


class PolicyError(ValueError):
    """Declared rules or master data are ambiguous; no silent fallback."""


def text(value: Any) -> str:
    if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
        return ''
    if isinstance(value, Real) and not isinstance(value, bool) and float(value).is_integer():
        return str(int(value))
    return str(value).strip()


def number(value: Any) -> Decimal | None:
    value = text(value).replace(',', '').replace('￥', '').replace('¥', '')
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)', value):
        return None
    try:
        result = Decimal(value)
        return result if result.is_finite() else None
    except InvalidOperation:
        return None


def date(value: Any) -> datetime | None:
    if not text(value):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    # Task persistence serialises real Excel datetimes as ISO with a T and
    # fractional seconds. A restored task must retain the same business date;
    # numeric epochs are deliberately not guessed from ordinary numeric cells.
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?', text(value)):
        try:
            return datetime.fromisoformat(text(value).replace('Z', '+00:00')).replace(
                hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
        except ValueError:
            return None
    for fmt in ('%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y', '%Y/%m/%d', '%Y-%m'):
        try:
            return datetime.strptime(text(value), fmt)
        except ValueError:
            pass
    return None


def money(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def display_number(value: Decimal | None) -> str:
    """Readable evidence text; never change the exact calculation value."""
    return '待核验' if value is None else format(value, ',.2f').rstrip('0').rstrip('.')


def strict_sum(values) -> Decimal | None:
    """Missing prerequisite propagates; a legitimate empty population is zero."""
    values = list(values)
    return None if any(v is None for v in values) else sum(values, ZERO)


def serialise(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, datetime):
        return value.isoformat(sep=' ', timespec='seconds')
    if isinstance(value, dict):
        return {k: serialise(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [serialise(v) for v in value]
    return value


def frame(rows: list[dict], columns=()) -> pd.DataFrame:
    def cell(key, value):
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, datetime):
            return value
        if key.endswith('月份') and re.fullmatch(r'\d{4}-\d{2}', text(value)):
            return value  # a month label, not an invented day-of-month
        if re.search(r'日期$|日$|月份$|生效开始$|生效结束$', key):
            parsed = date(value)
            if parsed is not None:
                return parsed
        return value
    # Retain native Excel dates/numbers in report frames. JSON serialisation
    # belongs to metadata, not the cell writer.
    return pd.DataFrame([{k: cell(k, v) for k, v in r.items()} for r in rows], columns=list(columns) or None)


@dataclass(frozen=True)
class DeclaredPolicy:
    start: datetime
    cutoff: datetime
    good_orders: tuple[str, ...]
    bad_orders: tuple[str, ...]
    source: str
    digest: str
    clauses: tuple[tuple[str, str], ...]
    aliases: Mapping[str, str]
    roles: Mapping[str, tuple[int, ...]]
    stock_multiplier: Decimal | None
    ar_ratio: Decimal | None


def compile_declared_policy(frames: Sequence[pd.DataFrame], names: Sequence[str], request: str) -> DeclaredPolicy | None:
    """A bounded recogniser for explicitly supplied rules, not row-note prompts.

    User must point to a quoted sheet for rules; mere presence of a notes tab
    never authorises it. Unknown/conflicting clauses fail closed to review.
    """
    if not re.search(r'口径|规则', request):
        return None
    mentioned = re.findall(r'[“"「]([^”"」]{1,80})[”"」]', request)
    indices = [i for i, n in enumerate(names) if n.rsplit('__', 1)[-1] in mentioned]
    if len(indices) != 1:
        return None
    policy_index = indices[0]
    notes = frames[policy_index]
    if len(notes.columns) != 2 or len(notes) > 100:
        return None
    clauses = tuple((text(r[0]), text(r[1])) for r in notes.itertuples(index=False, name=None))
    content = '\n'.join(v for _, v in clauses)
    required = ('最高版本', '冲突', '隔离', '有效日期', '实付', '生效区间',
                '按订单内有效行实付比例', '退货', '截止', '不是', '人民币')
    if not all(t in content for t in required):
        return None
    if re.search(r'执行代码|运行代码|忽略.*安全|删除原文件', content):
        return None
    periods = re.findall(r'(\d{4}-\d{2}-\d{2})\s*至\s*(\d{4}-\d{2}-\d{2})', content)
    if len(set(periods)) != 1:
        return None
    start, end = map(date, periods[0])
    if not start or not end or start > end:
        return None
    declared_cuts = re.findall(r'(?:截至|截止(?:日)?[：:]?)[^\d\n]{0,5}(\d{4}-\d{2}-\d{2})', content)
    if any(date(v) != end for v in declared_cuts):
        return None
    positive = re.search(r'正常[：:]([^。\n]+)', content)
    negative = re.search(r'(未完成[^。\n]+?)均无效', content)
    if not positive or not negative:
        return None
    good = tuple(v.strip() for v in re.split(r'[、，,]', positive[1]))
    bad = tuple(v.strip() for v in re.split(r'[、，,]', negative[1]))
    if not good or set(good) & set(bad):
        return None
    aliases = dict(ALIASES)
    mapping_text = '\n'.join(v for k, v in clauses if re.search(r'字段映射|列名映射', k))
    for left, right in re.findall(r'([\w]+)\s*=\s*([\w]+)', mapping_text):
        # Header aliases are inert text; only known semantics are bound.
        if right in set().union(*SIGNATURES.values()) or right in ALIASES.values():
            aliases[left] = right
    roles = defaultdict(list)
    for i, table in enumerate(frames):
        if i == policy_index:
            continue
        headers = {aliases.get(str(c), str(c)) for c in table}
        for role, signature in SIGNATURES.items():
            if signature <= headers:
                roles[role].append(i)
    if not all(roles.get(r) for r in ('sales', 'products', 'orgs', 'costs', 'refunds', 'returns')):
        return None
    if any(len(v) > 1 for k, v in roles.items() if k != 'sales'):
        return None
    stock = re.search(r'目标天数\s*([\d.]+)\s*倍', content)
    ar = re.search(r'期末应收/净销售\s*>\s*([\d.]+)%', content)
    return DeclaredPolicy(start, end, good, bad, names[policy_index],
        hashlib.sha256(content.encode('utf-8')).hexdigest(), clauses, aliases,
        {k: tuple(v) for k, v in roles.items()}, number(stock[1]) if stock else None,
        number(ar[1]) / 100 if ar else None)


def exact_unique(rows: list[dict], exclude=('来源表', '源行号')) -> tuple[list[dict], list[dict]]:
    """Keep one exact row, preserve removed records for audit."""
    seen, kept, duplicates = set(), [], []
    for row in rows:
        sig = json.dumps({k: serialise(v) for k, v in row.items() if k not in exclude},
                         sort_keys=True, ensure_ascii=False, default=str)
        if sig in seen:
            duplicates.append(row)
        else:
            seen.add(sig)
            kept.append(row)
    return kept, duplicates


def version_select(rows: list[dict], keys: tuple[str, ...], version: str, critical: tuple[str, ...]):
    """Highest numeric revision; ambiguous critical payload quarantines a key."""
    groups, kept, removed = defaultdict(list), [], []
    for r in rows:
        if any(not text(r.get(k)) for k in keys):
            removed.append((r, '业务键缺失'))
            continue
        groups[tuple(text(r.get(k)) for k in keys)].append(r)
    for _, values in groups.items():
        versions = [number(r.get(version)) for r in values]
        if any(v is None for v in versions):
            removed.extend((r, '版本号无效') for r in values)
            continue
        high = max(versions)
        latest = [r for r, v in zip(values, versions) if v == high]
        old = [r for r, v in zip(values, versions) if v != high]
        removed.extend((r, '被修订替代') for r in old)
        payloads = {tuple(text(r.get(k)) for k in critical) for r in latest}
        if len(payloads) != 1:
            removed.extend((r, '最高版本关键字段冲突') for r in latest)
        else:
            kept.append(latest[0])
            removed.extend((r, '同版本同业务内容重复') for r in latest[1:])
    return kept, removed


def allocate_amount(amount: Decimal, rows: list[dict], weight: str) -> list[Decimal] | None:
    """Proportional cents, deterministic residual, no quantity-based shortcut."""
    total = strict_sum(r.get(weight) for r in rows)
    if total is None or total <= 0 or amount < 0 or amount != money(amount) or any(r[weight] < 0 for r in rows):
        return None
    allocations, remaining = [], amount
    for row in rows[:-1]:
        # Independent half-up rounding can over-allocate very small refunds.
        # Bound each allocation to remaining cash so the last row never gets
        # a negative refund. Stable row order supplies deterministic tails.
        value = min(money(amount * row[weight] / total), remaining)
        allocations.append(value)
        remaining -= value
    return [*allocations, amount - sum(allocations, ZERO)]


def effective_value(rows: list[dict], sku: str, when: datetime) -> Decimal | None:
    matches = [r for r in rows if text(r.get('商品编码')).upper() == sku
               and date(r.get('生效开始')) and date(r.get('生效结束'))
               and date(r['生效开始']) <= when <= date(r['生效结束'])]
    if len(matches) != 1:
        return None
    value = number(matches[0].get('标准基础单价'))
    return value if value is not None and value >= 0 else None


def compile_operation_plan(policy: DeclaredPolicy, names: Sequence[str]) -> list[dict]:
    """Expose the typed, bound local operations; never model-supplied code."""
    specs = (
        ('exact_unique', ('sales',), '完整行；保留排除来源'),
        ('version_select', ('sales',), '订单号+行号；最高版本；冲突隔离'),
        ('canonicalise_and_filter', ('sales', 'products', 'orgs'), '主数据规范；精确状态；有效日期/数量/实付'),
        ('effective_value', ('sales', 'costs'), '商品+销售日；唯一生效区间'),
        ('allocate_amount', ('sales', 'refunds'), '原订单/行；确认退款；实付权重；分守恒'),
        ('accepted_return_reversal', ('sales', 'returns'), '验收入库；原销售成本；累计数量校验'),
        ('unit_conversion', ('purchases', 'products'), '业务键含单位；明确包装系数；币种不混加'),
        ('stock_roll_forward', ('opening_stock', 'sales', 'purchases', 'returns', 'adjustments'), '期初+入库-销售+退货+批准调整；未知局部传播'),
        ('latest_snapshot', ('counts',), '组织+商品；最新盘点；不重复计为调整'),
        ('settlement_reconcile', ('settlements',), '结算基数-渠道费-退款冲减=实际到账；费用一次'),
        ('cash_application', ('opening_ar', 'sales', 'receipts', 'applications'), '核销匹配已到账回款及有效业务；余额守恒'),
        ('accepted_service', ('services',), '验收确认；收入-直接成本'),
        ('incurred_expense', ('expenses',), '已发生；经营费用与资本性支出分开'),
        ('period_end_balance', ('workforce',), '组织+月；人数取期末，不跨月累计'),
        ('business_aggregate', ('sales', 'targets'), '各事实保持粒度；汇总后对比，不直接宽表连接'),
    )
    return [{'操作': op, '绑定来源': '；'.join(names[i] for role in roles for i in policy.roles.get(role, ())),
        '规则': rule, '期间': f'{policy.start:%Y-%m-%d}至{policy.cutoff:%Y-%m-%d}',
        '未提供域': '；'.join(role for role in roles if role not in policy.roles),
        '规则指纹': policy.digest, '执行边界': '本地确定性操作；无原文件覆盖；未确定项不按零填充'}
        for op, roles, rule in specs if any(role in policy.roles for role in roles)]


def build_declared_report(frames: Sequence[pd.DataFrame], names: Sequence[str], policy: DeclaredPolicy):
    """Compose facts at their own grain, propagate unknowns per dependency."""
    exceptions, audits, rules_used, risks = [], [], [], []
    def rows(role):
        result = []
        for i in policy.roles.get(role, ()):
            for position, values in enumerate(frames[i].to_dict('records'), start=2):
                row = {policy.aliases.get(str(k), str(k)): v for k, v in values.items()}
                row.update({'来源表': names[i].rsplit('__', 1)[-1], '源行号': position})
                result.append(row)
        return result
    def exclude(row, reason, level='P2'):
        exceptions.append({'优先级': level, '来源表': row.get('来源表'), '源行号': row.get('源行号'),
            '业务标识': text(row.get('订单号') or row.get('原订单号') or row.get('入库单号') or row.get('核销ID')),
            '原因': reason, '原记录': json.dumps(serialise(row), ensure_ascii=False, default=str)})
    def dedup(role):
        raw = rows(role)
        kept, dup = exact_unique(raw)
        for r in dup:
            exclude(r, '完整重复')
        audits.append({'步骤': role + ':完整行去重', '输入行': len(raw), '输出行': len(kept), '排除行': len(dup)})
        return kept
    def unique_master(role, key):
        raw = rows(role)
        result = {}
        for r in raw:
            value = text(r.get(key)).upper()
            if not value or value in result:
                raise PolicyError(f'{role}主数据键为空或重复：{value}')
            result[value] = r
        return result
    def risk(priority, obj, problem, evidence, action, department, days, acceptance):
        risks.append({'优先级': priority, '对象': obj, '问题': problem, '数据证据': evidence,
            '影响': problem, '建议行动': action, '责任岗位': department,
            '建议期限': (policy.cutoff + timedelta(days=days)).strftime('%Y-%m-%d'),
            '验收指标': acceptance, '性质': '事实触发，行动为建议'})
    products = unique_master('products', '商品编码')
    orgs = unique_master('orgs', '分公司编码')
    customers = unique_master('customers', '客户编码')
    org_aliases = {}
    for code, r in orgs.items():
        for alias in [code, text(r.get('分公司标准名称')), *re.split(r'[；;、]', text(r.get('允许别名')))]:
            if alias in org_aliases and org_aliases[alias] != code:
                raise PolicyError('组织别名映射不唯一')
            org_aliases[alias] = code
    costs = rows('costs')
    sales_raw = dedup('sales')
    revisions, rejected = version_select(sales_raw, ('订单号', '行号'), '版本号',
        ('商品编码', '营业日期', '分公司名称', '数量', '实付金额', '订单状态', '销售单位', '客户编码'))
    for r, reason in rejected:
        exclude(r, reason, 'P0' if '冲突' in reason else 'P2')
    sales = []
    for r in revisions:
        status, when, qty, paid = text(r.get('订单状态')), date(r.get('营业日期')), number(r.get('数量')), number(r.get('实付金额'))
        sku = text(r.get('商品编码')).upper()
        branch = org_aliases.get(text(r.get('分公司名称')))
        reason = ('无效状态' if status in policy.bad_orders or re.search(r'取消|失败|未支付|未完成|待支付', status)
                  else '未知状态' if status not in policy.good_orders
                  else '未知商品' if sku not in products else '非正数量' if qty is None or qty <= 0
                  else '无效金额' if paid is None or paid < 0 else '缺失或无效日期' if when is None
                  else '超出经营期间' if not policy.start <= when <= policy.cutoff
                  else '未知组织' if branch is None else '销售单位与基础单位不一致'
                  if text(r.get('销售单位')) != text(products[sku].get('基础单位')) else '')
        if reason:
            exclude(r, reason)
            continue
        unit_cost = effective_value(costs, sku, when)
        if unit_cost is None:
            exclude(r, '成本生效区间缺失或冲突；保留销售收入，毛利未知', 'P0')
        row = {**r, '分公司编码': branch, '商品编码': sku,
            '原分公司名称': r.get('分公司名称'), '分公司名称': orgs[branch]['分公司标准名称'],
            '原商品编码': r.get('商品编码'), '原商品名称': r.get('商品名称'), '原商品类别': r.get('商品类别'),
            '商品名称': products[sku]['标准商品名称'], '商品类别': products[sku]['标准类别'],
            '营业日期': when, '月份': when.strftime('%Y-%m'), '数量': qty, '实付金额': paid,
            '折扣金额': number(r.get('折扣金额')), '标准单位成本': unit_cost,
            '标准销售成本': money(qty * unit_cost) if unit_cost is not None else None,
            '分摊退款': ZERO, '退货成本冲回': ZERO}
        sales.append(row)
    sales.sort(key=lambda r: (text(r['订单号']), text(r['行号'])))
    sale_lookup = {(text(r['订单号']), text(r['行号'])): r for r in sales}
    order_lookup = defaultdict(list)
    for r in sales:
        order_lookup[text(r['订单号'])].append(r)
    audit_sales = {'步骤': '销售版本/状态/质量/期间处理', '输入行': len(sales_raw),
        '输出行': len(sales), '排除行': len(sales_raw) - len(sales)}
    audits.append(audit_sales)
    refunds, future_refunds, orphan_refunds = [], [], []
    # Duplicate refund IDs with different payloads are quarantined, not summed.
    rf = dedup('refunds')
    refund_groups = defaultdict(list)
    for r in rf:
        refund_groups[text(r.get('退款单号'))].append(r)
    for refund_id, group in refund_groups.items():
        if not refund_id or len(group) != 1:
            for r in group:
                exclude(r, '退款业务键缺失或冲突', 'P0')
            continue
        r = group[0]
        when, amount = date(r.get('退款日期')), number(r.get('退款金额'))
        if classify_refund_status(r.get('退款状态')) != REFUND_CONFIRMED or when is None or amount is None or amount < 0:
            exclude(r, '退款未确认/失败或日期金额无效')
            continue
        row = {**r, '退款日期': when, '退款金额': amount}
        if when > policy.cutoff:
            future_refunds.append(row)
            continue
        candidates = order_lookup.get(text(r.get('原订单号')), [])
        if text(r.get('原行号')):
            candidates = [s for s in candidates if text(s['行号']) == text(r['原行号'])]
        if not candidates:
            orphan_refunds.append(row)
            exclude(r, '已确认退款无法关联有效销售', 'P2')
            continue
        if any(when < s['营业日期'] for s in candidates):
            exclude(r, '退款日期早于关联销售日期', 'P0')
            continue
        allocations = allocate_amount(amount, candidates, '实付金额')
        if allocations is None or any(s['分摊退款'] + v > s['实付金额'] for s, v in zip(candidates, allocations)):
            exclude(r, '退款超过有效实付或分摊权重无效', 'P0')
            continue
        for s, value in zip(candidates, allocations):
            s['分摊退款'] += value
        refunds.append({**row, '分公司编码': candidates[0]['分公司编码'],
            '客户编码': candidates[0].get('客户编码'), '关联订单月': candidates[0]['月份'],
            '分摊依据': '指定行' if text(r.get('原行号')) else '订单有效行实付比例',
            '关联销售源行': '；'.join(f"{s['来源表']}:{s['源行号']}" for s in candidates)})
    returns = []
    for r in dedup('returns'):
        s = sale_lookup.get((text(r.get('原订单号')), text(r.get('原行号'))))
        when, qty = date(r.get('验收日期')), number(r.get('数量'))
        if text(r.get('验收状态')) != '验收合格已入库' or when is None or not policy.start <= when <= policy.cutoff:
            exclude(r, '退货未验收入库或期后退货')
            continue
        if not s or qty is None or qty <= 0 or qty > s['数量'] or when < s['营业日期'] or text(r.get('商品编码')).upper() != s['商品编码'] or text(r.get('单位')) != text(products[s['商品编码']]['基础单位']):
            exclude(r, '退货关联、数量或单位无效', 'P0')
            continue
        returned = sum(x['数量'] for x in returns if (text(x['原订单号']), text(x['原行号'])) == (text(s['订单号']), text(s['行号'])))
        if returned + qty > s['数量']:
            exclude(r, '累计退货超过销售数量', 'P0')
            continue
        reversal = money(qty * s['标准单位成本']) if s['标准单位成本'] is not None else None
        s['退货成本冲回'] = None if reversal is None or s['退货成本冲回'] is None else s['退货成本冲回'] + reversal
        returns.append({**r, '数量': qty, '验收日期': when, '成本冲回': reversal})
    for s in sales:
        s['净商品销售'] = s['实付金额'] - s['分摊退款']
        s['退款后标准毛利'] = None if s['标准销售成本'] is None or s['退货成本冲回'] is None else s['净商品销售'] - s['标准销售成本'] + s['退货成本冲回']
    # Remaining independent facts are compiled below; never join their rows
    # directly into sales and duplicate amounts.
    return _compose_other_facts(frames, names, policy, rows, dedup, exclude, risk,
        products, orgs, customers, costs, sales, sale_lookup, order_lookup,
        refunds, future_refunds, orphan_refunds, returns, exceptions, audits, risks, rules_used)


def _compose_other_facts(frames, names, policy, rows, dedup, exclude, risk,
        products, orgs, customers, costs, sales, sale_lookup, order_lookup,
        refunds, future_refunds, orphan_refunds, returns, exceptions, audits, risks, rules_used):
    purchases, unknown_stock = [], defaultdict(list)
    pg = defaultdict(list)
    for r in dedup('purchases'):
        r['商品编码'] = text(r.get('商品编码')).upper()
        r['分公司编码'] = text(r.get('分公司编码')).upper()
        when = date(r.get('入库日期'))
        if text(r.get('入库状态')) != '已入库' or when is None or not policy.start <= when <= policy.cutoff:
            exclude(r, '采购未入库或不在经营期间；历史入库不重复计入期初库存')
            continue
        pg[tuple(text(r.get(k)) for k in ('入库单号', '分公司编码', '商品编码', '采购单位'))].append(r)
    for key, group in pg.items():
        b, sku = key[1:3]
        if len(group) != 1:
            unknown_stock[(b, sku)].append('采购业务键冲突')
            for r in group:
                exclude(r, '采购业务键冲突', 'P0')
            continue
        r = group[0]
        when, qty, price = date(r.get('入库日期')), number(r.get('采购数量')), number(r.get('采购单价'))
        product = products.get(sku)
        unit = text(r.get('采购单位'))
        factor = Decimal(1) if product and unit == text(product.get('基础单位')) else number(product.get('每箱件数')) if product and unit == '箱' else None
        if not product or qty is None or qty <= 0 or factor is None or factor <= 0:
            unknown_stock[(b, sku)].append('商品/单位换算/数量未确定')
            exclude(r, '采购基础数量无法确定', 'P0')
            continue
        cash = money(qty * price) if price is not None and price >= 0 and text(r.get('币种')).upper() in ('CNY', 'RMB', '人民币') else None
        if cash is None:
            exclude(r, '采购人民币金额无法估值；基础数量仍有效')
        purchases.append({**r, '入库日期': when, '基础数量': qty * factor,
            '换算系数': factor, '折算基础单价': price / factor if price is not None else None,
            '人民币采购金额': cash})
    services = []
    pending_services = []
    for r in dedup('services'):
        when = date(r.get('验收或预计验收日'))
        amounts = [number(r.get(k)) for k in ('合同金额', '直接材料费', '直接人工费', '差旅费')]
        row = {**r, **dict(zip(('合同金额', '直接材料费', '直接人工费', '差旅费'), amounts))}
        if text(r.get('验收状态')) != '已验收' or when is None or not policy.start <= when <= policy.cutoff:
            pending_services.append(row)
            exclude(r, '工单未验收或不在经营期，合同不计收入')
            continue
        direct = strict_sum(amounts[1:])
        services.append({**row, '月份': when.strftime('%Y-%m'), '服务直接成本': direct,
            '服务直接贡献': None if direct is None or amounts[0] is None else amounts[0] - direct})
    expenses, capex = [], []
    for r in dedup('expenses'):
        when, amount = date(r.get('发生月份')), number(r.get('金额'))
        if when is None or not policy.start <= when <= policy.cutoff or amount is None or text(r.get('币种')) not in ('CNY', 'RMB', '人民币'):
            exclude(r, '费用日期/金额/币种未确认')
            continue
        row = {**r, '金额': amount, '月份': when.strftime('%Y-%m')}
        if text(r.get('费用类别')) == '设备购置':
            capex.append(row)
        elif text(r.get('订单状态') or r.get('状态')) in ('已发生', '已支付'):
            expenses.append(row)
        else:
            exclude(r, '费用发生状态未确认')
    settlements = []
    for r in dedup('settlements'):
        when = date(r.get('结算月份'))
        nums = {k: number(r.get(k)) for k in ('结算基数', '佣金', '配送费用', '活动补贴承担', '退款冲减', '应结实际到账')}
        if when is None or not policy.start <= when <= policy.cutoff or any(v is None for v in nums.values()):
            exclude(r, '结算期间或必要金额缺失', 'P0')
            continue
        fee = nums['佣金'] + nums['配送费用'] + nums['活动补贴承担']
        diff = money(nums['结算基数'] - fee - nums['退款冲减'] - nums['应结实际到账'])
        cash_date = date(r.get('预计或实际到账日'))
        arrived = text(r.get('到账状态')) == '已到账' and cash_date is not None and cash_date <= policy.cutoff
        if diff:
            exclude(r, '渠道结算勾稽不平', 'P0')
        settlements.append({**r, **nums, '月份': when.strftime('%Y-%m'), '渠道费用': fee,
            '勾稽差额': diff, '截止已到账': arrived, '截止待到账': ZERO if arrived else nums['应结实际到账']})
    # AR ledger owns the roll-forward. Applications must match both bank cash
    # and a valid document; orphan applications do not consume allocated cash.
    opening_ar = dedup('opening_ar')
    ar_documents = {}
    for r in opening_ar:
        key = ('期初应收', text(r.get('应收单号')), '')
        amount = number(r.get('期初应收金额'))
        if key in ar_documents or amount is None:
            raise PolicyError('期初应收业务键或金额无效')
        ar_documents[key] = {**r, '应收基数': amount, '已核销': ZERO, '文档类型': '期初应收',
            '到期日': date(r.get('到期日')), '销售净额': ZERO}
    for s in sales:
        key = ('商品销售', text(s['订单号']), text(s['行号']))
        customer = customers.get(text(s.get('客户编码')).upper())
        term = number(customer.get('账期天数')) if customer else None
        ar_documents[key] = {**s, '应收基数': s['净商品销售'], '已核销': ZERO,
            '文档类型': '商品销售', '销售净额': s['净商品销售'],
            '到期日': s['营业日期'] + timedelta(days=float(term)) if term is not None else None}
    receipts = {}
    for r in dedup('receipts'):
        identifier = text(r.get('回款ID'))
        if not identifier or identifier in receipts:
            raise PolicyError('回款ID为空或有冲突')
        receipts[identifier] = {**r, '回款金额': number(r.get('回款金额')), '有效核销': ZERO}
    valid_apps = []
    for r in dedup('applications'):
        rec = receipts.get(text(r.get('回款ID')))
        key = ('期初应收', text(r.get('业务单号')), '') if text(r.get('业务类型')) == '期初应收' else ('商品销售', text(r.get('业务单号')), text(r.get('原行号')))
        document = ar_documents.get(key)
        amount, when = number(r.get('核销金额')), date(r.get('核销日期'))
        arrived = date(rec.get('到账日期')) if rec else None
        if not rec or not document or amount is None or amount < 0 or when is None or when > policy.cutoff or arrived is None or arrived > policy.cutoff or text(rec.get('到账状态')) != '已到账' or text(rec.get('币种')) not in ('CNY', 'RMB', '人民币'):
            exclude(r, '核销无法关联有效业务/到账回款或截止条件不符')
            continue
        if text(rec.get('分公司编码')) != text(document.get('分公司编码')) or rec['回款金额'] is None or rec['有效核销'] + amount > rec['回款金额']:
            exclude(r, '核销组织不符或超过已到账回款', 'P0')
            continue
        document_date = document.get('营业日期')
        if when < arrived or (document_date is not None and when < document_date):
            exclude(r, '核销早于到账或关联业务发生日期', 'P0')
            continue
        if document['已核销'] + amount > document['应收基数']:
            exclude(r, '核销超过业务应收；预收款用途未声明，不冲成负应收', 'P0')
            continue
        document['已核销'] += amount
        rec['有效核销'] += amount
        valid_apps.append({**r, '核销金额': amount, '分公司编码': document['分公司编码'],
            '客户编码': document.get('客户编码'), '业务来源表': document.get('来源表'),
            '业务源行': document.get('源行号')})
    ar_rows = []
    for doc in ar_documents.values():
        closing = doc['应收基数'] - doc['已核销']
        due = doc['到期日']
        ar_rows.append({**doc, '期末已知应收': closing,
            '逾期天数': max((policy.cutoff - due).days, 0) if due else None,
            '账龄边界': '客户主数据账期管理估算' if doc['文档类型'] == '商品销售' else '源到期日'})
        if closing < 0:
            exclude(doc, '业务单核销超过应收，负余额待核验', 'P0')
    cash_rows = []
    for r in receipts.values():
        when = date(r.get('到账日期'))
        arrived = text(r.get('到账状态')) == '已到账' and when and when <= policy.cutoff
        currency_ok = text(r.get('币种')).upper() in ('CNY', 'RMB', '人民币')
        if arrived and not currency_ok:
            exclude(r, '已到账外币回款无汇率，不合并为人民币')
        cash_rows.append({**r, '截止已到账': bool(arrived),
            '未分配回款': r['回款金额'] - r['有效核销'] if arrived and currency_ok and r['回款金额'] is not None else None})
    # Stock is opening + movements. Counts remain a separate latest snapshot.
    stock = {}
    for r in dedup('opening_stock'):
        key = (text(r.get('分公司编码')), text(r.get('商品编码')).upper())
        if key in stock or number(r.get('期初数量')) is None:
            raise PolicyError('期初库存键或数量无效')
        stock[key] = {'期初数量': number(r['期初数量']), '采购基础数量': ZERO,
            '销售数量': ZERO, '退货入库': ZERO, '已批准调整': ZERO, '来源表': r['来源表'], '源行号': r['源行号']}
        if date(r.get('库存日期')) != policy.start:
            unknown_stock[key].append('期初库存日期与经营起点不一致')
            exclude(r, '期初库存日期与经营起点不一致；不猜测期间间隔的出入库', 'P0')
    def movement(branch, sku, field, amount):
        key = (text(branch), text(sku).upper())
        if key not in stock:
            unknown_stock[key].append('缺少期初库存')
            stock[key] = {'期初数量': ZERO, '采购基础数量': ZERO, '销售数量': ZERO, '退货入库': ZERO, '已批准调整': ZERO}
        stock[key][field] += amount
    for r in purchases:
        movement(r['分公司编码'], r['商品编码'], '采购基础数量', r['基础数量'])
    for r in sales:
        movement(r['分公司编码'], r['商品编码'], '销售数量', r['数量'])
    for r in returns:
        movement(r['分公司编码'], r['商品编码'], '退货入库', r['数量'])
    for r in dedup('adjustments'):
        when, qty = date(r.get('日期')), number(r.get('调整数量'))
        p = products.get(text(r.get('商品编码')).upper())
        if text(r.get('审批状态')) != '已批准' or not when or not policy.start <= when <= policy.cutoff or qty is None or not p or text(r.get('单位')) != text(p.get('基础单位')):
            exclude(r, '调整未批准、日期或单位无效')
            continue
        movement(r['分公司编码'], r['商品编码'], '已批准调整', qty)
    counts = defaultdict(list)
    for r in rows('counts'):
        when = date(r.get('盘点日期'))
        if when and when <= policy.cutoff:
            counts[(text(r.get('分公司编码')), text(r.get('商品编码')).upper())].append(r)
    inventory = []
    days = Decimal((policy.cutoff - policy.start).days + 1)
    for (branch, sku), r in stock.items():
        product = products.get(sku, {})
        known = r['期初数量'] + r['采购基础数量'] - r['销售数量'] + r['退货入库'] + r['已批准调整']
        unknown = unknown_stock.get((branch, sku), [])
        closing = None if unknown else known
        latest_cost = effective_value(costs, sku, policy.cutoff)
        value = money(closing * latest_cost) if closing is not None and latest_cost is not None else None
        coverage = max(closing, ZERO) / (r['销售数量'] / days) if closing is not None and r['销售数量'] > 0 else None
        safe, target = number(product.get('安全库存_件')), number(product.get('目标库存天数'))
        candidates = sorted(counts.get((branch, sku), []), key=lambda x: date(x['盘点日期']))
        latest = candidates[-1] if candidates else None
        count = number(latest.get('盘点数量')) if latest else None
        same_day = [c for c in candidates if date(c['盘点日期']) == date(latest['盘点日期'])] if latest else []
        count_conflict = len({number(c.get('盘点数量')) for c in same_day}) > 1
        if count_conflict:
            count = None
            exclude(latest, '同日盘点数量冲突，不任选最后一条', 'P0')
        count_basis = ('同日盘点冲突' if count_conflict else '缺少盘点' if latest is None
                       else '同一截止日可比较' if date(latest['盘点日期']) == policy.cutoff
                       else '盘点日与期末不同，不直接计算差异')
        status = '待核验' if unknown else '负库存' if closing < 0 else '低于安全库存' if safe is not None and closing < safe else '积压关注' if coverage is not None and target is not None and policy.stock_multiplier is not None and coverage > target * policy.stock_multiplier else '无销量库存' if r['销售数量'] == 0 and closing > 0 else '正常'
        evidence = (f'已知数量={display_number(known)}；完整期末={display_number(closing)}；'
                    f'可售天数={display_number(coverage)}；目标天数={display_number(target)}')
        if status != '正常':
            risk('P0' if status in ('待核验', '负库存') else 'P1', f'{branch}/{sku}', status,
                 evidence + ('；' + '；'.join(unknown) if unknown else ''),
                 '核对采购、换算及出入库；确认后制定补货/消化计划，不自动下单', '仓库及采购负责人', 1 if status in ('负库存', '待核验') else 7,
                 '异常记录逐笔确认；单位和账面库存可勾稽')
        inventory.append({'分公司编码': branch, '商品编码': sku, '商品名称': product.get('标准商品名称'),
            **r, '已知部分期末数量': known, '完整期末数量': closing, '标准库存金额': value,
            '标准单位成本': latest_cost, '安全库存': safe, '目标库存天数': target,
            '可售天数': coverage, '状态': status, '未确定原因': '；'.join(unknown),
            '最新盘点日期': date(latest['盘点日期']) if latest else None, '最新盘点数量': count,
            '盘点差异': count - closing if count is not None and closing is not None and date(latest['盘点日期']) == policy.cutoff else None,
            '盘点比较口径': count_basis})
    return _report_outputs(frames, names, policy, products, orgs, customers, sales, order_lookup, refunds,
        future_refunds, orphan_refunds, returns, purchases, services, pending_services,
        expenses, capex, settlements, ar_rows, cash_rows, valid_apps, inventory,
        rows, exclude, risk, exceptions, audits, risks)


def _report_outputs(frames, names, policy, products, orgs, customers, sales, order_lookup, refunds,
        future_refunds, orphan_refunds, returns, purchases, services, pending_services,
        expenses, capex, settlements, ar_rows, cash_rows, valid_apps, inventory,
        rows, exclude, risk, exceptions, audits, risks):
    def total(population, column):
        return strict_sum(r.get(column) for r in population)
    def fact_total(role, population, column):
        # No supplied domain is not an empty, confirmed population.
        return total(population, column) if role in policy.roles else None
    def pick(population, key, value):
        return [r for r in population if text(r.get(key)) == text(value)]
    def minus(a, b):
        return None if a is None or b is None else a - b
    def add(a, b):
        return None if a is None or b is None else a + b
    def ratio(a, b):
        return a / b if a is not None and b is not None and b != 0 else None
    targets = rows('targets')
    workforce = rows('workforce')
    for r in [*targets, *workforce]:
        when = date(r.get('月份'))
        r['月份'] = when.strftime('%Y-%m') if when else ''
        for c in ('净销售目标', '月末人数', '正常工时', '加班工时'):
            if c in r:
                r[c] = number(r[c])
    targets = [r for r in targets if policy.start.strftime('%Y-%m') <= r['月份'] <= policy.cutoff.strftime('%Y-%m')]
    workforce = [r for r in workforce if policy.start.strftime('%Y-%m') <= r['月份'] <= policy.cutoff.strftime('%Y-%m')]
    latest_month = max((r['月份'] for r in workforce), default='')
    latest_people = fact_total('workforce', [r for r in workforce if r['月份'] == latest_month], '月末人数')
    goods, gp = total(sales, '净商品销售'), total(sales, '退款后标准毛利')
    service_gp = total(services, '服务直接贡献') if 'services' in policy.roles else None
    fees = total(settlements, '渠道费用') if 'settlements' in policy.roles else None
    opex = total(expenses, '金额') if 'expenses' in policy.roles else None
    contribution = minus(minus(add(gp, service_gp), fees), opex)
    head = [
        ('经营期间', f'{policy.start:%Y-%m-%d} 至 {policy.cutoff:%Y-%m-%d}', '', '客户声明期间'),
        ('有效订单', len({text(s['订单号']) for s in sales}), '笔', '清洗后订单号去重计数'),
        ('有效销售明细', len(sales), '行', '版本、状态、主数据、日期金额与单位校验'),
        ('商品实付', total(sales, '实付金额'), '元', '源实付，不由单价补算'),
        ('截止日已确认关联退款', total(refunds, '退款金额'), '元', '确认到账且关联有效订单，排除期后'),
        ('净商品销售', goods, '元', '商品实付减截止已确认关联退款'),
        ('标准销售成本', total(sales, '标准销售成本'), '元', '销售数量乘销售日唯一生效成本'),
        ('退货成本冲回', total(returns, '成本冲回'), '元', '仅验收合格入库，使用原销售日成本'),
        ('退款后标准毛利', gp, '元', '净商品销售减标准成本加退货成本冲回'),
        ('退款后标准毛利率', ratio(gp, goods), '%', '合计毛利除以合计净销售'),
        ('已验收服务收入', fact_total('services', services, '合同金额'), '元', '只计已验收工单'),
        ('已验收服务直接成本', fact_total('services', services, '服务直接成本'), '元', '材料、人工、差旅；不重复计费'),
        ('已验收服务贡献', service_gp, '元', '已验收服务收入减直接成本'),
        ('渠道费用', fees, '元', '佣金、配送、活动补贴承担扣一次'),
        ('运营费用', opex, '元', '经营期已发生；不含设备购置'),
        ('已知管理经营贡献', contribution, '元', '商品标准毛利+服务贡献-渠道费-运营费用，不是净利润'),
        ('期初应收', total([r for r in ar_rows if r['文档类型'] == '期初应收'], '应收基数'), '元', '独立期初单据'),
        ('有效核销', total(valid_apps, '核销金额'), '元', '有效业务单、已到账、截止日前，孤立核销不计'),
        ('期末已知应收', total(ar_rows, '期末已知应收'), '元', '期初+净销售-有效核销，不包括未提供的其他应收'),
        ('已到账未分配回款', total([r for r in cash_rows if r['截止已到账']], '未分配回款'), '元', '已到账回款减有效核销，不冲减孤立核销'),
        ('渠道截止待到账', fact_total('settlements', settlements, '截止待到账'), '元', '独立渠道结算口径，不能与回款流水重复相加'),
        ('已折算确认采购CNY金额', fact_total('purchases', [r for r in purchases if r['人民币采购金额'] is not None], '人民币采购金额'), '元', '已入库、已折算基础数量且人民币可估值部分'),
        ('完整库存待核验组合', sum(r['完整期末数量'] is None for r in inventory), '组', '未知换算或采购冲突局部传播'),
        ('已确认组合标准库存金额', total([r for r in inventory if r['标准库存金额'] is not None], '标准库存金额'), '元', '仅有完整数量及期末标准成本的组合，不是实际存货成本余额'),
        ('期末人数', latest_people, '人', f'{latest_month}各组织人数合计，不跨月累计'),
        ('待验收合同金额', fact_total('services', pending_services, '合同金额'), '元', '不计已确认收入'),
        ('期后退款', total(future_refunds, '退款金额'), '元', '不回扣截止日已知结果'),
        ('无法关联已确认退款', total(orphan_refunds, '退款金额'), '元', '不强行扣到任何订单'),
        ('设备购置现金', fact_total('expenses', capex, '金额'), '元', '资本性项目，不直接作为当期运营费用'),
    ]
    monthly = []
    # A refund-only/settlement-only month is a real operating period too.
    # Do not drop it just because there were no sales or service rows.
    month_keys = sorted({r['月份'] for r in sales} | {r['月份'] for r in services}
        | {r['月份'] for r in expenses} | {r['月份'] for r in settlements}
        | {r['退款日期'].strftime('%Y-%m') for r in refunds if policy.start <= r['退款日期'] <= policy.cutoff})
    for m in month_keys:
        ss, sv, st, ex = (pick(pop, '月份', m) for pop in (sales, services, settlements, expenses))
        mp, mg = total(ss, '退款后标准毛利'), total(ss, '净商品销售')
        month_gp = add(mp, fact_total('services', sv, '服务直接贡献'))
        monthly.append({'月份': m, '商品实付': total(ss, '实付金额'),
            '订单归属退款': total(ss, '分摊退款'),
            '发生月退款': total([r for r in refunds if r['退款日期'].strftime('%Y-%m') == m], '退款金额'),
            '净商品销售': mg, '标准销售成本': total(ss, '标准销售成本'),
            '退货成本冲回': total(ss, '退货成本冲回'), '退款后标准毛利': mp,
            '标准毛利率': ratio(mp, mg), '服务贡献': fact_total('services', sv, '服务直接贡献'),
            '渠道费用': fact_total('settlements', st, '渠道费用'), '运营费用': fact_total('expenses', ex, '金额'),
            '已知管理贡献': minus(minus(month_gp, fact_total('settlements', st, '渠道费用')), fact_total('expenses', ex, '金额')),
            '时间边界': '商品毛利按订单归属月；渠道/服务/费用按各自发生月，未强行分摊到订单'})
    branches = []
    for b, org in orgs.items():
        ss, sv, st, ex, ar, tg, wf = (pick(pop, '分公司编码', b) for pop in (sales, services, settlements, expenses, ar_rows, targets, workforce))
        net, profit, due = total(ss, '净商品销售'), total(ss, '退款后标准毛利'), total(ar, '期末已知应收')
        before_hq = minus(minus(add(profit, fact_total('services', sv, '服务直接贡献')), fact_total('settlements', st, '渠道费用')), fact_total('expenses', ex, '金额'))
        target = fact_total('targets', tg, '净销售目标')
        closing_people = fact_total('workforce', [r for r in wf if r['月份'] == latest_month], '月末人数')
        branches.append({'分公司编码': b, '分公司': org['分公司标准名称'], '净商品销售': net,
            '退款后标准毛利': profit, '毛利率': ratio(profit, net), '服务贡献': fact_total('services', sv, '服务直接贡献'),
            '渠道费用': fact_total('settlements', st, '渠道费用'), '直接运营费用': fact_total('expenses', ex, '金额'),
            '总部费用分摊前贡献': before_hq, '商品净销售目标': target, '目标完成率': ratio(net, target),
            '有效订单': len({text(s['订单号']) for s in ss}), '期末已知应收': due,
            '应收占净销售': ratio(due, net), '期末人数': closing_people,
            '总工时': add(total(wf, '正常工时'), total(wf, '加班工时')),
            '加班工时': total(wf, '加班工时'), '净销售每工时': ratio(net, add(total(wf, '正常工时'), total(wf, '加班工时'))),
            '费用边界': '未提供总部费用分摊规则，组织贡献不含总部公共费用'})
        if policy.ar_ratio is not None and ratio(due, net) is not None and ratio(due, net) > policy.ar_ratio:
            risk('P1', b, '应收占净销售超过客户阈值', f'应收={display_number(due)}元；净销售={display_number(net)}元；阈值={policy.ar_ratio:.1%}',
                '逐笔核对到期及核销，分层跟进回款', '财务及销售负责人', 7, '逾期业务单有负责人和承诺回款日')
        ordered = sorted(wf, key=lambda r: r['月份'])
        overtime = [r['加班工时'] for r in ordered]
        if len(overtime) >= 3 and all(x is not None for x in overtime) and all(y > x for x, y in zip(overtime, overtime[1:])):
            evidence = '月度加班：' + ' → '.join(f"{r['月份']} {display_number(r['加班工时'])}小时" for r in ordered)
            risk('P1', b, '加班工时持续增加', evidence, '核对订单、工时及排班效率；不能仅据加班判定员工表现', '运营及人事负责人', 14, '确认加班原因及同口径每工时销售变化')
    channels = []
    for c in sorted({text(r.get('渠道')) for r in sales}):
        ss, st = pick(sales, '渠道', c), pick(settlements, '渠道', c)
        net, profit, fee = total(ss, '净商品销售'), total(ss, '退款后标准毛利'), fact_total('settlements', st, '渠道费用')
        channels.append({'渠道': c, '净商品销售': net, '退款后标准毛利': profit, '渠道费用': fee,
            '扣渠道费后商品贡献': minus(profit, fee), '费用率': ratio(fee, net),
            '截止待到账': fact_total('settlements', st, '截止待到账'), '边界': '未分摊服务、总部及组织费用，不称渠道净利润'})
        if minus(profit, fee) is not None and minus(profit, fee) < 0:
            risk('P1', c, '扣渠道费后商品贡献为负', f'标准毛利={display_number(profit)}元；渠道费用={display_number(fee)}元',
                '核对费用合同及促销结构，讨论价格或渠道调整', '渠道及财务负责人', 7, '费用只扣一次；调整方案有贡献测算')
    product_rows = []
    for sku, p in products.items():
        ss = pick(sales, '商品编码', sku)
        net, profit = total(ss, '净商品销售'), total(ss, '退款后标准毛利')
        product_rows.append({'商品编码': sku, '商品': p['标准商品名称'], '类别': p['标准类别'],
            '销售数量': total(ss, '数量'), '净商品销售': net, '退款金额': total(ss, '分摊退款'),
            '退款后标准毛利': profit, '毛利率': ratio(profit, net),
            '边界': '按源单位及历史标准成本；未将无分摊依据渠道费复制给商品'})
    product_rows.sort(key=lambda r: r['净商品销售'], reverse=True)
    feedback = []
    for r in rows('feedback'):
        ss = order_lookup.get(text(r.get('订单号')), [])
        matched = bool(ss) and all(text(s.get('客户编码')) == text(r.get('客户编码')) and text(s['分公司编码']) == text(r.get('分公司编码')) for s in ss)
        if not matched:
            exclude(r, '反馈订单/客户/组织无法一致关联')
        feedback.append({**r, '关联状态': '一致匹配' if matched else '人工核验',
            '关联商品': '；'.join(sorted({s['商品名称'] for s in ss})) if matched else None,
            '关联退款金额': total(ss, '分摊退款') if matched else None,
            '源满意度': number(r.get('满意度')), '源高严重度': text(r.get('严重度')) == '高'})
    customer_rows = []
    for code, customer in customers.items():
        ss, ar = pick(sales, '客户编码', code), pick(ar_rows, '客户编码', code)
        closing = total(ar, '期末已知应收')
        credit = number(customer.get('信用额度'))
        customer_rows.append({'客户编码': code, '客户': customer.get('客户名称'),
            '净商品销售': total(ss, '净商品销售'), '标准毛利': total(ss, '退款后标准毛利'),
            '期末已知应收': closing, '信用额度': credit, '信用等级': customer.get('信用等级'),
            '逾期应收': total([r for r in ar if r['逾期天数'] is not None and r['逾期天数'] > 0], '期末已知应收'),
            '反馈一致关联数': sum(r['关联状态'] == '一致匹配' and text(r.get('客户编码')) == code for r in feedback)})
        if credit is not None and closing is not None and closing > credit:
            risk('P1', code, '应收超过主数据信用额度', f'应收={display_number(closing)}元；额度={display_number(credit)}元',
                '核验额度及到期日，讨论催收和授信调整，不自动停止业务', '财务及客户负责人', 7, '额度口径确认，超额业务逐单跟进')
    if len(monthly) >= 2:
        first, last = monthly[0], monthly[-1]
        if last['净商品销售'] > first['净商品销售'] and last['标准毛利率'] is not None and first['标准毛利率'] is not None and last['标准毛利率'] < first['标准毛利率']:
            risk('P1', '整体经营', '销售增长但标准毛利率下降',
                f"{first['月份']}净销售={display_number(first['净商品销售'])}元、毛利率={first['标准毛利率']:.2%}；{last['月份']}净销售={display_number(last['净商品销售'])}元、毛利率={last['标准毛利率']:.2%}",
                '分解商品结构、历史成本、折扣及渠道费用变化；贡献口径与净利润分开', '经营及财务负责人', 7, '价格/成本/结构变化有源明细支持，不把相关性称因果')
    for e in exceptions:
        if e['优先级'] == 'P0':
            risk('P0', f"{e['来源表']}:{e['源行号']}", e['原因'], e['业务标识'],
                 '核对原记录及业务负责人，不自动猜测缺失值', '对应业务负责人', 1, '冲突有审批结论及保留记录')
    risks.sort(key=lambda r: (r['优先级'], r['对象']))
    dash = []
    for i in range(max(len(monthly), len(branches), len(product_rows))):
        r = {}
        if i < len(monthly):
            m = monthly[i]
            r.update({'月份': m['月份'], '净销售': m['净商品销售'], '标准毛利': m['退款后标准毛利'], '管理贡献': m['已知管理贡献']})
        if i < len(branches):
            b = branches[i]
            r.update({'组织': b['分公司'], '组织净销售': b['净商品销售'], '组织分摊前贡献': b['总部费用分摊前贡献']})
        if i < len(product_rows):
            p = product_rows[i]
            r.update({'商品': p['商品'], '商品净销售': p['净商品销售'], '商品标准毛利率': p['毛利率']})
        dash.append(r)
    ar_total = total(ar_rows, '期末已知应收')
    ar_expected = minus(add(total([r for r in ar_rows if r['文档类型'] == '期初应收'], '应收基数'), goods), total(valid_apps, '核销金额'))
    checks = [
        {'校验项': '应收滚动勾稽', '差额': minus(ar_total, ar_expected), '判定': '通过' if ar_total == ar_expected else '不通过'},
        {'校验项': '关联退款分摊守恒', '差额': minus(total(refunds, '退款金额'), total(sales, '分摊退款')), '判定': '通过' if total(refunds, '退款金额') == total(sales, '分摊退款') else '不通过'},
        {'校验项': '渠道结算勾稽', '差额': total(settlements, '勾稽差额'), '判定': '通过' if all(r['勾稽差额'] == 0 for r in settlements) else '不通过'},
        {'校验项': '标准成本覆盖', '差额': sum(r['标准销售成本'] is None for r in sales), '判定': '通过' if all(r['标准销售成本'] is not None for r in sales) else '部分不可计算'},
        {'校验项': '完整库存覆盖', '差额': sum(r['完整期末数量'] is None for r in inventory), '判定': '部分待核验' if any(r['完整期末数量'] is None for r in inventory) else '通过'},
    ]
    # Concise management reading layer derived exclusively from computed
    # facts. This is presentation, not a second calculation or answer fixture.
    def display(value):
        return '不可确定' if value is None else f'{value:,.2f}'
    conclusions = [{'主题': '盈利边界', '已验证结论': '有标准成本管理贡献，不等于会计净利润',
        '数据证据': f'净商品销售{display(goods)}元；商品标准毛利{display(gp)}元；已知管理经营贡献{display(contribution)}元',
        '下一步建议': '复核标准成本与实际存货成本；补齐税费、折旧、利息等会计资料后再判断净利润。'}]
    if len(monthly) >= 2:
        a, b = monthly[0], monthly[-1]
        conclusions.append({'主题': '增长与贡献', '已验证结论': '用同口径月度差额解释贡献变化，不把相关性当因果',
            '数据证据': f"{a['月份']}→{b['月份']}：净销售{display(a['净商品销售'])}→{display(b['净商品销售'])}；标准毛利{display(a['退款后标准毛利'])}→{display(b['退款后标准毛利'])}；管理贡献{display(a['已知管理贡献'])}→{display(b['已知管理贡献'])}元。渠道费变化{display(minus(b['渠道费用'], a['渠道费用']))}元；运营费变化{display(minus(b['运营费用'], a['运营费用']))}元；服务贡献变化{display(minus(b['服务贡献'], a['服务贡献']))}元。",
            '下一步建议': '先核对变化最大的成本项，再沿商品和订单明细核实价格、折扣、成本及结构变化；按各事实域的时间边界解读。'})
    eligible = [r for r in branches if r['总部费用分摊前贡献'] is not None]
    if eligible:
        best = max(eligible, key=lambda r: r['总部费用分摊前贡献'])
        conclusions.append({'主题': '组织表现', '已验证结论': f"按公共费用分摊前贡献，{best['分公司']}最高",
            '数据证据': f"分摊前贡献{display(best['总部费用分摊前贡献'])}元；净销售{display(best['净商品销售'])}元；期末已知应收{display(best['期末已知应收'])}元",
            '下一步建议': '结合目标完成、回款及工时效率比较；未有公共费用分摊依据时不称组织净利润或据此奖惩。'})
    conclusions.append({'主题': '资金占用线索', '已验证结论': '应收、渠道结算和库存分别看，不拼成完整现金流',
        '数据证据': f"已知应收{display(ar_total)}元；渠道待到账{display(total(settlements, '截止待到账'))}元；可估值组合标准库存{display(total([r for r in inventory if r['标准库存金额'] is not None], '标准库存金额'))}元",
        '下一步建议': '逐单核对逾期及核销，核实渠道到账日；单独处理积压、负库存及未知换算，不自动补货或付款。'})
    conclusions.append({'主题': '交付未确定项', '已验证结论': '确定的部分已计算，不确定的部分没有填零',
        '数据证据': f"库存待核验{sum(r['完整期末数量'] is None for r in inventory)}组；异常/排除审计{len(exceptions)}条；规则来自{policy.source.rsplit('__', 1)[-1]}",
        '下一步建议': '在异常、库存、未确认事项页核验单位、币种、冲突及来源；无效状态与已批准去重无需再次手动处理。'})
    outputs = {
        '经营结论摘要': frame(conclusions),
        '管理层通用总览': frame([{'指标': k, '结果': v, '单位': u, '数据口径': n} for k, v, u, n in head]),
        '关键问题与行动': frame(summarise_risks(risks)), '经营图表': frame(dash),
        '月度经营分析': frame(monthly), '组织经营分析': frame(branches),
        '渠道贡献分析': frame(channels), '商品贡献分析': frame(product_rows),
        '应收业务明细': frame(ar_rows), '客户风险与反馈': frame(customer_rows),
        '库存滚动与风险': frame(inventory), '销售清洗明细': frame(sales),
        '退款关联与分摊': frame(refunds), '采购单位与金额': frame(purchases),
        '服务验收与贡献': frame(services), '渠道结算勾稽': frame(settlements),
        '回款与未分配': frame(cash_rows), '有效核销明细': frame(valid_apps),
        '评价关联明细': frame(feedback), '人员月度工时': frame(workforce),
        '费用发生明细': frame(expenses), '期后及未确认事项': frame([
            *[dict(r, 事项='期后退款') for r in future_refunds],
            *[dict(r, 事项='孤立已确认退款') for r in orphan_refunds],
            *[dict(r, 事项='未验收服务合同') for r in pending_services],
            *[dict(r, 事项='资本性购置，不计运营费用') for r in capex]]),
        '异常与排除记录': frame(exceptions), '风险行动明细': frame(risks),
        '编译操作计划': frame(compile_operation_plan(policy, names)),
        '处理步骤与审计': frame(audits),
        '数据口径与校验': frame([{'事项': k, '说明': v, '规则来源': policy.source.rsplit('__', 1)[-1]} for k, v in policy.clauses]),
        '业务勾稽校验': frame(checks),
        '输入角色与粒度': frame([{'源文件与工作表': n, '记录数': len(f),
            '识别角色': ','.join(k for k, v in policy.roles.items() if i in v) or '声明口径/未绑定参考',
            '字段数': len(f.columns)} for i, (n, f) in enumerate(zip(names, frames))]),
    }
    for name, output in outputs.items():
        if output.empty and not len(output.columns):
            outputs[name] = frame([{'说明': '本次没有适用记录；不表示其他未提供数据为零'}])
        outputs[name].attrs['toolbox_report_kind'] = 'declared_multifact_report'
    report = {'source_table_count': len(frames), 'source_rows': sum(map(len, frames)),
        'sheet_count': len(outputs), 'chart_count': (2 if len(monthly) >= 2 else 0) + bool(branches) + bool(product_rows), 'business_rules_applied': True,
        'policy_source': policy.source, 'policy_digest': policy.digest,
        'policy_origin': 'explicit_user_referenced_sheet',
        'compiled_operations': compile_operation_plan(policy, names),
        'valid_sales_rows': len(sales), 'valid_orders': len({text(s['订单号']) for s in sales}),
        'exception_count': len(exceptions), 'unknown_inventory_groups': sum(r['完整期末数量'] is None for r in inventory),
        'execution_status': 'completed_with_disclosed_limits',
        'business_checks': serialise(checks), 'metrics': {k: serialise(v) for k, v, _, _ in head},
        'request_coverage': [
            {'topic': '经营管理贡献', 'status': 'available' if contribution is not None else 'partial',
             'boundary': '标准成本管理口径，不是会计净利润；缺少依赖时保持空值'},
            {'topic': '完整库存', 'status': 'partial' if any(r['完整期末数量'] is None for r in inventory) else 'available',
             'boundary': '单位未确定或业务冲突仅影响有关商品/组织，需人工核验'},
        ],
        'limits': ['标准成本管理估算不是净利润；缺少税费、折旧等会计资料',
            '未知单位/币种、冲突采购不猜测；仅影响有关对象',
            '总部费用未分摊；渠道与回款不能重复计为银行现金流',
            '建议不触发采购、付款、授信或人员审批']}
    if 'services' not in policy.roles or 'settlements' not in policy.roles or 'expenses' not in policy.roles:
        report['limits'].append('服务、渠道结算或费用域未提供，管理经营贡献不能形成完整数值')
    return outputs, report
