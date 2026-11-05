"""Kernel properties are independent of industry-specific benchmark answers."""
from decimal import Decimal
import random

import pandas as pd
import pytest

from excel_data_toolbox.business_kernel import (allocate_money, complete_sum, numbers,
    deduplicate_events, unique_master, BusinessInvariantError)
from excel_data_toolbox.task_contract import compile_scope, apply_scope, TaskScope
from excel_data_toolbox.task_acceptance import assess_task


@pytest.mark.parametrize('method', ['largest_remainder', 'last_residual'])
def test_hundreds_of_allocations_conserve_money_and_never_go_negative(method):
    rng = random.Random(20261104)
    for _ in range(300):
        weights = [Decimal(rng.randrange(0, 10000)) / 100 for _ in range(rng.randrange(1, 12))]
        if not any(weights):
            weights[0] = Decimal(1)
        amount = Decimal(rng.randrange(0, 100000)) / 100
        shares = allocate_money(amount, weights, method=method)
        assert sum(shares) == amount
        assert all(x >= 0 and x == x.quantize(Decimal('.01')) for x in shares)
        assert all(share == 0 for share, weight in zip(shares, weights) if weight == 0)


@pytest.mark.parametrize('value', ['10%', '1箱', '10kg', '待核对', True, float('inf')])
def test_units_missing_values_and_booleans_are_not_silently_money(value):
    assert numbers(pd.Series([value], dtype=object)).isna().all()


def test_empty_event_population_is_zero_but_unknown_event_is_not():
    assert complete_sum([]) == 0
    assert pd.isna(complete_sum([1, None]))
    assert complete_sum([0, 0]) == 0
    assert pd.isna(complete_sum([pd.NA]))
    assert pd.isna(complete_sum(pd.Series([1, pd.NA], dtype='Float64')))


def test_line_identity_is_idempotent_and_conflicts_are_not_last_write_wins():
    events = pd.DataFrame({'单号': ['A', 'A'], '商品': ['P1', 'P2'], '数量': [1, 2]})
    clean, duplicates = deduplicate_events(pd.concat([events, events]), ['单号', '商品'], label='采购')
    pd.testing.assert_frame_equal(clean.reset_index(drop=True), events)
    assert len(duplicates) == 2
    conflict = events.iloc[[0]].copy()
    conflict['数量'] = 9
    with pytest.raises(BusinessInvariantError, match='冲突'):
        deduplicate_events(pd.concat([events, conflict]), ['单号', '商品'], label='采购')
    with pytest.raises(BusinessInvariantError, match='冲突'):
        unique_master(pd.DataFrame({'SKU': ['A', 'A'], '成本': [1, 2]}), ['SKU'], label='商品')


def source():
    return pd.DataFrame({'订单号': ['A', 'B'], '日期': ['2026-08-01', '2026-08-31 23:59:59'],
                         '地区': ['华东', '华西'], '销售额': [10, 20]})


@pytest.mark.parametrize('prompt', ['只分析2026年8月华东地区销售额', '统计华东销售额', '只分析8月华东地区销售额'])
def test_scope_binds_to_actual_columns_and_keeps_originals(prompt):
    raw = source()
    before = raw.copy(deep=True)
    scope = compile_scope(prompt, [raw], ['销售'])
    frames, excluded, evidence = apply_scope(scope, [raw], ['销售'])
    assert frames[0]['销售额'].sum() == 10
    assert evidence[0]['excluded_rows'] == 1 and len(excluded[0]) == 1
    pd.testing.assert_frame_equal(raw, before)
    assert TaskScope.from_dict(scope.to_dict()) == scope


def test_inclusive_final_day_and_top_one_are_explicit():
    raw = source()
    scope = compile_scope('分析2026-08-01至2026-08-31的趋势，显示前1名', [raw], ['销售'])
    frames, _, _ = apply_scope(scope, [raw], ['销售'])
    assert len(frames[0]) == 2 and scope.top_n == 1


@pytest.mark.parametrize('prompt', ['只看华北地区', '只分析2026年13月销售额', '排除华东地区后分析', '只分析华东和华西地区'])
def test_unbound_or_ambiguous_scope_never_becomes_all_data(prompt):
    with pytest.raises(BusinessInvariantError):
        compile_scope(prompt, [source()], ['销售'])


def test_one_fact_cannot_bypass_another_facts_missing_scope_field():
    a = source()
    b = a.drop(columns='地区').rename(columns={'销售额': '费用金额'})
    scope = compile_scope('只分析华东地区', [a, b], ['销售', '费用'])
    with pytest.raises(BusinessInvariantError, match='无法.*绑定'):
        apply_scope(scope, [a, b], ['销售', '费用'])


def test_failed_business_constraint_cannot_pass_file_acceptance():
    result = {'inputs_unchanged': True, 'steps_completed': 1, 'acceptance': {'status': 'passed'},
              'download_url': '/download/test', 'pending_reviews': 0, 'status': 'completed',
              'constraint_checks': [{'item': '完整成本证据', 'status': 'needs_review'}]}
    assert assess_task(result, expected_steps=1)['status'] == 'needs_review'


@pytest.mark.parametrize('start,end', [('2026-08-01', ''), ('NaT', '2026-09-01'), ('2026-09-01', '2026-08-01')])
def test_malformed_scope_is_rejected_before_execution(start, end):
    with pytest.raises(BusinessInvariantError):
        TaskScope.from_dict({'start': start, 'end_exclusive': end, 'entities': [], 'top_n': None})
