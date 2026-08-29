from __future__ import annotations

from decimal import Decimal
import json

import pandas as pd
import pandas.testing as pdt
import pytest

from excel_data_toolbox import (
    ReconciliationResult,
    reconcile_tables,
)


def test_primary_keys_classify_exact_amount_date_and_review_buckets() -> None:
    left = pd.DataFrame(
        {
            "流水号": ["A", "B", "C", "D", "L"],
            "日期": ["2026-01-01"] * 5,
            "金额": ["100", "100", "100", "100", "5"],
            "摘要": ["甲", "乙", "丙", "丁", "仅左"],
        }
    )
    right = pd.DataFrame(
        {
            "交易号": ["A", "B", "C", "D", "R"],
            "入账日期": [
                "2026-01-01",
                "2026-01-01",
                "2026-01-02",
                "2026-01-02",
                "2026-01-03",
            ],
            "入账金额": ["100.00", "100.05", "100", "100.05", "9"],
            "对方": ["甲", "乙", "丙", "丁", "仅右"],
        }
    )

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="入账金额",
        left_date="日期",
        right_date="入账日期",
        left_key_columns="流水号",
        right_key_columns="交易号",
        left_secondary_columns="摘要",
        right_secondary_columns="对方",
        amount_tolerance="0.10",
        date_tolerance_days=2,
    )

    assert isinstance(result, ReconciliationResult)
    assert result.matched["left__流水号"].tolist() == ["A"]
    assert result.amount_difference["left__流水号"].tolist() == ["B"]
    assert result.amount_difference.loc[0, "amount_difference_decimal"] == Decimal(
        "0.05"
    )
    assert result.date_difference["left__流水号"].tolist() == ["C"]
    assert result.date_difference.loc[0, "date_difference_days"] == 1
    assert result.review["left__流水号"].tolist() == ["D"]
    assert "金额差" in result.review.loc[0, "match_reason"]
    assert result.left_only["left__流水号"].tolist() == ["L"]
    assert result.right_only["right__交易号"].tolist() == ["R"]
    assert result.summary["matched_count"] == 1
    assert result.summary["review_group_count"] == 1


def test_keyless_auto_match_requires_mutual_unique_exact_secondary_evidence() -> None:
    left = pd.DataFrame(
        {
            "编号": [None, None],
            "日期": ["2026-02-01", "2026-02-01"],
            "金额": [10, 20],
            "客户": ["上海 公司", "乙公司"],
        }
    )
    right = pd.DataFrame(
        {
            "编号2": [None, None, None],
            "日期2": ["2026-02-01", "2026-02-01", "2026-02-01"],
            "金额2": ["10.00", "20", "20"],
            "客户2": ["上海  公司", "乙公司", "乙公司"],
        }
    )

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_date="日期",
        right_date="日期2",
        left_key_columns="编号",
        right_key_columns="编号2",
        left_secondary_columns="客户",
        right_secondary_columns="客户2",
    )

    assert result.matched["left_row_position"].tolist() == [0]
    assert result.matched.loc[0, "match_type"] == "secondary_unique_exact"
    assert result.matched.loc[0, "match_score"] <= 0.95
    ambiguous = result.review.loc[result.review["left_row_position"] == 1]
    assert ambiguous["right_row_position"].tolist() == [1, 2]
    assert ambiguous["candidate_rank"].tolist() == [1, 2]
    assert set(ambiguous["candidate_count"]) == {2}

    no_secondary = reconcile_tables(
        left.iloc[[0]],
        right.iloc[[0]],
        left_amount="金额",
        right_amount="金额2",
        left_date="日期",
        right_date="日期2",
    )
    assert no_secondary.matched.empty
    assert len(no_secondary.review) == 1
    assert "未配置次级字段" in no_secondary.review.loc[0, "match_reason"]


def test_duplicate_key_on_either_side_quarantines_all_affected_rows() -> None:
    left = pd.DataFrame(
        {
            "订单号": ["0001", "0001", "9007199254740993"],
            "金额": [10, 10, 30],
        }
    )
    right = pd.DataFrame(
        {
            "订单编码": ["0001", 9_007_199_254_740_993],
            "金额2": [10, 30],
        }
    )

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_key_columns="订单号",
        right_key_columns="订单编码",
    )

    assert len(result.duplicates) == 3
    duplicate = result.duplicates.loc[result.duplicates["key_values"] == "0001"]
    assert set(duplicate["source_side"]) == {"left", "right"}
    assert set(duplicate["duplicate_count_left"]) == {2}
    assert set(duplicate["duplicate_count_right"]) == {1}
    assert result.matched.loc[0, "left__订单号"] == "9007199254740993"
    assert result.to_dict()["matched"][0]["right__订单编码"] == "9007199254740993"


def test_leading_zero_keys_remain_distinct_from_numeric_keys() -> None:
    left = pd.DataFrame({"键": ["001"], "金额": [1]})
    right = pd.DataFrame({"键2": [1], "金额2": [2]})

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_key_columns="键",
        right_key_columns="键2",
    )

    assert result.matched.empty
    assert result.left_only.loc[0, "left__键"] == "001"
    assert result.right_only.loc[0, "right__键2"] == 1


def test_decimal_split_candidates_are_exact_bounded_and_review_only() -> None:
    left = pd.DataFrame(
        {
            "编号": ["L1", "L2", "L3"],
            "日期": ["2026-03-01"] * 3,
            "金额": ["0.3", "0.4", "0.6"],
        }
    )
    right = pd.DataFrame(
        {
            "编号2": ["R1", "R2", "R3"],
            "日期2": ["2026-03-01"] * 3,
            "金额2": ["0.1", "0.2", "1.0"],
        }
    )

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_date="日期",
        right_date="日期2",
        left_key_columns="编号",
        right_key_columns="编号2",
        enable_split_candidates=True,
        max_candidates_per_row=10,
        max_split_combinations=100,
    )

    assert result.matched.empty
    assert set(result.review["match_type"]) == {"split_1_to_2", "split_2_to_1"}
    assert result.summary["split_candidate_groups"] == 2
    assert result.summary["review_candidate_rows"] == 4
    assert set(result.review["group_amount_difference_decimal"]) == {Decimal("0.0")}
    assert all("人工确认" in reason for reason in result.review["match_reason"])
    assert result.left_only.empty
    assert result.right_only.empty


def test_split_search_stops_at_configured_combination_limit() -> None:
    left = pd.DataFrame({"金额": [100]})
    right = pd.DataFrame({"金额2": [1, 2, 3, 4, 5]})

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        enable_split_candidates=True,
        max_candidates_per_row=5,
        max_split_combinations=2,
    )

    assert result.summary["split_limit_hit"] is True
    assert result.summary["split_combinations_evaluated"] == 2
    assert result.review.empty


def test_inputs_are_not_modified_and_json_export_is_decimal_safe() -> None:
    left = pd.DataFrame(
        {"键": ["A", "B"], "日期": ["2026-01-01", "错误"], "金额": ["¥1,000.10", None]}
    )
    right = pd.DataFrame(
        {"键2": ["A"], "日期2": ["2026-01-01"], "金额2": ["1000.10"]}
    )
    left_before = left.copy(deep=True)
    right_before = right.copy(deep=True)

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_date="日期",
        right_date="日期2",
        left_key_columns="键",
        right_key_columns="键2",
    )

    pdt.assert_frame_equal(left, left_before)
    pdt.assert_frame_equal(right, right_before)
    assert result.matched.loc[0, "left_amount_decimal"] == Decimal("1000.10")
    assert "金额为空" in result.left_only.loc[0, "unmatched_reason"]
    encoded = json.dumps(result.to_dict(), ensure_ascii=False, allow_nan=False)
    assert "1000.10" in encoded


def test_candidate_pair_limit_and_argument_validation_fail_clearly() -> None:
    left = pd.DataFrame({"金额": [10, 10, 10], "日期": ["2026-01-01"] * 3})
    right = pd.DataFrame({"金额2": [10, 10, 10], "日期2": ["2026-01-01"] * 3})

    with pytest.raises(ValueError, match="max_candidate_pairs"):
        reconcile_tables(
            left,
            right,
            left_amount="金额",
            right_amount="金额2",
            max_candidate_pairs=4,
        )
    with pytest.raises(ValueError, match="必须同时提供"):
        reconcile_tables(
            left,
            right,
            left_amount="金额",
            right_amount="金额2",
            left_date="日期",
        )
    with pytest.raises(ValueError, match="必须同时提供"):
        reconcile_tables(
            left,
            right,
            left_amount="金额",
            right_amount="金额2",
            left_key_columns="日期",
        )
    with pytest.raises(ValueError, match="非负金额"):
        reconcile_tables(
            left,
            right,
            left_amount="金额",
            right_amount="金额2",
            amount_tolerance="-0.01",
        )


def test_candidate_truncation_keeps_every_source_row_visible() -> None:
    left = pd.DataFrame({"金额": [10], "名称": ["左"]})
    right = pd.DataFrame(
        {"金额2": [10, 10, 10, 10], "名称2": ["甲", "乙", "丙", "丁"]}
    )

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_secondary_columns="名称",
        right_secondary_columns="名称2",
        max_candidates_per_row=2,
    )

    assert result.review["right_row_position"].tolist() == [0, 1]
    assert result.right_only["right_row_position"].tolist() == [2, 3]
    assert all(
        "展示上限" in reason for reason in result.right_only["unmatched_reason"]
    )
    represented_right = set(result.review["right_row_position"]) | set(
        result.right_only["right_row_position"]
    )
    assert represented_right == set(range(len(right)))
    assert result.summary["candidate_groups_truncated"] == 1


def test_empty_tables_return_stable_schemas_and_zero_summary() -> None:
    left = pd.DataFrame(columns=["编号", "金额"])
    right = pd.DataFrame(columns=["编号2", "金额2"])

    result = reconcile_tables(
        left,
        right,
        left_amount="金额",
        right_amount="金额2",
        left_key_columns="编号",
        right_key_columns="编号2",
    )

    assert result.matched.empty
    assert result.review.empty
    assert result.duplicates.empty
    assert result.summary["left_rows"] == 0
    assert result.summary["right_rows"] == 0
    assert "left_row_position" in result.matched.columns
    assert "left_row_position" in result.left_only.columns
