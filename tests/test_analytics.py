from __future__ import annotations

import json

import pandas as pd
import pandas.testing as pdt
import pytest

from excel_data_toolbox import (
    aggregate_trend,
    assess_data_quality,
    category_contribution,
    compare_tables,
    correlation_matrix,
    cross_pivot,
    descriptive_statistics,
    detect_outliers,
    rfm_segmentation,
)


def test_quality_report_scores_and_lists_actionable_issues_without_mutation() -> None:
    frame = pd.DataFrame(
        {
            "客户编号": ["000001", "000001", "000003", ""],
            "金额": [10, 10, None, 30],
            "状态": ["完成", "完成", "完成", "完成"],
            "混合": [1, 1, "3", 4],
        }
    )
    untouched = frame.copy(deep=True)

    report = assess_data_quality(frame, key_columns="客户编号")

    pdt.assert_frame_equal(frame, untouched)
    assert 0 <= report.score < 100
    assert report.grade in {"良好", "需关注", "较差"}
    codes = {issue.code for issue in report.issues}
    assert {"MISSING_VALUES", "DUPLICATE_ROWS", "DUPLICATE_KEYS", "MISSING_KEYS"} <= codes
    assert "CONSTANT_COLUMN" in codes
    assert "MIXED_TYPES" in codes
    encoded = json.dumps(report.to_dict(), ensure_ascii=False)
    assert "客户编号" in encoded


def test_descriptive_statistics_and_correlation_protect_long_identifiers() -> None:
    frame = pd.DataFrame(
        {
            "订单号": ["9007199254740993", "9007199254740994", "9007199254740995"],
            "销售额": [10, 20, 30],
            "数量": [1, 2, 3],
            "数字文本": ["2.5", "5.0", "7.5"],
            "备注": ["甲", "甲", None],
        }
    )

    stats = descriptive_statistics(frame)
    stats_by_column = stats.set_index("column")

    assert stats_by_column.loc["销售额", "mean"] == 20
    assert stats_by_column.loc["数字文本", "sum"] == pytest.approx(15)
    assert pd.isna(stats_by_column.loc["订单号", "mean"])
    assert stats_by_column.loc["订单号", "semantic_type"] == "identifier"
    assert stats_by_column.loc["备注", "mode"] == "甲"

    correlations = correlation_matrix(frame)
    assert set(correlations.columns) == {"销售额", "数量", "数字文本"}
    assert "订单号" not in correlations.columns
    assert correlations.loc["销售额", "数量"] == pytest.approx(1.0)

    kendall = correlation_matrix(frame, method="kendall")
    assert kendall.loc["销售额", "数量"] == pytest.approx(1.0)
    reverse = correlation_matrix(
        pd.DataFrame({"正序": [1, 2, 3, 4], "倒序": [4, 3, 2, 1]}),
        method="kendall",
    )
    assert reverse.loc["正序", "倒序"] == pytest.approx(-1.0)


def test_outlier_detection_supports_iqr_and_zscore() -> None:
    frame = pd.DataFrame(
        {
            "客户编号": [f"{index:06d}" for index in range(10)],
            "金额": [10] * 9 + [1000],
            "数量": list(range(10)),
        }
    )

    iqr = detect_outliers(frame, columns=["金额"], method="iqr")

    assert iqr.outliers["row_position"].tolist() == [9]
    assert iqr.outliers.loc[0, "column"] == "金额"
    assert iqr.flagged_rows.loc[0, "客户编号"] == "000009"
    assert iqr.summary.loc[0, "outlier_count"] == 1
    assert frame.loc[9, "金额"] == 1000
    json.dumps(iqr.to_dict(), ensure_ascii=False, allow_nan=False)

    zscore = detect_outliers(
        frame, columns=["金额"], method="zscore", z_threshold=2.5
    )
    assert zscore.outliers["row_position"].tolist() == [9]

    collision = frame.assign(row_position="原值", outlier_columns="原字段")
    collision_result = detect_outliers(collision, columns=["金额"])
    assert collision_result.flagged_rows.loc[0, "row_position"] == 9
    assert collision_result.flagged_rows.loc[0, "row_position_source"] == "原值"

    finite_only = detect_outliers(
        pd.DataFrame({"金额": [1.0, 2.0, float("inf")]}), columns=["金额"]
    )
    assert finite_only.summary.loc[0, "valid_count"] == 2


def test_trend_aggregation_handles_chinese_columns_invalid_dates_and_groups() -> None:
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01", "2026-01-20", "2026-02-01", "错误日期", None],
            "区域": ["华东", "华东", "华东", "华东", "华南"],
            "销售额": [10, 20, 30, 40, 50],
        }
    )

    result = aggregate_trend(
        frame,
        date_column="日期",
        value_columns="销售额",
        frequency="月",
        aggregation="sum",
        group_by="区域",
        period_column="月份",
    )

    assert result.frequency == "month"
    assert result.invalid_date_count == 1
    assert result.used_rows == 3
    assert result.data["销售额"].tolist() == [30, 30]
    assert result.data["月份"].dt.month.tolist() == [1, 2]
    json.dumps(result.to_dict(), ensure_ascii=False, allow_nan=False)


def test_category_contribution_builds_pareto_and_preserves_category_text() -> None:
    frame = pd.DataFrame(
        {
            "商品编码": ["00001", "00002", "00003"],
            "销售额": [50, 30, 20],
        }
    )

    result = category_contribution(
        frame,
        category_columns="商品编码",
        value_column="销售额",
        pareto_threshold=0.8,
    )

    assert result.total == 100
    assert result.input_rows == 3
    assert result.used_rows == 3
    assert result.invalid_value_count == 0
    assert result.core_category_count == 2
    assert result.data["商品编码"].tolist() == ["00001", "00002", "00003"]
    assert result.data["contribution_pct"].tolist() == pytest.approx([0.5, 0.3, 0.2])
    assert result.data["pareto_group"].tolist() == ["核心贡献", "核心贡献", "长尾贡献"]
    json.dumps(result.to_dict(), ensure_ascii=False, allow_nan=False)


def test_cross_pivot_supports_counts_values_margins_and_blank_categories() -> None:
    frame = pd.DataFrame(
        {
            "区域": ["华东", "华东", "华南", None],
            "产品": ["A", "B", "A", "A"],
            "销售额": [10, 20, 30, 40],
        }
    )

    counts = cross_pivot(frame, index="区域", columns="产品", margins=True)
    assert set(counts["区域"]) == {"华东", "华南", "（空值）", "合计"}
    assert counts.loc[counts["区域"] == "华东", "A"].iloc[0] == 1
    assert counts.loc[counts["区域"] == "合计", "合计"].iloc[0] == 4

    values = cross_pivot(
        frame,
        index="区域",
        columns="产品",
        values="销售额",
        aggregation="sum",
    )
    assert values.loc[values["区域"] == "华东", "A"].iloc[0] == 10
    assert values.loc[values["区域"] == "华东", "B"].iloc[0] == 20

    collision = pd.DataFrame(
        {"区域": ["华东"], "产品": ["区域"], "销售额": [10]}
    )
    collision_pivot = cross_pivot(
        collision,
        index="区域",
        columns="产品",
        values="销售额",
        aggregation="sum",
    )
    assert list(collision_pivot.columns) == ["区域", "区域_2"]


def test_compare_tables_covers_all_buckets_and_keeps_long_keys_exact() -> None:
    old = pd.DataFrame(
        {
            "客户编号": [
                "9007199254740993",
                "0002",
                "0002",
                "0003",
                "0006",
                None,
            ],
            "姓名": ["甲", "乙", "乙旧", "丙", "己", "无键旧"],
            "金额": [10, 20, 21, 30, 60, 1],
        }
    )
    new = pd.DataFrame(
        {
            "客户编号": [
                "9007199254740993",
                "0003",
                "0004",
                "0004",
                "0005",
                None,
            ],
            "姓名": ["甲新", "丙", "丁", "丁新", "戊", "无键新"],
            "金额": [11, 30, 40, 41, 50, 2],
        }
    )

    result = compare_tables(old, new, key_columns="客户编号")

    assert result.summary["added_count"] == 1
    assert result.summary["removed_count"] == 1
    assert result.summary["modified_count"] == 1
    assert result.summary["unchanged_count"] == 1
    assert result.added["客户编号"].tolist() == ["0005"]
    assert result.removed["客户编号"].tolist() == ["0006"]
    assert result.modified.loc[0, "客户编号"] == "9007199254740993"
    assert result.modified.loc[0, "changed_columns"] == "姓名；金额"
    assert result.unchanged["客户编号"].tolist() == ["0003"]
    assert len(result.duplicate_keys_old) == 2
    assert len(result.duplicate_keys_new) == 2
    assert len(result.invalid_keys_old) == 1
    assert len(result.invalid_keys_new) == 1
    json.dumps(result.to_dict(), ensure_ascii=False, allow_nan=False)


def test_compare_tables_matches_safe_numeric_and_text_keys_but_not_leading_zero() -> None:
    old = pd.DataFrame({"键": [1, "001"], "值": [10, 20]})
    new = pd.DataFrame({"键": ["1", "002"], "值": [10, 99]})

    result = compare_tables(old, new, key_columns="键")

    assert result.summary["unchanged_count"] == 1
    assert result.summary["added_count"] == 1
    assert result.summary["removed_count"] == 1
    assert result.unchanged.loc[0, "键"] == "1"

    long_key_result = compare_tables(
        pd.DataFrame({"键": pd.Series(dtype="int64"), "值": pd.Series(dtype="int64")}),
        pd.DataFrame({"键": [9_007_199_254_740_993], "值": [1]}),
        key_columns="键",
    )
    assert long_key_result.to_dict()["added"][0]["键"] == "9007199254740993"


def test_rfm_segmentation_preserves_customer_ids_and_returns_invalid_rows() -> None:
    frame = pd.DataFrame(
        {
            "客户编号": ["000001", "000001", "000002", "000003", "000003", "000004"],
            "订单号": ["A1", "A2", "B1", "C1", "C2", "D1"],
            "日期": [
                "2026-01-01",
                "2026-03-01",
                "2025-01-01",
                "2026-02-20",
                "2026-02-21",
                "错误日期",
            ],
            "金额": [100, 200, 10, 80, 90, 50],
        }
    )

    result = rfm_segmentation(
        frame,
        customer_column="客户编号",
        date_column="日期",
        amount_column="金额",
        transaction_column="订单号",
        reference_date="2026-03-02",
        quantiles=3,
    )

    assert set(result.customers["客户编号"]) == {"000001", "000002", "000003"}
    customer_one = result.customers.loc[result.customers["客户编号"] == "000001"].iloc[0]
    assert customer_one["recency_days"] == 1
    assert customer_one["frequency"] == 2
    assert customer_one["monetary"] == 300
    assert len(result.invalid_rows) == 1
    assert result.invalid_rows.loc[0, "客户编号"] == "000004"
    assert "日期" in result.invalid_rows.loc[0, "invalid_reason"]
    assert int(result.segment_summary["customer_count"].sum()) == 3
    assert result.reference_date == pd.Timestamp("2026-03-02")
    json.dumps(result.to_dict(), ensure_ascii=False, allow_nan=False)


def test_analytics_validation_and_empty_inputs() -> None:
    empty = pd.DataFrame(columns=["编号", "金额"])
    quality = assess_data_quality(empty, key_columns="编号")
    assert quality.score == 0

    trend_empty = pd.DataFrame(columns=["日期", "金额"])
    assert aggregate_trend(
        trend_empty, date_column="日期", value_columns="金额"
    ).data.empty
    rfm_empty = pd.DataFrame(columns=["客户", "日期", "金额"])
    assert rfm_segmentation(
        rfm_empty,
        customer_column="客户",
        date_column="日期",
        amount_column="金额",
    ).customers.empty

    with pytest.raises(KeyError, match="不存在的列"):
        assess_data_quality(empty, key_columns="缺少")

    with pytest.raises(KeyError, match="不存在的列"):
        correlation_matrix(pd.DataFrame({"a": [1]}), columns=["b"])
    with pytest.raises(ValueError, match="非负"):
        category_contribution(
            pd.DataFrame({"类别": ["A", "B"], "金额": [10, -1]}),
            category_columns="类别",
            value_column="金额",
        )
    with pytest.raises(ValueError, match="重复列名"):
        detect_outliers(pd.DataFrame([[1, 2]], columns=["a", "a"]))

    huge_numeric_identifier = pd.DataFrame(
        {"序列": [9_007_199_254_740_992, 9_007_199_254_740_994], "金额": [1, 2]}
    )
    assert list(correlation_matrix(huge_numeric_identifier).columns) == ["金额"]
    huge_stats = descriptive_statistics(huge_numeric_identifier).set_index("column")
    assert huge_stats.loc["序列", "semantic_type"] == "identifier"
    assert pd.isna(huge_stats.loc["序列", "mean"])
