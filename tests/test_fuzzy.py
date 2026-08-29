from __future__ import annotations

import pandas as pd
import pandas.testing as pdt
import pytest

from excel_data_toolbox.fuzzy import (
    apply_value_mapping,
    cluster_similar_values,
    fuzzy_lookup,
    normalize_text,
)


def test_normalize_text_handles_case_whitespace_width_and_suffixes() -> None:
    assert normalize_text("  ＡＢＣ　科技有限责任公司  ") == "abc科技"
    assert normalize_text("星河工作室", company_suffixes=("工作室",)) == "星河"
    assert normalize_text("某某有限公司", company_suffixes=()) == "某某有限公司"
    assert normalize_text(None) == ""
    assert normalize_text(pd.NA) == ""


def test_cluster_similar_values_returns_reviewable_groups_without_mutation() -> None:
    frame = pd.DataFrame(
        {
            "供应商": [
                "上海星河科技有限公司",
                " 上海星河科技 ",
                "上海星河科技有限公司",
                "上海星河科枝有限公司",
                "远洋贸易有限公司",
                "远洋贸易",
                "完全无关",
                None,
            ],
            "金额": list(range(8)),
        }
    )
    untouched = frame.copy(deep=True)

    groups = cluster_similar_values(frame, "供应商", threshold=0.8, max_unique=20)

    pdt.assert_frame_equal(frame, untouched)
    assert list(groups.columns) == ["原值", "建议标准值", "相似度", "出现次数", "组ID"]
    assert "完全无关" not in groups["原值"].tolist()
    star_group = groups.loc[groups["原值"].astype(str).str.contains("星河")]
    assert star_group["组ID"].nunique() == 1
    assert set(star_group["原值"]) == {
        "上海星河科技有限公司",
        " 上海星河科技 ",
        "上海星河科枝有限公司",
    }
    assert set(star_group["建议标准值"]) == {"上海星河科技有限公司"}
    assert star_group.loc[
        star_group["原值"] == "上海星河科技有限公司", "出现次数"
    ].item() == 2
    assert star_group["相似度"].between(0.8, 1.0).all()

    ocean_group = groups.loc[groups["原值"].astype(str).str.contains("远洋")]
    assert ocean_group["组ID"].nunique() == 1
    assert ocean_group["相似度"].eq(1.0).all()


def test_cluster_similar_values_limits_work_and_validates_arguments() -> None:
    frame = pd.DataFrame({"名称": ["甲", "乙", "丙"]})

    with pytest.raises(ValueError, match="超过 max_unique"):
        cluster_similar_values(frame, "名称", max_unique=2)
    with pytest.raises(KeyError, match="不存在"):
        cluster_similar_values(frame, "缺失列")
    with pytest.raises(ValueError, match="0 到 1"):
        cluster_similar_values(frame, "名称", threshold=1.01)
    with pytest.raises(ValueError, match="正整数"):
        cluster_similar_values(frame, "名称", max_unique=0)

    empty = cluster_similar_values(pd.DataFrame({"名称": [None, "唯一值"]}), "名称")
    assert empty.empty
    assert list(empty.columns) == ["原值", "建议标准值", "相似度", "出现次数", "组ID"]


def test_fuzzy_lookup_matches_only_confident_rows_and_keeps_inputs() -> None:
    source = pd.DataFrame(
        {
            "录入供应商": ["上海星河科枝", "北京远洋贸易", "完全不存在", None],
            "订单号": [1, 2, 3, 4],
        },
        index=[10, 20, 30, 40],
    )
    lookup = pd.DataFrame(
        {
            "标准供应商": ["上海星河科技有限公司", "北京远洋贸易有限公司"],
            "供应商编号": ["XH001", "BJ001"],
            "区域": ["上海", "北京"],
        }
    )
    source_before = source.copy(deep=True)
    lookup_before = lookup.copy(deep=True)

    result = fuzzy_lookup(
        source,
        lookup,
        "录入供应商",
        "标准供应商",
        ["供应商编号", "区域"],
        threshold=0.8,
        ambiguous_gap=0.05,
    )

    pdt.assert_frame_equal(source, source_before)
    pdt.assert_frame_equal(lookup, lookup_before)
    assert result.index.tolist() == [10, 20, 30, 40]
    assert result["匹配状态"].tolist() == ["已匹配", "已匹配", "未匹配", "未匹配"]
    assert result.loc[10, "供应商编号"] == "XH001"
    assert result.loc[20, "区域"] == "北京"
    assert pd.isna(result.loc[30, "供应商编号"])
    assert result.loc[30, "候选值"] in lookup["标准供应商"].tolist()
    assert result.loc[10, "相似度"] >= 0.8
    assert pd.isna(result.loc[40, "候选值"])
    assert pd.isna(result.loc[40, "相似度"])


def test_fuzzy_lookup_marks_close_candidates_for_confirmation() -> None:
    source = pd.DataFrame({"名称": ["华兴科技"]})
    lookup = pd.DataFrame(
        {
            "标准名称": ["华星科技有限公司", "华新科技有限公司"],
            "代码": ["A", "B"],
        }
    )

    result = fuzzy_lookup(
        source,
        lookup,
        "名称",
        "标准名称",
        ["代码"],
        threshold=0.7,
        ambiguous_gap=0.05,
    )

    assert result.loc[0, "匹配状态"] == "待确认"
    assert pd.isna(result.loc[0, "代码"])
    assert result.loc[0, "候选值"] in lookup["标准名称"].tolist()
    assert result.loc[0, "次选候选值"] in lookup["标准名称"].tolist()
    assert result.loc[0, "候选值"] != result.loc[0, "次选候选值"]
    assert result.loc[0, "相似度"] == result.loc[0, "次选相似度"]


def test_fuzzy_lookup_does_not_silently_choose_conflicting_duplicate_key() -> None:
    source = pd.DataFrame({"名称": ["甲"]})
    lookup = pd.DataFrame(
        {
            "标准名称": ["甲公司", "甲有限公司"],
            "等级": ["金", "银"],
        }
    )

    result = fuzzy_lookup(
        source,
        lookup,
        "名称",
        "标准名称",
        "等级",
        threshold=1.0,
        ambiguous_gap=0.0,
    )

    assert result.loc[0, "匹配状态"] == "待确认"
    assert result.loc[0, "候选值"] == "甲公司"
    assert result.loc[0, "相似度"] == 1.0
    assert pd.isna(result.loc[0, "等级"])


def test_fuzzy_lookup_handles_output_collisions_empty_lookup_and_invalid_args() -> None:
    source = pd.DataFrame({"名称": ["甲"], "等级": ["原等级"]})
    lookup = pd.DataFrame({"标准名称": ["甲公司"], "等级": ["新等级"]})

    result = fuzzy_lookup(source, lookup, "名称", "标准名称", ["等级"])
    assert result.loc[0, "等级"] == "原等级"
    assert result.loc[0, "等级_查找"] == "新等级"

    empty_lookup = lookup.iloc[0:0].copy()
    unmatched = fuzzy_lookup(source, empty_lookup, "名称", "标准名称", ["等级"])
    assert unmatched.loc[0, "匹配状态"] == "未匹配"
    assert pd.isna(unmatched.loc[0, "候选值"])
    assert pd.isna(unmatched.loc[0, "等级_查找"])

    with pytest.raises(KeyError, match="不存在"):
        fuzzy_lookup(source, lookup, "缺失", "标准名称", ["等级"])
    with pytest.raises(ValueError, match="ambiguous_gap"):
        fuzzy_lookup(
            source,
            lookup,
            "名称",
            "标准名称",
            ["等级"],
            ambiguous_gap=-0.1,
        )
    with pytest.raises(ValueError, match="结果列"):
        fuzzy_lookup(
            pd.DataFrame({"名称": ["甲"], "匹配状态": ["旧状态"]}),
            lookup,
            "名称",
            "标准名称",
            ["等级"],
        )


def test_apply_value_mapping_only_applies_explicit_confirmed_values() -> None:
    frame = pd.DataFrame(
        {"供应商": ["上海星河科技", "上海星河科枝", "待确认", None], "金额": [1, 2, 3, 4]}
    )
    untouched = frame.copy(deep=True)

    result = apply_value_mapping(
        frame,
        "供应商",
        {"上海星河科枝": "上海星河科技有限公司"},
        output_column="标准供应商",
    )

    pdt.assert_frame_equal(frame, untouched)
    assert result["供应商"].tolist() == frame["供应商"].tolist()
    assert result.loc[1, "标准供应商"] == "上海星河科技有限公司"
    assert result.loc[0, "标准供应商"] == "上海星河科技"
    assert result.loc[2, "标准供应商"] == "待确认"
    assert pd.isna(result.loc[3, "标准供应商"])

    replaced = apply_value_mapping(frame, "供应商", {"待确认": "人工确认值"})
    assert replaced.loc[2, "供应商"] == "人工确认值"
    assert frame.loc[2, "供应商"] == "待确认"

    with pytest.raises(KeyError, match="不存在"):
        apply_value_mapping(frame, "缺失", {})
    with pytest.raises(TypeError, match="键值映射"):
        apply_value_mapping(frame, "供应商", [("a", "b")])  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="已存在"):
        apply_value_mapping(frame, "供应商", {}, output_column="金额")
