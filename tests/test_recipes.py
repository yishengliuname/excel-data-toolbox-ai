from __future__ import annotations

import json

import pandas as pd
import pandas.testing as pdt
import pytest

from excel_data_toolbox.recipes import (
    ProcessingRecipe,
    RecipeStep,
    fingerprint_dataframe,
    run_recipe,
)


def test_recipe_json_roundtrip_is_strict_safe_and_immutable() -> None:
    original_params = {
        "conditions": [{"column": "状态", "operator": "eq", "value": "完成"}],
        "combine": "and",
    }
    recipe = ProcessingRecipe(
        name="月度订单",
        description="只保留已完成订单",
        steps=[RecipeStep("filter", original_params, name="完成订单")],
    )
    original_params["conditions"][0]["value"] = "已取消"

    payload = recipe.to_json()
    restored = ProcessingRecipe.from_json(payload)

    assert restored == recipe
    assert restored.to_dict()["steps"][0]["params"]["conditions"][0]["value"] == "完成"
    assert "月度订单" in payload
    json.loads(payload)

    with pytest.raises(ValueError, match="不支持的配方操作"):
        RecipeStep("python", {"code": "__import__('os')"})
    with pytest.raises(ValueError, match="不支持的字段"):
        RecipeStep.from_dict({"operation": "clean", "params": {}, "callable": "x"})
    with pytest.raises(TypeError, match="JSON"):
        RecipeStep("fill_missing", {"default": lambda value: value})
    with pytest.raises(ValueError, match="不支持的筛选比较符"):
        RecipeStep(
            "filter",
            {"conditions": [{"column": "状态", "operator": "eval", "value": "x"}]},
        )
    with pytest.raises(ValueError, match="JSON"):
        ProcessingRecipe.from_json("{not-json")
    with pytest.raises(ValueError, match="NaN"):
        ProcessingRecipe.from_json(
            '{"name":"x","steps":[{"operation":"fill_missing","params":{"default":NaN}}]}'
        )


def test_run_recipe_chains_steps_without_mutating_long_identifiers() -> None:
    frame = pd.DataFrame(
        {
            "订单号": [
                "9007199254740993",
                "9007199254740994",
                "9007199254740994",
                "9007199254740995",
            ],
            "状态": [" 待处理 ", "完成", "完成", None],
            "金额": ["10", "20", "20", None],
            "备注": [None, "同一订单", "同一订单", "缺金额"],
        }
    )
    untouched = frame.copy(deep=True)
    recipe = ProcessingRecipe(
        "订单标准化",
        [
            RecipeStep(
                "clean",
                {
                    "drop_duplicates": False,
                    "missing_strategy": "keep",
                    "type_inference_threshold": 1.0,
                },
            ),
            RecipeStep(
                "replace",
                {
                    "replacements": [
                        {"column": "状态", "old": "待处理", "new": "处理中"}
                    ]
                },
            ),
            RecipeStep(
                "fill_missing", {"values": {"状态": "未知", "金额": 0}}
            ),
            RecipeStep(
                "drop_duplicates", {"subset": ["订单号"], "keep": "first"}
            ),
            RecipeStep(
                "filter",
                {
                    "conditions": [
                        {"column": "金额", "operator": "gte", "value": 10}
                    ]
                },
            ),
            RecipeStep(
                "select_rename_sort",
                {
                    "columns": ["订单号", "状态", "金额"],
                    "rename": {"金额": "成交额"},
                    "sort_by": "成交额",
                    "ascending": False,
                },
            ),
        ],
    )

    result, report = run_recipe(frame, recipe)

    pdt.assert_frame_equal(frame, untouched)
    assert result["订单号"].tolist() == ["9007199254740994", "9007199254740993"]
    assert result["成交额"].tolist() == [20, 10]
    assert result.loc[1, "状态"] == "处理中"
    assert report.input_fingerprint.row_count == 4
    assert report.output_fingerprint.row_count == 2
    assert [step.operation for step in report.steps] == [
        "clean",
        "replace",
        "fill_missing",
        "drop_duplicates",
        "filter",
        "select_rename_sort",
    ]
    assert report.steps[3].details["removed_rows"] == 1
    assert report.steps[4].rows_before == 3
    assert report.steps[4].rows_after == 2


def test_dry_run_returns_isolated_preview_skips_disabled_steps_and_warns() -> None:
    frame = pd.DataFrame({"编号": ["A", "B"], "状态": ["完成", "完成"]})
    recipe = ProcessingRecipe(
        "预览",
        [
            RecipeStep(
                "replace",
                {"replacements": [{"column": "状态", "old": "不存在", "new": "x"}]},
            ),
            RecipeStep(
                "filter",
                {"conditions": [{"column": "编号", "operator": "eq", "value": "A"}]},
                enabled=False,
            ),
        ],
    )

    preview, report = run_recipe(frame, recipe, dry_run=True)

    pdt.assert_frame_equal(preview, frame)
    assert preview is not frame
    assert report.dry_run is True
    assert report.steps[0].warnings == ("替换规则没有匹配任何单元格",)
    assert report.steps[1].status == "skipped"
    assert report.steps[1].warnings == ("步骤已禁用",)
    preview.loc[0, "编号"] = "已改"
    assert frame.loc[0, "编号"] == "A"


def test_filter_uses_only_structured_whitelisted_comparisons() -> None:
    frame = pd.DataFrame(
        {
            "名称": ["华北一区", "华东二区", "华北三区", None],
            "金额": [5, 10, 20, 30],
            "状态": ["新建", "完成", "完成", "取消"],
        }
    )
    recipe = ProcessingRecipe(
        "筛选",
        [
            RecipeStep(
                "filter",
                {
                    "combine": "and",
                    "conditions": [
                        {"column": "名称", "operator": "contains", "value": "华北"},
                        {"column": "金额", "operator": "between", "value": [10, 25]},
                        {"column": "状态", "operator": "in", "value": ["完成", "取消"]},
                    ],
                },
            )
        ],
    )

    result, report = run_recipe(frame, recipe)

    assert result.to_dict(orient="records") == [
        {"名称": "华北三区", "金额": 20, "状态": "完成"}
    ]
    details = report.steps[0].details
    assert details["removed_rows"] == 3
    assert details["conditions"] == (
        {"column": "名称", "operator": "contains"},
        {"column": "金额", "operator": "between"},
        {"column": "状态", "operator": "in"},
    )


def test_summary_allows_named_aggregations_but_not_callables() -> None:
    frame = pd.DataFrame(
        {"区域": ["华东", "华东", "华北"], "金额": [10, 20, 5], "订单": ["A", "B", "C"]}
    )
    recipe = ProcessingRecipe(
        "汇总",
        [
            RecipeStep(
                "summary",
                {
                    "by": "区域",
                    "aggregations": {"金额": ["sum", "mean"], "订单": "count"},
                },
            )
        ],
    )

    result, report = run_recipe(frame, recipe)

    assert list(result.columns) == ["区域", "金额_sum", "金额_mean", "订单_count"]
    assert report.steps[0].details["group_count"] == 2
    with pytest.raises(ValueError, match="不支持的汇总聚合"):
        RecipeStep(
            "summary", {"by": "区域", "aggregations": {"金额": "__getattribute__"}}
        )
    with pytest.raises(TypeError, match="JSON"):
        RecipeStep("summary", {"by": "区域", "aggregations": {"金额": sum}})


def test_dataframe_fingerprint_is_stable_non_reversible_and_order_sensitive() -> None:
    frame = pd.DataFrame(
        {"订单号": ["9007199254740993", "9007199254740994"], "金额": [10, 20]}
    )

    first = fingerprint_dataframe(frame)
    second = fingerprint_dataframe(frame.copy(deep=True))
    changed = fingerprint_dataframe(frame.iloc[::-1].reset_index(drop=True))

    assert first == second
    assert first.sha256 != changed.sha256
    assert len(first.sha256) == 64
    serialised = json.dumps(first.to_dict(), ensure_ascii=False)
    assert "9007199254740993" not in serialised
    assert first.to_dict()["schema"] == [
        {"name": "订单号", "dtype": "str"},
        {"name": "金额", "dtype": "int64"},
    ]


def test_recipe_rejects_ambiguous_columns_and_bad_step_parameters() -> None:
    duplicated = pd.DataFrame([[1, 2]], columns=["同名", "同名"])
    recipe = ProcessingRecipe("空配方", [])
    with pytest.raises(ValueError, match="重复列名"):
        run_recipe(duplicated, recipe)

    with pytest.raises(KeyError, match="不存在的列"):
        run_recipe(
            pd.DataFrame({"A": [1]}),
            ProcessingRecipe(
                "错误", [RecipeStep("drop_duplicates", {"subset": ["B"]})]
            ),
        )
    with pytest.raises(ValueError, match="不支持的参数"):
        RecipeStep("drop_duplicates", {"subset": ["A"], "inplace": True})
    with pytest.raises(TypeError, match="params 必须是 JSON 对象"):
        RecipeStep("clean", [])
    with pytest.raises(TypeError, match="必须是布尔值"):
        RecipeStep("clean", {"infer_types": "false"})
    with pytest.raises(TypeError, match="必须是数字"):
        RecipeStep("clean", {"type_inference_threshold": True})
    with pytest.raises(ValueError, match="空值筛选"):
        RecipeStep(
            "filter",
            {"conditions": [{"column": "A", "operator": "eq", "value": None}]},
        )
