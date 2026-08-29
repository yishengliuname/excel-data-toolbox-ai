from __future__ import annotations

import copy
import unittest
from unittest.mock import patch

import pandas as pd

from excel_data_toolbox.chart_agent import ChartSpecValidationError, validate_chart_spec
from excel_data_toolbox.nl_agent import build_table_catalog
from excel_data_toolbox.server import (
    SESSION,
    ToolboxHandler,
    _apply_chart_presentation,
    _chart_payload,
    _is_direct_chart_request,
)


class AiChartSpecTests(unittest.TestCase):
    def setUp(self) -> None:
        self.frame = pd.DataFrame(
            {
                "月份": ["1月", "2月", "3月", "4月"],
                "地区": ["华东", "华南", "华东", "华北"],
                "销售额": [680000, 820000, 930000, 710000],
            }
        )
        self.catalog = build_table_catalog({"sales": self.frame})
        self.payload = {
            "schema_version": 1,
            "status": "ready",
            "normalized_request": "按月份展示销售额并增加目标线",
            "message": "已选择折线图",
            "clarification_questions": [],
            "warnings": [],
            "chart": {
                "chart_type": "line", "dimension": "月份", "measure": "销售额",
                "series": None, "aggregation": "sum", "top_n": 12,
                "date_grain": "auto", "start": None, "end": None, "progress": None,
                "style_3d": False, "title": "月度销售趋势", "theme": "business_dark",
                "number_format": "wan", "sort": "source",
                "reference_lines": [{"value": 800000, "label": "月度目标", "color": "#FFB020"}],
                "highlight": {"field": "月份", "value": "3月", "color": "#FF6B35"},
                "show_labels": True, "show_legend": False,
            },
        }

    def test_valid_spec_renders_with_presentation(self) -> None:
        spec = validate_chart_spec(self.payload, self.catalog)
        chart = _apply_chart_presentation(_chart_payload(self.frame, spec["chart"]), spec["chart"])
        self.assertEqual(chart["title"], "月度销售趋势")
        self.assertEqual(chart["theme"], "business_dark")
        self.assertEqual(chart["reference_lines"][0]["value"], 800000.0)
        self.assertEqual(chart["highlight"]["value"], "3月")
        self.assertEqual(chart["number_format"], "wan")

    def test_dynamic_max_highlight_is_resolved_from_local_values(self) -> None:
        payload = copy.deepcopy(self.payload)
        payload["chart"]["highlight"] = {
            "field": "月份", "value": "__MAX__", "color": "#FF6B35",
        }
        spec = validate_chart_spec(payload, self.catalog)
        chart = _apply_chart_presentation(
            _chart_payload(self.frame, spec["chart"]), spec["chart"]
        )
        self.assertEqual(chart["highlight"]["value"], "3月")

    def test_monthly_chart_is_not_forced_through_data_plan(self) -> None:
        self.assertTrue(
            _is_direct_chart_request(
                "分析当前销售数据，按月份生成销售额折线图，增加80万元目标线并高亮最高月份"
            )
        )
        self.assertFalse(_is_direct_chart_request("清洗去重后按月份生成销售额折线图"))

    def test_unknown_field_is_rejected(self) -> None:
        unsafe = copy.deepcopy(self.payload)
        unsafe["chart"]["measure"] = "不存在字段"
        with self.assertRaises(ChartSpecValidationError):
            validate_chart_spec(unsafe, self.catalog)

    def test_equivalent_version_spellings_are_normalized(self) -> None:
        for version in ("1", "1.0", "v1", "V1.0", 1.0):
            compatible = copy.deepcopy(self.payload)
            compatible["schema_version"] = version
            self.assertEqual(validate_chart_spec(compatible, self.catalog)["schema_version"], 1)

        legacy = copy.deepcopy(self.payload)
        legacy["version"] = legacy.pop("schema_version")
        self.assertEqual(validate_chart_spec(legacy, self.catalog)["schema_version"], 1)

        omitted = copy.deepcopy(self.payload)
        omitted.pop("schema_version")
        self.assertEqual(validate_chart_spec(omitted, self.catalog)["schema_version"], 1)

    def test_genuinely_unknown_version_is_rejected(self) -> None:
        unsafe = copy.deepcopy(self.payload)
        unsafe["schema_version"] = 2
        with self.assertRaises(ChartSpecValidationError):
            validate_chart_spec(unsafe, self.catalog)

    def test_blank_non_executable_copy_gets_local_defaults(self) -> None:
        payload = copy.deepcopy(self.payload)
        payload["normalized_request"] = ""
        payload["message"] = None
        payload["chart"]["title"] = "  "
        payload["chart"]["background_color"] = None
        payload["chart"]["font_size"] = ""
        payload["warnings"] = None
        result = validate_chart_spec(payload, self.catalog)
        self.assertTrue(result["normalized_request"])
        self.assertTrue(result["message"])
        self.assertIsNone(result["chart"]["title"])
        self.assertEqual(result["chart"]["background_color"], "#FFFFFF")
        self.assertEqual(result["chart"]["font_size"], 12)
        self.assertEqual(result["warnings"], [])

    def test_missing_presentation_fields_get_safe_defaults(self) -> None:
        sparse = copy.deepcopy(self.payload)
        sparse["chart"] = {
            "chart_type": "line", "dimension": "月份", "measure": "销售额",
            "top_n": None,
        }
        result = validate_chart_spec(sparse, self.catalog)
        self.assertEqual(result["chart"]["top_n"], 20)
        self.assertEqual(result["chart"]["aggregation"], "sum")
        self.assertEqual(result["chart"]["theme"], "default")
        self.assertTrue(result["chart"]["show_labels"])

    def test_unknown_chart_field_stays_rejected_after_normalization(self) -> None:
        unsafe = copy.deepcopy(self.payload)
        unsafe["chart"]["javascript"] = "alert(1)"
        with self.assertRaises(ChartSpecValidationError):
            validate_chart_spec(unsafe, self.catalog)

    def test_series_field_list_becomes_wide_table_multi_measure_chart(self) -> None:
        frame = pd.DataFrame(
            {
                "序号": [1, 2, 3],
                "第1轮得分": [90, 80, 70],
                "第2轮得分": [85, None, 75],
                "第3轮得分": [88, 76, 81],
                "空轮得分": [None, None, None],
            }
        )
        catalog = build_table_catalog({"scores": frame})
        payload = copy.deepcopy(self.payload)
        payload["chart"].update(
            {
                "chart_type": "bar",
                "dimension": "序号",
                "measure": None,
                "series": ["第1轮得分", "第2轮得分", "第3轮得分", "空轮得分"],
                "title": "多轮得分对比",
                "x_axis_label": "序号",
                "y_axis_label": "分数",
                "series_colors": ["红", "绿", "黄", "紫"],
                "background_color": "#FFFDF7",
                "text_color": "#202020",
                "font_size": 15,
                "legend_position": "right",
                "label_rotation": -30,
                "show_grid": False,
                "opacity": 0.75,
                "bar_gap": 0.35,
                "chart_height": 420,
                "y_min": 0,
                "y_max": 100,
                "reference_lines": [],
                "highlight": None,
            }
        )
        spec = validate_chart_spec(payload, catalog)
        self.assertEqual(spec["chart"]["chart_type"], "grouped_bar")
        self.assertEqual(spec["chart"]["measures"], ["第1轮得分", "第2轮得分", "第3轮得分", "空轮得分"])
        self.assertEqual(spec["chart"]["series_colors"], ["#E53935", "#2E7D32", "#FBC02D", "#7B1FA2"])
        self.assertIsNone(spec["chart"]["series"])
        chart = _apply_chart_presentation(_chart_payload(frame, spec["chart"]), spec["chart"])
        self.assertEqual(chart["labels"], ["1", "2", "3"])
        self.assertEqual(len(chart["series"]), 3)
        self.assertIsNone(chart["series"][1]["values"][1])
        self.assertNotIn("空轮得分", [item["name"] for item in chart["series"]])
        self.assertEqual([item["color"] for item in chart["series"]], ["#E53935", "#2E7D32", "#FBC02D"])
        self.assertEqual(chart["x_axis_label"], "序号")
        self.assertEqual(chart["y_axis_label"], "分数")
        self.assertEqual(chart["legend_position"], "right")
        self.assertEqual(chart["chart_height"], 420)
        self.assertEqual(chart["y_max"], 100.0)
        self.assertFalse(chart["show_grid"])
        self.assertIn("已自动跳过", chart["summary"])

    def test_arbitrary_code_key_is_rejected(self) -> None:
        unsafe = copy.deepcopy(self.payload)
        unsafe["chart"]["python"] = "import os"
        with self.assertRaises(ChartSpecValidationError):
            validate_chart_spec(unsafe, self.catalog)

    def test_invalid_color_is_rejected(self) -> None:
        unsafe = copy.deepcopy(self.payload)
        unsafe["chart"]["highlight"]["color"] = "javascript:alert(1)"
        with self.assertRaises(ChartSpecValidationError):
            validate_chart_spec(unsafe, self.catalog)

    def test_clarification_has_no_executable_chart(self) -> None:
        payload = {
            "schema_version": 1, "status": "clarification",
            "normalized_request": "展示经营趋势", "message": "缺少指标",
            "clarification_questions": ["请问要分析销售额、利润还是订单量？"],
            "warnings": [], "chart": None,
        }
        result = validate_chart_spec(payload, self.catalog)
        self.assertIsNone(result["chart"])

    def test_endpoint_uses_schema_only_and_renders_locally(self) -> None:
        SESSION.reset()
        table_id = SESSION.add_table("销售生产数据", self.frame, source="测试")
        captured: dict[str, object] = {}

        def fake_create(_client, request, catalog, current_spec=None):
            captured["request"] = request
            captured["catalog"] = catalog
            captured["current_spec"] = current_spec
            return validate_chart_spec(self.payload, catalog)

        handler = object.__new__(ToolboxHandler)
        with patch("excel_data_toolbox.server.DeepSeekClient.create_chart_spec", fake_create):
            response = handler._ai_chart_plan(
                {
                    "prompt": "按月份画销售额趋势并加目标线",
                    "api_key": "sk-" + "x" * 32,
                    "model": "deepseek-v4-flash",
                    "table_id": table_id,
                    "current_spec": None,
                }
            )
        self.assertEqual(response["status"], "ready")
        self.assertEqual(response["chart"]["theme"], "business_dark")
        serialized = str(captured["catalog"])
        self.assertNotIn("680000", serialized)
        self.assertNotIn("华东", serialized)

    def test_unified_command_routes_to_chart_without_frontend_key(self) -> None:
        SESSION.reset()
        table_id = SESSION.add_table("销售生产数据", self.frame, source="测试")
        route = {
            "intent": "chart", "normalized_request": "按月展示销售额",
            "data_request": "", "chart_request": "按月份画销售额折线图",
            "engineering_category": None, "reason": "用户要求制作图表",
        }
        handler = object.__new__(ToolboxHandler)
        with (
            patch("excel_data_toolbox.server._project_ai_config", return_value={"configured": True, "api_key": "sk-" + "x" * 32, "model": "deepseek-v4-flash", "source": "project_env"}),
            patch("excel_data_toolbox.server.DeepSeekClient.classify_unified_request", return_value=route),
            patch("excel_data_toolbox.server.DeepSeekClient.create_chart_spec", return_value=self.payload),
        ):
            response = handler._ai_unified(
                {"prompt": "按月份画销售额折线图", "table_ids": [table_id], "current_chart_spec": None, "mode_hint": None}
            )
        self.assertEqual(response["mode"], "chart")
        self.assertEqual(response["chart"]["title"], "月度销售趋势")

    def test_unified_overrides_redundant_data_then_chart_route(self) -> None:
        SESSION.reset()
        table_id = SESSION.add_table("销售生产数据", self.frame, source="测试")
        route = {
            "intent": "data",
            "normalized_request": "按月份汇总销售额并生成折线图",
            "data_request": "按月份汇总销售额",
            "chart_request": "",
            "engineering_category": None,
            "reason": "包含汇总和图表",
        }
        dynamic_payload = copy.deepcopy(self.payload)
        dynamic_payload["chart"]["highlight"] = {
            "field": "月份", "value": "__MAX__", "color": "#FF6B35",
        }
        handler = object.__new__(ToolboxHandler)
        with (
            patch("excel_data_toolbox.server._project_ai_config", return_value={"configured": True, "api_key": "sk-" + "x" * 32, "model": "deepseek-v4-flash", "source": "project_env"}),
            patch("excel_data_toolbox.server.DeepSeekClient.classify_unified_request", return_value=route),
            patch("excel_data_toolbox.server.DeepSeekClient.create_chart_spec", return_value=dynamic_payload),
            patch("excel_data_toolbox.server.DeepSeekClient.create_plan") as create_plan,
        ):
            response = handler._ai_unified(
                {
                    "prompt": "分析当前销售数据，按月份生成销售额折线图，使用深色经营主题，增加80万元目标线并高亮最高月份。",
                    "table_ids": [table_id], "current_chart_spec": None, "mode_hint": None,
                }
            )
        create_plan.assert_not_called()
        self.assertEqual(response["mode"], "chart")
        self.assertEqual(response["chart"]["highlight"]["value"], "3月")


if __name__ == "__main__":
    unittest.main()
