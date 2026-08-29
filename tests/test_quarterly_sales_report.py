from __future__ import annotations

import unittest
from unittest.mock import patch

import pandas as pd

import excel_data_toolbox.server as server_module
from excel_data_toolbox.io_utils import _promote_embedded_sales_header
from excel_data_toolbox.sales_report import build_quarterly_sales_management_report
from excel_data_toolbox.server import AppSession, ToolboxHandler


def _monthly_tables() -> list[pd.DataFrame]:
    january = pd.DataFrame(
        {
            "订单编号": ["J001", "J002", "J001"],
            "下单日期": ["2026-01-03", "2026/01/05", "2026/01/03"],
            "产品类别 ": ["智能设备", "软件服务", "智能设备 "],
            "销售区域": ["华东", "华南", "华东 "],
            "业务员": ["张伟", "李娜", "张伟"],
            "数量": [12, "8件", 12],
            "销售金额": [36000, "￥24,000", "36,000元"],
            "成本": [22000, "9,000", 22000],
            "客户满意度": [5, "4分", 5],
            "订单状态": ["已完成", "已取消", "完成"],
            "备注": ["", "客户取消", "重复导出"],
        }
    )
    february = pd.DataFrame(
        {
            "流水号": ["F001", "F002", "F003"],
            "日期": [46060, "2026/02/10", "2026-02-22"],
            "产品": ["数据服务", "软件服务", "数据服务"],
            "销售人员": ["王强", "张伟", "赵敏"],
            "地区": ["华北", "华南", "华东"],
            "成交额": [62500, -15000, None],
            "采购成本": [22000, -6000, 14000],
            "件数": [25, 5, 16],
            "评分": [4, 4, 4],
            "状态": ["已完成", "已退款", "已完成"],
            "备注": ["", "整单退款", "成交金额缺失"],
        }
    )
    march = pd.DataFrame(
        {
            "单号": ["M001", "M002", "M002"],
            "业务日期": ["2026年3月3日", "2026.03.12", "2026/03/12"],
            "品类": ["数据服务", "智能设备", "智能设备 "],
            "区域": ["华东", "西 南", "西南"],
            "负责人": ["赵敏", "陈 浩", "陈浩"],
            "成交数量": [22, 11, 11],
            "含税销售额": ["55,000", "￥33,000", "33,000元"],
            "采购/服务成本": [20000, 21000, 21000],
            "满意度评分": [5, 2, "2分"],
            "是否有效": ["有效", "是", "是"],
            "来源": ["CRM", "人工补录", "人工重复录入"],
            "临时列": ["", "投诉待处理", ""],
        }
    )
    return [january, february, march]


class QuarterlySalesReportTests(unittest.TestCase):
    def test_embedded_sales_header_is_promoted(self) -> None:
        raw = pd.DataFrame(
            [["订单编号", "下单日期", "产品类别", "销售区域", "业务员", "销售金额"], ["J001", "2026-01-01", "A", "华东", "张伟", 100]],
            columns=["报表标题", "Unnamed: 1", "Unnamed: 2", "Unnamed: 3", "Unnamed: 4", "Unnamed: 5"],
        )
        promoted = _promote_embedded_sales_header(raw)
        self.assertEqual(list(promoted.columns), ["订单编号", "下单日期", "产品类别", "销售区域", "业务员", "销售金额"])
        self.assertEqual(promoted.iloc[0]["订单编号"], "J001")

    def test_quarterly_cleaning_and_report_are_auditable(self) -> None:
        result = build_quarterly_sales_management_report(
            _monthly_tables(),
            source_names=["1月销售流水", "2月销售明细", "3月临时补录"],
        )
        self.assertEqual(result.report["raw_rows"], 9)
        self.assertEqual(result.report["duplicate_rows_removed"], 2)
        self.assertEqual(result.report["invalid_rows_removed"], 3)
        self.assertEqual(result.report["valid_rows"], 4)
        self.assertEqual(result.report["attention_rows"], 1)
        self.assertEqual(len(result.outputs), 8)
        self.assertEqual(result.outputs["季度合并数据"]["订单编号"].tolist(), ["J001", "F001", "M001", "M002"])
        self.assertEqual(result.outputs["季度合并数据"].loc[3, "地区"], "西南")
        self.assertEqual(len(result.outputs["清洗审计"]), 9)

    def test_unified_command_routes_without_deepseek(self) -> None:
        session = AppSession()
        handler = object.__new__(ToolboxHandler)
        try:
            ids = [
                session.add_table(name, frame, source="单元测试", original=True)
                for name, frame in zip(["1月销售流水", "2月销售明细", "3月临时补录"], _monthly_tables())
            ]
            prompt = "这是一季度三个销售表，请清洗、去重、排除无效订单并合并，分析产品、地区、销售人员、月度趋势和重点关注，生成老板看的 Excel 经营报表。"
            with (
                patch.object(server_module, "SESSION", session),
                patch.object(server_module, "_project_ai_config", return_value={"configured": False, "api_key": "", "model": "deepseek-v4-flash"}),
                patch.object(server_module.DeepSeekClient, "classify_unified_request", side_effect=AssertionError("季度标准流程不应访问 DeepSeek")),
            ):
                response = handler._ai_unified({"prompt": prompt, "table_ids": ids})
            self.assertEqual(response["mode"], "data")
            self.assertTrue(response["auto_execute"])
            self.assertEqual(response["plan"]["steps"][0]["operation"], "quarterly_sales_report")
        finally:
            session.close()


if __name__ == "__main__":
    unittest.main()
