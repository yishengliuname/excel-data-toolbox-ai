from __future__ import annotations

import unittest
from unittest.mock import patch

import pandas as pd

import excel_data_toolbox.server as server_module
from excel_data_toolbox.inventory_report import (
    build_inventory_management_report,
    infer_inventory_table_roles,
)
from excel_data_toolbox.server import AppSession, ToolboxHandler


def _inventory_tables() -> list[pd.DataFrame]:
    products = pd.DataFrame({
        "商品编码": ["P001", "P002", "P003"], "商品名称": ["快销品", "慢销品", "停售品"],
        "品类": ["A", "B", "B"], "供应商": ["甲", "乙", "乙"], "采购单价": [10, 20, 5],
        "零售价": [20, 40, 10], "安全库存": [10, 5, 0], "采购提前期(天)": [5, 5, 0],
        "目标库存天数": [20, 20, 0], "商品状态": ["正常", "正常", "停售"],
    })
    opening = pd.DataFrame({
        "商品编码": [" P001 ", "P002", "P003"], "商品名称": ["快销品", "慢销品", "停售品"],
        "仓库": ["主仓"] * 3, "期初库存": [20, 100, 50], "已锁定": [2, 0, 0], "不良品": [0, 0, 0], "备注": ["", "", ""],
    })
    purchases = pd.DataFrame({
        "入库日期": ["2026/08/01", "2026/08/01", "2026/08/20"], "入库单号": ["R1", "R1", "R2"],
        "商品编码": ["P001", "P001", "P002"], "入库数量": [10, 10, None], "采购价": [10, 10, 20],
        "供应商": ["甲", "甲", "乙"], "状态": ["已入库", "已入库", "已入库"], "备注": ["", "重复", "漏填"],
    })
    sales = pd.DataFrame({
        "出库日期": ["2026/08/05", "2026/08/15", "2026/08/22"], "出库单号": ["S1", "S2", "S3"],
        "商品编码": ["P001", "P001", "P002"], "商品名称": ["快销品", "快销品", "慢销品"],
        "数量": [15, 10, -2], "销售额": [300, 200, -80], "渠道": ["电商"] * 3,
        "状态": ["已完成", "已完成", "已退货"],
    })
    adjustments = pd.DataFrame({
        "日期": ["2026/08/10", "2026/08/20"], "商品编码": ["P001", "P002"], "调整数量": [-1, -3],
        "类型": ["盘亏", "报损"], "状态": ["已确认", "待确认"], "说明": ["差异", "待批"],
    })
    notes = pd.DataFrame({"仓库补充说明": ["当前分析截止日期：2026-08-23。", "已锁定和不良品不能作为可销售库存。"]})
    return [products, opening, purchases, sales, adjustments, notes]


class InventoryReportTests(unittest.TestCase):
    def test_roles_are_inferred_from_columns(self) -> None:
        roles = infer_inventory_table_roles(_inventory_tables())
        self.assertEqual(roles, {"products": 0, "opening": 1, "purchases": 2, "sales": 3, "adjustments": 4})

    def test_stock_math_and_audit(self) -> None:
        frames = _inventory_tables()
        result = build_inventory_management_report(
            frames,
            source_names=["商品资料", "期初库存", "采购入库", "销售出库", "库存调整", "仓库说明"],
        )
        detail = result.outputs["商品库存分析"].set_index("商品编码")
        self.assertEqual(detail.loc["P001", "当前账面库存"], 4)
        self.assertEqual(detail.loc["P001", "可销售库存"], 2)
        self.assertEqual(detail.loc["P003", "库存状态"], "停售积压")
        self.assertEqual(len(result.outputs), 9)
        self.assertEqual(result.report["manual_review_count"], 3)
        self.assertEqual(result.report["audit_issue_count"], 4)

    def test_unified_command_routes_without_deepseek(self) -> None:
        session = AppSession()
        handler = object.__new__(ToolboxHandler)
        try:
            names = ["商品资料", "6月期初库存", "采购入库", "销售出库", "库存调整", "仓库说明"]
            ids = [session.add_table(name, frame, source="单元测试", original=True) for name, frame in zip(names, _inventory_tables())]
            prompt = "请把采购、销售和库存理清楚，按期初+入库-出库+调整计算当前库存，识别补货、缺货和积压，生成老板看的 Excel 库存经营报表。"
            with (
                patch.object(server_module, "SESSION", session),
                patch.object(server_module, "_project_ai_config", return_value={"configured": False, "api_key": "", "model": "deepseek-v4-flash"}),
                patch.object(server_module.DeepSeekClient, "classify_unified_request", side_effect=AssertionError("库存标准流程不应访问 DeepSeek")),
            ):
                response = handler._ai_unified({"prompt": prompt, "table_ids": ids})
            self.assertEqual(response["mode"], "data")
            self.assertTrue(response["auto_execute"])
            self.assertEqual(response["plan"]["steps"][0]["operation"], "inventory_management_report")
        finally:
            session.close()


if __name__ == "__main__":
    unittest.main()
