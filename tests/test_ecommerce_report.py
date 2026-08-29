from __future__ import annotations

import math
import unittest
from pathlib import Path
from unittest.mock import patch

import excel_data_toolbox.server as server_module
from excel_data_toolbox.ecommerce_report import build_ecommerce_diagnosis_report
from excel_data_toolbox.io_utils import load_tables_from_files
from excel_data_toolbox.server import AppSession, ToolboxHandler

SOURCE = Path(r"D:\Users\liuyisheng\真实项目_多平台电商经营诊断终极测试.xlsx")


@unittest.skipUnless(SOURCE.exists(), "真实电商压力测试文件未提供")
class EcommerceDiagnosisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.tables = load_tables_from_files(SOURCE)

    def test_cross_fact_metrics_reconcile_to_customer_baseline(self) -> None:
        result = build_ecommerce_diagnosis_report(
            list(self.tables.values()), source_names=list(self.tables), user_request="全面诊断利润、现金和库存"
        )
        expected = {
            "valid_order_count": 32,
            "valid_line_count": 34,
            "buyer_paid": 26821,
            "refund_amount": 2478,
            "net_management_sales": 24343,
            "standard_cost": 16322,
            "product_gross_profit": 8021,
            "platform_fees": 2507,
            "actual_arrival": 21836,
            "ad_spend": 10440,
            "management_contribution": -4926,
            "latest_inventory_value": 167104,
        }
        for name, value in expected.items():
            self.assertTrue(math.isclose(float(result.report[name]), value, abs_tol=0.01), name)
        self.assertTrue(math.isclose(result.report["roas"], 2.132183908, rel_tol=1e-8))
        self.assertTrue(math.isclose(result.report["inventory_growth"], 0.605256585, rel_tol=1e-8))

    def test_channel_product_and_customer_risks_are_business_risks(self) -> None:
        result = build_ecommerce_diagnosis_report(list(self.tables.values()), source_names=list(self.tables))
        channel = result.outputs["渠道与广告诊断"].set_index("渠道")
        self.assertEqual(channel.loc["抖音", "管理贡献"], -6384)
        self.assertTrue(math.isclose(channel.loc["抖音", "ROAS"], 1.284210526, rel_tol=1e-8))
        product = result.outputs["商品利润质量"].set_index("SKU")
        self.assertEqual(product.loc["KIT01", "退款后商品毛利"], -242)
        customer = result.outputs["客户与回款风险"].set_index("客户ID")
        self.assertEqual(customer.loc["C025", "风险优先级"], "P0")
        self.assertTrue(math.isclose(customer.loc["C025", "退款率"], 0.5))
        action_names = result.outputs["风险行动计划"]["风险事项"].astype(str).tolist()
        self.assertTrue(any("抖音" in item and "负管理贡献" in item for item in action_names))
        self.assertTrue(any("库存资金占用" in item for item in action_names))

    def test_audit_rejects_amount_field_relationships_and_optional_notes_noise(self) -> None:
        result = build_ecommerce_diagnosis_report(list(self.tables.values()), source_names=list(self.tables))
        audit = result.outputs["数据口径与验收"]
        relations = audit.loc[audit["审计类型"].eq("关系证据"), "审计项"].astype(str)
        self.assertTrue(relations.str.contains("原订单号.*订单号", regex=True).any())
        self.assertFalse(relations.str.contains("退款金额.*退款金额", regex=True).any())
        self.assertFalse(audit["审计项"].astype(str).str.contains("备注缺失", regex=False).any())

    def test_natural_language_routes_to_domain_plugin_without_api(self) -> None:
        prompt = "公司销售增长但利润、现金和库存都恶化，请完整分析退款、平台费用、广告和客户风险并生成老板报表。"
        session = AppSession()
        handler = object.__new__(ToolboxHandler)
        try:
            ids = [session.add_table(name, frame, source=str(SOURCE), original=True) for name, frame in self.tables.items()]
            with (
                patch.object(server_module, "SESSION", session),
                patch.object(server_module, "_project_ai_config", return_value={"configured": False, "api_key": "", "model": "deepseek-v4-flash"}),
                patch.object(server_module.DeepSeekClient, "classify_unified_request", side_effect=AssertionError("不应依赖API")),
            ):
                routed = handler._ai_unified({"prompt": prompt, "table_ids": ids})
        finally:
            session.close()
        self.assertEqual(routed["plan"]["steps"][0]["operation"], "enterprise_diagnosis_report")
        self.assertIn("多平台电商", routed["plan"]["message"])

    def test_enterprise_plan_supports_direct_server_script_context(self) -> None:
        session = AppSession()
        try:
            ids = [session.add_table(name, frame, source=str(SOURCE), original=True) for name, frame in self.tables.items()]
            entries = [session.get(table_id) for table_id in ids]
            with patch.object(server_module, "__package__", ""):
                payload = server_module._enterprise_diagnosis_plan_payload(
                    entries,
                    "分析利润、现金、广告、退款与库存风险，并生成老板经营诊断报表。",
                )
        finally:
            session.close()
        self.assertEqual(payload["steps"][0]["operation"], "enterprise_diagnosis_report")
        self.assertIn("多平台电商", payload["message"])


if __name__ == "__main__":
    unittest.main()
