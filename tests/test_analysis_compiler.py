from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

import pandas as pd

from excel_data_toolbox.analysis_compiler import compile_analysis


class AnalysisCompilerTests(unittest.TestCase):
    def test_sales_structure_compiles_metrics_trend_ranking_and_profitability(self) -> None:
        frame = pd.DataFrame({
            "订单日期": ["2026-01-01", "2026-02-01", "2026-02-08"],
            "订单号": ["A1", "A2", "A3"],
            "地区": ["华东", "华南", "华东"],
            "产品": ["甲", "乙", "甲"],
            "销售额": [1000, 1600, 900],
            "成本": [600, 1000, 450],
            "利润": [400, 600, 450],
        })
        plan = compile_analysis(
            [frame], source_names=["订单明细"],
            user_request="分析整体销售、利润、产品排名、月度趋势和需要关注的问题，生成老板报表",
        )
        self.assertEqual(plan.domain_id, "sales")
        self.assertIn("profitability", plan.capabilities)
        self.assertIn("ranking", plan.capabilities)
        self.assertIn("trend", plan.capabilities)
        self.assertEqual(plan.metrics[0], "利润")
        self.assertIn("订单日期", plan.dates)
        self.assertGreaterEqual(len(plan.charts), 3)

    def test_inventory_structure_reports_missing_evidence_instead_of_guessing(self) -> None:
        frame = pd.DataFrame({
            "商品编码": ["P1", "P2", "P3"],
            "仓库": ["一仓", "一仓", "二仓"],
            "当前库存": [10, 80, 2],
            "安全库存": [12, 20, 5],
            "采购数量": [0, 50, 3],
        })
        plan = compile_analysis(
            [frame], source_names=["库存台账"],
            user_request="判断哪些商品缺货或积压，并分析采购、销售、库存情况",
        )
        self.assertEqual(plan.domain_id, "inventory")
        self.assertIn("inventory", plan.capabilities)
        self.assertIn("库存", plan.metrics[0])
        self.assertFalse(any("库存判断" in item for item in plan.missing_evidence))

    def test_workforce_structure_is_detected_without_request_specific_python(self) -> None:
        frame = pd.DataFrame({
            "月份": ["2026-01", "2026-01", "2026-02"],
            "员工姓名": ["甲", "乙", "甲"],
            "部门": ["销售", "运营", "销售"],
            "绩效得分": [92, 76, 88],
            "加班工时": [8, 35, 18],
            "薪资": [9000, 8500, 9200],
        })
        plan = compile_analysis(
            [frame], source_names=["员工月度记录"],
            user_request="整合员工绩效、薪资与加班，找出表现好和需要关注的人员",
        )
        self.assertEqual(plan.domain_id, "workforce")
        self.assertIn("workforce", plan.capabilities)
        self.assertIn("ranking", plan.capabilities)
        self.assertIn("绩效得分", plan.metrics)

    def test_generated_summary_is_never_selected_over_raw_fact(self) -> None:
        summary = pd.DataFrame({"指标": ["总收入", "总利润"], "结果": [1000, 200], "说明": ["汇总", "汇总"]})
        fact = pd.DataFrame({
            "日期": ["2026-01-01", "2026-01-02", "2026-01-03"],
            "订单号": ["A", "B", "C"], "客户": ["甲", "乙", "丙"], "销售额": [300, 400, 300],
        })
        plan = compile_analysis(
            [summary, fact], source_names=["管理层诊断总览", "POS销售明细"], user_request="全面分析"
        )
        self.assertEqual(plan.primary_table, "POS销售明细")
        self.assertEqual(plan.table_profiles[0].role, "summary")

    def test_new_domain_can_be_added_only_by_configuration(self) -> None:
        payload = {
            "schema_version": 1,
            "domains": [{
                "id": "education", "label": "教育培训",
                "anchors": ["课程", "学员", "班级"],
                "concepts": {
                    "date": ["上课日期"], "customer": ["学员"], "product": ["课程"],
                    "score": ["结课成绩"], "revenue": ["学费收入"],
                },
            }],
        }
        frame = pd.DataFrame({
            "上课日期": ["2026-01-01", "2026-02-01"], "学员": ["甲", "乙"],
            "课程": ["数据", "财务"], "结课成绩": [90, 82], "学费收入": [3000, 2600],
        })
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "packs.json"
            path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            plan = compile_analysis(
                [frame], source_names=["培训记录"], user_request="分析课程和学员表现", domain_pack_path=path
            )
        self.assertEqual(plan.domain_id, "education")
        self.assertEqual(plan.domain_label, "教育培训")
        self.assertIn("结课成绩", plan.metrics)


if __name__ == "__main__":
    unittest.main()
