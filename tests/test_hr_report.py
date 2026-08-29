from __future__ import annotations

import unittest
from unittest.mock import patch

import pandas as pd

import excel_data_toolbox.server as server_module
from excel_data_toolbox.hr_report import build_hr_management_report, infer_hr_table_roles
from excel_data_toolbox.server import AppSession, ToolboxHandler


def _hr_tables() -> list[pd.DataFrame]:
    employees = pd.DataFrame({
        "员工编号": ["E001", "E002", "E003", "E004", "E005", "E006"],
        "姓名": ["张伟", "李娜", "王强", "赵敏", "陈浩", "刘洋"],
        "部门": ["研发部", "销售部", "研发部", "市场部", "销售部", "行政部"],
        "岗位": ["工程师", "销售经理", "工程师", "运营专员", "销售专员", "行政"],
        "入职日期": ["2024-03-15", "2023-06-20", "2025-01-10", "2024-08-01", "2025-02-18", "2023-11-12"],
        "基本工资": [12000, 15000, 11000, 9000, 8000, 7000],
        "状态": ["在职", "在职", "在职", "在职", "在职", "离职"],
    })
    attendance = pd.DataFrame({
        "员工编号": ["E001", "E002", "E003", "E004", "E005", "E006"], "月份": ["2026-07"] * 6,
        "出勤天数": [21, 22, 18, 22, 15, 10], "迟到次数": [1, 0, 5, 0, 8, 3],
        "早退次数": [0, 1, 1, 0, 2, 1], "请假天数": [1, 0, 3, 1, 5, 8], "加班小时": [12, 8, 20, 6, 2, 0],
    })
    performance = pd.DataFrame({
        "员工编号": ["E002", "E005", "E001", "E003", "E004"], "销售额": [320000, 85000, 0, 0, 50000],
        "完成目标比例": [1.15, 0.62, 1.0, 0.95, 0.88], "客户评分": [4.8, 3.6, 4.7, 4.5, 4.2],
    })
    adjustments = pd.DataFrame({
        "员工编号": ["E001", "E002", "E003", "E005"], "调整类型": ["奖金", "销售提成", "加班补贴", "扣款"],
        "金额": [3000, 12000, 1500, -500], "备注": ["项目奖励", "季度业绩", "研发任务", "考勤异常"],
    })
    notes = pd.DataFrame({"老板说：": ["重点关注长期迟到、低绩效、高离职风险人员。"]})
    return [employees, attendance, performance, adjustments, notes]


class HRReportTests(unittest.TestCase):
    def test_roles_are_inferred_by_columns(self) -> None:
        roles = infer_hr_table_roles(_hr_tables())
        self.assertEqual(roles["employees"], 0)
        self.assertEqual(roles["attendance"], (1,))
        self.assertEqual(roles["performance"], (2,))
        self.assertEqual(roles["adjustments"], (3,))

    def test_scores_payroll_and_attention(self) -> None:
        result = build_hr_management_report(
            _hr_tables(), source_names=["员工基础信息", "月度考勤记录", "绩效数据", "薪资调整记录", "老板需求说明"]
        )
        detail = result.outputs["员工综合分析"].set_index("员工编号")
        self.assertEqual(detail.loc["E002", "预计薪资"], 27000)
        self.assertEqual(detail.loc["E005", "管理分类"], "重点关注")
        self.assertEqual(detail.loc["E005", "离职风险代理等级"], "高")
        self.assertEqual(detail.loc["E002", "管理分类"], "表现优秀")
        self.assertEqual(result.report["active_employee_count"], 5)
        self.assertEqual(result.report["excellent_count"], 3)
        self.assertEqual(result.report["attention_count"], 2)
        self.assertEqual(result.report["manual_review_count"], 1)
        self.assertEqual(len(result.outputs), 10)

    def test_unified_command_routes_without_deepseek(self) -> None:
        session = AppSession()
        handler = object.__new__(ToolboxHandler)
        try:
            names = ["员工基础信息", "月度考勤记录", "绩效数据", "薪资调整记录", "老板需求说明"]
            ids = [session.add_table(name, frame, source="单元测试", original=True) for name, frame in zip(names, _hr_tables())]
            prompt = "最近人员管理有点混乱，请把员工考勤、绩效、薪资情况整合，看看哪些员工表现好、哪些需要重点关注，生成老板能直接看的 Excel 人事经营分析报表。"
            with (
                patch.object(server_module, "SESSION", session),
                patch.object(server_module, "_project_ai_config", return_value={"configured": False, "api_key": "", "model": "deepseek-v4-flash"}),
                patch.object(server_module.DeepSeekClient, "classify_unified_request", side_effect=AssertionError("员工标准流程不应访问 DeepSeek")),
            ):
                response = handler._ai_unified({"prompt": prompt, "table_ids": ids})
            self.assertEqual(response["mode"], "data")
            self.assertTrue(response["auto_execute"])
            self.assertEqual(response["plan"]["steps"][0]["operation"], "hr_management_report")
        finally:
            session.close()


if __name__ == "__main__":
    unittest.main()
