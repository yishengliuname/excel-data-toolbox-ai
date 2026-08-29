from __future__ import annotations

import json
import unittest
from unittest.mock import patch

import pandas as pd

import excel_data_toolbox.nl_agent as nl
from excel_data_toolbox.nl_agent import DeepSeekClient, build_table_catalog


class JsonRecoveryTests(unittest.TestCase):
    def test_parser_extracts_json_from_model_prose(self) -> None:
        payload = nl._parse_model_json(
            '以下是结果：```json\n{"intent":"data"}\n```', label="测试"
        )
        self.assertEqual(payload, {"intent": "data"})

    def test_unified_route_retries_invalid_json_once(self) -> None:
        client = DeepSeekClient("sk-test")
        valid = {
            "intent": "data",
            "normalized_request": "汇总数据",
            "data_request": "汇总数据",
            "chart_request": "",
            "engineering_category": None,
            "reason": "数据处理",
        }
        responses = [
            json.dumps({"choices": [{"finish_reason": "stop", "message": {"content": "抱歉，我没有返回 JSON"}}]}).encode(),
            json.dumps({"choices": [{"finish_reason": "stop", "message": {"content": json.dumps(valid, ensure_ascii=False)}}]}, ensure_ascii=False).encode(),
        ]
        catalog = build_table_catalog({"t": pd.DataFrame({"金额": [1]})})
        with patch.object(client, "_request", side_effect=responses) as request:
            routed = client.classify_unified_request("请汇总当前数据", catalog)
        self.assertEqual(routed["intent"], "data")
        self.assertEqual(request.call_count, 2)
        retry_body = request.call_args_list[1].args[0]
        self.assertIn("上一条响应未通过", retry_body["messages"][-1]["content"])


if __name__ == "__main__":
    unittest.main()
