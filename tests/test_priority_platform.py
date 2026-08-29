from __future__ import annotations

import json
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path

import pandas as pd
from excel_data_toolbox.ai_evaluation import EvaluationScenario, run_evaluation
from excel_data_toolbox.conversation import ConversationStore
from excel_data_toolbox.data_contracts import DataContractStore, infer_data_contract, validate_contract
from excel_data_toolbox.database_connections import ConnectionProfileStore
from excel_data_toolbox.engine_router import group_summary_auto
from excel_data_toolbox.lineage import LineageStore, dataset_metadata
from excel_data_toolbox.task_engine import PersistentJobEngine
from excel_data_toolbox.tool_registry import build_builtin_registry


class _MemorySecrets:
    def __init__(self) -> None:
        self.values: dict[str, str] = {}

    def set(self, key: str, value: str) -> None:
        self.values[key] = value

    def get(self, key: str) -> str | None:
        return self.values.get(key)

    def delete(self, key: str) -> None:
        self.values.pop(key, None)


class PriorityPlatformTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_contract_versions_and_schema_drift(self) -> None:
        frame = pd.DataFrame({"订单编号": ["A1", "A2"], "销售额": [100.0, 200.0], "地区": ["华东", "华南"]})
        store = DataContractStore(self.root / "contracts")
        contract = infer_data_contract(frame, name="销售订单合同")
        store.save(contract)
        self.assertTrue(validate_contract(frame, store.load(contract.contract_id)).passed)
        changed = frame.rename(columns={"销售额": "成交额"})
        result = validate_contract(changed, contract)
        self.assertFalse(result.passed)
        self.assertEqual(result.schema_drift["missing_columns"], ["销售额"])
        self.assertEqual(store.list()[0]["active_version"], 1)

    def test_lineage_is_privacy_preserving_and_exportable(self) -> None:
        store = LineageStore(self.root / "lineage.sqlite3")
        frame = pd.DataFrame({"客户": ["绝密客户甲"], "金额": [1234]})
        store.append_completed(
            task_id="TASK-1", job_name="分组汇总",
            inputs=[dataset_metadata("销售表", frame, source="input:source")],
            outputs=[], parameters={"method": "sum"},
        )
        destination = store.export_evidence("TASK-1", self.root / "evidence.json")
        text = destination.read_text(encoding="utf-8")
        self.assertNotIn("绝密客户甲", text)
        self.assertEqual(json.loads(text)["completed"], 1)

    def test_job_queue_progress_retry_and_completion(self) -> None:
        engine = PersistentJobEngine(self.root / "jobs.sqlite3", workers=1)

        def handler(context, payload):
            context.update(50, "处理中")
            return {"answer": int(payload["value"]) * 2}

        engine.register("double", handler)
        engine.start()
        try:
            record = engine.submit("TASK-1", "double", {"value": 21})
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                record = engine.get(record.job_id)
                if record.status in {"completed", "failed"}:
                    break
                time.sleep(0.02)
            self.assertEqual(record.status, "completed")
            self.assertEqual(record.result["answer"], 42)
            self.assertEqual(engine.counts("TASK-1")["completed"], 1)
        finally:
            engine.stop()

    def test_engine_route_preserves_summary_schema(self) -> None:
        frame = pd.DataFrame({"地区": ["华东", "华东", "华南"], "销售额": [1, 2, 4]})
        result, decision = group_summary_auto(frame, by=["地区"], aggregations={"销售额": "sum"})
        self.assertEqual(list(result.columns), ["地区", "销售额"])
        self.assertEqual(result.set_index("地区").loc["华东", "销售额"], 3)
        self.assertIn(decision.engine, {"pandas", "duckdb", "polars"})

    def test_read_only_database_connection_center(self) -> None:
        database = self.root / "orders.sqlite3"
        connection = sqlite3.connect(database)
        try:
            connection.execute("CREATE TABLE orders (region TEXT, amount REAL)")
            connection.executemany("INSERT INTO orders VALUES (?,?)", [("华东", 10), ("华南", 20)])
            connection.commit()
        finally:
            connection.close()
        store = ConnectionProfileStore(self.root / "profiles.sqlite3", _MemorySecrets())
        profile = store.save(name="订单库", kind="sqlite", secret={"path": str(database)})
        result = store.query(profile.profile_id, "SELECT region, SUM(amount) AS total FROM orders GROUP BY region")
        self.assertEqual(len(result), 2)
        with self.assertRaises(ValueError):
            store.query(profile.profile_id, "DELETE FROM orders")

    def test_conversation_followup_is_scoped_and_secret_redacted(self) -> None:
        store = ConversationStore(self.root / "conversation.json")
        store.append(user_request="按月份画销售额折线图", route="chart", status="completed", output_names=["月度趋势"])
        resolved, followup = store.resolve("把它改成柱状图")
        self.assertTrue(followup)
        self.assertIn("按月份画销售额折线图", resolved)
        store.append(user_request="密钥 sk-1234567890abcdef 不要保存")
        self.assertNotIn("sk-1234567890abcdef", self.root.joinpath("conversation.json").read_text(encoding="utf-8"))

    def test_tool_registry_and_ai_regression_report(self) -> None:
        registry = build_builtin_registry(["clean", "summary", "chart"])
        self.assertEqual(registry.names(), frozenset({"clean", "summary", "chart"}))
        report = run_evaluation(
            [EvaluationScenario("sales_case", "请清洗并汇总销售数据", expected_operations=("clean",))],
            lambda _: {"status": "ready", "steps": [{"operation": "clean"}], "summary": "安全清洗"},
        )
        self.assertEqual(report.passed, 1)
        self.assertEqual(report.failed, 0)


if __name__ == "__main__":
    unittest.main()
