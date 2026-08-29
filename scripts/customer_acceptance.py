"""Run real customer orders sequentially and produce auditable deliverables.

This is intentionally sequential: the next order starts only after the current
workbook has been exported, reopened, compared and inspected for native charts.
No DeepSeek call is required for specialist routes, so the suite is repeatable.
"""
# ruff: noqa: E402

from __future__ import annotations

import hashlib
import json
import sys
import time
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Mapping

import pandas as pd
from openpyxl import load_workbook

PROJECT = Path(__file__).resolve().parents[1]
WORKSPACE = PROJECT.parent
if str(WORKSPACE) not in sys.path:
    sys.path.insert(0, str(WORKSPACE))

from excel_data_toolbox.adaptive_report import build_adaptive_analysis_report
from excel_data_toolbox.core import export_tables, load_tables
from excel_data_toolbox.delivery_qa import verify_delivery
from excel_data_toolbox.enterprise_report import build_enterprise_diagnosis_report
from excel_data_toolbox.hr_report import build_hr_management_report
from excel_data_toolbox.inventory_report import build_inventory_management_report
from excel_data_toolbox.sales_report import (
    build_quarterly_sales_management_report,
    build_sales_management_report,
    infer_sales_report_columns,
)
from excel_data_toolbox.selection_report import build_selection_recommendation_report


@dataclass(frozen=True)
class AcceptanceCase:
    case_id: str
    customer_request: str
    source: Path
    output_name: str
    builder: Callable[[list[pd.DataFrame], list[str]], Any]


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load(path: Path) -> tuple[list[pd.DataFrame], list[str]]:
    tables = load_tables([path])
    return list(tables.values()), [str(name).rsplit("::", 1)[-1] for name in tables]


def _sales(frames: list[pd.DataFrame], names: list[str]) -> Any:
    del names
    candidates: list[tuple[pd.DataFrame, dict[str, str]]] = []
    for frame in frames:
        try:
            candidates.append((frame, infer_sales_report_columns(frame)))
        except ValueError:
            continue
    if len(candidates) != 1:
        raise ValueError(f"销售基础分析应唯一识别一张业务表，实际识别 {len(candidates)} 张")
    return build_sales_management_report(candidates[0][0], **candidates[0][1])


def _quarter(frames: list[pd.DataFrame], names: list[str]) -> Any:
    selected_frames: list[pd.DataFrame] = []
    selected_names: list[str] = []
    for frame, name in zip(frames, names, strict=True):
        try:
            infer_sales_report_columns(frame)
        except ValueError:
            continue
        selected_frames.append(frame)
        selected_names.append(name)
    return build_quarterly_sales_management_report(selected_frames, source_names=selected_names)


def _inventory(frames: list[pd.DataFrame], names: list[str]) -> Any:
    return build_inventory_management_report(frames, source_names=names)


def _hr(frames: list[pd.DataFrame], names: list[str]) -> Any:
    return build_hr_management_report(frames, source_names=names)


def _selection(frames: list[pd.DataFrame], names: list[str]) -> Any:
    return build_selection_recommendation_report(
        frames, source_names=names,
        user_request="按照需求和每一个组的数据，选取最优秀的八个组参加比赛",
        top_n=8, include_charts=True,
    )


def _enterprise(frames: list[pd.DataFrame], names: list[str]) -> Any:
    return build_enterprise_diagnosis_report(
        frames, source_names=names,
        user_request="全面分析客户、销售、成本和库存风险，并给出下一步行动",
    )


def _adaptive(frames: list[pd.DataFrame], names: list[str]) -> Any:
    return build_adaptive_analysis_report(
        frames, source_names=names,
        user_request="不知道问题在哪里，请全面分析、说明风险并给出行动建议",
    )


def _chart_integrity(path: Path) -> dict[str, Any]:
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        charts = {sheet.title: len(sheet._charts) for sheet in workbook.worksheets if sheet._charts}
        merged = sum(len(sheet.merged_cells.ranges) for sheet in workbook.worksheets)
        return {"native_chart_count": sum(charts.values()), "charts_by_sheet": charts, "merged_ranges": merged}
    finally:
        workbook.close()


def run(output_root: Path) -> dict[str, Any]:
    candidate_homes = [Path.home(), Path("D:/Users") / Path.home().name]
    home = next(
        (item for item in candidate_homes if (item / "Excel工具客户测试数据.xlsx").is_file()),
        candidate_homes[0],
    )
    cases = [
        AcceptanceCase("01_sales", "销售经营五表分析", home / "Excel工具客户测试数据.xlsx", "01_销售经营分析.xlsx", _sales),
        AcceptanceCase("02_quarter", "一季度脏销售表清洗合并与老板报表", home / "第二轮_脏数据Excel客户测试.xlsx", "02_季度销售经营分析.xlsx", _quarter),
        AcceptanceCase("03_inventory", "采购销售库存、补货和积压分析", home / "第三轮_库存采购_真实客户测试 (1).xlsx", "03_采购销售库存经营报告.xlsx", _inventory),
        AcceptanceCase("04_hr", "考勤绩效薪资与人员关注分析", home / "第四轮_员工考勤薪资分析客户测试 (1).xlsx", "04_员工经营分析.xlsx", _hr),
        AcceptanceCase("05_selection", "按多轮数据选出最优秀八组", home / "Desktop" / "问题记录_第五轮.xlsx", "05_八组评选报告.xlsx", _selection),
        AcceptanceCase("06_enterprise", "企业客户销售成本库存综合诊断", home / "第六轮_企业集团经营诊断终极压力测试.xlsx", "06_企业集团经营诊断.xlsx", _enterprise),
        AcceptanceCase("07_adaptive", "陌生业务数据通用自适应分析", home / "第五轮_企业经营异常诊断客户测试 (1).xlsx", "07_通用自适应分析.xlsx", _adaptive),
    ]
    output_root.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    suite_started = time.perf_counter()
    for index, case in enumerate(cases, start=1):
        started = time.perf_counter()
        item: dict[str, Any] = {
            "sequence": index, "case_id": case.case_id,
            "customer_request": case.customer_request, "source": str(case.source),
            "status": "failed", "checks": {}, "error": "",
        }
        try:
            if not case.source.is_file():
                raise FileNotFoundError(f"客户测试文件不存在：{case.source}")
            frames, names = _load(case.source)
            result = case.builder(frames, names)
            outputs: Mapping[str, pd.DataFrame] = result.outputs
            destination = output_root / case.output_name
            export_tables(outputs, destination, include_log=False, overwrite=True)
            acceptance = verify_delivery(destination, outputs)
            chart_check = _chart_integrity(destination)
            item["checks"] = {"reopen_and_compare": acceptance.to_dict(), "workbook_integrity": chart_check}
            if acceptance.status != "passed":
                raise RuntimeError("导出文件重新打开后的数据一致性验收未通过")
            item.update({
                "status": "passed", "artifact": str(destination), "artifact_sha256": _sha256(destination),
                "source_table_count": len(frames), "output_sheet_count": len(outputs),
                "business_report": dict(result.report),
            })
        except Exception as exc:
            item["error"] = f"{type(exc).__name__}: {exc}"
            item["traceback"] = traceback.format_exc(limit=5)
        item["duration_ms"] = round((time.perf_counter() - started) * 1000)
        results.append(item)
        # Strictly sequential: never hide an earlier failure by running in parallel.
    report = {
        "suite": "真实客户逐单专业验收",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "execution_mode": "sequential",
        "case_count": len(results),
        "passed": sum(item["status"] == "passed" for item in results),
        "failed": sum(item["status"] != "passed" for item in results),
        "duration_ms": round((time.perf_counter() - suite_started) * 1000),
        "results": results,
    }
    (output_root / "customer_acceptance.json").write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    lines = ["# 真实客户逐单专业验收报告", "", f"- 通过：{report['passed']}/{report['case_count']}", "- 执行：逐项串行", ""]
    for item in results:
        mark = "通过" if item["status"] == "passed" else "失败"
        lines.extend([f"## {item['sequence']}. {item['customer_request']} — {mark}", "", f"- 输入：`{item['source']}`", f"- 耗时：{item['duration_ms']} ms"])
        if item.get("artifact"):
            lines.append(f"- 交付：`{item['artifact']}`")
            lines.append(f"- 工作表：{item['output_sheet_count']}；原生图表：{item['checks']['workbook_integrity']['native_chart_count']}")
        if item.get("error"):
            lines.append(f"- 错误：{item['error']}")
        lines.append("")
    (output_root / "customer_acceptance.md").write_text("\n".join(lines), encoding="utf-8")
    return report


if __name__ == "__main__":
    destination = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else PROJECT / "acceptance_reports"
    result = run(destination)
    print(json.dumps({key: result[key] for key in ("case_count", "passed", "failed", "duration_ms")}, ensure_ascii=False))
    raise SystemExit(0 if result["failed"] == 0 else 1)
