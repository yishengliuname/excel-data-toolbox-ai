"""Local-only web server for 表格快处.

The server intentionally binds to 127.0.0.1 and keeps each customer order in an
isolated, recoverable local task directory.  It uses the Python standard library
for HTTP/UI delivery; pandas and openpyxl power the spreadsheet engine.
"""

from __future__ import annotations

import argparse
import atexit
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from email import policy as email_policy
from email.parser import BytesParser
import io
import json
import math
import mimetypes
import os
from pathlib import Path
import random
import re
import secrets
import shutil
import socket
import sqlite3
import sys
import tempfile
import threading
import time
from typing import Any, Mapping
from urllib.parse import unquote, urlparse
import uuid
import webbrowser
import zipfile

import pandas as pd

if __package__:
    from .analytics import (
        aggregate_trend,
        assess_data_quality,
        category_contribution,
        correlation_matrix,
        cross_pivot,
        descriptive_statistics,
        detect_outliers,
        rfm_segmentation,
    )
    from .core import (
        concat_tables,
        export_tables,
        group_summary,
        join_tables,
        load_tables,
        mask_columns,
        profile_dataframe,
        select_rename_sort,
        smart_clean,
        split_dataframe,
    )
    from .fuzzy import cluster_similar_values, fuzzy_lookup
    from .models import CleaningConfig, OperationLog
    from .nl_agent import (
        ALLOWED_AGENT_OPERATIONS,
        ENGINEERING_CATEGORIES,
        SUPPORTED_DEEPSEEK_MODELS,
        AgentExecutionError,
        AgentPlan,
        DeepSeekAPIError,
        DeepSeekClient,
        PlanValidationError,
        build_table_catalog,
        execute_plan,
        preview_plan,
        validate_plan,
    )
    from .recipes import ProcessingRecipe, run_recipe
    from .reconciliation import reconcile_tables
    from .validation import ValidationRule, validate_dataframe
    from .chart_agent import ChartSpecValidationError, validate_chart_spec
    from .power_bi_automation import (
        PowerBIAutomationError,
        PowerBIConfig,
        build_power_bi_bundle,
        fallback_power_bi_brief,
        publish_bundle_if_configured,
    )
    from .delivery_qa import acceptance_frame, verify_delivery, write_acceptance_json
    from .secure_secrets import SecureSecretStore, SecretStoreError
    from .task_store import TaskRepository
    from .workbook_fidelity import preserve_workbook_export, workbook_feature_inventory
    from .advanced_automation import (
        build_vba_bundle,
        document_capabilities,
        extract_image_text,
        extract_pdf_tables,
        query_sqlite_read_only,
    )
    from .large_data import LargeDataUnavailable, duckdb_available, query_files
    from .order_intake import quote_order
    from .inventory_report import can_build_inventory_report
    from .hr_report import can_build_hr_report
    from .adaptive_report import can_build_adaptive_report
    from .selection_report import can_build_selection_report, explicit_selection_count, parse_selection_count
    from .enterprise_report import can_build_enterprise_diagnosis_report
    from .sales_report import infer_sales_report_columns
    from .scheduler import LocalScheduler
    from .ai_evaluation import AITraceStore, ScenarioStore, default_scenarios, run_evaluation
    from .conversation import ConversationStore
    from .data_contracts import DataContractStore, infer_data_contract, issues_frame, validate_contract
    from .database_connections import ConnectionProfileStore
    from .engine_router import available_engines, choose_engine, group_summary_auto
    from .lineage import LineageStore, dataset_metadata
    from .session_registry import SessionProxy, SessionRegistry
    from .task_engine import PersistentJobEngine
    from .tool_registry import build_builtin_registry
    from .source_guard import (
        assess_prompt_data_alignment,
        detect_generated_workbook,
        source_confirmation_frame,
    )
else:  # Supports: python server.py
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from excel_data_toolbox.analytics import (
        aggregate_trend,
        assess_data_quality,
        category_contribution,
        correlation_matrix,
        cross_pivot,
        descriptive_statistics,
        detect_outliers,
        rfm_segmentation,
    )
    from excel_data_toolbox.core import (
        concat_tables,
        export_tables,
        group_summary,
        join_tables,
        load_tables,
        mask_columns,
        profile_dataframe,
        select_rename_sort,
        smart_clean,
        split_dataframe,
    )
    from excel_data_toolbox.fuzzy import cluster_similar_values, fuzzy_lookup
    from excel_data_toolbox.models import CleaningConfig, OperationLog
    from excel_data_toolbox.nl_agent import (
        ALLOWED_AGENT_OPERATIONS,
        ENGINEERING_CATEGORIES,
        SUPPORTED_DEEPSEEK_MODELS,
        AgentExecutionError,
        AgentPlan,
        DeepSeekAPIError,
        DeepSeekClient,
        PlanValidationError,
        build_table_catalog,
        execute_plan,
        preview_plan,
        validate_plan,
    )
    from excel_data_toolbox.recipes import ProcessingRecipe, run_recipe
    from excel_data_toolbox.reconciliation import reconcile_tables
    from excel_data_toolbox.validation import ValidationRule, validate_dataframe
    from excel_data_toolbox.chart_agent import ChartSpecValidationError, validate_chart_spec
    from excel_data_toolbox.power_bi_automation import (
        PowerBIAutomationError,
        PowerBIConfig,
        build_power_bi_bundle,
        fallback_power_bi_brief,
        publish_bundle_if_configured,
    )
    from excel_data_toolbox.delivery_qa import acceptance_frame, verify_delivery, write_acceptance_json
    from excel_data_toolbox.secure_secrets import SecureSecretStore, SecretStoreError
    from excel_data_toolbox.task_store import TaskRepository
    from excel_data_toolbox.workbook_fidelity import preserve_workbook_export, workbook_feature_inventory
    from excel_data_toolbox.advanced_automation import (
        build_vba_bundle,
        document_capabilities,
        extract_image_text,
        extract_pdf_tables,
        query_sqlite_read_only,
    )
    from excel_data_toolbox.large_data import LargeDataUnavailable, duckdb_available, query_files
    from excel_data_toolbox.order_intake import quote_order
    from excel_data_toolbox.inventory_report import can_build_inventory_report
    from excel_data_toolbox.hr_report import can_build_hr_report
    from excel_data_toolbox.adaptive_report import can_build_adaptive_report
    from excel_data_toolbox.selection_report import (
        can_build_selection_report,
        explicit_selection_count,
        parse_selection_count,
    )
    from excel_data_toolbox.enterprise_report import can_build_enterprise_diagnosis_report
    from excel_data_toolbox.sales_report import infer_sales_report_columns
    from excel_data_toolbox.scheduler import LocalScheduler
    from excel_data_toolbox.ai_evaluation import AITraceStore, ScenarioStore, default_scenarios, run_evaluation
    from excel_data_toolbox.conversation import ConversationStore
    from excel_data_toolbox.data_contracts import (
        DataContractStore,
        infer_data_contract,
        issues_frame,
        validate_contract,
    )
    from excel_data_toolbox.database_connections import ConnectionProfileStore
    from excel_data_toolbox.engine_router import available_engines, choose_engine, group_summary_auto
    from excel_data_toolbox.lineage import LineageStore, dataset_metadata
    from excel_data_toolbox.session_registry import SessionProxy, SessionRegistry
    from excel_data_toolbox.task_engine import PersistentJobEngine
    from excel_data_toolbox.tool_registry import build_builtin_registry
    from excel_data_toolbox.source_guard import (
        assess_prompt_data_alignment,
        detect_generated_workbook,
        source_confirmation_frame,
    )


APP_DIR = Path(__file__).resolve().parent
WEB_DIR = APP_DIR / "web"
APP_VERSION = "9.8.0"
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_REQUEST_BYTES = 180 * 1024 * 1024
MAX_ROWS_PER_TABLE = 300_000
PREVIEW_ROWS = 100
ALLOWED_SUFFIXES = {
    ".xlsx",
    ".xlsm",
    ".csv",
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
    ".bmp",
    ".db",
    ".sqlite",
    ".sqlite3",
    ".parquet",
}
TABULAR_SUFFIXES = {".xlsx", ".xlsm", ".csv"}
SAFE_NAME = re.compile(r"[^\w\-. ()\[\]（）\u4e00-\u9fff]+", re.UNICODE)
FORMULA_PREFIXES = ("=", "+", "-", "@")
_USER_DATA_OVERRIDE = os.environ.get("BIAOGE_USER_DATA", "").strip()
if _USER_DATA_OVERRIDE:
    USER_DATA_DIR = Path(_USER_DATA_OVERRIDE).expanduser().resolve()
elif getattr(sys, "frozen", False):
    USER_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "BiaogeKuaichu"
else:
    USER_DATA_DIR = APP_DIR / "user_data"
RECIPE_STORE_DIR = USER_DATA_DIR / "recipes"
MAX_SAVED_RECIPES = 100
MAX_STORED_RECIPE_BYTES = 1_200_000
RECIPE_ID_PATTERN = re.compile(r"^[a-f0-9]{12}$")
AI_PLAN_TTL_SECONDS = 15 * 60
AI_MAX_PROMPT_CHARS = 8_000
# Real customer workbooks commonly contain more than twelve related sheets.
# The privacy-safe catalogue already enforces a 100-table ceiling and sends
# only schema/profile metadata (never raw cell values), so keep the UI/API
# scope aligned with that validated catalogue limit.
AI_MAX_SELECTED_TABLES = 100
AI_MAX_OUTPUT_TABLES = 80
AI_MAX_OUTPUT_CELLS = 5_000_000
AI_SECRET_PATTERN = re.compile(r"\bsk-[A-Za-z0-9_-]{16,}\b")
PROJECT_ENV_PATH = APP_DIR / ".env"


def _bounded_environment_int(name: str, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(os.environ.get(name, str(default)) or default)
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(value, maximum))


TASK_RETENTION_DAYS = _bounded_environment_int("TASK_RETENTION_DAYS", 30, 1, 3650)
TASK_REPOSITORY = TaskRepository(USER_DATA_DIR / "tasks", retention_days=TASK_RETENTION_DAYS)
SECRET_STORE = SecureSecretStore(USER_DATA_DIR / "secrets.dpapi")
SCHEDULER = LocalScheduler(USER_DATA_DIR / "schedules.sqlite3")
CONTRACT_STORE = DataContractStore(USER_DATA_DIR / "contracts")
LINEAGE_STORE = LineageStore(USER_DATA_DIR / "lineage.sqlite3")
AI_TRACE_STORE = AITraceStore(USER_DATA_DIR / "ai_traces.jsonl")
AI_SCENARIO_STORE = ScenarioStore(USER_DATA_DIR / "ai_scenarios.json")
AI_SCENARIO_STORE.ensure_defaults()
DATABASE_CONNECTIONS = ConnectionProfileStore(USER_DATA_DIR / "database_connections.sqlite3", SECRET_STORE)
JOB_ENGINE = PersistentJobEngine(USER_DATA_DIR / "jobs.sqlite3", workers=2)
TOOL_REGISTRY = build_builtin_registry(ALLOWED_AGENT_OPERATIONS)


def _project_ai_config() -> dict[str, Any]:
    """Read the project-local AI credential without exposing its value."""

    values: dict[str, str] = {}
    if PROJECT_ENV_PATH.exists():
        try:
            for raw_line in PROJECT_ENV_PATH.read_text(encoding="utf-8-sig").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if key in {"DEEPSEEK_API_KEY", "DEEPSEEK_MODEL"}:
                    values[key] = value.strip().strip('"').strip("'")
        except OSError:
            values = {}
    vault_key = ""
    try:
        vault_key = SECRET_STORE.get("DEEPSEEK_API_KEY")
    except SecretStoreError:
        vault_key = ""
    environment_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    api_key = environment_key or vault_key or values.get("DEEPSEEK_API_KEY", "").strip()
    model = (
        os.environ.get("DEEPSEEK_MODEL", "").strip() or values.get("DEEPSEEK_MODEL", "").strip() or "deepseek-v4-flash"
    )
    if model not in SUPPORTED_DEEPSEEK_MODELS:
        model = "deepseek-v4-flash"
    valid = bool(api_key) and len(api_key) <= 512 and not any(character.isspace() for character in api_key)
    source = "environment" if environment_key else ("windows_vault" if vault_key else "project_env")
    return {"configured": valid, "api_key": api_key if valid else "", "model": model, "source": source}


def _task_id() -> str:
    return f"{datetime.now():%Y%m%d}-{uuid.uuid4().hex[:4].upper()}"


def _safe_filename(raw_name: str, *, fallback: str = "文件") -> str:
    name = Path(str(raw_name)).name.strip().replace("\x00", "")
    name = SAFE_NAME.sub("_", name).strip(". ")
    return name[:120] or fallback


def _specific_report_filename(operation: str, file_names: list[str] | tuple[str, ...]) -> str:
    """Build a customer-recognisable filename from the current task input.

    A report must not be called ``通用自适应经营分析报告 (4).xlsx``: that
    hides which customer/project it belongs to and makes accidental
    self-ingestion much more likely.  The name is derived only from the
    explicitly uploaded files for the current task, never from an output
    table name or a directory scan.
    """

    descriptors = {
        "enterprise_diagnosis_report": "经营诊断报告",
        "adaptive_analysis_report": "自适应经营分析报告",
        "selection_recommendation_report": "结构化评选报告",
        "hr_management_report": "人效经营分析报告",
        "inventory_management_report": "库存经营分析报告",
        "quarterly_sales_report": "季度销售经营分析报告",
        "sales_management_report": "销售经营分析报告",
    }
    source_name = next((str(name) for name in file_names if str(name).strip()), "")
    if not source_name:
        # Keep the stable unit-test/API fallback for programmatic callers that
        # do not have an uploaded filename; uploaded customer tasks always use
        # the specific branch below.
        legacy = {
            "enterprise_diagnosis_report": "企业集团经营诊断报告.xlsx",
            "adaptive_analysis_report": "通用自适应经营分析报告.xlsx",
            "selection_recommendation_report": "候选对象结构化评选报告.xlsx",
            "hr_management_report": "员工考勤绩效薪资经营分析报告.xlsx",
            "inventory_management_report": "采购销售库存经营报告.xlsx",
            "quarterly_sales_report": "季度销售经营分析报告.xlsx",
            "sales_management_report": "销售经营分析报告.xlsx",
        }
        return legacy.get(operation, "本次任务_分析报告.xlsx")
    source = Path(source_name).stem
    source = re.sub(r"\s*\(\d+\)\s*$", "", source).strip()
    source = re.sub(
        r"(?:[_\- ]?(?:终极)?(?:压力|客户|真实客户)?测试(?:数据|案例)?|[_\- ]?原始数据)$",
        "",
        source,
        flags=re.IGNORECASE,
    ).strip(" _-")
    descriptor = descriptors.get(operation, "分析报告")
    if source.endswith(descriptor):
        return _safe_filename(f"{source}.xlsx", fallback=f"本次任务_{descriptor}.xlsx")
    return _safe_filename(f"{source}_{descriptor}.xlsx", fallback=f"本次任务_{descriptor}.xlsx")


def _decode_multipart_text(part: Any) -> str:
    """Decode browser FormData text as UTF-8 instead of email's ASCII default.

    Browsers send FormData text in UTF-8 but usually omit a per-part charset.
    ``email.message.get_content()`` then falls back to ASCII and replaces every
    Chinese character with ``�``.  Reading the original bytes preserves task
    names and any other future Unicode text fields.
    """

    payload = part.get_payload(decode=True)
    if isinstance(payload, bytes):
        for encoding in ("utf-8", "gb18030"):
            try:
                return payload.decode(encoding)
            except UnicodeDecodeError:
                continue
        return payload.decode("utf-8", errors="replace")
    return str(payload or "")


def _safe_table_name(raw_name: str, *, fallback: str = "处理结果") -> str:
    name = str(raw_name or "").strip()
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return (name or fallback)[:80]


def _unique_columns(columns: list[Any]) -> tuple[list[str], list[str]]:
    """Return stable string column names plus any import warnings."""

    result: list[str] = []
    used: set[str] = set()
    warnings: list[str] = []
    for position, value in enumerate(columns, start=1):
        base = str(value).strip() if value is not None else ""
        if not base or base.lower().startswith("unnamed:"):
            base = f"未命名列_{position}"
        candidate = base
        counter = 2
        while candidate.casefold() in used:
            candidate = f"{base}_{counter}"
            counter += 1
        if candidate != base:
            warnings.append(f"重复列名“{base}”已重命名为“{candidate}”")
        used.add(candidate.casefold())
        result.append(candidate)
    return result, warnings


def _json_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return None if math.isnan(value) or math.isinf(value) else value
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, dict):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_value(item) for item in value]
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            return _json_value(value.item())
        except (TypeError, ValueError):
            pass
    return str(value)


def _escape_spreadsheet_formulas(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy(deep=True)
    for column in result.columns:
        if pd.api.types.is_object_dtype(result[column].dtype) or pd.api.types.is_string_dtype(result[column].dtype):
            result[column] = result[column].map(
                lambda value: (
                    f"'{value}" if isinstance(value, str) and value.lstrip().startswith(FORMULA_PREFIXES) else value
                )
            )
    return result


def _archive_member_contains(archive: zipfile.ZipFile, name: str, token: bytes) -> bool:
    """Search a ZIP member without loading a potentially huge worksheet XML."""

    tail = b""
    with archive.open(name) as stream:
        while True:
            chunk = stream.read(256 * 1024)
            if not chunk:
                return False
            payload = tail + chunk
            if token in payload:
                return True
            tail = payload[-max(len(token) - 1, 0) :]


def _audit_xlsx_structure(path: Path) -> list[str]:
    """Return privacy-safe workbook capability warnings before data extraction."""

    if path.suffix.lower() != ".xlsx":
        return []
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            sheet_names = sorted(name for name in names if name.startswith("xl/worksheets/") and name.endswith(".xml"))
            has_formulas = any(_archive_member_contains(archive, name, b"<f") for name in sheet_names)
            has_merged = any(_archive_member_contains(archive, name, b"<mergeCell") for name in sheet_names)
            workbook_xml = archive.read("xl/workbook.xml") if "xl/workbook.xml" in names else b""
            flags = {
                "公式": has_formulas,
                "外部链接": any(name.startswith("xl/externalLinks/") for name in names),
                "宏": "xl/vbaProject.bin" in names,
                "图表": any(name.startswith("xl/charts/") for name in names),
                "数据透视表": any(name.startswith(("xl/pivotTables/", "xl/pivotCache/")) for name in names),
                "切片器": any(name.startswith(("xl/slicers/", "xl/slicerCaches/")) for name in names),
                "隐藏工作表": b'state="hidden"' in workbook_xml or b'state="veryHidden"' in workbook_xml,
                "合并单元格": has_merged,
                "批注": any(name.startswith("xl/comments") for name in names),
            }
    except (OSError, KeyError, zipfile.BadZipFile):
        return [f"文件“{path.name}”的工作簿结构无法完整预检，请确认它是有效的 .xlsx 文件"]

    detected = [label for label, enabled in flags.items() if enabled]
    if not detected:
        return []
    warnings = [f"文件“{path.name}”检测到：{'、'.join(detected)}"]
    if has_formulas:
        warnings.append("当前数据提取模式读取公式缓存值；新导出文件不会保留原公式")
    if any(flags[label] for label in ("外部链接", "宏", "图表", "数据透视表", "切片器")):
        warnings.append("该工作簿含高级Excel对象；请保留原件，本工具仅生成新的数据交付文件")
    return warnings


def _recipe_file(recipe_id: str) -> Path:
    value = str(recipe_id or "").strip().lower()
    if not RECIPE_ID_PATTERN.fullmatch(value):
        raise ApiError("处理方案编号无效")
    return RECIPE_STORE_DIR / f"{value}.json"


def _read_stored_recipe(recipe_id: str) -> tuple[ProcessingRecipe, dict[str, Any]]:
    """Load one declarative recipe without ever evaluating stored content."""

    path = _recipe_file(recipe_id)
    try:
        size = path.stat().st_size
    except FileNotFoundError as exc:
        raise ApiError("处理方案不存在，可能已被移除", 404) from exc
    if size > MAX_STORED_RECIPE_BYTES:
        raise ApiError("处理方案文件异常，请重新保存")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ApiError("处理方案文件损坏，请重新保存") from exc
    if not isinstance(payload, dict) or not isinstance(payload.get("recipe"), dict):
        raise ApiError("处理方案文件格式无效")
    recipe = ProcessingRecipe.from_dict(payload["recipe"])
    metadata = {
        "id": str(payload.get("id") or recipe_id),
        "created_at": str(payload.get("created_at") or ""),
        "updated_at": str(payload.get("updated_at") or ""),
    }
    return recipe, metadata


def _list_saved_recipes() -> list[dict[str, Any]]:
    if not RECIPE_STORE_DIR.exists():
        return []
    recipes: list[dict[str, Any]] = []
    for path in sorted(RECIPE_STORE_DIR.glob("*.json")):
        if not RECIPE_ID_PATTERN.fullmatch(path.stem):
            continue
        try:
            recipe, metadata = _read_stored_recipe(path.stem)
        except (ApiError, TypeError, ValueError):
            continue
        recipes.append(
            {
                **metadata,
                "name": recipe.name,
                "description": recipe.description,
                "step_count": len(recipe.steps),
                "steps": [step.to_dict() for step in recipe.steps],
                "recipe": recipe.to_dict(),
            }
        )
    recipes.sort(key=lambda item: (item["updated_at"], item["id"]), reverse=True)
    return recipes


def _save_recipe(recipe: ProcessingRecipe) -> dict[str, Any]:
    """Persist only the recipe definition; task data and cell values are excluded."""

    existing = _list_saved_recipes()
    if len(existing) >= MAX_SAVED_RECIPES:
        raise ApiError(f"本机最多保存 {MAX_SAVED_RECIPES} 个处理方案")
    recipe_id = uuid.uuid4().hex[:12]
    timestamp = datetime.now().isoformat(timespec="seconds")
    payload = {
        "id": recipe_id,
        "created_at": timestamp,
        "updated_at": timestamp,
        "recipe": recipe.to_dict(),
    }
    raw = json.dumps(
        payload,
        ensure_ascii=False,
        allow_nan=False,
        indent=2,
        sort_keys=True,
    ).encode("utf-8")
    if len(raw) > MAX_STORED_RECIPE_BYTES:
        raise ApiError("处理方案内容过大，无法保存")
    RECIPE_STORE_DIR.mkdir(parents=True, exist_ok=True)
    destination = _recipe_file(recipe_id)
    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_bytes(raw)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
    return {
        "id": recipe_id,
        "created_at": timestamp,
        "updated_at": timestamp,
        "name": recipe.name,
        "description": recipe.description,
        "step_count": len(recipe.steps),
        "steps": [step.to_dict() for step in recipe.steps],
        "recipe": recipe.to_dict(),
    }


def _scheduled_recipe_job(payload: Mapping[str, Any]) -> None:
    """Run one previously approved declarative recipe in an isolated task."""

    task_id = str(payload.get("task_id") or "")
    table_id = str(payload.get("table_id") or "")
    recipe_id = str(payload.get("recipe_id") or "")
    output_name = _safe_table_name(str(payload.get("output_name") or "定时任务结果"))
    restored = TASK_REPOSITORY.load(task_id)
    tables = dict(restored["loaded_tables"])
    if table_id not in tables:
        raise ValueError("定时任务输入表不存在")
    recipe, _ = _read_stored_recipe(recipe_id)
    source_name, source_frame, _, _ = tables[table_id]
    result, report = run_recipe(source_frame, recipe, dry_run=False)
    new_id = uuid.uuid4().hex[:12]
    tables[new_id] = (output_name, result, "定时复用处理方案", False)
    operations = list(restored.get("operations") or [])
    operations.append(
        {
            "name": "定时运行处理方案",
            "detail": f"方案 {recipe.name}；{len(source_frame):,} 行 → {len(result):,} 行",
            "inputs": [source_name],
            "outputs": [output_name],
            "before_rows": len(source_frame),
            "after_rows": len(result),
            "time": datetime.now().strftime("%H:%M:%S"),
        }
    )
    TASK_REPOSITORY.save(
        task_id,
        task_name=str(restored.get("task_name") or "定时任务"),
        tables=tables,
        active_table=new_id,
        operations=operations,
        file_names=list(restored.get("file_names") or []),
        import_warnings=list(restored.get("import_warnings") or []),
    )


SCHEDULER.register_job("recipe", _scheduled_recipe_job)


@dataclass
class TableEntry:
    id: str
    name: str
    frame: pd.DataFrame
    source: str
    original: bool = False


_LONG_TEXT_COLUMN_HINT = re.compile(
    r"问题|说明|描述|备注|意见|评价|原因|内容|详情|comment|description|remark|note",
    re.IGNORECASE,
)


def _long_text_detail_frame(entries: list[TableEntry]) -> pd.DataFrame:
    """Expand dense narrative cells into one auditable item per row.

    The source sheets remain unchanged. This companion view is only added to
    professional exports so long issue lists are readable without horizontally
    stretching the original table.
    """

    detail_rows: list[dict[str, Any]] = []
    for entry in entries:
        frame = entry.frame
        if frame.empty:
            continue
        identifier = next(
            (
                column
                for column in frame.columns
                if re.search(r"序号|编号|订单号|ID$|代码|姓名|名称", str(column), re.IGNORECASE)
            ),
            frame.columns[0] if len(frame.columns) else None,
        )
        for column in frame.columns:
            source = frame[column]
            if not (pd.api.types.is_object_dtype(source.dtype) or pd.api.types.is_string_dtype(source.dtype)):
                continue
            samples = [str(value).strip() for value in source.dropna().head(250) if str(value).strip()]
            if not samples:
                continue
            if not _LONG_TEXT_COLUMN_HINT.search(str(column)) and max(map(len, samples)) < 80:
                continue
            for position, (_, row) in enumerate(frame.iterrows(), start=2):
                raw_value = row[column]
                if pd.isna(raw_value) or not str(raw_value).strip():
                    continue
                text = str(raw_value).strip()
                parts = [part.strip() for part in re.split(r"(?<=[；;。！？!?])\s*|[\r\n]+", text) if part.strip()] or [
                    text
                ]
                record_id = row[identifier] if identifier is not None else position - 1
                for item_number, part in enumerate(parts, start=1):
                    detail_rows.append(
                        {
                            "来源数据表": entry.name,
                            "原始行号": position,
                            "记录标识": record_id,
                            "字段": str(column),
                            "条目序号": item_number,
                            "内容": part,
                        }
                    )
    return pd.DataFrame(
        detail_rows,
        columns=["来源数据表", "原始行号", "记录标识", "字段", "条目序号", "内容"],
    )


def _is_blank_cell(value: Any) -> bool:
    """Treat nulls and whitespace-only strings as auditable empty cells."""

    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _blank_cell_detail_frame(entries: list[TableEntry]) -> pd.DataFrame:
    """List source blanks explicitly so they cannot be mistaken for export loss."""

    blank_rows: list[dict[str, Any]] = []
    for entry in entries:
        frame = entry.frame
        if frame.empty:
            continue
        identifier = next(
            (
                column
                for column in frame.columns
                if re.search(r"序号|编号|订单号|ID$|代码|姓名|名称", str(column), re.IGNORECASE)
            ),
            frame.columns[0] if len(frame.columns) else None,
        )
        for position, (_, row) in enumerate(frame.iterrows(), start=2):
            record_id = row[identifier] if identifier is not None else position - 1
            for column in frame.columns:
                if _is_blank_cell(row[column]):
                    blank_rows.append(
                        {
                            "来源数据表": entry.name,
                            "原始行号": position,
                            "记录标识": record_id,
                            "空值字段": str(column),
                            "说明": "源数据为空，导出时保持为空",
                        }
                    )
    return pd.DataFrame(
        blank_rows,
        columns=["来源数据表", "原始行号", "记录标识", "空值字段", "说明"],
    )


@dataclass(frozen=True)
class AiPlanTicket:
    """Short-lived, one-time permission to execute one validated AI plan."""

    task_id: str
    table_ids: tuple[str, ...]
    table_signatures: tuple[tuple[Any, ...], ...]
    plan: AgentPlan
    model: str
    created_at: float
    expires_at: float


class AppSession:
    """One local task session. All mutations are protected by a re-entrant lock."""

    def __init__(self, restore_task_id: str | None = None) -> None:
        self.lock = threading.RLock()
        self.temp_root: Path
        self.task_dir: Path
        self.upload_dir: Path
        self.output_dir: Path
        self.downloads: dict[str, Path] = {}
        if restore_task_id:
            self.restore(restore_task_id)
        else:
            self.reset(initial=True)

    def _make_temp(self) -> None:
        # Tasks are intentionally durable: one task id maps to one folder.  The
        # retention policy removes expired customer data instead of deleting a
        # live task merely because the server restarted.
        self.temp_root = TASK_REPOSITORY.root
        self.task_dir = TASK_REPOSITORY.create(self.task_id, self.task_name)
        self.upload_dir = self.task_dir / "source_files"
        self.output_dir = self.task_dir / "deliverables"
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def reset(self, *, initial: bool = False) -> None:
        with getattr(self, "lock", threading.RLock()):
            del initial
            self.task_id = _task_id()
            self.task_name = "新建数据处理任务"
            self._make_temp()
            self.tables: dict[str, TableEntry] = {}
            self.active_table: str | None = None
            self.operations: list[dict[str, Any]] = []
            self.history: list[dict[str, Any]] = []
            self.redo_stack: list[dict[str, Any]] = []
            self.file_names: set[str] = set()
            self.import_warnings: list[str] = []
            self.review_items: dict[str, dict[str, Any]] = {}
            self.ai_plans: dict[str, AiPlanTicket] = {}
            self.chart_history: list[dict[str, Any]] = []
            self.chart_redo_stack: list[dict[str, Any]] = []
            self.downloads = {}
            self.persist()

    def close(self) -> None:
        try:
            self.persist()
        except (OSError, ValueError, TypeError):
            pass

    def persist(self) -> Path:
        """Persist the current task without serialising executable objects."""

        serialisable_tables = {
            table_id: (entry.name, entry.frame, entry.source, entry.original) for table_id, entry in self.tables.items()
        }
        return TASK_REPOSITORY.save(
            self.task_id,
            task_name=self.task_name,
            tables=serialisable_tables,
            active_table=self.active_table,
            operations=self.operations,
            file_names=sorted(self.file_names),
            import_warnings=self.import_warnings,
        )

    def restore(self, task_id: str) -> None:
        restored = TASK_REPOSITORY.load(task_id)
        self.task_id = str(restored["task_id"])
        self.task_name = str(restored.get("task_name") or "恢复任务")
        self.temp_root = TASK_REPOSITORY.root
        self.task_dir = TASK_REPOSITORY.task_dir(self.task_id)
        self.upload_dir = self.task_dir / "source_files"
        self.output_dir = self.task_dir / "deliverables"
        self.tables = {
            table_id: TableEntry(table_id, name, frame, source, original)
            for table_id, (name, frame, source, original) in restored["loaded_tables"].items()
        }
        active = restored.get("active_table")
        self.active_table = active if active in self.tables else next(reversed(self.tables), None)
        self.operations = list(restored.get("operations") or [])
        self.history = []
        self.redo_stack = []
        self.file_names = set(restored.get("file_names") or [])
        self.import_warnings = list(restored.get("import_warnings") or [])
        self.review_items = {}
        self.ai_plans = {}
        self.chart_history = []
        self.chart_redo_stack = []
        self.downloads = {}

    def add_table(self, name: str, frame: pd.DataFrame, *, source: str, original: bool = False) -> str:
        base = _safe_table_name(name)
        existing = {entry.name.casefold() for entry in self.tables.values()}
        candidate = base
        counter = 2
        while candidate.casefold() in existing:
            candidate = f"{base}_{counter}"
            counter += 1
        table_id = uuid.uuid4().hex[:12]
        clean_columns, warnings = _unique_columns(list(frame.columns))
        stored = frame.copy(deep=True)
        stored.columns = clean_columns
        self.import_warnings.extend(warnings)
        self.tables[table_id] = TableEntry(table_id, candidate, stored, source, original)
        self.active_table = table_id
        return table_id

    def get(self, table_id: str) -> TableEntry:
        try:
            return self.tables[str(table_id)]
        except KeyError as exc:
            raise KeyError("所选数据表不存在，可能已被清空") from exc

    def record(
        self,
        name: str,
        detail: str,
        *,
        inputs: list[str],
        produced: list[str],
        before_rows: int | None = None,
        after_rows: int | None = None,
    ) -> None:
        operation = {
            "name": name,
            "detail": detail,
            "inputs": inputs,
            "outputs": [self.tables[item].name for item in produced if item in self.tables],
            "before_rows": before_rows,
            "after_rows": after_rows,
            "time": datetime.now().strftime("%H:%M:%S"),
        }
        self.operations.append(operation)
        self.history.append({"produced": list(produced), "operation": operation})
        self.redo_stack.clear()
        self.persist()
        try:
            input_metadata = [
                dataset_metadata(self.tables[table_id].name, self.tables[table_id].frame, source=f"input:{table_id}")
                for table_id in inputs
                if table_id in self.tables
            ]
            output_metadata = [
                dataset_metadata(self.tables[table_id].name, self.tables[table_id].frame, source=f"output:{table_id}")
                for table_id in produced
                if table_id in self.tables
            ]
            LINEAGE_STORE.append_completed(
                task_id=self.task_id,
                job_name=name,
                inputs=input_metadata,
                outputs=output_metadata,
                parameters={"detail": str(detail)[:800], "before_rows": before_rows, "after_rows": after_rows},
            )
        except Exception:
            pass

    def undo(self) -> None:
        if not self.history:
            raise ValueError("当前没有可撤销的处理操作")
        item = self.history.pop()
        produced = item["produced"]
        removed_tables: dict[str, TableEntry] = {}
        for table_id in produced:
            removed = self.tables.pop(table_id, None)
            if removed is not None:
                removed_tables[table_id] = removed
        if self.operations and self.operations[-1] is item["operation"]:
            self.operations.pop()
        self.redo_stack.append({**item, "removed_tables": removed_tables})
        if self.active_table not in self.tables:
            self.active_table = next(reversed(self.tables), None)
        self.persist()

    def redo(self) -> None:
        if not self.redo_stack:
            raise ValueError("当前没有可重做的处理操作")
        item = self.redo_stack.pop()
        removed_tables = item.get("removed_tables", {})
        for table_id in item["produced"]:
            entry = removed_tables.get(table_id)
            if entry is not None:
                self.tables[table_id] = entry
        self.operations.append(item["operation"])
        self.history.append({"produced": list(item["produced"]), "operation": item["operation"]})
        if item["produced"]:
            self.active_table = item["produced"][-1]
        self.persist()

    def record_chart(self, spec: dict[str, Any]) -> None:
        self.chart_history.append(json.loads(json.dumps(spec, ensure_ascii=False)))
        self.chart_history = self.chart_history[-100:]
        self.chart_redo_stack.clear()

    def operation_log(self) -> OperationLog:
        log = OperationLog()
        for operation in self.operations:
            log.record(
                operation["name"],
                input_tables=operation.get("inputs", ()),
                output_tables=operation.get("outputs", ()),
                details={
                    "说明": operation.get("detail", ""),
                    "处理前行数": operation.get("before_rows"),
                    "处理后行数": operation.get("after_rows"),
                },
            )
        return log

    def register_download(self, path: Path) -> str:
        token = uuid.uuid4().hex
        self.downloads[token] = path.resolve()
        return f"/download/{token}"

    def issue_ai_plan(
        self,
        *,
        table_ids: list[str],
        table_signatures: tuple[tuple[Any, ...], ...],
        plan: AgentPlan,
        model: str,
    ) -> str:
        """Store a validated plan without storing its prompt, API key, or cell values."""

        now = time.monotonic()
        self.ai_plans = {token: ticket for token, ticket in self.ai_plans.items() if ticket.expires_at > now}
        token = secrets.token_urlsafe(32)
        self.ai_plans[token] = AiPlanTicket(
            task_id=self.task_id,
            table_ids=tuple(table_ids),
            table_signatures=table_signatures,
            plan=plan,
            model=model,
            created_at=now,
            expires_at=now + AI_PLAN_TTL_SECONDS,
        )
        return token

    def consume_ai_plan(self, token: str) -> AiPlanTicket:
        """Consume a plan exactly once, whether execution succeeds or fails."""

        ticket = self.ai_plans.pop(token, None)
        if ticket is None:
            raise ApiError("AI 计划凭证不存在、已使用或已被任务清空", 410)
        if ticket.expires_at <= time.monotonic():
            raise ApiError("AI 计划已过期，请重新生成并确认", 410)
        return ticket

    def add_review_items(
        self,
        category: str,
        source: str,
        items: list[dict[str, Any]],
        *,
        table_id: str | None = None,
        limit: int = 500,
    ) -> list[str]:
        """Add local-only human-review tasks and return their identifiers."""

        created: list[str] = []
        available = max(0, 2000 - len(self.review_items))
        for raw in items[: min(limit, available)]:
            review_id = uuid.uuid4().hex[:12]
            entry = {
                "id": review_id,
                "category": str(category)[:40],
                "source": str(source)[:80],
                "title": str(raw.get("title") or "待人工确认")[:120],
                "detail": str(raw.get("detail") or "")[:500],
                "reason": str(raw.get("reason") or raw.get("detail") or "")[:500],
                "record_key": str(raw.get("record_key") or "")[:120],
                "original": _json_value(raw.get("original")),
                "candidate": _json_value(raw.get("candidate")),
                "score": _json_value(raw.get("score")),
                "evidence": _json_value(raw.get("evidence") or {}),
                "table_id": table_id,
                "status": "pending",
                "decision_note": "",
                "created_at": datetime.now().isoformat(timespec="seconds"),
            }
            self.review_items[review_id] = entry
            created.append(review_id)
        return created

    def review_payload(self) -> dict[str, Any]:
        items = list(self.review_items.values())
        items.sort(key=lambda item: item["created_at"], reverse=True)
        counts = {"pending": 0, "accepted": 0, "rejected": 0, "total": len(items)}
        for item in items:
            status = item.get("status", "pending")
            if status in counts:
                counts[status] += 1
        return {"items": items, "counts": counts}

    def decide_reviews(self, ids: list[str], decision: str, note: str = "") -> int:
        if decision not in {"accepted", "rejected", "pending"}:
            raise ValueError("核验决定必须是接受、拒绝或恢复待确认")
        changed = 0
        for review_id in ids:
            item = self.review_items.get(str(review_id))
            if not item:
                continue
            item["status"] = decision
            item["decision_note"] = str(note)[:300]
            item["decided_at"] = datetime.now().isoformat(timespec="seconds")
            changed += 1
        return changed

    def state_payload(self) -> dict[str, Any]:
        entries = []
        for entry in self.tables.values():
            entries.append(
                {
                    "id": entry.id,
                    "name": entry.name,
                    "rows": len(entry.frame),
                    "columns": [str(column) for column in entry.frame.columns],
                    "source": entry.source,
                    "original": entry.original,
                }
            )
        profile_payload: dict[str, Any] = {}
        preview_payload: dict[str, Any] = {"columns": [], "rows": []}
        warnings = list(dict.fromkeys(self.import_warnings[-8:]))
        if self.active_table and self.active_table in self.tables:
            frame = self.tables[self.active_table].frame
            profile = profile_dataframe(frame).to_dict()
            profile_payload = {
                "rows": profile["row_count"],
                "columns": profile["column_count"],
                "missing_cells": profile["missing_cell_count"],
                "duplicate_rows": profile["duplicate_row_count"],
                "column_profiles": [
                    {
                        "name": column["name"],
                        "dtype": column["semantic_type"],
                        "missing": column["missing_count"],
                    }
                    for column in profile["columns"]
                ],
            }
            if profile["missing_cell_count"]:
                warnings.append(f"当前表有 {profile['missing_cell_count']:,} 个空值单元格")
            if profile["duplicate_row_count"]:
                warnings.append(f"当前表有 {profile['duplicate_row_count']:,} 行完全重复")
            columns = [str(column) for column in frame.columns]
            rows: list[dict[str, Any]] = []
            for values in frame.head(PREVIEW_ROWS).itertuples(index=False, name=None):
                rows.append({column: _json_value(value) for column, value in zip(columns, values)})
            preview_payload = {"columns": columns, "rows": rows}
        return {
            "task_id": self.task_id,
            "task_name": self.task_name,
            "file_count": len(self.file_names),
            "tables": entries,
            "active_table": self.active_table,
            "profile": profile_payload,
            "preview": preview_payload,
            "warnings": list(dict.fromkeys(warnings)),
            "operations": self.operations,
            "can_undo": bool(self.history),
            "can_redo": bool(self.redo_stack),
            "review_counts": self.review_payload()["counts"],
            "job_counts": JOB_ENGINE.counts(self.task_id),
            "conversation": ConversationStore(self.task_dir / "conversation.json").context(),
        }


_INITIAL_SESSION = AppSession()
SESSION_REGISTRY = SessionRegistry(lambda task_id: AppSession(task_id), _INITIAL_SESSION)
SESSION = SessionProxy(SESSION_REGISTRY)
atexit.register(SESSION.close)


class ApiError(Exception):
    def __init__(self, message: str, status: int = 400) -> None:
        super().__init__(message)
        self.status = status


def _normalise_ai_table_scope(raw_value: Any, *, allow_empty: bool) -> list[str]:
    """Validate the browser's selected table-id list with actionable errors."""

    if raw_value is None:
        values: list[Any] = []
    elif isinstance(raw_value, list):
        values = raw_value
    else:
        raise ApiError("数据表范围必须是表 ID 列表；请刷新页面后重新选择工作表")
    if len(values) > AI_MAX_SELECTED_TABLES:
        raise ApiError(
            f"本次选择了 {len(values)} 张表，一次最多支持 {AI_MAX_SELECTED_TABLES} 张；请分批处理"
        )
    if not all(isinstance(item, str) and item.strip() for item in values):
        raise ApiError("数据表范围包含无效表 ID；请刷新页面后重新选择工作表")
    table_ids = list(dict.fromkeys(item.strip() for item in values))
    if not allow_empty and not table_ids:
        raise ApiError("请至少选择一张允许 AI 使用的数据表")
    return table_ids


_NUMERIC_SEMANTIC_TYPES = {"integer", "number", "numeric_text"}
_DATE_SEMANTIC_TYPES = {"datetime", "datetime_text"}
_CHART_TYPES = {
    "bar",
    "horizontal_bar",
    "grouped_bar",
    "stacked_bar",
    "line",
    "area",
    "pie",
    "radar",
    "funnel",
    "waterfall",
    "treemap",
    "histogram",
    "scatter",
    "box",
    "heatmap",
    "gantt",
}
_CHART_AGGREGATIONS = {"sum", "count", "mean", "nunique", "max", "min"}
_DATE_GRAINS = {"auto", "day", "week", "month", "quarter", "year"}
_DATE_GRAIN_LABELS = {"day": "日", "week": "周", "month": "月", "quarter": "季度", "year": "年"}
_ANOMALY_METHODS = {"iqr", "zscore"}
_PIVOT_AGGREGATIONS = {"sum", "count", "mean", "nunique"}
_AGGREGATION_LABELS = {
    "sum": "求和",
    "count": "记录数",
    "mean": "平均值",
    "nunique": "去重计数",
    "max": "最大值",
    "min": "最小值",
}
_DATE_NAME_HINT = re.compile(r"日期|时间|年月|月份|季度|date|time", re.IGNORECASE)
_MEASURE_NAME_HINT = re.compile(
    r"订单金额|销售额|收入|金额|实付|利润|成本|价格|数量|件数|amount|sales|revenue|profit|cost|price|qty|quantity",
    re.IGNORECASE,
)
_DIMENSION_NAME_HINT = re.compile(
    r"地区|区域|渠道|类别|分类|产品|商品|部门|门店|省|市|region|channel|category|product|department|store",
    re.IGNORECASE,
)
_IDENTIFIER_NAME_HINT = re.compile(
    r"(^|[_\s])(id|uuid|code)([_\s]|$)|编号|编码|代码|单号|序号|账号|号码|身份证|银行卡|电话|手机|邮箱",
    re.IGNORECASE,
)
_SENSITIVE_NAME_HINT = re.compile(
    r"姓名|客户|会员|用户|买家|顾客|联系人|电话|手机|邮箱|身份证|银行卡|地址|name|customer|member|user|buyer|phone|email|address",
    re.IGNORECASE,
)


def _require_choice(raw: Any, allowed: set[str], *, label: str) -> str:
    value = str(raw or "").strip().casefold()
    if value not in allowed:
        raise ApiError(f"{label}参数无效")
    return value


def _require_column(frame: pd.DataFrame, raw: Any, *, label: str) -> str:
    column = str(raw or "").strip()
    if not column:
        raise ApiError(f"请选择{label}")
    if column not in frame.columns:
        raise ApiError(f"{label}“{column}”不存在，请重新选择")
    return column


def _optional_column(frame: pd.DataFrame, raw: Any, *, label: str) -> str | None:
    column = str(raw or "").strip()
    if not column:
        return None
    if column not in frame.columns:
        raise ApiError(f"{label}“{column}”不存在，请重新选择")
    return column


def _validated_output_name(raw: Any, *, fallback: str) -> str:
    value = str(raw or "").strip()
    if len(value) > 80:
        raise ApiError("结果名称不能超过 80 个字符")
    return _safe_table_name(value, fallback=fallback)


def _bounded_integer(raw: Any, *, label: str, minimum: int, maximum: int) -> int:
    if isinstance(raw, bool):
        raise ApiError(f"{label}必须是整数")
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise ApiError(f"{label}必须是整数") from exc
    if isinstance(raw, float) and not raw.is_integer():
        raise ApiError(f"{label}必须是整数")
    if not minimum <= value <= maximum:
        raise ApiError(f"{label}必须在 {minimum} 到 {maximum} 之间")
    return value


def _bounded_probability(raw: Any, *, label: str, minimum: float = 0.5, maximum: float = 1.0) -> float:
    if isinstance(raw, bool):
        raise ApiError(f"{label}必须是数字")
    try:
        value = float(raw)
    except (TypeError, ValueError) as exc:
        raise ApiError(f"{label}必须是数字") from exc
    if not math.isfinite(value) or not minimum <= value <= maximum:
        raise ApiError(f"{label}必须在 {minimum:.0%} 到 {maximum:.0%} 之间")
    return value


def _validate_payload_keys(payload: dict[str, Any], allowed: set[str]) -> None:
    unknown = sorted(str(key) for key in payload if key not in allowed)
    if unknown:
        raise ApiError(f"请求包含不支持的字段：{'、'.join(unknown)}")


def _ai_table_signature(entry: TableEntry) -> tuple[Any, ...]:
    """Return a privacy-safe identity for immutable session tables."""

    return (
        entry.id,
        len(entry.frame),
        entry.frame.shape[1],
        tuple((str(column), str(dtype)) for column, dtype in zip(entry.frame.columns, entry.frame.dtypes)),
    )


_AI_OPERATION_LABELS: dict[str, str] = {
    "clean": "智能清洗",
    "select_rename_sort": "字段整理与排序",
    "concat": "多表追加",
    "join": "关联连接",
    "lookup": "精确查找匹配",
    "summary": "分组汇总",
    "split": "批量拆分",
    "mask": "敏感信息脱敏",
    "validate": "质量规则验收",
    "reconcile": "金额/日期容差对账",
    "fuzzy_cluster": "相似名称聚类",
    "fuzzy_lookup": "两表模糊匹配",
    "quality": "数据质量体检",
    "describe": "描述统计",
    "correlation": "相关性分析",
    "outliers": "异常值检测",
    "trend": "时间趋势分析",
    "contribution": "分类贡献分析",
    "pivot": "交叉透视",
    "compare": "新旧数据比对",
    "rfm": "RFM 客户分群",
    "recipe": "安全处理配方",
    "finance": "专业财务分析",
    "sales_management_report": "销售经营管理报告",
    "quarterly_sales_report": "季度多表销售经营报告",
    "inventory_management_report": "采购销售库存经营报告",
    "hr_management_report": "员工考勤绩效薪资经营报告",
    "adaptive_analysis_report": "通用自适应经营分析报告",
    "selection_recommendation_report": "候选对象结构化评选报告",
    "enterprise_diagnosis_report": "企业集团经营诊断报告",
}


def _is_sales_management_report_request(prompt: str) -> bool:
    """Recognise a complete sales-management deliverable, not a single chart."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    signals = (
        "总销售额",
        "总成本",
        "总利润",
        "利润率",
        "产品分析",
        "产品销售",
        "销售人员",
        "人员分析",
        "月度销售",
        "地区销售",
        "客户满意度",
        "异常数据",
        "重点关注",
        "管理层",
        "图表展示",
    )
    score = sum(signal in folded for signal in signals)
    asks_for_workbook = any(token in folded for token in ("excel", "工作表", "sheet", "新的表", "输出一个"))
    return score >= 5 and asks_for_workbook


def _is_quarterly_sales_report_request(prompt: str) -> bool:
    """Recognise multi-sheet dirty-sales consolidation and management reporting."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    sales_context = any(token in folded for token in ("销售表", "销售数据", "订单", "成交明细"))
    multi_period = any(token in folded for token in ("一季度", "季度", "1月", "2月", "3月", "多个销售表", "三个销售表"))
    cleaning_score = sum(
        token in folded for token in ("清洗", "合并", "重复", "去重", "无效订单", "格式统一", "格式不统一")
    )
    analysis_score = sum(
        token in folded for token in ("产品", "地区", "销售人员", "月度趋势", "重点关注", "经营报表", "老板")
    )
    asks_for_workbook = any(token in folded for token in ("excel", "工作表", "报表", "输出"))
    return sales_context and multi_period and cleaning_score >= 2 and analysis_score >= 3 and asks_for_workbook


def _is_inventory_management_report_request(prompt: str) -> bool:
    """Recognise a complete procurement-sales-inventory management order."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    inventory_context = any(token in folded for token in ("库存", "期初库存", "可销售库存", "库存天数"))
    movement_score = sum(token in folded for token in ("采购", "入库", "销售", "出库", "库存调整"))
    decision_score = sum(
        token in folded for token in ("补货", "缺货", "积压", "安全库存", "目标库存", "老板", "经营报表")
    )
    formula_signal = any(
        token in folded for token in ("期初+入库-出库+调整", "当前库存", "采购、销售和库存", "采购销售库存")
    )
    asks_for_workbook = any(token in folded for token in ("excel", "报表", "输出", "老板能直接看"))
    return inventory_context and movement_score >= 2 and decision_score >= 2 and formula_signal and asks_for_workbook


def _is_hr_management_report_request(prompt: str) -> bool:
    """Recognise an attendance-performance-payroll management deliverable."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    people_context = any(token in folded for token in ("员工", "人员", "人事", "人力"))
    data_score = sum(token in folded for token in ("考勤", "出勤", "迟到", "绩效", "业绩", "薪资", "工资", "奖金"))
    decision_score = sum(
        token in folded for token in ("表现好", "需要关注", "重点关注", "离职风险", "老板", "经营分析", "管理报表")
    )
    asks_for_workbook = any(token in folded for token in ("excel", "报表", "输出", "生成一份", "老板能直接看"))
    return people_context and data_score >= 3 and decision_score >= 2 and asks_for_workbook


def _is_adaptive_analysis_report_request(prompt: str) -> bool:
    """Recognise a broad analysis deliverable after specialist routes fail."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    analysis_score = sum(
        token in folded
        for token in (
            "分析",
            "趋势",
            "排名",
            "异常",
            "问题",
            "数据质量",
            "整体情况",
            "表现",
            "重点关注",
            "指标",
            "洞察",
            "清洗",
            "合并",
        )
    )
    deliverable_score = sum(
        token in folded
        for token in (
            "excel",
            "报表",
            "看板",
            "老板",
            "管理层",
            "输出",
            "生成一份",
            "直接看",
        )
    )
    strong_phrase = any(
        token in folded
        for token in (
            "帮我看看这个表",
            "全面分析",
            "综合分析",
            "经营分析",
            "老板能直接看",
        )
    )
    chart_only = any(token in folded for token in ("画一个", "画张", "改图表", "修改图表")) and analysis_score < 2
    return not chart_only and ((analysis_score >= 2 and deliverable_score >= 1) or strong_phrase)


def _is_selection_recommendation_request(prompt: str) -> bool:
    """Recognise requests to choose N candidates from scored/reviewed rows."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    action = any(
        token in folded
        for token in (
            "选择",
            "选出",
            "选取",
            "挑出",
            "挑选",
            "筛选",
            "遴选",
            "推荐",
            "入选",
            "晋级",
            "参加比赛",
            "参赛",
        )
    )
    candidate_context = any(
        token in folded
        for token in (
            "序号",
            "编号",
            "候选",
            "选手",
            "人员",
            "作品",
            "项目",
            "团队",
            "队伍",
            "名额",
            "比赛",
            "竞赛",
        )
    )
    count_signal = explicit_selection_count(folded) is not None
    return action and candidate_context and count_signal


def _is_enterprise_diagnosis_request(prompt: str) -> bool:
    """Recognise cross-domain operating diagnosis even without the word report."""

    folded = re.sub(r"\s+", "", str(prompt or "")).casefold()
    company_context = any(token in folded for token in ("公司", "集团", "企业", "经营", "门店", "餐饮"))
    domain_score = sum(
        token in folded
        for token in (
            "增长",
            "利润",
            "客户",
            "销售",
            "成本",
            "费用",
            "库存",
            "回款",
            "风险",
            "问题",
            "全面分析",
            "下一步",
            "怎么做",
            "折扣",
            "退款",
            "平台费用",
            "损耗",
            "人工",
            "评价",
        )
    )
    chart_only = any(token in folded for token in ("画一个", "画张", "改图表", "修改图表"))
    restaurant_signal = any(token in folded for token in ("门店", "餐饮", "菜品", "食材", "外卖"))
    return company_context and (domain_score >= 4 or (restaurant_signal and domain_score >= 3)) and not chart_only


def _sales_report_source(
    entries: list[TableEntry],
) -> tuple[TableEntry, dict[str, str]] | None:
    candidates: list[tuple[TableEntry, dict[str, str]]] = []
    for entry in entries:
        try:
            candidates.append((entry, infer_sales_report_columns(entry.frame)))
        except (TypeError, ValueError):
            continue
    if len(candidates) != 1:
        return None
    return candidates[0]


def _quarterly_sales_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    candidates = 0
    for entry in entries:
        try:
            infer_sales_report_columns(entry.frame)
        except (TypeError, ValueError):
            continue
        candidates += 1
    return entries if candidates >= 2 else None


def _inventory_report_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    frames = [entry.frame for entry in entries]
    return entries if can_build_inventory_report(frames) else None


def _hr_report_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    frames = [entry.frame for entry in entries]
    return entries if can_build_hr_report(frames) else None


def _adaptive_report_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    # Never feed tables generated by a previous execution back into a new
    # analysis.  The active task's explicit uploads are the only source of
    # truth whenever original entries are available.
    originals = [entry for entry in entries if entry.original]
    candidates = originals or entries
    frames = [entry.frame for entry in candidates]
    return candidates if can_build_adaptive_report(frames) else None


def _selection_report_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    frames = [entry.frame for entry in entries]
    return entries if can_build_selection_report(frames) else None


def _enterprise_diagnosis_sources(entries: list[TableEntry]) -> list[TableEntry] | None:
    # A task may already contain prior derived tables.  They are outputs, not
    # new evidence; feeding them back creates self-ingestion and can exceed
    # the plan input limit.  Prefer the original upload set whenever present.
    originals = [entry for entry in entries if entry.original]
    if originals:
        entries = originals
    frames = [entry.frame for entry in entries]
    names = [entry.name for entry in entries]
    if __package__:
        from .restaurant_report import can_build_restaurant_diagnosis_report
    else:
        from excel_data_toolbox.restaurant_report import can_build_restaurant_diagnosis_report
    if can_build_restaurant_diagnosis_report(frames, names):
        return entries
    return entries if can_build_enterprise_diagnosis_report(frames) else None


def _sales_report_plan_payload(
    table_id: str,
    columns: dict[str, str],
) -> dict[str, Any]:
    params: dict[str, Any] = {
        **columns,
        "satisfaction_threshold": 4,
    }
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：已识别的销售数据表；处理动作：计算经营指标、产品与销售人员排名、"
            "月度/产品/地区可视化数据及异常提醒；关键字段/规则/阈值：销售额减成本得到利润，"
            "平均利润率按总利润除以总销售额，客户满意度低于4分列为重点关注；输出：管理层数据总览、"
            "产品分析、销售人员分析、异常数据提醒、图表展示五张工作表和三张原生Excel图表；"
            "人工核验边界：业务口径与异常处置建议由用户复核。"
        ),
        "message": "已识别为完整销售经营报告，将在本机自动计算并生成五张工作表。",
        "clarification_questions": [],
        "assumptions": ["平均利润率采用整体加权口径：总利润÷总销售额。"],
        "warnings": ["异常提醒用于辅助复核，不替代人工业务判断。"],
        "steps": [
            {
                "id": "sales_report_1",
                "operation": "sales_management_report",
                "input_ids": [table_id],
                "output_name": "销售经营分析报告",
                "params": params,
            }
        ],
    }


def _quarterly_sales_plan_payload(entries: list[TableEntry]) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：多张月度销售表及可选部门备注表；处理动作：自动识别表头，统一字段、日期、金额和业务文本，"
            "按订单编号去重，排除取消、退款、无效、非正金额及金额/成本缺失订单，合并后生成季度经营分析；"
            "关键字段/规则/阈值：订单编号作为去重键，满意度低于4分或评分缺失列为重点关注，利润=销售额-成本；"
            "输出：管理层总览、季度合并数据、产品/地区/销售人员分析、异常提醒、清洗审计和图表看板；"
            "人工核验边界：无效订单与重点关注项保留完整审计明细，业务责任和最终处置由用户复核。"
        ),
        "message": "已识别为多表季度销售经营报告，将在本机完成清洗、去重、无效排除、合并、分析和 Excel 交付。",
        "clarification_questions": [],
        "assumptions": [
            "依据部门备注默认口径：管理报表仅统计有效且未取消/未退款的订单。",
            "重复订单保留字段更完整且更早出现的一条；所有剔除明细写入清洗审计表。",
        ],
        "warnings": ["重点关注和剔除清单用于管理复核，不替代业务责任判定或会计凭证审核。"],
        "steps": [
            {
                "id": "quarterly_sales_1",
                "operation": "quarterly_sales_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "季度销售经营报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "satisfaction_threshold": 4,
                },
            }
        ],
    }


def _inventory_report_plan_payload(entries: list[TableEntry]) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：商品资料、期初库存、采购入库、销售出库、库存调整及仓库说明；"
            "处理动作：自动统一商品编码、日期、数量、金额和状态，按单据号去重，计算账面库存和可销售库存，"
            "分析采购、销售、补货、积压及数据质量；关键字段/规则/阈值：账面库存=期初+已入库-已完成出库+已确认调整，"
            "可销售库存=账面库存-已锁定-不良品，采用近30天销量，补货线=安全库存+日均销量×采购提前期，"
            "积压阈值为可售天数超过目标库存天数1.5倍，停售商品不补货；输出：管理层总览、商品库存分析、补货建议、"
            "积压清单、采购分析、销售分析、人工核验、数据审计和库存图表看板；人工核验边界：退货是否重新入库、"
            "缺失数量、待确认调整、未知商品和负库存保留人工核验。"
        ),
        "message": "已识别为采购销售库存联动分析，将在本机完成清洗、核算、预警、审计和九表 Excel 交付。",
        "clarification_questions": [],
        "assumptions": [
            "依据仓库说明：仅已入库采购、已完成销售和已确认调整参与库存计算。",
            "在未提供企业专属阈值时，使用近30天销量与1.5倍目标库存天数作为保守积压口径，并在报表中披露。",
        ],
        "warnings": ["退货是否重新入库、待确认调整和缺失数量不自动推断，统一进入人工核验。"],
        "steps": [
            {
                "id": "inventory_management_1",
                "operation": "inventory_management_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "采购销售库存经营报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "recent_days": 30,
                    "overstock_multiplier": 1.5,
                },
            }
        ],
    }


def _hr_report_plan_payload(entries: list[TableEntry]) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：员工基础信息、月度考勤、绩效、薪资调整及老板备注；处理动作：按员工编号整合，"
            "计算考勤、绩效、薪资和综合表现，识别优秀员工、重点关注人员与数据冲突；"
            "关键字段/规则/阈值：默认每人月应出勤22天，考勤得分=100-迟到×5-早退×5-请假×2-缺勤×10，"
            "绩效得分=目标完成率×70+客户评分/5×30，综合得分=考勤30%+绩效70%；"
            "输出：管理层人效总览、员工综合分析、优秀/关注员工、考勤、绩效、薪资、人工核验、审计和图表看板；"
            "人工核验边界：离职风险仅为考勤绩效代理预警，最终评价、奖惩和劳动人事决定必须由主管与人事复核。"
        ),
        "message": "已识别为员工考勤、绩效与薪资经营分析，将在本机完成整合、评分、预警、审计和十表 Excel 交付。",
        "clarification_questions": [],
        "assumptions": [
            "未提供企业制度时，采用22天/人月和披露在数据审计中的保守评分口径。",
            "销售额仅作为岗位业绩展示，不直接用于跨岗位综合评分。",
        ],
        "warnings": ["离职风险代理等级不代表真实离职意愿，不得单独用于处分、解雇或薪酬决定。"],
        "steps": [
            {
                "id": "hr_management_1",
                "operation": "hr_management_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "员工考勤绩效薪资经营分析报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "expected_workdays": 22,
                    "excellent_score": 85,
                    "attention_score": 70,
                },
            }
        ],
    }


def _adaptive_report_plan_payload(entries: list[TableEntry], prompt: str) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：本次任务显式上传的全部非空原始数据表；处理动作：由通用分析编译器识别领域、表角色、粒度、"
            "标准业务概念、指标聚合语义和用户意图，再按证据动态启用核心指标、分类表现、结构、趋势、关系、质量和风险分析；"
            "关键规则：历史输出永不作为输入，只合并字段集合一致的事实表，比例按分子分母重算，余额按期末，"
            "证据不足时列出缺口而不强行计算；"
            "输出：管理层通用总览、主数据、数据字典、质量、关系建议、分类排名、时间趋势、异常和图表看板九表；"
            "人工核验边界：业务口径、关联键语义和异常处置必须由用户确认。"
        ),
        "message": "已启用本地通用分析编译器，将按当前需求和字段证据动态生成九表 Excel 经营报告。",
        "clarification_questions": [],
        "assumptions": [
            "只读取本次任务的原始上传表；同构表按规范化字段集合一致判定并纵向合并，完全重复行自动删除。",
            "数值异常采用 1.5 倍 IQR 作为通用核验线索；不会据此自动删除记录。",
        ],
        "warnings": ["领域、字段角色和表关系属于数据与需求共同驱动的推断，已披露置信度和证据缺口，不能替代业务口径确认。"],
        "steps": [
            {
                "id": "adaptive_analysis_1",
                "operation": "adaptive_analysis_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "通用自适应经营分析报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "user_request": prompt,
                    "top_n": 10,
                    "outlier_multiplier": 1.5,
                },
            }
        ],
    }


def _selection_report_plan_payload(
    entries: list[TableEntry],
    prompt: str,
    *,
    include_charts: bool = True,
) -> dict[str, Any]:
    top_n = parse_selection_count(prompt, default=8)
    output_description = (
        "评选管理总览、建议入选名单、全部候选排序、风险复核清单、评选规则与字段、图表看板六表"
        if include_charts
        else "评选管理总览、建议入选名单、全部候选排序、风险复核清单、评选规则与字段五表"
    )
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            f"输入范围：当前选中的候选评分或问题记录表；处理动作：自动识别候选标识、各轮得分和评语字段，"
            f"按有效均分、最新表现、完整率、文本风险和正向依据计算可解释排序并选出{top_n}个；"
            "关键字段/规则/阈值：基础表现分=有效均分×70%+最新有效得分×30%，缺失分数、重复/诚信、"
            "核心结果错误、严重完成质量和纯AI等评语只形成披露的扣分与复核提示；"
            f"输出：{output_description}；"
            "人工核验边界：程序提供数据驱动推荐，最终参赛资格、原创性和专业结论须由组织者复核。"
        ),
        "message": f"已识别为候选对象结构化评选，将在本机选出{top_n}个并生成可追溯 Excel。",
        "clarification_questions": [],
        "assumptions": [
            "未给出正式评选办法时，使用公开披露的70%有效均分+30%最新得分作为基础表现口径。",
            "文字评语只用于风险扣分与人工复核提示，不把任何单一关键词作为自动淘汰决定。",
        ],
        "warnings": ["入选名单是辅助推荐，不替代赛事资格、原创性、专业正确性和最终评委决定。"],
        "steps": [
            {
                "id": "selection_recommendation_1",
                "operation": "selection_recommendation_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "候选对象结构化评选报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "user_request": prompt,
                    "top_n": top_n,
                    "include_charts": include_charts,
                },
            }
        ],
    }


def _enterprise_diagnosis_plan_payload(entries: list[TableEntry], prompt: str) -> dict[str, Any]:
    if __package__:
        from .ecommerce_report import can_build_ecommerce_diagnosis_report
        from .restaurant_report import can_build_restaurant_diagnosis_report, restaurant_diagnosis_profile
    else:  # Supports: python server.py
        from excel_data_toolbox.ecommerce_report import can_build_ecommerce_diagnosis_report
        from excel_data_toolbox.restaurant_report import can_build_restaurant_diagnosis_report, restaurant_diagnosis_profile

    restaurant_profile = restaurant_diagnosis_profile(
        [entry.frame for entry in entries], [entry.name for entry in entries]
    )
    is_restaurant = restaurant_profile is not None
    is_ecommerce = can_build_ecommerce_diagnosis_report(
        [entry.frame for entry in entries], [entry.name for entry in entries]
    )
    if restaurant_profile == "compact_store_period_pnl":
        summary = (
            "输入范围：门店-月份经营事实表及可选汇总校验表；处理动作：先按工作表语义区分事实表、汇总表和说明表，"
            "汇总表只用于校验且禁止重复计入，再对营业额、退款、食材、人工、平台、租金、水电、营销和管理利润执行加权经营分析；"
            "关键规则：所有金额求和，管理利润率=管理利润合计÷净营业收入合计，时点余额取期末，任何比例不得默认平均；"
            "输出：管理总览、数据源事实域、门店与月度诊断、成本结构、能力边界、风险行动、底稿、验收和四图看板；"
            "人工核验边界：未提供菜品/BOM、渠道结算、采购盘点、工时及评价明细的主题不强行判断。"
        )
        message = "已识别为门店-月份经营事实表，将排除汇总校验表并按加权口径生成经营诊断；不会因工作表较少降级为通用模板。"
        assumptions = [
            "净营业收入优先使用源字段并与营业额-退款勾稽。",
            "管理利润优先使用源字段并与已提供成本项目重算值勾稽；差异进入人工核验。",
        ]
    elif is_restaurant:
        summary = (
            "输入范围：门店、菜品、POS销售、退款、外卖平台结算、原料主数据、BOM、采购入库、盘点损耗、人工工时、固定费用和顾客评价；"
            "处理动作：按独立事实域识别粒度和业务键，保留多菜品订单，不按订单号误删明细，分别勾稽销售、退款、平台到账、标准食材成本、采购、损耗、人工和评价；"
            "关键规则：已退款才冲减收入，处理中退款只列风险，平台到账不冒充销售，BOM/损耗为管理代理指标，若销售与人工/固定费用尺度不匹配则不输出确定盈亏；"
            "输出：管理层总览、门店、渠道外卖、菜品、原料损耗、人工、评价退款、月度利润驱动、风险行动、底稿、数据口径和看板；"
            "人工核验边界：完整期间、单位、工资是否年化、固定费用归属、BOM版本和盘点差异需财务/运营确认。"
        )
        message = "已识别为餐饮门店多事实域诊断，将分别关联销售、退款、平台、菜品、原料、人工和评价后生成可追溯经营报告。"
        assumptions = [
            "POS精确重复才删除；同一订单的多菜品行保留。",
            "标准食材成本和损耗用于经营线索，不替代财务结转和完整库存流水。",
        ]
    elif is_ecommerce:
        summary = (
            "输入范围：商品、订单明细、售后退款、平台结算、广告投放、采购入库、月末库存、客户会员及口径说明；"
            "处理动作：按独立事实域识别粒度和业务键，清洗SKU与数量格式，保留多SKU订单，仅删除完整重复，排除取消/待付款/关闭订单，"
            "再建立订单→退款→结算→标准成本→广告→库存→客户的可追溯经营链；关键规则：已退款才冲减管理收入，处理中退款只披露风险，"
            "实际到账不冒充销售收入，广告7日归因不与订单收入相加，标准成本仅用于管理毛利，退货成本不自动冲回；"
            "输出：经营总览、利润驱动、渠道广告、商品利润、退款售后、平台回款、广告、采购、库存、客户、行动、审计与看板十三表；"
            "人工核验边界：管理贡献不等同财务净利润，季节性、广告归因、退货可回收入库及标准成本更新需业务/财务确认。"
        )
        message = "已识别为多平台电商多事实域诊断，将在本机重建利润、现金、投放、退款和库存经营链并生成十三表 Excel 交付。"
        assumptions = [
            "退款后管理收入=有效订单买家实付-已发生退款；处理中退款不提前冲减。",
            "趋势性管理贡献=平台实际到账-标准商品成本-广告费，仅用于管理诊断。",
        ]
    else:
        summary = (
            "输入范围：订单/收入、客户、人员绩效、费用、库存及可选生产成本等多张事实表和主数据；处理动作："
            "先识别各表业务角色与字段语义，再分别聚合收入、绩效毛利、期间费用、回款、客户、人员、库存和生产成本，"
            "禁止把不同粒度事实表直接笛卡尔连接；关键规则：优先使用可勾稽的绩效毛利或流水毛利，生产成本不默认等于"
            "同期销售成本；退款同时展示原始、排除和负数冲减情景；未结清订单只作为风险订单金额，不冒充应收余额；"
            "表关系同时披露行覆盖率和唯一键覆盖率，产品别名仅建议人工确认；输出：管理层诊断总览、利润驱动、客户与回款、销售团队、成本费用、库存风险、"
            "行动计划、诊断底稿、数据口径与验收、经营诊断看板十表；人工核验边界：税费、折旧、部分回款金额、"
            "安全库存、采购提前期及企业审批阈值未提供，不自动补造；缺失成本保留为空，未知客户不判定为低风险，"
            "综合风险不得低于源业务风险；看板按KPI、核心诊断、优先风险和自适应图表组织。"
        )
        message = "已识别为多事实域企业经营诊断，将在事实与建议分离的前提下定位利润、现金、客户和库存风险并生成十表 Excel 交付。"
        assumptions = [
            "退款口径未确认时同时展示三种情景，不静默改写原始金额。",
            "成本费用表作为期间费用；生产成本与销售成本缺少结转关系时分别披露，不重复扣减。",
        ]
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": summary,
        "message": message,
        "clarification_questions": [],
        "assumptions": assumptions,
        "warnings": [
            "经营利润为基于已提供数据的估算口径，不替代法定财务报表。",
            "行动建议只形成责任与审批清单，不自动执行授信、停发、降价、采购或人事决定。",
        ],
        "steps": [
            {
                "id": "enterprise_diagnosis_1",
                "operation": "enterprise_diagnosis_report",
                "input_ids": [entry.id for entry in entries],
                "output_name": "企业集团经营诊断报告",
                "params": {
                    "source_names": [entry.name for entry in entries],
                    "user_request": prompt,
                    "low_margin_threshold": 0.15,
                },
            }
        ],
    }


def _ai_review_category(output_name: str) -> str | None:
    lowered = output_name.casefold()
    patterns = (
        ("_review", "AI 对账待确认"),
        ("_failures", "AI 质量验收失败"),
        ("_amount_difference", "AI 金额差异"),
        ("_date_difference", "AI 日期差异"),
        ("_outliers", "AI 异常值"),
        ("_left_only", "AI 左表未匹配"),
        ("_right_only", "AI 右表未匹配"),
        ("_duplicates", "AI 重复键"),
        ("_invalid_rows", "AI 无效记录"),
    )
    for suffix, category in patterns:
        if lowered.endswith(suffix):
            return category
    return None


def _normalised_missing(series: pd.Series) -> pd.Series:
    return series.map(lambda value: pd.NA if isinstance(value, str) and not value.strip() else value)


def _numeric_column(frame: pd.DataFrame, column: str, *, minimum_ratio: float = 0.8) -> pd.Series:
    if _IDENTIFIER_NAME_HINT.search(column):
        raise ApiError(f"字段“{column}”看起来是标识符，不适合作为数值指标")
    source = _normalised_missing(frame[column])
    if pd.api.types.is_bool_dtype(source.dtype):
        raise ApiError(f"字段“{column}”是布尔字段，不能作为数值指标")
    converted = pd.to_numeric(source, errors="coerce")
    non_missing = source.notna()
    denominator = int(non_missing.sum())
    if denominator == 0:
        raise ApiError(f"字段“{column}”没有可用数值")
    valid = converted.notna() & converted.map(lambda value: bool(pd.notna(value)) and math.isfinite(float(value)))
    success_ratio = int((valid & non_missing).sum()) / denominator
    if success_ratio < minimum_ratio:
        raise ApiError(f"字段“{column}”只有 {success_ratio:.0%} 的非空内容可解析为数字，请先清洗类型")
    return converted.where(valid).astype("Float64")


def _column_metadata(frame: pd.DataFrame) -> pd.DataFrame:
    try:
        return descriptive_statistics(frame, include_text=True)
    except (TypeError, ValueError) as exc:
        raise ApiError(f"无法生成描述统计：{exc}") from exc


def _privacy_safe_statistics(statistics: pd.DataFrame) -> pd.DataFrame:
    result = statistics.copy(deep=True)
    if result.empty:
        return result
    for field in ("mode", "min", "max"):
        if field in result.columns:
            result[field] = result[field].astype("object")
    for position, row in result.iterrows():
        semantic = str(row.get("semantic_type") or "")
        column = str(row.get("column") or "")
        sensitive = bool(_SENSITIVE_NAME_HINT.search(column)) or semantic == "identifier"
        if "mode" in result.columns and (sensitive or semantic in {"text", "category", "mixed"}):
            if pd.notna(result.at[position, "mode"]):
                result.at[position, "mode"] = "（隐私保护：已隐藏）"
        if sensitive:
            for field in ("min", "max"):
                if field in result.columns:
                    result.at[position, field] = None
    return result


def _analysis_columns(
    frame: pd.DataFrame, statistics: pd.DataFrame
) -> tuple[list[str], list[str], str | None, str | None, str | None]:
    numeric_columns = [
        str(row["column"])
        for _, row in statistics.iterrows()
        if str(row.get("semantic_type")) in _NUMERIC_SEMANTIC_TYPES
        and not _IDENTIFIER_NAME_HINT.search(str(row["column"]))
    ]
    date_columns = [
        str(row["column"]) for _, row in statistics.iterrows() if str(row.get("semantic_type")) in _DATE_SEMANTIC_TYPES
    ]
    measure = next((column for column in numeric_columns if _MEASURE_NAME_HINT.search(column)), None)
    if measure is None and numeric_columns:
        measure = numeric_columns[0]
    date_column = next((column for column in date_columns if _DATE_NAME_HINT.search(column)), None)
    if date_column is None and date_columns:
        date_column = date_columns[0]

    candidates: list[tuple[int, int, str]] = []
    row_limit = max(20, min(200, max(1, len(frame) // 2)))
    for _, row in statistics.iterrows():
        column = str(row.get("column") or "")
        semantic = str(row.get("semantic_type") or "")
        unique_count = int(row.get("unique_count") or 0)
        if semantic not in {"category", "text"}:
            continue
        if _SENSITIVE_NAME_HINT.search(column) or _IDENTIFIER_NAME_HINT.search(column):
            continue
        if unique_count <= 0 or unique_count > row_limit:
            continue
        priority = 0 if _DIMENSION_NAME_HINT.search(column) else 1
        candidates.append((priority, unique_count, column))
    candidates.sort()
    dimension = candidates[0][2] if candidates else None
    return numeric_columns, date_columns, measure, date_column, dimension


def _issue_title(code: str, column: str | None) -> str:
    titles = {
        "EMPTY_DATASET": "数据表为空",
        "DUPLICATE_COLUMN_NAMES": "列名重复",
        "MISSING_VALUES": "字段存在空值",
        "EMPTY_COLUMN": "字段完全为空",
        "CONSTANT_COLUMN": "字段取值单一",
        "MIXED_TYPES": "字段类型混杂",
        "DUPLICATE_ROWS": "存在重复记录",
        "MISSING_KEYS": "业务键缺失",
        "DUPLICATE_KEYS": "业务键重复",
        "UNHASHABLE_VALUES": "存在复杂单元格内容",
    }
    title = titles.get(code, "数据质量提示")
    return f"{title} · {column}" if column else title


def _analysis_payload(frame: pd.DataFrame) -> dict[str, Any]:
    quality_report = assess_data_quality(frame)
    statistics = _column_metadata(frame)
    numeric_columns, date_columns, measure, date_column, dimension = _analysis_columns(frame, statistics)
    correlation = correlation_matrix(frame, columns=numeric_columns[:12])
    metrics = dict(quality_report.metrics)
    rows, columns = frame.shape
    missing_cells = int(metrics.get("missing_cell_count", 0))
    duplicate_rows = int(metrics.get("duplicate_row_count", 0))
    missing_rate = 1.0 - float(metrics.get("completeness_rate", 1.0))
    duplicate_rate = float(metrics.get("duplicate_row_rate", 0.0))
    memory_mb = float(frame.memory_usage(index=True, deep=True).sum()) / 1024 / 1024

    severity_map = {"high": "danger", "medium": "warning", "low": "info", "info": "info"}
    issues = [
        {
            "severity": severity_map.get(issue.severity, "info"),
            "title": _issue_title(issue.code, issue.column),
            "detail": issue.message,
            "recommendation": issue.suggestion,
            "code": issue.code,
            "column": issue.column,
            "count": issue.count,
            "ratio": issue.ratio,
        }
        for issue in quality_report.issues[:60]
    ]

    insights: list[dict[str, str]] = [
        {
            "title": f"数据质量为“{quality_report.grade}”",
            "detail": f"综合得分 {quality_report.score:.1f} 分，共识别 {len(quality_report.issues)} 项质量提示。",
        },
        {
            "title": "数据规模概览",
            "detail": f"当前数据表共 {rows:,} 行、{columns:,} 列，占用约 {memory_mb:.2f} MB 内存。",
        },
    ]
    if missing_cells:
        missing_issue = max(
            (issue for issue in quality_report.issues if issue.code == "MISSING_VALUES"),
            key=lambda item: item.count,
            default=None,
        )
        if missing_issue is not None:
            insights.append(
                {
                    "title": "优先处理缺失值",
                    "detail": f"共 {missing_cells:,} 个空值；字段“{missing_issue.column}”缺失最多（{missing_issue.count:,} 个）。",
                }
            )
    else:
        insights.append({"title": "字段完整度较好", "detail": "未检测到空白字符串或空值单元格。"})
    if duplicate_rows:
        insights.append(
            {
                "title": "重复记录可能影响统计",
                "detail": f"有 {duplicate_rows:,} 行处于完全重复记录组，占总行数 {duplicate_rate:.1%}。",
            }
        )

    if len(correlation.columns) >= 2:
        strongest: tuple[float, str, str, float] | None = None
        corr_columns = [str(column) for column in correlation.columns]
        for row_index, left in enumerate(corr_columns):
            for column_index in range(row_index + 1, len(corr_columns)):
                value = correlation.iloc[row_index, column_index]
                if pd.isna(value) or not math.isfinite(float(value)):
                    continue
                candidate = (abs(float(value)), left, corr_columns[column_index], float(value))
                if strongest is None or candidate[0] > strongest[0]:
                    strongest = candidate
        if strongest is not None:
            insights.append(
                {
                    "title": "数值字段相关关系",
                    "detail": f"“{strongest[1]}”与“{strongest[2]}”的相关系数为 {strongest[3]:.2f}；相关不代表因果，建议结合业务复核。",
                }
            )

    if date_column and measure:
        try:
            trend = aggregate_trend(
                frame,
                date_column=date_column,
                value_columns=measure,
                frequency="month",
                aggregation="sum",
            ).data
            if len(trend) >= 2:
                first = float(trend[measure].iloc[0])
                last = float(trend[measure].iloc[-1])
                change = None if first == 0 else (last - first) / abs(first)
                detail = f"按月汇总后，首期为 {first:,.2f}，末期为 {last:,.2f}。"
                if change is not None:
                    detail += f" 变化幅度 {change:+.1%}。"
                insights.append({"title": f"{measure}趋势", "detail": detail})
        except (TypeError, ValueError):
            pass

    if dimension and measure:
        try:
            contribution = category_contribution(
                frame,
                category_columns=dimension,
                value_column=measure,
                aggregation="sum",
                top_n=10,
            )
            if not contribution.data.empty:
                top = contribution.data.iloc[0]
                insights.append(
                    {
                        "title": f"{dimension}贡献集中度",
                        "detail": f"贡献最高的是“{top[dimension]}”，占整体 {float(top['contribution_pct']):.1%}；达到 80% 累计贡献约需 {contribution.core_category_count} 个类别。",
                    }
                )
        except (TypeError, ValueError):
            pass

    correlation_payload = {
        "columns": [str(column) for column in correlation.columns],
        "matrix": correlation.to_numpy().tolist() if not correlation.empty else [],
    }
    return {
        "quality": {
            "score": quality_report.score,
            "grade": quality_report.grade,
            "summary": f"完成度、重复、结构和类型一致性综合扫描；当前有 {len(quality_report.issues)} 项提示。",
        },
        "overview": {
            "rows": rows,
            "columns": columns,
            "missing_cells": missing_cells,
            "missing_rate": round(missing_rate, 6),
            "duplicate_rows": duplicate_rows,
            "duplicate_rate": round(duplicate_rate, 6),
            "numeric_columns": len(numeric_columns),
            "date_columns": len(date_columns),
            "memory_mb": round(memory_mb, 4),
        },
        "issues": issues,
        "insights": insights[:10],
        "correlations": correlation_payload,
    }


def _date_series_for_chart(series: pd.Series, *, required: bool) -> pd.Series | None:
    source = _normalised_missing(series)
    parsed = pd.to_datetime(source, errors="coerce", format="mixed")
    non_missing = int(source.notna().sum())
    success = int(parsed.notna().sum())
    if non_missing == 0 or success == 0:
        if required:
            raise ApiError("所选维度没有可用日期")
        return None
    if success / non_missing < 0.8:
        if required:
            raise ApiError("所选维度不足 80% 的内容可解析为日期，请先清洗")
        return None
    if getattr(parsed.dt, "tz", None) is not None:
        parsed = parsed.dt.tz_localize(None)
    return parsed


def _automatic_date_grain(dates: pd.Series) -> str:
    valid = dates.dropna()
    if valid.empty:
        return "month"
    span_days = max(0, int((valid.max() - valid.min()).days))
    if span_days <= 60:
        return "day"
    if span_days <= 240:
        return "week"
    if span_days <= 1095:
        return "month"
    if span_days <= 2555:
        return "quarter"
    return "year"


def _period_values(dates: pd.Series, grain: str) -> pd.Series:
    aliases = {"day": "D", "week": "W", "month": "M", "quarter": "Q", "year": "Y"}
    return dates.dt.to_period(aliases[grain]).dt.start_time


def _period_label(value: Any, grain: str) -> str:
    timestamp = pd.Timestamp(value)
    if grain == "day":
        return timestamp.strftime("%Y-%m-%d")
    if grain == "week":
        return f"{timestamp:%Y-%m-%d} 周"
    if grain == "month":
        return timestamp.strftime("%Y-%m")
    if grain == "quarter":
        return f"{timestamp.year} Q{timestamp.quarter}"
    return str(timestamp.year)


def _aggregate_chart(
    frame: pd.DataFrame,
    *,
    dimension: str,
    measure: str,
    aggregation: str,
    chart_type: str,
    top_n: int,
    date_grain: str,
) -> tuple[list[str], list[float], str | None]:
    explicit_date = date_grain != "auto"
    hinted_date = bool(_DATE_NAME_HINT.search(dimension)) or pd.api.types.is_datetime64_any_dtype(
        frame[dimension].dtype
    )
    dates = _date_series_for_chart(frame[dimension], required=explicit_date) if (explicit_date or hinted_date) else None
    grain: str | None = None
    if dates is not None:
        grain = _automatic_date_grain(dates) if date_grain == "auto" else date_grain
        dimension_values: pd.Series = _period_values(dates, grain)
    else:
        dimension_values = _normalised_missing(frame[dimension]).map(
            lambda value: "（空值）" if pd.isna(value) else str(value).strip() or "（空值）"
        )

    work = pd.DataFrame({"__dimension__": dimension_values}, index=frame.index)
    if aggregation == "count":
        work["__value__"] = 1.0
        grouped = (
            work.dropna(subset=["__dimension__"])
            .groupby("__dimension__", sort=bool(grain), dropna=False, observed=True)["__value__"]
            .sum()
        )
    elif aggregation == "nunique":
        work["__raw__"] = _normalised_missing(frame[measure])
        grouped = (
            work.dropna(subset=["__dimension__"])
            .groupby("__dimension__", sort=bool(grain), dropna=False, observed=True)["__raw__"]
            .nunique(dropna=True)
        )
    else:
        work["__value__"] = _numeric_column(frame, measure)
        valid_work = work.dropna(subset=["__dimension__", "__value__"])
        if valid_work.empty:
            raise ApiError("所选字段没有可用于图表的数据")
        grouped = valid_work.groupby("__dimension__", sort=bool(grain), dropna=False, observed=True)["__value__"].agg(
            aggregation
        )
    grouped = grouped.dropna()
    if grouped.empty:
        raise ApiError("所选字段没有可用于图表的数据")
    if grain and chart_type in {"line", "area", "waterfall"}:
        grouped = grouped.sort_index(kind="stable").tail(top_n)
    elif chart_type in {"funnel", "waterfall"}:
        # Funnel stages and waterfall movements are sequences.  Sorting them
        # by magnitude changes the business meaning, so keep first appearance.
        grouped = grouped.head(top_n)
    else:
        grouped = grouped.sort_values(ascending=False, kind="stable").head(top_n)
    labels = [_period_label(value, grain) if grain else str(value) for value in grouped.index.tolist()]
    values = [float(value) for value in grouped.tolist()]
    return labels, values, grain


def _chart_dimension_values(frame: pd.DataFrame, dimension: str, date_grain: str) -> tuple[pd.Series, str | None]:
    explicit_date = date_grain != "auto"
    hinted_date = bool(_DATE_NAME_HINT.search(dimension)) or pd.api.types.is_datetime64_any_dtype(
        frame[dimension].dtype
    )
    dates = _date_series_for_chart(frame[dimension], required=explicit_date) if explicit_date or hinted_date else None
    if dates is not None:
        grain = _automatic_date_grain(dates) if date_grain == "auto" else date_grain
        return _period_values(dates, grain), grain
    values = _normalised_missing(frame[dimension]).map(
        lambda value: "（空值）" if pd.isna(value) else str(value).strip() or "（空值）"
    )
    return values, None


def _multi_series_chart(
    frame: pd.DataFrame,
    *,
    dimension: str,
    series_column: str,
    measure: str,
    aggregation: str,
    top_n: int,
    date_grain: str,
) -> dict[str, Any]:
    if series_column == dimension:
        raise ApiError("系列字段不能与横轴字段相同")
    dimension_values, grain = _chart_dimension_values(frame, dimension, date_grain)
    series_values = _normalised_missing(frame[series_column]).map(
        lambda value: "（空值）" if pd.isna(value) else str(value).strip() or "（空值）"
    )
    work = pd.DataFrame(
        {"__dimension__": dimension_values, "__series__": series_values},
        index=frame.index,
    )
    group_keys = ["__dimension__", "__series__"]
    if aggregation == "count":
        work["__value__"] = 1.0
        grouped = work.dropna(subset=group_keys).groupby(group_keys, sort=bool(grain), observed=True)["__value__"].sum()
    elif aggregation == "nunique":
        work["__raw__"] = _normalised_missing(frame[measure])
        grouped = (
            work.dropna(subset=group_keys)
            .groupby(group_keys, sort=bool(grain), observed=True)["__raw__"]
            .nunique(dropna=True)
        )
    else:
        work["__value__"] = _numeric_column(frame, measure)
        valid = work.dropna(subset=[*group_keys, "__value__"])
        if valid.empty:
            raise ApiError("所选字段没有可用于多系列图表的数据")
        grouped = valid.groupby(group_keys, sort=bool(grain), observed=True)["__value__"].agg(aggregation)
    if grouped.empty:
        raise ApiError("所选字段没有可用于多系列图表的数据")
    pivot = grouped.unstack(fill_value=0).astype(float)
    if grain:
        pivot = pivot.sort_index(kind="stable").tail(top_n)
    else:
        row_order = pivot.abs().sum(axis=1).sort_values(ascending=False).head(top_n).index
        pivot = pivot.loc[row_order]
    column_order = pivot.abs().sum(axis=0).sort_values(ascending=False).head(8).index
    pivot = pivot.loc[:, column_order]
    labels = [_period_label(value, grain) if grain else str(value) for value in pivot.index.tolist()]
    series = [
        {
            "name": str(column),
            "values": [float(value) for value in pivot[column].tolist()],
        }
        for column in pivot.columns
    ]
    return {
        "labels": labels,
        "series": series,
        "badge": f"{len(labels)} 个横轴项 · {len(series)} 个系列",
        "summary": (
            f"按“{series_column}”拆分为 {len(series)} 个系列；展示值合计 {float(pivot.to_numpy().sum()):,.2f}。"
        ),
    }


def _multi_measure_chart(
    frame: pd.DataFrame,
    *,
    dimension: str,
    measures: list[str],
    aggregation: str,
    top_n: int,
    date_grain: str,
) -> dict[str, Any]:
    """Build aligned series from several numeric columns in a wide table."""

    if not 2 <= len(measures) <= 8:
        raise ApiError("多指标图表需要选择 2~8 个数值字段")
    if dimension in measures:
        raise ApiError("横轴字段不能同时作为指标字段")
    dimension_values, grain = _chart_dimension_values(frame, dimension, date_grain)
    grouped_columns: dict[str, pd.Series] = {}
    skipped_measures: list[str] = []
    for measure in measures:
        if int(_normalised_missing(frame[measure]).notna().sum()) == 0:
            skipped_measures.append(measure)
            continue
        values = _numeric_column(frame, measure)
        work = pd.DataFrame({"__dimension__": dimension_values, "__value__": values})
        valid = work.dropna(subset=["__dimension__", "__value__"])
        if valid.empty:
            skipped_measures.append(measure)
            continue
        grouped = valid.groupby("__dimension__", sort=bool(grain), observed=True)["__value__"]
        if aggregation == "sum":
            result = grouped.sum(min_count=1)
        elif aggregation == "count":
            result = grouped.count()
        elif aggregation == "nunique":
            result = grouped.nunique(dropna=True)
        else:
            result = grouped.agg(aggregation)
        grouped_columns[measure] = result.astype(float).rename(measure)
    if not grouped_columns:
        raise ApiError("所选字段没有可用于多指标图表的数据")
    matrix = pd.concat(grouped_columns.values(), axis=1)
    matrix.columns = list(grouped_columns)
    matrix = matrix.dropna(how="all")
    if matrix.empty:
        raise ApiError("所选字段没有可用于多指标图表的数据")
    matrix = matrix.sort_index(kind="stable").tail(top_n) if grain else matrix.head(top_n)
    labels = [_period_label(value, grain) if grain else str(value) for value in matrix.index.tolist()]
    series = [
        {
            "name": measure,
            "values": [None if pd.isna(value) else float(value) for value in matrix[measure].tolist()],
        }
        for measure in grouped_columns
    ]
    valid_count = int(matrix.notna().sum().sum())
    return {
        "labels": labels,
        "series": series,
        "badge": f"{len(labels)} 个横轴项 · {len(series)} 个指标系列",
        "summary": (
            f"已将 {len(series)} 个得分/指标字段按“{dimension}”对齐，共展示 {valid_count} 个有效数值；"
            f"空白单元格不绘制。"
            + (f" 已自动跳过整列为空的字段：{'、'.join(skipped_measures)}。" if skipped_measures else "")
        ),
    }


def _linear_regression_payload(paired: pd.DataFrame) -> dict[str, Any] | None:
    if len(paired) < 2:
        return None
    x_values = paired["x"].astype(float)
    y_values = paired["y"].astype(float)
    x_mean = float(x_values.mean())
    y_mean = float(y_values.mean())
    denominator = float(((x_values - x_mean) ** 2).sum())
    if denominator == 0:
        return None
    slope = float(((x_values - x_mean) * (y_values - y_mean)).sum() / denominator)
    intercept = y_mean - slope * x_mean
    predictions = slope * x_values + intercept
    total = float(((y_values - y_mean) ** 2).sum())
    residual = float(((y_values - predictions) ** 2).sum())
    r_squared = 1.0 - residual / total if total > 0 else 1.0
    min_x = float(x_values.min())
    max_x = float(x_values.max())
    return {
        "slope": slope,
        "intercept": intercept,
        "r_squared": max(0.0, min(1.0, r_squared)),
        "equation": f"y = {slope:.6g}x {intercept:+.6g}",
        "points": [
            {"x": min_x, "y": slope * min_x + intercept},
            {"x": max_x, "y": slope * max_x + intercept},
        ],
    }


def _box_chart_payload(frame: pd.DataFrame, *, dimension: str | None, measure: str, top_n: int) -> dict[str, Any]:
    values = _numeric_column(frame, measure)
    if dimension:
        dimensions = _normalised_missing(frame[dimension]).map(
            lambda value: "（空值）" if pd.isna(value) else str(value).strip() or "（空值）"
        )
    else:
        dimensions = pd.Series("全部", index=frame.index)
    work = pd.DataFrame({"group": dimensions, "value": values}).dropna(subset=["value"])
    if work.empty:
        raise ApiError(f"字段“{measure}”没有有效数值")
    group_order = work["group"].value_counts().head(min(top_n, 15)).index
    boxes: list[dict[str, Any]] = []
    for group in group_order:
        current = work.loc[work["group"] == group, "value"].astype(float)
        if current.empty:
            continue
        q1 = float(current.quantile(0.25))
        median = float(current.quantile(0.5))
        q3 = float(current.quantile(0.75))
        iqr = q3 - q1
        low_fence, high_fence = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        normal = current[(current >= low_fence) & (current <= high_fence)]
        boxes.append(
            {
                "label": str(group),
                "min": float(normal.min()) if not normal.empty else float(current.min()),
                "q1": q1,
                "median": median,
                "q3": q3,
                "max": float(normal.max()) if not normal.empty else float(current.max()),
                "outliers": [
                    float(value) for value in current[(current < low_fence) | (current > high_fence)].head(30)
                ],
                "count": int(len(current)),
            }
        )
    return {
        "chart_type": "box",
        "title": f"{measure}分组箱线分布" if dimension else f"{measure}箱线分布",
        "badge": f"{len(boxes)} 个分组",
        "boxes": boxes,
        "summary": f"箱体展示 Q1、中位数、Q3 和 1.5×IQR 须线；共分析 {len(work):,} 个有效值。",
    }


def _gantt_chart_payload(
    frame: pd.DataFrame,
    *,
    task_column: str,
    start_column: str,
    end_column: str,
    progress_column: str | None,
    top_n: int,
) -> dict[str, Any]:
    starts = pd.to_datetime(frame[start_column], errors="coerce", format="mixed")
    ends = pd.to_datetime(frame[end_column], errors="coerce", format="mixed")
    tasks = _normalised_missing(frame[task_column]).map(
        lambda value: "（未命名任务）" if pd.isna(value) else str(value).strip() or "（未命名任务）"
    )
    valid = starts.notna() & ends.notna() & (ends >= starts)
    if not valid.any():
        raise ApiError("甘特图需要有效的任务名称、开始日期和结束日期，且结束日期不能早于开始日期")
    progress = (
        pd.to_numeric(frame[progress_column], errors="coerce").fillna(0).clip(0, 100)
        if progress_column
        else pd.Series(0.0, index=frame.index)
    )
    work = (
        pd.DataFrame({"task": tasks, "start": starts, "end": ends, "progress": progress})
        .loc[valid]
        .sort_values(["start", "end"], kind="stable")
        .head(top_n)
    )
    items = [
        {
            "task": str(row.task),
            "start": row.start.strftime("%Y-%m-%d"),
            "end": row.end.strftime("%Y-%m-%d"),
            "progress": float(row.progress),
            "duration_days": max(1, int((row.end - row.start).days) + 1),
        }
        for row in work.itertuples(index=False)
    ]
    return {
        "chart_type": "gantt",
        "title": f"{task_column}项目进度甘特图",
        "badge": f"{len(items)} 项任务",
        "items": items,
        "summary": f"时间范围 {work['start'].min():%Y-%m-%d} 至 {work['end'].max():%Y-%m-%d}；进度限定在 0–100%。",
    }


def _chart_payload(frame: pd.DataFrame, payload: dict[str, Any]) -> dict[str, Any]:
    chart_type = _require_choice(payload.get("chart_type"), _CHART_TYPES, label="图表类型")
    aggregation = _require_choice(payload.get("aggregation", "sum"), _CHART_AGGREGATIONS, label="统计方式")
    date_grain = _require_choice(payload.get("date_grain", "auto"), _DATE_GRAINS, label="日期粒度")
    top_n = _bounded_integer(payload.get("top_n", 10), label="展示数量", minimum=1, maximum=50)
    style_3d = payload.get("style_3d") is True

    if chart_type == "gantt":
        task_column = _require_column(frame, payload.get("dimension"), label="任务名称字段")
        start_column = _require_column(frame, payload.get("start"), label="开始日期字段")
        end_column = _require_column(frame, payload.get("end"), label="结束日期字段")
        raw_progress = payload.get("progress")
        progress_column = _require_column(frame, raw_progress, label="进度字段") if raw_progress else None
        result = _gantt_chart_payload(
            frame,
            task_column=task_column,
            start_column=start_column,
            end_column=end_column,
            progress_column=progress_column,
            top_n=top_n,
        )
        result["style_3d"] = style_3d
        return result

    raw_measures = payload.get("measures") or []
    if raw_measures:
        if not isinstance(raw_measures, list):
            raise ApiError("多指标字段格式无效")
        if chart_type not in {"grouped_bar", "stacked_bar", "radar", "heatmap"}:
            raise ApiError("多指标对比仅支持分组柱状图、堆叠柱状图、雷达图或热力图")
        dimension = _require_column(frame, payload.get("dimension"), label="维度或横轴字段")
        measures = [
            _require_column(frame, item, label=f"指标字段 {index + 1}") for index, item in enumerate(raw_measures)
        ]
        multi = _multi_measure_chart(
            frame,
            dimension=dimension,
            measures=measures,
            aggregation=aggregation,
            top_n=top_n,
            date_grain=date_grain,
        )
        multi.update(
            {
                "chart_type": chart_type,
                "title": f"{'、'.join(measures)}按{dimension}对比",
                "style_3d": style_3d,
            }
        )
        return multi

    measure = _require_column(frame, payload.get("measure"), label="指标或纵轴字段")

    if chart_type == "histogram":
        values = _numeric_column(frame, measure).dropna()
        if values.empty:
            raise ApiError(f"字段“{measure}”没有有效数值")
        bin_count = min(30, max(5, min(top_n, int(math.sqrt(len(values))) or 5)))
        buckets = pd.cut(values.astype(float), bins=bin_count, duplicates="drop")
        counts = buckets.value_counts(sort=False)
        finite_values = values.astype(float)
        return {
            "chart_type": "histogram",
            "title": f"{measure}数值分布",
            "badge": f"{len(values):,} 个有效值",
            "labels": [str(label) for label in counts.index],
            "values": [int(value) for value in counts.tolist()],
            "summary": f"中位数 {finite_values.median():,.2f}，平均值 {finite_values.mean():,.2f}，范围 {finite_values.min():,.2f} 至 {finite_values.max():,.2f}。",
            "style_3d": style_3d,
        }

    if chart_type == "box":
        raw_dimension = payload.get("dimension")
        dimension = _require_column(frame, raw_dimension, label="分组字段") if raw_dimension else None
        result = _box_chart_payload(frame, dimension=dimension, measure=measure, top_n=top_n)
        result["style_3d"] = style_3d
        return result

    dimension = _require_column(frame, payload.get("dimension"), label="维度或横轴字段")
    if chart_type in {"stacked_bar", "grouped_bar", "radar", "heatmap"}:
        series_column = _require_column(frame, payload.get("series"), label="系列字段")
        multi = _multi_series_chart(
            frame,
            dimension=dimension,
            series_column=series_column,
            measure=measure,
            aggregation=aggregation,
            top_n=top_n,
            date_grain=date_grain,
        )
        multi.update(
            {
                "chart_type": chart_type,
                "title": (
                    f"{measure} · {dimension} × {series_column}热力矩阵"
                    if chart_type == "heatmap"
                    else (
                        f"{measure}按{dimension}与{series_column}雷达对比"
                        if chart_type == "radar"
                        else (
                            f"{measure}按{dimension}与{series_column}分组对比"
                            if chart_type == "grouped_bar"
                            else f"{measure}按{dimension}与{series_column}堆叠分析"
                        )
                    )
                ),
                "style_3d": style_3d,
            }
        )
        return multi

    if chart_type == "scatter":
        if dimension == measure:
            raise ApiError("散点图的横轴和纵轴字段不能相同")
        x_values = _numeric_column(frame, dimension)
        y_values = _numeric_column(frame, measure)
        valid = x_values.notna() & y_values.notna()
        positions = [position for position, flag in enumerate(valid.tolist()) if flag][:top_n]
        if not positions:
            raise ApiError("两个字段之间没有成对的有效数值")
        points = [
            {
                "x": float(x_values.iloc[position]),
                "y": float(y_values.iloc[position]),
                "label": f"数据点 {index + 1}",
            }
            for index, position in enumerate(positions)
        ]
        paired = pd.DataFrame({"x": x_values.loc[valid], "y": y_values.loc[valid]}).astype(float)
        correlation = paired["x"].corr(paired["y"]) if len(paired) >= 2 else float("nan")
        regression = _linear_regression_payload(paired)
        correlation_text = "样本不足" if pd.isna(correlation) else f"相关系数 {float(correlation):.2f}"
        regression_text = f"；线性回归 {regression['equation']}，R²={regression['r_squared']:.3f}" if regression else ""
        return {
            "chart_type": "scatter",
            "title": f"{dimension} × {measure} 散点关系",
            "badge": f"展示 {len(points)} 个数据点",
            "points": points,
            "trendline": regression,
            "summary": f"有效成对记录 {len(paired):,} 行；{correlation_text}{regression_text}。相关性和回归不代表因果关系。",
            "style_3d": style_3d,
        }

    labels, values, grain = _aggregate_chart(
        frame,
        dimension=dimension,
        measure=measure,
        aggregation=aggregation,
        chart_type=chart_type,
        top_n=top_n,
        date_grain=date_grain,
    )
    if chart_type in {"pie", "funnel", "treemap"} and any(value < 0 for value in values):
        raise ApiError("占比、漏斗和矩形树图不支持负数，请改用柱状图或瀑布图")
    aggregation_label = _AGGREGATION_LABELS[aggregation]
    title = f"{measure}按{dimension}{aggregation_label}"
    if grain:
        title = f"{measure}{aggregation_label} · {dimension}按{_DATE_GRAIN_LABELS.get(grain, grain)}趋势"
    if chart_type in {"line", "area"} and len(values) >= 2 and values[0] != 0:
        change = (values[-1] - values[0]) / abs(values[0])
        summary = f"首个展示周期为 {values[0]:,.2f}，末期为 {values[-1]:,.2f}，变化 {change:+.1%}。"
    elif chart_type == "funnel" and len(values) >= 2 and values[0] != 0:
        conversion = values[-1] / values[0]
        summary = f"首阶段“{labels[0]}”为 {values[0]:,.2f}，末阶段“{labels[-1]}”为 {values[-1]:,.2f}，整体转化率 {conversion:.1%}。"
    elif chart_type == "waterfall":
        summary = f"按业务顺序展示 {len(values)} 项增减，累计结果 {sum(values):,.2f}。"
    else:
        total = sum(values)
        summary = f"展示 {len(values)} 项；最高项为“{labels[0]}”，数值 {values[0]:,.2f}，展示项合计 {total:,.2f}。"
    return {
        "chart_type": chart_type,
        "title": title,
        "badge": f"{len(labels)} 个汇总项",
        "labels": labels,
        "values": values,
        "summary": summary,
        "style_3d": style_3d,
    }


_CHART_THEMES = {"default", "business_dark", "economist", "swiss", "finance", "warm", "minimal"}
_CHART_NUMBER_FORMATS = {"auto", "number", "currency", "percent", "wan", "yi"}
_CHART_SORT_MODES = {"auto", "asc", "desc", "source"}
_HEX_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _is_direct_chart_request(text: str) -> bool:
    """Return whether aggregation can safely stay inside the chart engine.

    Natural-language routers sometimes classify requests such as "按月汇总并画
    折线图" as a two-stage data transformation.  The chart engine already performs
    bounded sum/count/mean grouping locally, so creating an intermediate AI table
    only adds an unnecessary plan-schema failure point.  Explicit workbook mutation
    requests still use the normal data-plan path.
    """

    if not isinstance(text, str):
        return False
    has_chart = bool(
        re.search(
            r"(图表|画图|绘图|可视化|折线图|柱状图|条形图|饼图|散点图|面积图|雷达图|漏斗图|瀑布图|热力图|甘特图|看板)",
            text,
            flags=re.IGNORECASE,
        )
    )
    has_mutation = bool(
        re.search(
            r"(清洗|去重|删除|替换|填充|合并表|拼接表|匹配表|关联表|拆分表|新增列|修改原表|覆盖|脱敏|对账)",
            text,
            flags=re.IGNORECASE,
        )
    )
    return has_chart and not has_mutation


def _apply_chart_presentation(result: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    """Attach safe presentation controls; no model-produced code is accepted."""

    output = dict(result)
    title = payload.get("title")
    if title is not None:
        if not isinstance(title, str) or not title.strip() or len(title.strip()) > 80:
            raise ApiError("图表标题必须是 1~80 个字符")
        output["title"] = title.strip()
    output["theme"] = _require_choice(payload.get("theme", "default"), _CHART_THEMES, label="图表主题")
    output["number_format"] = _require_choice(
        payload.get("number_format", "auto"), _CHART_NUMBER_FORMATS, label="数字格式"
    )
    sort_mode = _require_choice(payload.get("sort", "auto"), _CHART_SORT_MODES, label="排序方式")
    output["sort"] = sort_mode
    for key in ("x_axis_label", "y_axis_label"):
        value = payload.get(key)
        if value is not None:
            if not isinstance(value, str) or not value.strip() or len(value.strip()) > 40:
                raise ApiError(f"{key} 必须是 1~40 个字符")
            output[key] = value.strip()
        else:
            output[key] = None
    raw_series_colors = payload.get("series_colors", [])
    if not isinstance(raw_series_colors, list) or len(raw_series_colors) > 8:
        raise ApiError("系列颜色必须是不超过 8 项的列表")
    series_colors: list[str] = []
    for color in raw_series_colors:
        if not isinstance(color, str) or not _HEX_COLOR.fullmatch(color):
            raise ApiError("系列颜色必须是 #RRGGBB")
        series_colors.append(color.upper())
    output["series_colors"] = series_colors
    measures = payload.get("measures") or []
    if measures and series_colors:
        if len(measures) != len(series_colors):
            raise ApiError("系列颜色数量必须与指标字段数量一致")
        color_by_measure = dict(zip(measures, series_colors))
        if isinstance(output.get("series"), list):
            output["series"] = [
                {**item, "color": color_by_measure.get(item.get("name"))}
                for item in output["series"]
                if isinstance(item, dict)
            ]
    for key, default in (("background_color", "#FFFFFF"), ("text_color", "#243831")):
        color = payload.get(key, default)
        if not isinstance(color, str) or not _HEX_COLOR.fullmatch(color):
            raise ApiError(f"{key} 必须是 #RRGGBB")
        output[key] = color.upper()
    integer_limits = {
        "font_size": (10, 24, 12),
        "label_rotation": (-90, 90, 0),
        "chart_height": (240, 600, 340),
    }
    for key, (minimum, maximum, default) in integer_limits.items():
        value = payload.get(key, default)
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or int(value) != value
            or not minimum <= int(value) <= maximum
        ):
            raise ApiError(f"{key} 必须是 {minimum}~{maximum} 的整数")
        output[key] = int(value)
    legend_position = payload.get("legend_position", "bottom")
    if legend_position not in {"top", "bottom", "left", "right"}:
        raise ApiError("legend_position 必须是 top、bottom、left 或 right")
    output["legend_position"] = legend_position
    show_grid = payload.get("show_grid", True)
    if not isinstance(show_grid, bool):
        raise ApiError("show_grid 必须是布尔值")
    output["show_grid"] = show_grid
    for key, minimum, maximum, default in (
        ("opacity", 0.2, 1.0, 0.92),
        ("bar_gap", 0.0, 0.8, 0.22),
    ):
        value = payload.get(key, default)
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(float(value))
            or not minimum <= float(value) <= maximum
        ):
            raise ApiError(f"{key} 必须在 {minimum}~{maximum} 之间")
        output[key] = float(value)
    for key in ("y_min", "y_max"):
        value = payload.get(key)
        if value is not None and (
            isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value))
        ):
            raise ApiError(f"{key} 必须是有限数字或 null")
        output[key] = None if value is None else float(value)
    if output["y_min"] is not None and output["y_max"] is not None and output["y_min"] >= output["y_max"]:
        raise ApiError("y_min 必须小于 y_max")
    for key in ("show_labels", "show_legend"):
        value = payload.get(key, True)
        if not isinstance(value, bool):
            raise ApiError(f"{key} 必须是布尔值")
        output[key] = value

    raw_lines = payload.get("reference_lines", [])
    if not isinstance(raw_lines, list) or len(raw_lines) > 5:
        raise ApiError("参考线最多 5 条")
    safe_lines: list[dict[str, Any]] = []
    for item in raw_lines:
        if not isinstance(item, dict) or set(item) != {"value", "label", "color"}:
            raise ApiError("参考线结构无效")
        value = item.get("value")
        color = item.get("color")
        label = item.get("label")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
            raise ApiError("参考线数值无效")
        if not isinstance(color, str) or not _HEX_COLOR.fullmatch(color):
            raise ApiError("参考线颜色必须是 #RRGGBB")
        if not isinstance(label, str) or not label.strip() or len(label.strip()) > 80:
            raise ApiError("参考线标签必须是 1~80 个字符")
        safe_lines.append({"value": float(value), "label": label.strip(), "color": color.upper()})
    output["reference_lines"] = safe_lines

    highlight = payload.get("highlight")
    output["highlight"] = None
    if highlight is not None:
        if not isinstance(highlight, dict) or set(highlight) != {"field", "value", "color"}:
            raise ApiError("高亮规则结构无效")
        field, value, color = highlight.get("field"), highlight.get("value"), highlight.get("color")
        if not isinstance(field, str) or not field or len(field) > 200:
            raise ApiError("高亮字段无效")
        if not isinstance(value, str) or not value or len(value) > 200:
            raise ApiError("高亮值无效")
        if not isinstance(color, str) or not _HEX_COLOR.fullmatch(color):
            raise ApiError("高亮颜色必须是 #RRGGBB")
        dynamic_max_values = {
            "__max__",
            "max",
            "maximum",
            "最大",
            "最大值",
            "最高",
            "最高值",
            "最高月份",
            "最大月份",
            "峰值",
            "峰值月份",
        }
        if (
            value.strip().casefold() in dynamic_max_values
            and isinstance(output.get("labels"), list)
            and isinstance(output.get("values"), list)
            and output["labels"]
            and len(output["labels"]) == len(output["values"])
        ):
            valid_positions = [
                index
                for index, item in enumerate(output["values"])
                if isinstance(item, (int, float)) and not isinstance(item, bool) and math.isfinite(float(item))
            ]
            if valid_positions:
                max_position = max(valid_positions, key=lambda index: float(output["values"][index]))
                value = str(output["labels"][max_position])
        output["highlight"] = {"field": field, "value": value, "color": color.upper()}

    if (
        sort_mode in {"asc", "desc"}
        and isinstance(output.get("labels"), list)
        and isinstance(output.get("values"), list)
    ):
        pairs = list(zip(output["labels"], output["values"]))
        pairs.sort(key=lambda item: float(item[1]), reverse=sort_mode == "desc")
        output["labels"] = [item[0] for item in pairs]
        output["values"] = [item[1] for item in pairs]
    return output


def _demo_sales_frame() -> pd.DataFrame:
    """Create deterministic, entirely fictional sales data for local demos."""

    rng = random.Random(20260821)
    customers = [f"虚构客户{index:03d}" for index in range(1, 46)]
    regions = ["华东", "华南", "华北", "西南", "华中"]
    channels = ["线上商城", "线下门店", "企业直销", "经销商"]
    products = {
        "轻享办公椅": (899.0, 520.0),
        "云影显示器": (1599.0, 1050.0),
        "星河键鼠套装": (299.0, 148.0),
        "远山升降桌": (2399.0, 1510.0),
        "清风护眼灯": (459.0, 235.0),
        "拾光收纳柜": (1099.0, 680.0),
    }
    start = datetime(2025, 1, 1)
    rows: list[dict[str, Any]] = []
    for index in range(360):
        product = rng.choice(list(products))
        unit_price, unit_cost = products[product]
        quantity = rng.choices([1, 2, 3, 4, 5, 6], weights=[38, 27, 16, 10, 6, 3])[0]
        discount = rng.choice([0.82, 0.88, 0.92, 0.95, 1.0])
        order_amount = round(unit_price * quantity * discount, 2)
        cost = round(unit_cost * quantity, 2)
        rows.append(
            {
                "订单编号": f"DEMO-{index + 1:05d}",
                "日期": start + timedelta(days=rng.randrange(0, 545)),
                "客户": rng.choice(customers),
                "地区": rng.choices(regions, weights=[30, 24, 19, 14, 13])[0],
                "渠道": rng.choices(channels, weights=[38, 25, 21, 16])[0],
                "产品": product,
                "订单金额": order_amount,
                "数量": quantity,
                "成本": cost,
            }
        )
    frame = pd.DataFrame(rows)
    # Deliberately add a few quality issues so every analysis feature is visible.
    frame.loc[17, ["订单金额", "数量", "成本"]] = [128000.0, 80, 42600.0]
    frame.loc[119, ["订单金额", "数量", "成本"]] = [96000.0, 65, 33800.0]
    frame.loc[5, "地区"] = pd.NA
    frame.loc[73, "客户"] = pd.NA
    frame.loc[144, "订单金额"] = pd.NA
    frame.loc[201, "渠道"] = " "
    frame = pd.concat([frame, frame.iloc[[10, 28]].copy(deep=True)], ignore_index=True)
    return frame


def _analysis_export_tables(frame: pd.DataFrame) -> dict[str, pd.DataFrame]:
    payload = _analysis_payload(frame)
    statistics = _column_metadata(frame)
    safe_statistics = _privacy_safe_statistics(statistics)
    numeric_columns, _, measure, date_column, dimension = _analysis_columns(frame, statistics)

    overview_rows = [
        {"项目": "质量得分", "内容": payload["quality"]["score"]},
        {"项目": "质量等级", "内容": payload["quality"]["grade"]},
        {"项目": "数据行数", "内容": payload["overview"]["rows"]},
        {"项目": "字段数", "内容": payload["overview"]["columns"]},
        {"项目": "空值单元格", "内容": payload["overview"]["missing_cells"]},
        {"项目": "缺失率", "内容": payload["overview"]["missing_rate"]},
        {"项目": "重复记录行", "内容": payload["overview"]["duplicate_rows"]},
        {"项目": "重复率", "内容": payload["overview"]["duplicate_rate"]},
        {"项目": "数值字段数", "内容": payload["overview"]["numeric_columns"]},
        {"项目": "日期字段数", "内容": payload["overview"]["date_columns"]},
        {"项目": "内存占用 MB", "内容": payload["overview"]["memory_mb"]},
        {"项目": "隐私说明", "内容": "报告仅在本机生成；描述统计不展示文本众数、标识符最值或原始样本。"},
    ]
    issue_rows = [
        {
            "严重程度": item["severity"],
            "问题": item["title"],
            "说明": item["detail"],
            "建议": item["recommendation"],
            "数量": item["count"],
            "比例": item["ratio"],
        }
        for item in payload["issues"]
    ] or [
        {
            "严重程度": "info",
            "问题": "未发现明显问题",
            "说明": "当前规则扫描通过",
            "建议": "仍建议结合业务规则复核",
            "数量": 0,
            "比例": 0.0,
        }
    ]
    insight_rows = payload["insights"] or [{"title": "暂无自动洞察", "detail": "当前字段不足。"}]

    correlation = correlation_matrix(frame, columns=numeric_columns[:12])
    if correlation.empty:
        correlation_sheet = pd.DataFrame([{"说明": "至少需要两个可用数值字段。"}])
    else:
        correlation_sheet = correlation.copy(deep=True)
        correlation_sheet.insert(0, "字段", [str(value) for value in correlation_sheet.index])
        correlation_sheet = correlation_sheet.reset_index(drop=True)

    if numeric_columns:
        outlier_result = detect_outliers(frame, columns=numeric_columns[:20], method="iqr")
        anomaly_sheet = outlier_result.outliers.copy(deep=True)
        if anomaly_sheet.empty:
            anomaly_sheet = pd.DataFrame([{"说明": "按 IQR 方法未检测到异常数值。"}])
    else:
        anomaly_sheet = pd.DataFrame([{"说明": "没有可用于异常检测的数值字段。"}])

    if date_column and measure:
        try:
            trend_sheet = aggregate_trend(
                frame,
                date_column=date_column,
                value_columns=measure,
                frequency="month",
                aggregation="sum",
            ).data
            if trend_sheet.empty:
                trend_sheet = pd.DataFrame([{"说明": "没有足够的有效日期和数值生成趋势。"}])
        except (TypeError, ValueError) as exc:
            trend_sheet = pd.DataFrame([{"说明": f"趋势分析未生成：{exc}"}])
    else:
        trend_sheet = pd.DataFrame([{"说明": "需要至少一个日期字段和一个数值字段。"}])

    if dimension and measure:
        try:
            contribution_sheet = category_contribution(
                frame,
                category_columns=dimension,
                value_column=measure,
                aggregation="sum",
                top_n=50,
                include_other=True,
            ).data
        except (TypeError, ValueError) as exc:
            contribution_sheet = pd.DataFrame([{"说明": f"分类贡献未生成：{exc}"}])
    else:
        contribution_sheet = pd.DataFrame([{"说明": "需要一个非敏感分类字段和一个数值字段。"}])

    tables = {
        "原数据": frame.copy(deep=True),
        "质量概览": pd.DataFrame(overview_rows),
        "质量问题": pd.DataFrame(issue_rows),
        "自动洞察": pd.DataFrame(insight_rows).rename(columns={"title": "发现", "detail": "说明"}),
        "描述统计": safe_statistics,
        "相关性": correlation_sheet,
        "异常明细": anomaly_sheet,
        "趋势分析": trend_sheet,
        "分类贡献": contribution_sheet,
    }
    return {name: _escape_spreadsheet_formulas(table) for name, table in tables.items()}


def _load_advanced_asset(path: Path) -> dict[str, pd.DataFrame]:
    """Convert one approved non-Excel asset into safe local table context."""

    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return {f"{path.name}::{name}": frame for name, frame in extract_pdf_tables(path).items()}
    if suffix in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}:
        return {f"{path.name}::OCR文字": extract_image_text(path)}
    if suffix in {".db", ".sqlite", ".sqlite3"}:
        connection = sqlite3.connect(f"file:{path.resolve().as_posix()}?mode=ro", uri=True, timeout=10)
        try:
            rows = connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            ).fetchall()
            result: dict[str, pd.DataFrame] = {}
            for (table_name,) in rows[:100]:
                escaped = str(table_name).replace('"', '""')
                columns = [item[1] for item in connection.execute(f'PRAGMA table_info("{escaped}")').fetchall()]
                result[f"{path.name}::{table_name}"] = pd.DataFrame(columns=columns)
            if not result:
                raise ValueError("SQLite 数据库没有可查询的业务表")
            return result
        finally:
            connection.close()
    if suffix == ".parquet":
        return {f"{path.name}::Parquet": query_files([path], "SELECT * FROM input_1", max_rows=MAX_ROWS_PER_TABLE)}
    raise ValueError(f"不支持的高级文件类型：{suffix}")


from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer  # noqa: E402


class ToolboxHandler(BaseHTTPRequestHandler):
    server_version = f"BiaogeKuaichu/{APP_VERSION}"

    def log_message(self, format_string: str, *args: Any) -> None:
        # Do not log customer filenames, URL parameters, or payloads.
        message = format_string % args
        print(f"[{datetime.now():%H:%M:%S}] {self.client_address[0]} {message}")

    def _json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(_json_value(payload), ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def _error(self, error: Exception) -> None:
        if isinstance(error, ApiError):
            status = error.status
            message = str(error)
        elif isinstance(error, KeyError):
            status = 400
            message = str(error).strip("'")
        else:
            status = 400
            message = str(error) or "处理失败，请检查文件与参数"
        self._json({"error": message}, status=status)

    def _read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length > 2 * 1024 * 1024:
            raise ApiError("请求参数过大", 413)
        raw = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8") or "{}")
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ApiError("请求参数不是有效 JSON") from exc
        if not isinstance(payload, dict):
            raise ApiError("请求参数必须是对象")
        return payload

    def _serve_static(self, path: Path) -> None:
        try:
            resolved = path.resolve(strict=True)
        except FileNotFoundError:
            self.send_error(404)
            return
        if WEB_DIR.resolve() not in resolved.parents and resolved != WEB_DIR.resolve():
            self.send_error(403)
            return
        content = resolved.read_bytes()
        mime, _ = mimetypes.guess_type(resolved.name)
        self.send_response(200)
        self.send_header("Content-Type", f"{mime or 'application/octet-stream'}; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:  # noqa: N802
        path = unquote(urlparse(self.path).path)
        try:
            SESSION.bind(self.headers.get("X-Task-ID", "").strip() or None)
            if path == "/health":
                self._json(
                    {
                        "status": "ok",
                        "app": "表格快处 Pro",
                        "version": APP_VERSION,
                        "storage": "local-persistent-per-task",
                        "retention_days": TASK_RETENTION_DAYS,
                        "port": int(self.server.server_address[1]),
                    }
                )
            elif path == "/api/state":
                with SESSION.lock:
                    self._json(SESSION.state_payload())
            elif path == "/api/reviews":
                with SESSION.lock:
                    self._json(SESSION.review_payload())
            elif path == "/api/recipes":
                self._json({"recipes": _list_saved_recipes()})
            elif path == "/api/tasks":
                self._json({"tasks": TASK_REPOSITORY.list_tasks(), "retention_days": TASK_RETENTION_DAYS})
            elif path == "/api/contracts":
                self._json({"contracts": CONTRACT_STORE.list()})
            elif path == "/api/lineage":
                self._json(LINEAGE_STORE.evidence(SESSION.task_id))
            elif path == "/api/jobs":
                self._json({"jobs": [item.to_dict() for item in JOB_ENGINE.list(SESSION.task_id)]})
            elif path == "/api/tool-manifest":
                self._json({"tools": TOOL_REGISTRY.manifest(), "engines": available_engines()})
            elif path == "/api/conversation":
                self._json(ConversationStore(SESSION.task_dir / "conversation.json").context())
            elif path == "/api/database-connections":
                self._json({"profiles": [item.to_dict() for item in DATABASE_CONNECTIONS.list()]})
            elif path == "/api/ai-evaluations":
                self._json(
                    {
                        "scenarios": [item.to_dict() for item in AI_SCENARIO_STORE.list()],
                        "traces": AI_TRACE_STORE.list(SESSION.task_id, limit=50),
                    }
                )
            elif path == "/api/schedules":
                self._json({"schedules": [item.to_dict() for item in SCHEDULER.list()]})
            elif path == "/api/system-capabilities":
                self._json(
                    {
                        "duckdb": duckdb_available(),
                        "documents": document_capabilities(),
                        "secure_vault": SECRET_STORE.available,
                        "persistent_tasks": True,
                        "delivery_reopen_verification": True,
                        "workbook_fidelity_mode": True,
                        "engines": available_engines(),
                        "persistent_jobs": True,
                        "data_contracts": True,
                        "lineage_evidence": True,
                        "database_connection_center": True,
                        "tool_registry": len(TOOL_REGISTRY.manifest()),
                    }
                )
            elif path == "/api/chart-history":
                with SESSION.lock:
                    self._json({"history": SESSION.chart_history, "can_redo": bool(SESSION.chart_redo_stack)})
            elif path == "/api/ai-capabilities":
                self._json(self._ai_capabilities())
            elif path == "/api/config-status":
                config = _project_ai_config()
                try:
                    power_bi = PowerBIConfig.from_environment(env_file=PROJECT_ENV_PATH).public_status()
                except PowerBIAutomationError as exc:
                    power_bi = {"configured": False, "missing": [], "workspace_id": None, "error": str(exc)}
                self._json(
                    {
                        "configured": config["configured"],
                        "model": config["model"],
                        "source": config["source"],
                        "message": "项目级 AI 已配置"
                        if config["configured"]
                        else "请运行安全配置脚本保存 DeepSeek API Key",
                        "power_bi": power_bi,
                    }
                )
            elif path.startswith("/download/"):
                self._download(path.rsplit("/", 1)[-1])
            elif path == "/" or path == "/index.html":
                self._serve_static(WEB_DIR / "unified.html")
            elif path.startswith("/static/"):
                relative = path[len("/static/") :]
                self._serve_static(WEB_DIR / relative)
            else:
                self.send_error(404)
        except Exception as exc:  # pragma: no cover - last-resort HTTP boundary
            self._error(exc)

    def _download(self, token: str) -> None:
        with SESSION.lock:
            path = SESSION.downloads.get(token)
        if not path or not path.exists() or SESSION.output_dir.resolve() not in path.parents:
            raise ApiError("下载文件不存在或已清空", 404)
        content = path.read_bytes()
        quoted = path.name.encode("utf-8").hex()
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(len(content)))
        self.send_header(
            "Content-Disposition",
            f"attachment; filename=download{path.suffix}; filename*=UTF-8''{_urlquote(path.name)}",
        )
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Download-Name-Hex", quoted)
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        try:
            SESSION.bind(self.headers.get("X-Task-ID", "").strip() or None)
            if path == "/api/upload":
                self._upload()
                return
            payload = self._read_json()
            handlers = {
                "/api/task": self._task,
                "/api/select": self._select,
                "/api/demo": self._demo,
                "/api/analysis": self._analysis,
                "/api/chart": self._chart,
                "/api/anomalies": self._anomalies,
                "/api/pivot": self._pivot,
                "/api/rfm": self._rfm,
                "/api/analysis-export": self._analysis_export,
                "/api/fuzzy-cluster": self._fuzzy_cluster,
                "/api/fuzzy-lookup": self._fuzzy_lookup,
                "/api/recipe-save": self._recipe_save,
                "/api/recipe-run": self._recipe_run,
                "/api/validate": self._validate,
                "/api/reconcile-advanced": self._reconcile_advanced,
                "/api/ai-diagnose": self._ai_diagnose,
                "/api/ai-plan": self._ai_plan,
                "/api/ai-chart-plan": self._ai_chart_plan,
                "/api/ai-unified": self._ai_unified,
                "/api/ai-execute": self._ai_execute,
                "/api/ai-engineering": self._ai_engineering,
                "/api/clean": self._clean,
                "/api/columns": self._columns,
                "/api/replace": self._replace,
                "/api/concat": self._concat,
                "/api/join": self._join,
                "/api/compare": self._compare,
                "/api/summary": self._summary,
                "/api/split": self._split,
                "/api/mask": self._mask,
                "/api/export": self._export,
                "/api/review-decision": self._review_decision,
                "/api/undo": self._undo,
                "/api/redo": self._redo,
                "/api/reset": self._reset,
                "/api/task-open": self._task_open,
                "/api/task-delete": self._task_delete,
                "/api/task-purge": self._task_purge,
                "/api/order-quote": self._order_quote,
                "/api/database-query": self._database_query,
                "/api/vba-bundle": self._vba_bundle,
                "/api/schedule": self._schedule,
                "/api/chart-history": self._chart_history,
                "/api/contract-generate": self._contract_generate,
                "/api/contract-validate": self._contract_validate,
                "/api/lineage-export": self._lineage_export,
                "/api/job-submit": self._job_submit,
                "/api/job-cancel": self._job_cancel,
                "/api/job-retry": self._job_retry,
                "/api/database-profile-save": self._database_profile_save,
                "/api/database-profile-delete": self._database_profile_delete,
                "/api/database-profile-test": self._database_profile_test,
                "/api/database-profile-schema": self._database_profile_schema,
                "/api/database-profile-query": self._database_profile_query,
                "/api/conversation-clear": self._conversation_clear,
                "/api/ai-evaluation-run": self._ai_evaluation_run,
            }
            handler = handlers.get(path)
            if not handler:
                raise ApiError("接口不存在", 404)
            result = handler(payload)
            self._json(result or {"message": "操作完成"})
        except Exception as exc:
            self._error(exc)

    def _upload(self) -> None:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length <= 0:
            raise ApiError("没有收到上传文件")
        if length > MAX_REQUEST_BYTES:
            raise ApiError("本次上传总大小超过 180 MB 限制", 413)
        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data"):
            raise ApiError("上传格式错误")
        raw_body = self.rfile.read(length)
        message = BytesParser(policy=email_policy.default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + raw_body
        )
        fields: list[tuple[str, bytes]] = []
        requested_name = ""
        for part in message.iter_parts():
            field_name = part.get_param("name", header="content-disposition")
            if field_name == "task_name":
                requested_name = _decode_multipart_text(part).strip()[:60]
                continue
            if field_name != "files":
                continue
            filename = str(part.get_filename() or "")
            payload_bytes = part.get_payload(decode=True)
            if filename and isinstance(payload_bytes, bytes):
                fields.append((filename, payload_bytes))
        if not fields:
            raise ApiError("请选择至少一个文件")
        staged_files: list[tuple[str, bytes]] = []
        with tempfile.TemporaryDirectory(prefix="biaoge_upload_stage_") as stage_name:
            stage_dir = Path(stage_name)
            staged_paths: list[Path] = []
            staged_warnings: list[str] = []
            used_names: set[str] = set()
            for raw_filename, payload_bytes in fields:
                filename = _safe_filename(raw_filename)
                suffix = Path(filename).suffix.lower()
                if suffix not in ALLOWED_SUFFIXES:
                    raise ApiError(f"文件“{filename}”不支持；首版仅支持 .xlsx 和 .csv")
                data = payload_bytes
                if len(data) > MAX_FILE_BYTES:
                    raise ApiError(f"文件“{filename}”超过 50 MB 限制", 413)
                candidate_name = filename
                counter = 2
                while candidate_name.casefold() in used_names:
                    candidate_name = f"{Path(filename).stem}_{counter}{suffix}"
                    counter += 1
                used_names.add(candidate_name.casefold())
                candidate = stage_dir / candidate_name
                candidate.write_bytes(data)
                if suffix in {".xlsx", ".xlsm"}:
                    staged_warnings.extend(_audit_xlsx_structure(candidate))
                staged_paths.append(candidate)
                staged_files.append((candidate_name, data))

            # Historical toolbox reports are deliverables, not raw business
            # facts.  Exclude them before parsing so an earlier dashboard can
            # never be selected as the next task's primary dataset.
            generated_outputs: list[tuple[Path, str, tuple[str, ...]]] = []
            accepted_paths: list[Path] = []
            for candidate in staged_paths:
                if candidate.suffix.lower() in {".xlsx", ".xlsm"}:
                    assessment = detect_generated_workbook(candidate)
                    if assessment.generated:
                        generated_outputs.append(
                            (candidate, assessment.report_kind, assessment.matched_sheets)
                        )
                        continue
                accepted_paths.append(candidate)
            if generated_outputs and not accepted_paths:
                names = "、".join(item[0].name for item in generated_outputs)
                raise ApiError(
                    f"已阻止历史输出文件回流：{names}。请上传本次客户的原始业务数据；"
                    "系统生成的经营报告默认不能再次作为原始输入。",
                    422,
                )
            if generated_outputs:
                blocked_names = {item[0].name.casefold() for item in generated_outputs}
                staged_paths = accepted_paths
                staged_files = [item for item in staged_files if item[0].casefold() not in blocked_names]
                for path, report_kind, matches in generated_outputs:
                    staged_warnings.append(
                        f"已排除疑似系统历史输出“{path.name}”（{report_kind}；匹配工作表："
                        + "、".join(matches[:6])
                        + "），未纳入本次分析。"
                    )
            try:
                loaded: dict[str, pd.DataFrame] = {}
                tabular_paths = [path for path in staged_paths if path.suffix.lower() in TABULAR_SUFFIXES]
                if tabular_paths:
                    loaded.update(load_tables(tabular_paths))
                for asset_path in staged_paths:
                    if asset_path.suffix.lower() not in TABULAR_SUFFIXES:
                        loaded.update(_load_advanced_asset(asset_path))
            except ImportError as exc:
                raise ApiError("缺少读取该文件所需组件，请运行健康检查并安装可选组件") from exc
            except (RuntimeError, LargeDataUnavailable, ValueError) as exc:
                raise ApiError(f"高级文件导入失败：{exc}") from exc
            total_rows = sum(len(frame) for frame in loaded.values())
            if any(len(frame) > MAX_ROWS_PER_TABLE for frame in loaded.values()):
                raise ApiError("某个工作表超过 300,000 行安全上限，请先拆分后再导入")
            if total_rows > MAX_ROWS_PER_TABLE * 3:
                raise ApiError("本次任务数据量过大，请减少文件或分批处理")

        if not requested_name or requested_name in {"AI 一句话 Excel 任务", "新建数据处理任务"}:
            requested_name = Path(staged_files[0][0]).stem[:60]
        with SESSION.lock:
            # One upload batch is one isolated task. Related files should be
            # selected together in the same upload action.
            SESSION.reset()
            SESSION.task_name = requested_name
            for filename, data in staged_files:
                (SESSION.upload_dir / filename).write_bytes(data)
            SESSION.import_warnings.extend(staged_warnings)
            for name, frame in loaded.items():
                SESSION.add_table(name, frame, source="导入文件", original=True)
            SESSION.file_names.update(filename for filename, _ in staged_files)
            SESSION.persist()
            task_id = SESSION.task_id
            task_name = SESSION.task_name
        self._json(
            {
                "message": f"已新建独立任务“{task_name}”，导入 {len(staged_files)} 个文件、{len(loaded)} 张数据表",
                "task_id": task_id,
                "task_name": task_name,
                "new_task": True,
            }
        )

    def _task(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            SESSION.task_name = str(payload.get("name") or "新建数据处理任务").strip()[:60]
            SESSION.persist()
        return {"message": "任务名称已更新"}

    def _select(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            entry = SESSION.get(payload.get("table", ""))
            SESSION.active_table = entry.id
        return {"message": f"已切换到“{entry.name}”"}

    def _ai_capabilities(self) -> dict[str, Any]:
        operations = [
            {
                "id": operation,
                "label": _AI_OPERATION_LABELS.get(operation, operation),
            }
            for operation in sorted(ALLOWED_AGENT_OPERATIONS)
        ]
        return {
            "available": True,
            "version": APP_VERSION,
            "provider": "DeepSeek",
            "models": sorted(SUPPORTED_DEEPSEEK_MODELS),
            "default_model": "deepseek-v4-flash",
            "operations": operations,
            "workflow": "自然语言规划、本地安全校验、自动执行、交付前自动验收",
            "connection_diagnosis": "可单独检测网络、密钥、余额和模型，不发送表格数据",
            "engineering_orders": [
                "VBA 安全代码包、使用说明与静态风险扫描",
                "Power BI 星型模型、DAX、Power Query、PBIP 工程包与可选发布",
                "SQLite/ODBC 外部数据库只读查询与结果交付",
                "业务判断决策矩阵、例外清单与可审计结果包",
                "PDF 结构化表格提取、图片中英文 OCR、CSV/Parquet 大数据汇总",
            ],
            "finance_orders": [
                "应收账款账龄、逾期结构与客户账龄汇总",
                "预算对实际、差异额、差异率与有利/不利判断",
                "月度现金流入、流出、净额、累计额与分类现金流",
                "盈利、偿债、营运与现金流财务比率",
                "会计凭证借贷平衡、异常分录与不平衡凭证审计",
            ],
            "privacy": {
                "sent_to_provider": ["用户需求", "所选表名", "字段名", "数据类型", "行列规模"],
                "not_sent_by_default": ["单元格原值", "文件内容", "预览行", "处理结果"],
                "api_key_storage": "优先使用 Windows DPAPI 本机加密保险箱；不写入任务、配方或日志",
                "local_execution": True,
            },
            "unsupported": [
                "任意 Shell、危险 VBA、写库 SQL 或未登记的外部网络操作",
                "伪造票据、流水、证明或业务数据",
                "绕过 Microsoft 登录、MFA、租户授权、宏信任中心或数据库权限",
                "保证第三方私有格式、损坏工作簿、切片器缓存和所有外部插件 100% 无损",
            ],
            "limits": {
                "prompt_characters": AI_MAX_PROMPT_CHARS,
                "selected_tables": AI_MAX_SELECTED_TABLES,
                "plan_ttl_seconds": AI_PLAN_TTL_SECONDS,
                "max_output_tables": AI_MAX_OUTPUT_TABLES,
            },
        }

    def _ai_engineering(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"category", "prompt", "api_key", "model", "table_ids"})
        category = str(payload.get("category") or "").strip()
        if category not in ENGINEERING_CATEGORIES:
            raise ApiError("请选择 VBA、Power BI、外部数据库或业务决策类别")
        prompt = payload.get("prompt")
        if not isinstance(prompt, str) or len(prompt.strip()) < 20:
            raise ApiError("请至少用 20 个字符说明运行环境、输入、规则和期望交付物")
        prompt = prompt.strip()
        if len(prompt) > AI_MAX_PROMPT_CHARS:
            raise ApiError(f"工程需求不能超过 {AI_MAX_PROMPT_CHARS:,} 个字符")
        if AI_SECRET_PATTERN.search(prompt):
            raise ApiError("工程需求中检测到疑似密钥；请删除后重试")
        raw_key = payload.get("api_key")
        if not isinstance(raw_key, str) or not raw_key.strip():
            raise ApiError("请填写 DeepSeek API Key")
        api_key = raw_key.strip()
        if len(api_key) > 512 or any(character.isspace() for character in api_key):
            raise ApiError("DeepSeek API Key 格式无效")
        model = str(payload.get("model") or "deepseek-v4-pro").strip()
        if model not in SUPPORTED_DEEPSEEK_MODELS:
            raise ApiError("模型不可用")
        table_ids = _normalise_ai_table_scope(payload.get("table_ids"), allow_empty=True)
        with SESSION.lock:
            entries = [SESSION.get(table_id) for table_id in table_ids]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in entries}
            display_names = {entry.id: entry.name for entry in entries}
        if not tables:
            tables = {"order_context": pd.DataFrame({"待客户提供样例字段": pd.Series(dtype="string")})}
            display_names = {"order_context": "尚未提供样例表"}
        catalog = build_table_catalog(tables, display_names=display_names)
        used_deterministic_fallback = False
        try:
            brief = DeepSeekClient(api_key, model=model, timeout_seconds=90).create_engineering_brief(
                category, prompt, catalog
            )
        except DeepSeekAPIError as exc:
            safe_message = AI_SECRET_PATTERN.sub("[API Key 已隐藏]", str(exc).replace(api_key, "[API Key 已隐藏]"))
            raise ApiError(safe_message or "高级工程方案生成失败", 502) from None
        except (PlanValidationError, TypeError, ValueError) as exc:
            # Power BI execution does not trust or execute the model's JSON.
            # If DeepSeek varies a descriptive field shape, use a stable local
            # brief and continue with the deterministic compiler instead of
            # repeatedly rejecting the customer's otherwise valid order.
            if category == "power_bi":
                brief = fallback_power_bi_brief(prompt)
                used_deterministic_fallback = True
            else:
                raise ApiError(f"高级工程方案未通过本地安全校验：{exc}", 422) from None
        finally:
            api_key = ""
            raw_key = ""

        if category == "power_bi" and not used_deterministic_fallback:
            # Preserve the model's interpretation of the request, but replace
            # advisory artifacts with the exact capabilities the local
            # compiler actually delivers.  This prevents stale AI prose such
            # as “需手动发布” from contradicting the unattended publisher.
            interpreted_request = str(brief.get("normalized_request") or prompt)
            brief = fallback_power_bi_brief(prompt)
            brief["normalized_request"] = interpreted_request[:800]

        automation: dict[str, Any] | None = None
        if category == "power_bi":
            if not entries:
                raise ApiError("全自动 Power BI 任务需要先加载演示数据或上传一张销售明细表", 422)
            try:
                bundle = build_power_bi_bundle(
                    entries[0].frame,
                    SESSION.output_dir,
                    task_id=SESSION.task_id,
                    source_name=entries[0].name,
                    engineering_brief=brief,
                )
                try:
                    power_bi_config = PowerBIConfig.from_environment(env_file=PROJECT_ENV_PATH)
                    publish_result = publish_bundle_if_configured(
                        bundle,
                        config=power_bi_config,
                        display_name=f"销售分析 {SESSION.task_id}",
                    )
                except PowerBIAutomationError as exc:
                    publish_result = {
                        "status": "configuration_invalid",
                        "published": False,
                        "message": str(exc),
                        "missing": [],
                    }
                with SESSION.lock:
                    download_url = SESSION.register_download(Path(bundle["zip_path"]))
                automation = {
                    "status": publish_result["status"],
                    "published": publish_result.get("published", False),
                    "message": publish_result["message"],
                    "missing": publish_result.get("missing", []),
                    "download_url": download_url,
                    "download_name": Path(bundle["zip_path"]).name,
                    "validation": bundle["validation"],
                    "model": {
                        "fact_rows": bundle["model_spec"]["tables"]["FactSales"]["rows"],
                        "dimension_tables": 5,
                        "measures": len(bundle["model_spec"]["measures"]),
                        "pages": len(bundle["report_spec"]["pages"]),
                        "visuals": sum(page["visual_count"] for page in bundle["report_spec"]["pages"]),
                    },
                    **{
                        key: publish_result[key]
                        for key in ("workspace_id", "semantic_model_id", "report_id", "report_url", "verified")
                        if key in publish_result
                    },
                }
            except PowerBIAutomationError as exc:
                raise ApiError(f"Power BI 自动化编译失败：{exc}", 422) from None

        elif category == "vba":
            artifacts = list(brief.get("artifacts") or [])
            vba_artifact = next(
                (
                    item
                    for item in artifacts
                    if str(item.get("language") or "").casefold() in {"vba", "vb", "visual basic"}
                    and str(item.get("content") or "").strip()
                ),
                None,
            )
            if vba_artifact is not None:
                destination = SESSION.output_dir / f"{SESSION.task_id}_VBA安全交付包.zip"
                try:
                    result = build_vba_bundle(
                        str(vba_artifact["content"]),
                        destination,
                        module_name="BiaogeAutomation",
                    )
                except (OSError, ValueError) as exc:
                    raise ApiError(f"VBA自动化包未通过安全编译：{exc}", 422) from None
                with SESSION.lock:
                    download_url = SESSION.register_download(destination)
                    SESSION.record(
                        "生成VBA安全交付包",
                        "AI代码已通过本地危险指令扫描并生成哈希与回滚清单",
                        inputs=[entry.name for entry in entries],
                        produced=[],
                        before_rows=None,
                        after_rows=None,
                    )
                automation = {
                    "status": "package_built_and_scanned",
                    "published": False,
                    "message": result.message,
                    "download_url": download_url,
                    "checks": list(result.checks),
                    "warnings": list(result.warnings),
                }

        elif category == "database":
            database_files = [
                path for path in SESSION.upload_dir.iterdir() if path.suffix.lower() in {".db", ".sqlite", ".sqlite3"}
            ]
            sql_artifact = next(
                (
                    item
                    for item in list(brief.get("artifacts") or [])
                    if str(item.get("language") or "").casefold() in {"sql", "sqlite"}
                    and str(item.get("content") or "").strip()
                ),
                None,
            )
            if len(database_files) == 1 and sql_artifact is not None:
                try:
                    result_frame = query_sqlite_read_only(database_files[0], str(sql_artifact["content"]))
                except (OSError, ValueError, sqlite3.Error) as exc:
                    raise ApiError(f"AI只读SQL执行失败：{exc}", 422) from None
                with SESSION.lock:
                    result_id = SESSION.add_table("数据库AI查询结果", result_frame, source="AI只读SQL")
                    SESSION.record(
                        "执行AI数据库只读查询",
                        "SQL已通过只读语法校验并在SQLite只读连接中执行",
                        inputs=[database_files[0].name],
                        produced=[result_id],
                        before_rows=None,
                        after_rows=len(result_frame),
                    )
                automation = {
                    "status": "query_executed_read_only",
                    "published": False,
                    "message": f"AI只读查询已执行，生成 {len(result_frame):,} 行结果",
                    "table_id": result_id,
                    "checks": ["单条SELECT/WITH校验通过", "数据库以只读模式打开", "结果行数上限已检查"],
                }

        elif category == "business_decision":
            decision_tables = {
                "决策范围": pd.DataFrame([{"项目": "范围", "内容": brief.get("scope", "")}]),
                "实施步骤": pd.DataFrame({"实施步骤": list(brief.get("implementation_steps") or [])}),
                "风险清单": pd.DataFrame({"风险": list(brief.get("risks") or [])}),
                "人工审批点": pd.DataFrame({"人工审批点": list(brief.get("human_approval_points") or [])}),
                "验收清单": pd.DataFrame({"验收项": list(brief.get("test_checklist") or [])}),
            }
            destination = SESSION.output_dir / f"{SESSION.task_id}_业务决策与审批包.xlsx"
            export_tables(decision_tables, destination, include_log=False, overwrite=True)
            acceptance = verify_delivery(destination, decision_tables)
            if acceptance.status != "passed":
                raise ApiError("业务决策包导出后验收失败", 500)
            with SESSION.lock:
                download_url = SESSION.register_download(destination)
            automation = {
                "status": "decision_package_built",
                "published": False,
                "message": "决策矩阵、风险、审批点和验收清单已生成",
                "download_url": download_url,
                "validation": acceptance.to_dict(),
            }

        if automation:
            message = (
                "Power BI 语义模型和报表已自动生成、校验、发布并回读验证。"
                if category == "power_bi" and automation["published"]
                else automation.get("message", "高级工程自动化已完成")
            )
            if category == "power_bi" and not automation["published"]:
                message = "Power BI 自包含交付包已自动生成并通过本地校验；微软发布凭据未就绪，未对外发布。"
            if category == "power_bi" and used_deterministic_fallback:
                message += " DeepSeek 描述结构不兼容时已自动切换本地稳定方案。"
            return {
                "message": message,
                "brief": brief,
                "privacy": "仅发送需求与所选表结构给 DeepSeek；Power BI 数据和发布载荷均由本机确定性程序生成。",
                "execution": (
                    ("published" if automation.get("published") else "package_built_and_validated")
                    if category == "power_bi"
                    else str(automation.get("status") or "package_built_and_validated")
                ),
                "automation": automation,
            }
        return {
            "message": "高级工程订单方案已生成；代码与查询仅供人工审查，不会自动运行。",
            "brief": brief,
            "privacy": "仅发送需求与所选表结构，不发送单元格原值、数据库凭据或文件内容。",
            "execution": "not_executed",
        }

    def _ai_diagnose(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"api_key", "model"})
        raw_key = payload.get("api_key")
        config = _project_ai_config()
        if raw_key is not None and not isinstance(raw_key, str):
            raise ApiError("DeepSeek API Key 格式无效")
        api_key = raw_key.strip() if isinstance(raw_key, str) else ""
        if not api_key:
            if not config["configured"]:
                raise ApiError("尚未配置 DeepSeek API Key；请先运行安全配置脚本")
            api_key = config["api_key"]
        if len(api_key) > 512 or any(character.isspace() for character in api_key):
            raise ApiError("DeepSeek API Key 格式无效")
        model = str(payload.get("model") or config.get("model") or "deepseek-v4-flash").strip()
        if model not in SUPPORTED_DEEPSEEK_MODELS:
            raise ApiError("模型不可用；请选择 deepseek-v4-flash 或 deepseek-v4-pro")
        try:
            result = DeepSeekClient(api_key, model=model, timeout_seconds=20).check_connection()
        except DeepSeekAPIError as exc:
            safe_message = str(exc).replace(api_key, "[API Key 已隐藏]")
            safe_message = AI_SECRET_PATTERN.sub("[API Key 已隐藏]", safe_message)
            raise ApiError(safe_message or "DeepSeek 连接检测失败", 502) from None
        finally:
            api_key = ""
            raw_key = ""
        return result

    def _ai_plan(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"prompt", "api_key", "model", "table_ids"})
        prompt = payload.get("prompt")
        if not isinstance(prompt, str):
            raise ApiError("处理需求必须是文本")
        prompt = prompt.strip()
        if len(prompt) < 12:
            raise ApiError("请把需求写得更具体，至少填写 12 个字符")
        if len(prompt) > AI_MAX_PROMPT_CHARS:
            raise ApiError(f"处理需求不能超过 {AI_MAX_PROMPT_CHARS:,} 个字符")
        if AI_SECRET_PATTERN.search(prompt):
            raise ApiError("需求描述中检测到疑似 API Key；请从需求文字中删除，只填写在密钥框")

        raw_key = payload.get("api_key")
        if not isinstance(raw_key, str) or not raw_key.strip():
            raise ApiError("请填写 DeepSeek API Key")
        api_key = raw_key.strip()
        if len(api_key) > 512 or any(character.isspace() for character in api_key):
            raise ApiError("DeepSeek API Key 格式无效")

        model = str(payload.get("model") or "deepseek-v4-flash").strip()
        if model not in SUPPORTED_DEEPSEEK_MODELS:
            raise ApiError("模型不可用；请选择 deepseek-v4-flash 或 deepseek-v4-pro")
        raw_table_ids = payload.get("table_ids")
        table_ids = _normalise_ai_table_scope(raw_table_ids, allow_empty=False)
        if len(table_ids) != len(raw_table_ids):
            raise ApiError("数据表范围不能包含重复项")

        with SESSION.lock:
            entries = [SESSION.get(table_id) for table_id in table_ids]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in entries}
            display_names = {entry.id: entry.name for entry in entries}
            signatures = tuple(_ai_table_signature(entry) for entry in entries)
            task_id = SESSION.task_id

        catalog = build_table_catalog(tables, display_names=display_names)
        try:
            client = DeepSeekClient(api_key, model=model, timeout_seconds=75)
            plan = client.create_plan(prompt, catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True) if plan.executable else None
        except DeepSeekAPIError as exc:
            safe_message = str(exc).replace(api_key, "[API Key 已隐藏]")
            safe_message = AI_SECRET_PATTERN.sub("[API Key 已隐藏]", safe_message)
            raise ApiError(safe_message or "DeepSeek API 请求失败", 502) from None
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"AI 计划未通过本地安全校验：{exc}", 422) from None
        finally:
            # Do not retain the credential beyond this request handler.
            api_key = ""
            raw_key = ""

        token = ""
        if plan.executable:
            with SESSION.lock:
                if SESSION.task_id != task_id:
                    raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
                current_entries = [SESSION.get(table_id) for table_id in table_ids]
                current_signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
                if current_signatures != signatures:
                    raise ApiError("规划期间数据表范围发生变化，请重新生成计划", 409)
                token = SESSION.issue_ai_plan(
                    table_ids=table_ids,
                    table_signatures=signatures,
                    plan=plan,
                    model=model,
                )

        response: dict[str, Any] = {
            "status": plan.status,
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token or None,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS if token else 0,
            "data_scope": [
                {
                    "id": entry.id,
                    "name": entry.name,
                    "rows": len(entry.frame),
                    "columns": entry.frame.shape[1],
                }
                for entry in entries
            ],
            "privacy": "规划已发送需求与表结构；未发送单元格原值。执行与预演均在本机完成。",
        }
        if dry_run_result is not None:
            response["dry_run"] = dry_run_result.to_dict()
        return response

    def _local_sales_management_plan(
        self,
        entry: TableEntry,
        columns: dict[str, str],
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic five-sheet sales report plan."""

        with SESSION.lock:
            current = SESSION.get(entry.id)
            tables = {current.id: current.frame.copy(deep=True)}
            signatures = (_ai_table_signature(current),)
            task_id = SESSION.task_id
        catalog = build_table_catalog(tables, display_names={entry.id: entry.name})
        try:
            plan = validate_plan(_sales_report_plan_payload(entry.id, columns), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"销售经营报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            current = SESSION.get(entry.id)
            if (_ai_table_signature(current),) != signatures:
                raise ApiError("规划期间销售数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {
                    "id": entry.id,
                    "name": entry.name,
                    "rows": len(entry.frame),
                    "columns": entry.frame.shape[1],
                }
            ],
            "privacy": "DeepSeek 仅负责识别任务类型；字段映射、计算、异常规则、图表数据和Excel生成均在本机完成。",
        }

    def _local_quarterly_sales_plan(
        self,
        entries: list[TableEntry],
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic multi-sheet quarterly report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        display_names = {entry.id: entry.name for entry in current_entries}
        catalog = build_table_catalog(tables, display_names=display_names)
        try:
            plan = validate_plan(_quarterly_sales_plan_payload(current_entries), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"季度销售报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间销售数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "季度多表识别、表头升格、字段标准化、去重、无效排除、指标计算和 Excel 生成全部在本机完成。",
        }

    def _local_inventory_management_plan(
        self,
        entries: list[TableEntry],
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic procurement-sales-inventory report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        display_names = {entry.id: entry.name for entry in current_entries}
        catalog = build_table_catalog(tables, display_names=display_names)
        try:
            plan = validate_plan(_inventory_report_plan_payload(current_entries), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"库存经营报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间库存数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "表角色识别、格式统一、单据去重、库存核算、补货与积压判断、审计和 Excel 生成全部在本机完成。",
        }

    def _local_hr_management_plan(
        self,
        entries: list[TableEntry],
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic attendance/performance/payroll report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        catalog = build_table_catalog(tables, display_names={entry.id: entry.name for entry in current_entries})
        try:
            plan = validate_plan(_hr_report_plan_payload(current_entries), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"员工经营报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间员工数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "员工表角色识别、考勤绩效薪资整合、评分、预警、审计和 Excel 生成全部在本机完成。",
        }

    def _local_enterprise_diagnosis_plan(
        self,
        entries: list[TableEntry],
        prompt: str,
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic enterprise diagnosis report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        catalog = build_table_catalog(tables, display_names={entry.id: entry.name for entry in current_entries})
        try:
            plan = validate_plan(_enterprise_diagnosis_plan_payload(current_entries, prompt), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"企业经营诊断未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间经营数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "财务、客户、人员、成本、库存的识别、计算、风险诊断、行动计划和 Excel 生成全部在本机完成。",
        }

    def _local_selection_recommendation_plan(
        self,
        entries: list[TableEntry],
        prompt: str,
        *,
        model: str,
        include_charts: bool = True,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic candidate-selection report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        catalog = build_table_catalog(tables, display_names={entry.id: entry.name for entry in current_entries})
        try:
            plan = validate_plan(
                _selection_report_plan_payload(current_entries, prompt, include_charts=include_charts),
                catalog,
            )
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"候选评选报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间候选数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "候选标识、得分、评语、排名、风险与入选建议的计算和 Excel 生成全部在本机完成。",
        }

    def _local_adaptive_analysis_plan(
        self,
        entries: list[TableEntry],
        prompt: str,
        *,
        model: str,
    ) -> dict[str, Any]:
        """Build and dry-run the deterministic general adaptive report."""

        with SESSION.lock:
            current_entries = [SESSION.get(entry.id) for entry in entries]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in current_entries}
            signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            task_id = SESSION.task_id
        catalog = build_table_catalog(tables, display_names={entry.id: entry.name for entry in current_entries})
        try:
            plan = validate_plan(_adaptive_report_plan_payload(current_entries, prompt), catalog)
            preview = preview_plan(plan, tables)
            dry_run_result = execute_plan(plan, tables, dry_run=True)
        except (PlanValidationError, AgentExecutionError, TypeError, ValueError) as exc:
            raise ApiError(f"通用自适应报告未通过本地安全校验：{exc}", 422) from None
        with SESSION.lock:
            if SESSION.task_id != task_id:
                raise ApiError("规划期间任务已被清空或切换，请重新生成计划", 409)
            refreshed = [SESSION.get(entry.id) for entry in current_entries]
            if tuple(_ai_table_signature(entry) for entry in refreshed) != signatures:
                raise ApiError("规划期间数据发生变化，请重新生成计划", 409)
            token = SESSION.issue_ai_plan(
                table_ids=[entry.id for entry in current_entries],
                table_signatures=signatures,
                plan=plan,
                model=model,
            )
        return {
            "status": "ready",
            "normalized_request": plan.summary,
            "plan": plan.to_dict(),
            "preview": preview.to_dict(),
            "dry_run": dry_run_result.to_dict(),
            "warnings": list(plan.warnings),
            "plan_token": token,
            "expires_in_seconds": AI_PLAN_TTL_SECONDS,
            "auto_execute": True,
            "data_scope": [
                {"id": entry.id, "name": entry.name, "rows": len(entry.frame), "columns": entry.frame.shape[1]}
                for entry in current_entries
            ],
            "privacy": "字段角色识别、多事实域语义建模、同构合并、关系覆盖率、指标分析、异常检测和 Excel 生成全部在本机完成。",
        }

    def _ai_chart_plan(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Plan and render one read-only chart from natural language."""

        _validate_payload_keys(payload, {"prompt", "api_key", "model", "table_id", "current_spec"})
        prompt = payload.get("prompt")
        if not isinstance(prompt, str) or len(prompt.strip()) < 8:
            raise ApiError("请至少用 8 个字符说明想看的图表或修改要求")
        prompt = prompt.strip()
        if len(prompt) > AI_MAX_PROMPT_CHARS:
            raise ApiError(f"可视化需求不能超过 {AI_MAX_PROMPT_CHARS:,} 个字符")
        if AI_SECRET_PATTERN.search(prompt):
            raise ApiError("需求中检测到疑似 API Key；请只填写在密钥框")
        raw_key = payload.get("api_key")
        if not isinstance(raw_key, str) or not raw_key.strip():
            raise ApiError("请填写 DeepSeek API Key")
        api_key = raw_key.strip()
        if len(api_key) > 512 or any(character.isspace() for character in api_key):
            raise ApiError("DeepSeek API Key 格式无效")
        model = str(payload.get("model") or "deepseek-v4-flash").strip()
        if model not in SUPPORTED_DEEPSEEK_MODELS:
            raise ApiError("模型不可用")
        current_spec = payload.get("current_spec")
        if current_spec is not None and not isinstance(current_spec, dict):
            raise ApiError("当前图表上下文无效")

        with SESSION.lock:
            entry = SESSION.get(payload.get("table_id", ""))
            frame = entry.frame.copy(deep=True)
        catalog = build_table_catalog({entry.id: frame}, display_names={entry.id: entry.name})
        try:
            spec = DeepSeekClient(api_key, model=model, timeout_seconds=75).create_chart_spec(
                prompt, catalog, current_spec=current_spec
            )
        except DeepSeekAPIError as exc:
            safe_message = AI_SECRET_PATTERN.sub("[API Key 已隐藏]", str(exc).replace(api_key, "[API Key 已隐藏]"))
            raise ApiError(safe_message or "DeepSeek 图表规划失败", 502) from None
        except (ChartSpecValidationError, PlanValidationError, TypeError, ValueError) as exc:
            raise ApiError(f"AI 图表计划未通过本地安全校验：{exc}", 422) from None
        finally:
            api_key = ""
            raw_key = ""

        chart = None
        if spec["status"] == "ready":
            chart_parameters = dict(spec["chart"])
            chart = _apply_chart_presentation(_chart_payload(frame, chart_parameters), chart_parameters)
            with SESSION.lock:
                SESSION.record_chart({"table_id": entry.id, "spec": spec})
        return {
            "status": spec["status"],
            "normalized_request": spec["normalized_request"],
            "message": spec["message"],
            "clarification_questions": spec["clarification_questions"],
            "warnings": spec["warnings"],
            "spec": spec,
            "chart": chart,
            "history_position": len(SESSION.chart_history),
            "privacy": "仅发送自然语言需求、表名、字段名、类型和行列规模；未发送单元格原值。图表数据在本机计算。",
        }

    def _semantic_local_fallback(
        self,
        entries: list[TableEntry],
        prompt: str,
        *,
        model: str,
        original_route: dict[str, Any],
    ) -> dict[str, Any] | None:
        """Run specialist local engines after DeepSeek normalises colloquial text.

        This is the bridge between the two AI stages.  DeepSeek stage one may
        rewrite a phrase that the fast literal recognisers did not understand;
        the rewritten request is checked again against deterministic report
        engines before stage two is allowed to generate a generic operation
        plan.  No model output bypasses the local table-shape checks.
        """

        def response(result: dict[str, Any], reason: str) -> dict[str, Any]:
            route = {
                **original_route,
                "intent": "data",
                "data_request": prompt,
                "chart_request": "",
                "engineering_category": None,
                "reason": "DeepSeek 第一阶段已理解并标准化口语；" + reason,
            }
            return {
                "mode": "data",
                "route": route,
                "ai_pipeline": {
                    "stage_1": "DeepSeek 口语与业务意图标准化",
                    "stage_2": "本地白名单操作、字段参数映射与安全执行计划",
                    "execution": "本地确定性数据引擎",
                    "visualization_decision": original_route.get("visualization_need", "uncertain"),
                    "visualization_reason": original_route.get("visualization_reason", "由专业模块按数据结构判断"),
                },
                **result,
            }

        structured_action = original_route.get("business_action")
        structured_count = original_route.get("target_count")
        structured_confidence = original_route.get("interpretation_confidence")
        include_selection_charts = original_route.get("visualization_need") != "not_needed"
        if (
            structured_action == "select_candidates"
            and isinstance(structured_count, int)
            and not isinstance(structured_count, bool)
            and structured_count > 0
            and structured_confidence in {"high", "medium"}
        ):
            sources = _selection_report_sources(entries)
            if sources is not None:
                subject = str(original_route.get("business_subject") or "候选对象").strip()[:80]
                structured_prompt = (
                    f"根据{subject}的评分、成绩、评语和风险记录，选取综合表现最好的{structured_count}个候选组参加比赛"
                )
                return response(
                    self._local_selection_recommendation_plan(
                        sources,
                        structured_prompt,
                        model=model,
                        include_charts=include_selection_charts,
                    ),
                    "结构化语义识别为候选评选，已映射数量、评分字段、评语、排名和风险复核流程",
                )

        if _is_enterprise_diagnosis_request(prompt):
            sources = _enterprise_diagnosis_sources(entries)
            if sources is not None:
                return response(
                    self._local_enterprise_diagnosis_plan(sources, prompt, model=model),
                    "已映射为企业集团经营诊断和十表交付流程",
                )
        if _is_hr_management_report_request(prompt):
            sources = _hr_report_sources(entries)
            if sources is not None:
                return response(
                    self._local_hr_management_plan(sources, model=model),
                    "已映射为员工考勤、绩效、薪资整合分析流程",
                )
        if _is_inventory_management_report_request(prompt):
            sources = _inventory_report_sources(entries)
            if sources is not None:
                return response(
                    self._local_inventory_management_plan(sources, model=model),
                    "已映射为采购、销售、库存联动分析流程",
                )
        if _is_quarterly_sales_report_request(prompt):
            sources = _quarterly_sales_sources(entries)
            if sources is not None:
                return response(
                    self._local_quarterly_sales_plan(sources, model=model),
                    "已映射为多表销售清洗、合并与季度经营报告流程",
                )
        if _is_sales_management_report_request(prompt):
            source = _sales_report_source(entries)
            if source is not None:
                source_entry, column_mapping = source
                return response(
                    self._local_sales_management_plan(source_entry, column_mapping, model=model),
                    "已映射为销售经营分析和五表交付流程",
                )
        if _is_selection_recommendation_request(prompt):
            sources = _selection_report_sources(entries)
            if sources is not None:
                return response(
                    self._local_selection_recommendation_plan(
                        sources,
                        prompt,
                        model=model,
                        include_charts=include_selection_charts,
                    ),
                    "已映射为候选对象结构化评选、排名和风险复核流程",
                )
        if _is_adaptive_analysis_report_request(prompt):
            sources = _adaptive_report_sources(entries)
            if sources is not None:
                return response(
                    self._local_adaptive_analysis_plan(sources, prompt, model=model),
                    "已由通用分析编译器映射为领域、语义、粒度、证据门控和动态可视化流程",
                )
        return None

    def _selection_visualization_policy(
        self,
        entries: list[TableEntry],
        prompt: str,
        *,
        config: dict[str, Any],
        model: str,
    ) -> dict[str, str]:
        """Let AI judge chart value for a locally recognised selection task.

        The decision is advisory and bounded to an enum.  A network/API error
        never prevents the deterministic selection workflow from running.
        """

        fallback = {
            "need": "recommended",
            "reason": "多轮得分、排名和风险对比通常适合使用简洁图表辅助复核",
            "source": "professional_default",
        }
        if not config.get("configured"):
            return fallback
        api_key = str(config.get("api_key") or "")
        try:
            tables = {entry.id: entry.frame.copy(deep=True) for entry in entries}
            names = {entry.id: entry.name for entry in entries}
            catalog = build_table_catalog(tables, display_names=names)
            route = DeepSeekClient(api_key, model=model, timeout_seconds=75).classify_unified_request(prompt, catalog)
            return {
                "need": str(route.get("visualization_need") or "uncertain"),
                "reason": str(route.get("visualization_reason") or "AI未提供具体理由")[:500],
                "source": "deepseek_stage_1",
            }
        except (DeepSeekAPIError, PlanValidationError, TypeError, ValueError):
            return fallback
        finally:
            api_key = ""

    def _ai_unified(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Single-command entry point for all AI capabilities."""

        _validate_payload_keys(payload, {"prompt", "table_ids", "current_chart_spec", "mode_hint"})
        prompt = payload.get("prompt")
        if not isinstance(prompt, str) or len(prompt.strip()) < 8:
            raise ApiError("请至少用 8 个字符说明要完成的任务")
        original_prompt = prompt.strip()
        conversation = ConversationStore(SESSION.task_dir / "conversation.json")
        try:
            prompt, is_followup = conversation.resolve(original_prompt)
        except ValueError as exc:
            raise ApiError(str(exc)) from exc
        conversation.append(
            user_request=original_prompt,
            normalized_request=original_prompt,
            route="follow_up" if is_followup else "auto",
            status="received",
            plan_summary="已进入统一 AI 路由与本地安全执行流程",
            chart_spec=payload.get("current_chart_spec") if isinstance(payload.get("current_chart_spec"), dict) else {},
        )
        if len(prompt) > AI_MAX_PROMPT_CHARS:
            raise ApiError(f"任务描述不能超过 {AI_MAX_PROMPT_CHARS:,} 个字符")
        if AI_SECRET_PATTERN.search(prompt):
            raise ApiError("命令中检测到疑似 API Key；密钥应只保存在项目 .env 中")
        table_ids = _normalise_ai_table_scope(payload.get("table_ids"), allow_empty=True)
        with SESSION.lock:
            if not table_ids and SESSION.active_table:
                table_ids = [SESSION.active_table]
            entries = [SESSION.get(table_id) for table_id in table_ids]
            tables = {entry.id: entry.frame.copy(deep=True) for entry in entries}
            display_names = {entry.id: entry.name for entry in entries}
        mode_hint = payload.get("mode_hint")
        if mode_hint not in {None, "chart"}:
            raise ApiError("内部续接模式无效")
        if entries and mode_hint is None and payload.get("current_chart_spec") is None:
            alignment = assess_prompt_data_alignment(
                prompt,
                [entry.frame for entry in entries],
                [entry.name for entry in entries],
            )
            if not alignment.aligned:
                raise ApiError(
                    "需求与数据源语义不一致，已停止分析以防串单："
                    + alignment.reason
                    + "。请点击“新建任务”并重新上传本次项目原始文件。",
                    409,
                )

        # Complete sales-management orders are a stable, allow-listed local
        # workflow.  Run this recogniser before contacting DeepSeek so the
        # customer can still produce the five-sheet deliverable when the API,
        # firewall or network is unavailable.
        config = _project_ai_config()
        model = str(config.get("model") or "deepseek-v4-flash")
        if mode_hint is None and payload.get("current_chart_spec") is None and _is_enterprise_diagnosis_request(prompt):
            enterprise_sources = _enterprise_diagnosis_sources(entries)
            if enterprise_sources is not None:
                result = self._local_enterprise_diagnosis_plan(enterprise_sources, prompt, model=model)
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已识别为企业集团经营诊断，使用本地确定性跨表勾稽、利润驱动、客户回款、销售质量、成本库存风险和十表交付流程",
                    },
                    **result,
                }
        if mode_hint is None and payload.get("current_chart_spec") is None and _is_hr_management_report_request(prompt):
            hr_sources = _hr_report_sources(entries)
            if hr_sources is not None:
                result = self._local_hr_management_plan(hr_sources, model=model)
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已识别为员工考勤绩效薪资经营报告，使用本地确定性整合、评分、预警、审计和十表交付流程",
                    },
                    **result,
                }
        if (
            mode_hint is None
            and payload.get("current_chart_spec") is None
            and _is_inventory_management_report_request(prompt)
        ):
            inventory_sources = _inventory_report_sources(entries)
            if inventory_sources is not None:
                result = self._local_inventory_management_plan(inventory_sources, model=model)
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已识别为采购销售库存联动报告，使用本地确定性清洗、库存核算、补货积压判断和九表交付流程",
                    },
                    **result,
                }
        if (
            mode_hint is None
            and payload.get("current_chart_spec") is None
            and _is_quarterly_sales_report_request(prompt)
        ):
            quarterly_sources = _quarterly_sales_sources(entries)
            if quarterly_sources is not None:
                result = self._local_quarterly_sales_plan(quarterly_sources, model=model)
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已识别为多表季度销售经营报告，使用本地确定性清洗、去重、无效排除、分析与八表交付流程",
                    },
                    **result,
                }
        if (
            mode_hint is None
            and payload.get("current_chart_spec") is None
            and _is_sales_management_report_request(prompt)
        ):
            report_source = _sales_report_source(entries)
            if report_source is not None:
                source_entry, column_mapping = report_source
                result = self._local_sales_management_plan(
                    source_entry,
                    column_mapping,
                    model=model,
                )
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已识别为完整销售经营报告，使用本地确定性五表交付流程",
                    },
                    **result,
                }

        # Candidate selection is a local, structured data task rather than an
        # engineering/business-decision order.  Keep it ahead of model routing
        # so harmless requests such as “从序号里选8个参赛” cannot be sent to the
        # engineering schema and fail on descriptive code artifacts.
        if (
            mode_hint is None
            and payload.get("current_chart_spec") is None
            and _is_selection_recommendation_request(prompt)
        ):
            selection_sources = _selection_report_sources(entries)
            if selection_sources is not None:
                visual_policy = self._selection_visualization_policy(
                    selection_sources,
                    prompt,
                    config=config,
                    model=model,
                )
                result = self._local_selection_recommendation_plan(
                    selection_sources,
                    prompt,
                    model=model,
                    include_charts=visual_policy["need"] != "not_needed",
                )
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "visualization_need": visual_policy["need"],
                        "visualization_reason": visual_policy["reason"],
                        "reason": "已识别为候选对象结构化评选，使用本地确定性字段识别、综合评分、风险复核和结构化交付流程",
                    },
                    "ai_pipeline": {
                        "stage_1": "本地识别业务动作；DeepSeek判断可视化价值",
                        "stage_2": "本地白名单字段参数映射与安全执行计划",
                        "execution": "本地确定性数据引擎",
                        "visualization_decision": visual_policy["need"],
                        "visualization_reason": visual_policy["reason"],
                        "visualization_source": visual_policy["source"],
                    },
                    **result,
                }

        # Last deterministic layer: if no specialist route matched, build an
        # auditable report from the observed structure instead of forcing every
        # unfamiliar workbook through model-generated JSON.
        if (
            mode_hint is None
            and payload.get("current_chart_spec") is None
            and _is_adaptive_analysis_report_request(prompt)
        ):
            adaptive_sources = _adaptive_report_sources(entries)
            if adaptive_sources is not None:
                result = self._local_adaptive_analysis_plan(adaptive_sources, prompt, model=model)
                return {
                    "mode": "data",
                    "route": {
                        "intent": "data",
                        "normalized_request": prompt,
                        "data_request": prompt,
                        "chart_request": "",
                        "engineering_category": None,
                        "reason": "已启用本地通用分析编译器，按需求意图和当前数据证据动态选择分析与可视化，不依赖客户专用模板",
                    },
                    **result,
                }

        if not config["configured"]:
            raise ApiError("项目尚未配置 DeepSeek API Key；请运行安全配置脚本后重试", 503)
        api_key = config["api_key"]
        routing_tables = tables or {"order_context": pd.DataFrame({"待上传Excel字段": pd.Series(dtype="string")})}
        routing_names = display_names or {"order_context": "尚未上传数据表"}
        catalog = build_table_catalog(routing_tables, display_names=routing_names)
        ai_started_text = datetime.now().isoformat(timespec="milliseconds")
        ai_started = time.perf_counter()
        try:
            client = DeepSeekClient(api_key, model=model, timeout_seconds=75)
            if mode_hint == "chart" or payload.get("current_chart_spec") is not None:
                route = {
                    "intent": "chart",
                    "normalized_request": prompt,
                    "data_request": "",
                    "chart_request": prompt,
                    "engineering_category": None,
                    "reason": "继续创建或修改当前图表",
                }
            else:
                route = client.classify_unified_request(prompt, catalog)
            AI_TRACE_STORE.record(
                task_id=SESSION.task_id,
                kind="unified_router",
                model=model,
                prompt=prompt,
                status="success",
                started_at=ai_started_text,
                duration_ms=round((time.perf_counter() - ai_started) * 1000),
                metadata={"intent": route.get("intent"), "table_count": len(tables), "follow_up": is_followup},
            )
        except DeepSeekAPIError as exc:
            AI_TRACE_STORE.record(
                task_id=SESSION.task_id,
                kind="unified_router",
                model=model,
                prompt=prompt,
                status="failed",
                started_at=ai_started_text,
                duration_ms=round((time.perf_counter() - ai_started) * 1000),
                error_code=type(exc).__name__,
                metadata={"table_count": len(tables)},
            )
            safe_message = AI_SECRET_PATTERN.sub("[API Key 已隐藏]", str(exc).replace(api_key, "[API Key 已隐藏]"))
            raise ApiError(safe_message or "AI 无法理解当前任务", 502) from None
        except (PlanValidationError, TypeError, ValueError) as exc:
            AI_TRACE_STORE.record(
                task_id=SESSION.task_id,
                kind="unified_router",
                model=model,
                prompt=prompt,
                status="invalid",
                started_at=ai_started_text,
                duration_ms=round((time.perf_counter() - ai_started) * 1000),
                error_code=type(exc).__name__,
                metadata={"table_count": len(tables)},
            )
            raise ApiError(f"统一命令未通过安全校验：{exc}", 422) from None

        intent = route["intent"]
        if mode_hint is None and payload.get("current_chart_spec") is None and intent in {"data", "data_then_chart"}:
            semantic_prompt = str(route.get("data_request") or route.get("normalized_request") or prompt).strip()
            semantic_result = self._semantic_local_fallback(
                entries,
                semantic_prompt,
                model=model,
                original_route=dict(route),
            )
            if semantic_result is not None:
                return semantic_result
        if intent in {"data", "data_then_chart"} and _is_direct_chart_request(prompt):
            # Date/category aggregation for a requested chart is performed by the
            # deterministic local chart engine.  This avoids a fragile, redundant
            # AI-generated summary table while preserving every style instruction.
            route = {
                **route,
                "intent": "chart",
                "data_request": "",
                "chart_request": prompt,
                "reason": "当前需求只需在本地图表引擎内汇总并绘图，无需修改或生成中间数据表",
            }
            intent = "chart"
        if intent == "unsupported":
            return {"mode": "unsupported", "route": route, "message": "该需求超出当前安全能力范围，请拆分或改写任务。"}
        if intent in {"chart", "data_then_chart"} and not tables and intent == "chart":
            raise ApiError("制作图表前请先上传或生成一张数据表")
        try:
            if intent == "chart":
                selected = entries[0]
                result = self._ai_chart_plan(
                    {
                        "prompt": route["chart_request"],
                        "api_key": api_key,
                        "model": model,
                        "table_id": selected.id,
                        "current_spec": payload.get("current_chart_spec"),
                    }
                )
                return {"mode": "chart", "route": route, **result}
            if intent in {"data", "data_then_chart"}:
                if not table_ids:
                    raise ApiError("处理数据前请先上传 Excel 或 CSV 文件")
                result = self._ai_plan(
                    {
                        "prompt": route["data_request"],
                        "api_key": api_key,
                        "model": model,
                        "table_ids": table_ids,
                    }
                )
                return {
                    "mode": "data_then_chart" if intent == "data_then_chart" else "data",
                    "route": route,
                    "follow_up_chart_request": route["chart_request"] if intent == "data_then_chart" else None,
                    **result,
                }
            category = route["engineering_category"]
            result = self._ai_engineering(
                {
                    "category": category,
                    "prompt": route["normalized_request"],
                    "api_key": api_key,
                    "model": model,
                    "table_ids": table_ids,
                }
            )
            return {"mode": "engineering", "route": route, **result}
        finally:
            api_key = ""

    def _ai_execute(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"plan_token", "confirmed"})
        if payload.get("confirmed") is not True:
            raise ApiError("必须先人工核对计划并明确确认执行")
        token = payload.get("plan_token")
        if not isinstance(token, str) or not token or len(token) > 200:
            raise ApiError("AI 计划凭证无效")

        with SESSION.lock:
            ticket = SESSION.consume_ai_plan(token)
            if SESSION.task_id != ticket.task_id:
                raise ApiError("任务已发生变化，原 AI 计划不能执行", 409)
            entries = [SESSION.get(table_id) for table_id in ticket.table_ids]
            signatures = tuple(_ai_table_signature(entry) for entry in entries)
            if signatures != ticket.table_signatures:
                raise ApiError("数据表范围已发生变化，原 AI 计划不能执行", 409)
            tables = {entry.id: entry.frame.copy(deep=True) for entry in entries}
            input_names = [entry.name for entry in entries]
            task_file_names = sorted(SESSION.file_names)

        try:
            result = execute_plan(ticket.plan, tables, dry_run=False)
        except (AgentExecutionError, PlanValidationError, TypeError, ValueError) as exc:
            raise ApiError(f"AI 计划执行失败，任务数据未修改：{exc}", 422) from None

        generated = list(result.tables.items())
        if not generated:
            raise ApiError("AI 计划没有生成结果表，任务数据未修改", 422)
        if len(generated) > AI_MAX_OUTPUT_TABLES:
            raise ApiError(
                f"计划将生成 {len(generated)} 张表，超过 {AI_MAX_OUTPUT_TABLES} 张安全上限；请缩小拆分范围",
                422,
            )
        total_cells = sum(len(frame) * max(1, frame.shape[1]) for _, frame in generated)
        if total_cells > AI_MAX_OUTPUT_CELLS:
            raise ApiError("AI 计划结果超过 5,000,000 个单元格安全上限，请拆分任务", 422)
        if any(len(frame) > MAX_ROWS_PER_TABLE for _, frame in generated):
            raise ApiError("AI 计划某张结果表超过 300,000 行安全上限，请拆分任务", 422)

        produced: list[str] = []
        output_tables: list[dict[str, Any]] = []
        created_review_ids: list[str] = []
        with SESSION.lock:
            if SESSION.task_id != ticket.task_id:
                raise ApiError("执行期间任务已被清空，结果没有写入", 409)
            current_entries = [SESSION.get(table_id) for table_id in ticket.table_ids]
            current_signatures = tuple(_ai_table_signature(entry) for entry in current_entries)
            if current_signatures != ticket.table_signatures:
                raise ApiError("执行期间数据表范围发生变化，结果没有写入", 409)
            previous_active = SESSION.active_table
            previous_warning_count = len(SESSION.import_warnings)
            try:
                for output_name, frame in generated:
                    table_id = SESSION.add_table(
                        output_name,
                        frame,
                        source=f"AI 一句话执行（{ticket.model}）",
                    )
                    produced.append(table_id)
                    stored = SESSION.tables[table_id]
                    output_tables.append(
                        {
                            "id": table_id,
                            "name": stored.name,
                            "rows": len(stored.frame),
                            "columns": stored.frame.shape[1],
                        }
                    )
                    category = _ai_review_category(output_name)
                    if category and not stored.frame.empty:
                        review_items: list[dict[str, Any]] = []
                        evidence_columns = list(stored.frame.columns[:8])
                        for position in range(min(100, len(stored.frame))):
                            row = stored.frame.iloc[position]
                            review_items.append(
                                {
                                    "title": f"{stored.name} · 第 {position + 1} 条",
                                    "detail": "AI 计划已识别此项；请结合业务凭证人工确认",
                                    "record_key": f"结果行 {position + 1}",
                                    "evidence": {
                                        str(column): _json_value(row.iloc[index])
                                        for index, column in enumerate(evidence_columns)
                                    },
                                }
                            )
                        created_review_ids.extend(
                            SESSION.add_review_items(
                                category,
                                "AI 一句话执行",
                                review_items,
                                table_id=table_id,
                                limit=100,
                            )
                        )
                SESSION.record(
                    "AI 一句话执行",
                    f"已人工确认；模型 {ticket.model}；白名单计划 {len(ticket.plan.steps)} 步；生成 {len(produced)} 张结果表",
                    inputs=input_names,
                    produced=produced,
                    before_rows=sum(len(frame) for frame in tables.values()),
                    after_rows=sum(len(frame) for _, frame in generated),
                )
            except Exception:
                for table_id in produced:
                    SESSION.tables.pop(table_id, None)
                for review_id in created_review_ids:
                    SESSION.review_items.pop(review_id, None)
                del SESSION.import_warnings[previous_warning_count:]
                SESSION.active_table = previous_active
                raise
            review_counts = SESSION.review_payload()["counts"]

        management_report_download_url: str | None = None
        sales_operation = next(
            (
                step.operation
                for step in ticket.plan.steps
                if step.operation
                in {
                    "sales_management_report",
                    "quarterly_sales_report",
                    "inventory_management_report",
                    "hr_management_report",
                    "adaptive_analysis_report",
                    "selection_recommendation_report",
                    "enterprise_diagnosis_report",
                }
            ),
            None,
        )
        if sales_operation is not None:
            if sales_operation == "enterprise_diagnosis_report":
                # Domain fact plugins may compose different evidence sheets.
                # Export exactly the generated contract instead of assuming the
                # manufacturing ten-sheet template.
                expected_names = tuple(name for name, _ in generated if name != "primary")
                if "门店经营诊断" in expected_names:
                    # Restaurant plugin contract is intentionally discovered
                    # from its generated outputs, so new restaurant fact
                    # sheets can be added without changing the generic route.
                    expected_names = tuple(name for name in expected_names if name != "数据源确认")
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            elif sales_operation == "adaptive_analysis_report":
                expected_names = (
                    "管理层通用总览",
                    "主数据分析",
                    "数据字典",
                    "数据质量",
                    "表关系建议",
                    "分类排名",
                    "时间趋势",
                    "异常数据",
                    "自适应图表看板",
                )
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            elif sales_operation == "selection_recommendation_report":
                selection_step = next(
                    step for step in ticket.plan.steps if step.operation == "selection_recommendation_report"
                )
                expected_names = (
                    "评选管理总览",
                    "建议入选名单",
                    "全部候选排序",
                    "风险复核清单",
                    "评选规则与字段",
                )
                if selection_step.params.get("include_charts", True):
                    expected_names = (*expected_names, "评选图表看板")
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            elif sales_operation == "hr_management_report":
                expected_names = (
                    "管理层人效总览",
                    "员工综合分析",
                    "表现优秀员工",
                    "重点关注员工",
                    "考勤分析",
                    "绩效分析",
                    "薪资分析",
                    "人工核验",
                    "数据审计",
                    "人力图表看板",
                )
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            elif sales_operation == "inventory_management_report":
                expected_names = (
                    "管理层库存总览",
                    "商品库存分析",
                    "补货建议",
                    "积压清单",
                    "采购分析",
                    "销售分析",
                    "人工核验",
                    "数据审计",
                    "库存图表看板",
                )
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            elif sales_operation == "quarterly_sales_report":
                expected_names = (
                    "管理层数据总览",
                    "季度合并数据",
                    "产品分析",
                    "地区分析",
                    "销售人员分析",
                    "异常数据提醒",
                    "清洗审计",
                    "图表展示",
                )
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            else:
                expected_names = (
                    "管理层数据总览",
                    "产品分析",
                    "销售人员分析",
                    "异常数据提醒",
                    "图表展示",
                )
                destination_name = _specific_report_filename(sales_operation, task_file_names)
            with SESSION.lock:
                by_name = {
                    SESSION.tables[table_id].name: SESSION.tables[table_id].frame.copy(deep=True)
                    for table_id in produced
                    if table_id in SESSION.tables
                }
                missing = [name for name in expected_names if name not in by_name]
                if missing:
                    raise ApiError("经营报告缺少交付工作表：" + "、".join(missing), 422)
                report_request = next(
                    (
                        str(step.params.get("user_request") or "")
                        for step in ticket.plan.steps
                        if isinstance(step.params, Mapping) and step.params.get("user_request")
                    ),
                    "",
                )
                confirmation = source_confirmation_frame(
                    [entry.frame for entry in entries],
                    [entry.name for entry in entries],
                    file_names=task_file_names,
                    task_id=ticket.task_id,
                    user_request=report_request,
                )
                report_tables = {}
                for position, name in enumerate(expected_names):
                    report_tables[name] = by_name[name]
                    if position == 0:
                        report_tables["数据源确认"] = confirmation
                destination = SESSION.output_dir / destination_name
            try:
                export_tables(
                    report_tables,
                    destination,
                    include_log=False,
                    overwrite=True,
                )
            except (OSError, TypeError, ValueError) as exc:
                raise ApiError(f"经营报告 Excel 生成失败：{exc}", 422) from None
            with SESSION.lock:
                management_report_download_url = SESSION.register_download(destination)

        result_metadata = result.to_dict()
        reports = result_metadata.get("reports", {})
        step_results = [
            {
                "id": step.id,
                "operation": step.operation,
                "name": _AI_OPERATION_LABELS.get(step.operation, step.operation),
                "status": "completed",
                "summary": reports.get(step.id, {}),
            }
            for step in ticket.plan.steps
        ]
        pending_added = len(created_review_ids)
        return {
            "status": "needs_review" if pending_added else "completed",
            "message": (
                f"AI 计划已完成 {len(ticket.plan.steps)} 步，生成 {len(output_tables)} 张结果表"
                + (f"，新增 {pending_added} 条人工核验项" if pending_added else "")
            ),
            "steps_completed": len(ticket.plan.steps),
            "tables_created": len(output_tables),
            "rows_processed": sum(len(frame) for frame in tables.values()),
            "pending_reviews": pending_added,
            "review_counts": review_counts,
            "output_tables": output_tables,
            "step_results": step_results,
            "warnings": list(result.warnings),
            "download_url": management_report_download_url,
            "auto_download": management_report_download_url is not None,
        }

    def _recipe_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"name", "description", "steps"})
        name = str(payload.get("name") or "").strip()
        description = str(payload.get("description") or "").strip()
        if not name:
            raise ApiError("请填写处理方案名称")
        if len(name) > 60:
            raise ApiError("处理方案名称不能超过 60 个字符")
        if len(description) > 300:
            raise ApiError("处理方案说明不能超过 300 个字符")
        steps = payload.get("steps")
        if not isinstance(steps, list) or not steps:
            raise ApiError("处理方案至少需要一个步骤")
        recipe = ProcessingRecipe.from_dict({"name": name, "description": description, "steps": steps})
        saved = _save_recipe(recipe)
        return {
            "message": f"处理方案“{recipe.name}”已安全保存到本机",
            "saved": saved,
            "recipes": _list_saved_recipes(),
        }

    def _recipe_run(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"recipe_id", "table_id", "output_name", "dry_run"})
        dry_run = payload.get("dry_run", False)
        if not isinstance(dry_run, bool):
            raise ApiError("预演参数必须是布尔值")
        recipe, metadata = _read_stored_recipe(str(payload.get("recipe_id") or ""))
        with SESSION.lock:
            table_id = payload.get("table_id") or SESSION.active_table or ""
            source = SESSION.get(str(table_id))
            result, report = run_recipe(source.frame, recipe, dry_run=dry_run)
            response: dict[str, Any] = {
                "recipe": {**metadata, **recipe.to_dict()},
                "report": report.to_dict(),
                "steps_count": sum(step.status == "applied" for step in report.steps),
                "before_rows": len(source.frame),
                "after_rows": len(result),
                "step_results": [step.to_dict() for step in report.steps],
            }
            if dry_run:
                columns = [str(column) for column in result.columns]
                preview_rows = [
                    {column: _json_value(value) for column, value in zip(columns, values)}
                    for values in result.head(20).itertuples(index=False, name=None)
                ]
                response.update(
                    {
                        "message": (
                            f"方案“{recipe.name}”预演完成：{len(source.frame):,} 行 → {len(result):,} 行，未写入任务"
                        ),
                        "preview": {"columns": columns, "rows": preview_rows},
                    }
                )
                return response
            output_name = _validated_output_name(payload.get("output_name"), fallback=f"{recipe.name}_结果")
            new_id = SESSION.add_table(output_name, result, source="复用处理方案")
            SESSION.record(
                "运行复用处理方案",
                f"方案：{recipe.name}；共 {len(recipe.steps)} 步；输入指纹已记录但不保存原始数据",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
            response.update(
                {
                    "message": f"方案“{recipe.name}”运行完成，已生成“{SESSION.tables[new_id].name}”",
                    "table_id": new_id,
                }
            )
            return response

    def _demo(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, set())
        frame = _demo_sales_frame()
        with SESSION.lock:
            # Demo data is a complete disposable task, not another table to
            # append to a customer's live task. Repeated clicks therefore
            # always produce exactly one clean demo table.
            SESSION.reset()
            SESSION.task_name = "销售分析演示任务"
            new_id = SESSION.add_table(
                "虚构销售演示数据",
                frame,
                source="本地生成的虚构演示数据",
                original=False,
            )
            SESSION.record(
                "生成演示数据",
                "本机生成虚构销售数据；内含少量空值、重复记录和异常数值",
                inputs=[],
                produced=[new_id],
                before_rows=0,
                after_rows=len(frame),
            )
            table_name = SESSION.tables[new_id].name
        return {
            "message": (f"已新建演示任务并生成“{table_name}”，共 {len(frame):,} 行；所有名称和订单均为虚构数据"),
            "table_id": new_id,
            "task_id": SESSION.task_id,
            "new_task": True,
        }

    def _analysis(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            frame = source.frame.copy(deep=True)
        return _analysis_payload(frame)

    def _chart(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(
            payload,
            {
                "table",
                "chart_type",
                "dimension",
                "measure",
                "aggregation",
                "top_n",
                "date_grain",
                "series",
                "start",
                "end",
                "progress",
                "style_3d",
                "title",
                "theme",
                "number_format",
                "sort",
                "reference_lines",
                "highlight",
                "show_labels",
                "show_legend",
            },
        )
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            frame = source.frame.copy(deep=True)
        return _apply_chart_presentation(_chart_payload(frame, payload), payload)

    def _anomalies(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table", "column", "method", "output_name"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            column = _require_column(source.frame, payload.get("column"), label="检测字段")
            method = _require_choice(payload.get("method", "iqr"), _ANOMALY_METHODS, label="异常检测算法")
            output_name = _validated_output_name(payload.get("output_name"), fallback="异常值明细")
            try:
                detection = detect_outliers(source.frame, columns=[column], method=method)
            except (TypeError, ValueError) as exc:
                raise ApiError(str(exc)) from exc
            result = detection.flagged_rows.copy(deep=True)
            if not detection.outliers.empty:
                details = detection.outliers.loc[
                    :, ["row_position", "value", "score", "lower_bound", "upper_bound"]
                ].rename(
                    columns={
                        "value": "异常检测值",
                        "score": "异常强度",
                        "lower_bound": "正常下界",
                        "upper_bound": "正常上界",
                    }
                )
                result = result.merge(details, on="row_position", how="left", validate="one_to_one")
            new_id = SESSION.add_table(output_name, result, source="异常值检测")
            count = len(result)
            review_items: list[dict[str, Any]] = []
            for row in detection.outliers.head(300).to_dict("records"):
                position = row.get("row_position")
                review_items.append(
                    {
                        "title": f"{column} 异常值待核验",
                        "detail": f"算法 {method} 检测到超出正常范围的记录",
                        "record_key": f"原始行 {position}",
                        "evidence": {
                            "原始行位置": position,
                            "检测值": row.get("value"),
                            "异常强度": row.get("score"),
                            "正常下界": row.get("lower_bound"),
                            "正常上界": row.get("upper_bound"),
                        },
                    }
                )
            SESSION.add_review_items("异常值", "异常值检测", review_items, table_id=new_id, limit=300)
            SESSION.record(
                "异常值检测",
                f"字段 {column}；算法 {method}；发现 {count} 行异常",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=count,
            )
            table_name = SESSION.tables[new_id].name
        return {
            "message": f"异常检测完成，发现 {count:,} 行，已生成“{table_name}”",
            "table_id": new_id,
            "outlier_count": count,
        }

    def _pivot(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table", "index", "columns", "value", "aggregation", "output_name"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            index_column = _require_column(source.frame, payload.get("index"), label="行维度")
            column_dimension = _optional_column(source.frame, payload.get("columns"), label="列维度")
            value_column = _require_column(source.frame, payload.get("value"), label="数值字段")
            aggregation = _require_choice(payload.get("aggregation", "sum"), _PIVOT_AGGREGATIONS, label="透视统计方式")
            if column_dimension == index_column:
                raise ApiError("行维度和列维度不能选择同一个字段")
            output_name = _validated_output_name(payload.get("output_name"), fallback="交叉透视结果")

            if column_dimension:
                row_cardinality = max(1, int(source.frame[index_column].nunique(dropna=False)))
                column_cardinality = max(1, int(source.frame[column_dimension].nunique(dropna=False)))
                if row_cardinality * column_cardinality > 250_000:
                    raise ApiError("所选维度组合可能生成超过 250,000 个透视单元格，请改用较低基数字段或先筛选")
                try:
                    result = cross_pivot(
                        source.frame,
                        index=index_column,
                        columns=column_dimension,
                        values=value_column,
                        aggregation=aggregation,
                        fill_value=0,
                        margins=True,
                        margins_name="合计",
                    )
                except (TypeError, ValueError) as exc:
                    raise ApiError(str(exc)) from exc
                detail = f"{index_column} × {column_dimension}；{value_column} {_AGGREGATION_LABELS[aggregation]}"
            else:
                work = pd.DataFrame(
                    {"__index__": _normalised_missing(source.frame[index_column])},
                    index=source.frame.index,
                )
                work["__index__"] = work["__index__"].map(lambda value: "（空值）" if pd.isna(value) else value)
                if aggregation == "count":
                    grouped = work.groupby("__index__", dropna=False, observed=True).size()
                elif aggregation == "nunique":
                    work["__value__"] = _normalised_missing(source.frame[value_column])
                    grouped = work.groupby("__index__", dropna=False, observed=True)["__value__"].nunique(dropna=True)
                else:
                    work["__value__"] = _numeric_column(source.frame, value_column)
                    grouped = (
                        work.dropna(subset=["__value__"])
                        .groupby("__index__", dropna=False, observed=True)["__value__"]
                        .agg(aggregation)
                    )
                result = grouped.reset_index(name=f"{value_column}_{aggregation}").rename(
                    columns={"__index__": index_column}
                )
                result = result.sort_values(
                    f"{value_column}_{aggregation}", ascending=False, kind="stable"
                ).reset_index(drop=True)
                detail = f"按 {index_column} 汇总；{value_column} {_AGGREGATION_LABELS[aggregation]}"

            if result.shape[0] * max(1, result.shape[1]) > 500_000:
                raise ApiError("透视结果超过 500,000 个单元格，请缩小维度范围")
            new_id = SESSION.add_table(output_name, result, source="交叉透视分析")
            SESSION.record(
                "交叉透视分析",
                detail,
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
            table_name = SESSION.tables[new_id].name
        return {
            "message": f"透视分析完成，已生成“{table_name}”（{len(result):,} 行）",
            "table_id": new_id,
        }

    def _rfm(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table", "customer", "date", "amount", "output_name"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            customer = _require_column(source.frame, payload.get("customer"), label="客户标识字段")
            date_column = _require_column(source.frame, payload.get("date"), label="交易日期字段")
            amount = _require_column(source.frame, payload.get("amount"), label="交易金额字段")
            if len({customer, date_column, amount}) != 3:
                raise ApiError("客户标识、交易日期和交易金额必须选择三个不同字段")
            output_name = _validated_output_name(payload.get("output_name"), fallback="RFM客户分群")
            try:
                analysis = rfm_segmentation(
                    source.frame,
                    customer_column=customer,
                    date_column=date_column,
                    amount_column=amount,
                )
            except (TypeError, ValueError) as exc:
                raise ApiError(str(exc)) from exc
            if analysis.customers.empty:
                raise ApiError("没有同时具备客户、有效日期和有效金额的记录，无法生成 RFM 分群")

            produced: list[str] = []
            summary_id = SESSION.add_table(
                f"{output_name}_分群汇总", analysis.segment_summary, source="RFM客户价值分析"
            )
            produced.append(summary_id)
            if not analysis.invalid_rows.empty:
                invalid_id = SESSION.add_table(
                    f"{output_name}_无效记录", analysis.invalid_rows, source="RFM客户价值分析"
                )
                produced.append(invalid_id)
            customer_id = SESSION.add_table(output_name, analysis.customers, source="RFM客户价值分析")
            produced.append(customer_id)
            SESSION.record(
                "RFM客户价值分群",
                f"客户 {len(analysis.customers)}；分群 {len(analysis.segment_summary)}；无效记录 {len(analysis.invalid_rows)}",
                inputs=[source.name],
                produced=produced,
                before_rows=len(source.frame),
                after_rows=len(analysis.customers),
            )
            table_name = SESSION.tables[customer_id].name
        return {
            "message": f"RFM 分析完成，识别 {len(analysis.customers):,} 位客户，已生成“{table_name}”及分群汇总",
            "table_id": customer_id,
            "summary_table_id": summary_id,
            "customer_count": len(analysis.customers),
            "invalid_row_count": len(analysis.invalid_rows),
        }

    def _analysis_export(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table", "filename"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            tables = _analysis_export_tables(source.frame)
            raw_name = payload.get("filename") or f"{source.name}_智能分析报告"
            base = Path(_safe_filename(raw_name, fallback="智能分析报告")).stem
            destination = SESSION.output_dir / f"{base}.xlsx"
            counter = 2
            while destination.exists():
                destination = SESSION.output_dir / f"{base}_{counter}.xlsx"
                counter += 1
            export_tables(tables, destination, include_log=False, overwrite=True)
            url = SESSION.register_download(destination)
            SESSION.record(
                "导出智能分析报告",
                "包含原数据、质量概览、问题、洞察、描述统计、相关性、异常、趋势和分类贡献",
                inputs=[source.name],
                produced=[],
                before_rows=len(source.frame),
                after_rows=len(source.frame),
            )
        return {
            "message": "智能分析交付包已生成，浏览器将开始下载",
            "download_url": url,
        }

    def _fuzzy_cluster(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table", "column", "threshold", "output_name"})
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            column = _require_column(source.frame, payload.get("column"), label="目标字段")
            threshold = _bounded_probability(payload.get("threshold", 0.88), label="相似度阈值")
            output_name = _validated_output_name(payload.get("output_name"), fallback="相似值待确认")
            try:
                candidates = cluster_similar_values(
                    source.frame,
                    column,
                    threshold=threshold,
                    max_unique=1_000,
                )
            except (KeyError, TypeError, ValueError) as exc:
                raise ApiError(f"相似值扫描失败：{str(exc).strip(chr(39))}") from exc
            new_id = SESSION.add_table(output_name, candidates, source="相似值候选分组")
            group_count = (
                int(candidates["组ID"].nunique(dropna=True))
                if "组ID" in candidates.columns and not candidates.empty
                else 0
            )
            review_items = []
            for _, row in candidates.head(500).iterrows():
                review_items.append(
                    {
                        "title": f"候选组 {row.get('组ID', '')}：{row.get('原值', '')}",
                        "detail": "请确认是否统一为建议标准值",
                        "record_key": str(row.get("组ID") or ""),
                        "evidence": {
                            "原值": row.get("原值"),
                            "建议标准值": row.get("建议标准值"),
                            "相似度": row.get("相似度"),
                            "出现次数": row.get("出现次数"),
                        },
                    }
                )
            SESSION.add_review_items("相似值聚类", "相似值候选分组", review_items, table_id=new_id, limit=500)
            SESSION.record(
                "相似值候选分组",
                f"字段 {column}；阈值 {threshold:.0%}；候选 {len(candidates)} 个；分组 {group_count} 个",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(candidates),
            )
            table_name = SESSION.tables[new_id].name
        return {
            "message": f"扫描完成，发现 {group_count:,} 个候选组、{len(candidates):,} 个名称，已生成“{table_name}”",
            "table_id": new_id,
            "group_count": group_count,
            "candidate_count": len(candidates),
        }

    def _fuzzy_lookup(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(
            payload,
            {"source", "lookup", "source_key", "lookup_key", "threshold", "output_name"},
        )
        with SESSION.lock:
            source = SESSION.get(payload.get("source", ""))
            lookup = SESSION.get(payload.get("lookup", ""))
            if source.id == lookup.id:
                raise ApiError("来源表和标准表必须选择两张不同的数据表")
            source_key = _require_column(source.frame, payload.get("source_key"), label="来源名称字段")
            lookup_key = _require_column(lookup.frame, payload.get("lookup_key"), label="标准名称字段")
            threshold = _bounded_probability(payload.get("threshold", 0.88), label="最低相似度")
            output_name = _validated_output_name(payload.get("output_name"), fallback="模糊匹配结果")
            try:
                source_unique = int(
                    source.frame[source_key].dropna().astype(str).str.strip().replace("", pd.NA).nunique()
                )
                lookup_unique = int(
                    lookup.frame[lookup_key].dropna().astype(str).str.strip().replace("", pd.NA).nunique()
                )
            except (TypeError, ValueError) as exc:
                raise ApiError("名称字段包含无法比较的复杂内容，请先转成普通文本") from exc
            if source_unique > 5_000:
                raise ApiError(f"来源名称字段有 {source_unique:,} 个唯一值，超过 5,000 个安全上限；请先筛选或分批处理")
            if lookup_unique > 1_000:
                raise ApiError(f"标准名称字段有 {lookup_unique:,} 个唯一值，超过 1,000 个安全上限；请先筛选标准表")
            try:
                result = fuzzy_lookup(
                    source.frame,
                    lookup.frame,
                    source_key,
                    lookup_key,
                    value_columns=[],
                    threshold=threshold,
                    ambiguous_gap=0.03,
                )
            except (KeyError, TypeError, ValueError) as exc:
                raise ApiError(f"模糊匹配失败：{str(exc).strip(chr(39))}") from exc
            counts = result["匹配状态"].value_counts(dropna=False).to_dict()
            matched = int(counts.get("已匹配", 0))
            review = int(counts.get("待确认", 0))
            unmatched = int(counts.get("未匹配", 0))
            new_id = SESSION.add_table(output_name, result, source="两表模糊匹配")
            review_rows = result[result["匹配状态"].isin(["待确认", "未匹配"])].head(500)
            review_items = []
            evidence_columns = [
                column
                for column in (
                    source_key,
                    "候选值",
                    "相似度",
                    "次选候选值",
                    "次选相似度",
                    "匹配状态",
                )
                if column in review_rows.columns
            ]
            for position, row in review_rows.iterrows():
                status = str(row.get("匹配状态") or "待确认")
                review_items.append(
                    {
                        "title": f"{status}：{row.get(source_key, '')}",
                        "detail": "请比较最佳候选与次选候选后决定是否接受",
                        "record_key": f"来源行 {position}",
                        "evidence": {column: row.get(column) for column in evidence_columns},
                    }
                )
            SESSION.add_review_items("模糊匹配", "两表模糊匹配", review_items, table_id=new_id, limit=500)
            SESSION.record(
                "两表模糊匹配",
                f"{source_key} ↔ {lookup_key}；阈值 {threshold:.0%}；已匹配 {matched}；待确认 {review}；未匹配 {unmatched}",
                inputs=[source.name, lookup.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
            table_name = SESSION.tables[new_id].name
        return {
            "message": f"模糊匹配完成：已匹配 {matched:,}，待确认 {review:,}，未匹配 {unmatched:,}；已生成“{table_name}”",
            "table_id": new_id,
            "matched_count": matched,
            "review_count": review,
            "unmatched_count": unmatched,
        }

    def _validate(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table_id", "rules", "output_name"})
        raw_rules = payload.get("rules")
        if not isinstance(raw_rules, list) or not raw_rules:
            raise ApiError("请至少添加一条质量验收规则")
        if len(raw_rules) > 100:
            raise ApiError("单次最多运行 100 条质量验收规则")

        prepared_rules: list[ValidationRule] = []
        compact_keys = {
            "type",
            "column",
            "min",
            "max",
            "pattern",
            "values",
            "severity",
            "message",
            "enabled",
        }
        for index, raw in enumerate(raw_rules, start=1):
            if not isinstance(raw, dict):
                raise ApiError(f"第 {index} 条验收规则必须是对象")
            if "rule_type" in raw:
                direct = dict(raw)
                direct.setdefault("rule_id", f"R{index:03d}")
                try:
                    prepared_rules.append(ValidationRule.from_dict(direct))
                except (TypeError, ValueError) as exc:
                    raise ApiError(f"第 {index} 条验收规则无效：{exc}") from exc
                continue
            unknown = sorted(set(raw) - compact_keys)
            if unknown:
                raise ApiError(f"第 {index} 条验收规则包含未知字段：{unknown}")
            rule_type = str(raw.get("type") or "").strip()
            column = str(raw.get("column") or "").strip()
            if not rule_type or not column:
                raise ApiError(f"第 {index} 条验收规则缺少类型或字段")
            params: dict[str, Any] = {}
            if rule_type == "not_null":
                params = {"blank_as_null": True}
            elif rule_type == "unique":
                params = {"ignore_nulls": True, "blank_as_null": True}
            elif rule_type == "range":
                if "min" in raw:
                    params["min"] = raw["min"]
                if "max" in raw:
                    params["max"] = raw["max"]
                params.update({"value_type": "numeric", "ignore_nulls": True})
            elif rule_type == "regex":
                params = {
                    "pattern": raw.get("pattern"),
                    "mode": "fullmatch",
                    "ignore_nulls": True,
                }
            elif rule_type == "allowed_values":
                params = {"values": raw.get("values"), "ignore_nulls": True}
            else:
                raise ApiError(f"第 {index} 条验收规则类型“{rule_type}”不受支持")
            try:
                prepared_rules.append(
                    ValidationRule(
                        rule_id=f"R{index:03d}",
                        rule_type=rule_type,
                        column=column,
                        severity=str(raw.get("severity") or "error"),
                        params=params,
                        message=raw.get("message"),
                        enabled=raw.get("enabled", True),
                    )
                )
            except (TypeError, ValueError) as exc:
                raise ApiError(f"第 {index} 条验收规则无效：{exc}") from exc

        with SESSION.lock:
            source = SESSION.get(payload.get("table_id", ""))
            try:
                report = validate_dataframe(
                    source.frame,
                    prepared_rules,
                    include_values=True,
                    max_value_chars=80,
                )
            except (KeyError, TypeError, ValueError) as exc:
                raise ApiError(f"质量验收失败：{str(exc).strip(chr(39))}") from exc
            output_name = _validated_output_name(payload.get("output_name"), fallback="质量验收明细")
            results_frame = report.rule_results_frame()
            failures_frame = report.failures_frame()
            result_id = SESSION.add_table(f"{output_name}_规则汇总", results_frame, source="质量验收")
            failure_id = SESSION.add_table(f"{output_name}_失败明细", failures_frame, source="质量验收")
            review_items: list[dict[str, Any]] = []
            for _, row in failures_frame.head(500).iterrows():
                value_preview = row.get("value_preview")
                other_preview = row.get("other_value_preview")
                review_items.append(
                    {
                        "title": f"规则 {row.get('rule_id', '')}：{row.get('column', '')}",
                        "detail": str(row.get("message") or "质量规则未通过"),
                        "reason": str(row.get("message") or "质量规则未通过"),
                        "record_key": f"源数据第 {int(row.get('row_position', 0)) + 1} 行",
                        "original": value_preview,
                        "candidate": other_preview,
                        "evidence": {
                            "规则": row.get("rule_type"),
                            "严重性": row.get("severity"),
                            "失败代码": row.get("code"),
                            "字段": row.get("column"),
                            "另一字段": row.get("other_column"),
                        },
                    }
                )
            SESSION.add_review_items("质量验收", source.name, review_items, table_id=failure_id, limit=500)
            SESSION.record(
                "运行质量验收规则",
                (
                    f"规则 {report.rule_count} 条；通过 {report.passed_rule_count} 条；"
                    f"失败记录 {report.failure_count} 条；阻断级失败 {report.blocking_failure_count} 条"
                ),
                inputs=[source.name],
                produced=[result_id, failure_id],
                before_rows=len(source.frame),
                after_rows=len(results_frame) + len(failures_frame),
            )
        total_rules = report.rule_count
        pass_rate = report.passed_rule_count / total_rules if total_rules else 1.0
        report_payload = report.to_dict(include_failures=False)
        return {
            "message": (
                f"质量验收完成：{report.passed_rule_count}/{total_rules} 条规则通过，"
                f"发现 {report.failure_count:,} 条失败记录"
            ),
            "passed": report.passed,
            "pass_rate": pass_rate,
            "failed_count": report.failure_count,
            "passed_rules": report.passed_rule_count,
            "total_rules": total_rules,
            "rule_results": report_payload["rule_results"],
            "report": report_payload,
            "table_ids": {"规则汇总": result_id, "失败明细": failure_id},
        }

    def _reconcile_advanced(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"left_id", "right_id", "config", "output_name"})
        config = payload.get("config")
        if not isinstance(config, dict):
            raise ApiError("高级对账配置必须是对象")
        allowed_config = {
            "left_keys",
            "right_keys",
            "left_secondary_columns",
            "right_secondary_columns",
            "amount",
            "date",
            "enable_split_candidates",
        }
        _validate_payload_keys(config, allowed_config)

        def checked_columns(frame: pd.DataFrame, raw: Any, label: str) -> list[str]:
            if raw in (None, ""):
                return []
            values = [raw] if isinstance(raw, str) else raw
            if not isinstance(values, list) or not all(isinstance(item, str) for item in values):
                raise ApiError(f"{label}必须是字段列表")
            if len(values) != len(set(values)):
                raise ApiError(f"{label}不能包含重复字段")
            return [_require_column(frame, item, label=label) for item in values]

        with SESSION.lock:
            left = SESSION.get(payload.get("left_id", ""))
            right = SESSION.get(payload.get("right_id", ""))
            if left.id == right.id:
                raise ApiError("高级对账必须选择两张不同的数据表")
            left_keys = checked_columns(left.frame, config.get("left_keys"), "左侧匹配键")
            right_keys = checked_columns(right.frame, config.get("right_keys"), "右侧匹配键")
            if len(left_keys) != len(right_keys):
                raise ApiError("左右匹配键数量必须一致")
            left_secondary = checked_columns(left.frame, config.get("left_secondary_columns"), "左侧辅助字段")
            right_secondary = checked_columns(right.frame, config.get("right_secondary_columns"), "右侧辅助字段")
            if len(left_secondary) != len(right_secondary):
                raise ApiError("左右辅助字段数量必须一致")

            amount = config.get("amount")
            if not isinstance(amount, dict):
                raise ApiError("高级对账必须同时选择左右金额字段")
            _validate_payload_keys(amount, {"left_column", "right_column", "tolerance"})
            left_amount = _require_column(left.frame, amount.get("left_column"), label="左侧金额字段")
            right_amount = _require_column(right.frame, amount.get("right_column"), label="右侧金额字段")
            try:
                amount_tolerance = Decimal(str(amount.get("tolerance", 0)))
            except (InvalidOperation, ValueError) as exc:
                raise ApiError("金额容差必须是非负数字") from exc
            if not amount_tolerance.is_finite() or amount_tolerance < 0:
                raise ApiError("金额容差必须是非负数字")

            date_config = config.get("date")
            left_date: str | None = None
            right_date: str | None = None
            date_tolerance_days = 0
            if date_config is not None:
                if not isinstance(date_config, dict):
                    raise ApiError("日期容差配置必须是对象")
                _validate_payload_keys(date_config, {"left_column", "right_column", "tolerance_days"})
                left_date = _require_column(left.frame, date_config.get("left_column"), label="左侧日期字段")
                right_date = _require_column(right.frame, date_config.get("right_column"), label="右侧日期字段")
                date_tolerance_days = _bounded_integer(
                    date_config.get("tolerance_days", 0),
                    label="日期容差",
                    minimum=0,
                    maximum=3650,
                )
            enable_split = config.get("enable_split_candidates", False)
            if not isinstance(enable_split, bool):
                raise ApiError("拆分候选参数必须是布尔值")

            try:
                result = reconcile_tables(
                    left.frame,
                    right.frame,
                    left_amount=left_amount,
                    right_amount=right_amount,
                    left_date=left_date,
                    right_date=right_date,
                    left_key_columns=left_keys or None,
                    right_key_columns=right_keys or None,
                    left_secondary_columns=left_secondary or None,
                    right_secondary_columns=right_secondary or None,
                    amount_tolerance=amount_tolerance,
                    date_tolerance_days=date_tolerance_days,
                    enable_split_candidates=enable_split,
                )
            except (KeyError, TypeError, ValueError) as exc:
                raise ApiError(f"高级对账失败：{str(exc).strip(chr(39))}") from exc

            output_name = _validated_output_name(payload.get("output_name"), fallback="高级对账结果")
            summary_labels = {
                "left_rows": "左侧总行数",
                "right_rows": "右侧总行数",
                "matched_count": "自动匹配",
                "amount_difference_count": "金额差异",
                "date_difference_count": "日期差异",
                "review_candidate_rows": "待确认候选行",
                "review_group_count": "待确认候选组",
                "left_only_count": "左侧独有",
                "right_only_count": "右侧独有",
                "duplicate_rows_count": "重复键隔离",
                "amount_tolerance": "金额容差",
                "date_tolerance_days": "日期容差（天）",
            }
            summary_frame = pd.DataFrame(
                [
                    {
                        "指标": summary_labels.get(str(key), str(key)),
                        "结果": _json_value(value),
                    }
                    for key, value in result.summary.items()
                ]
            )
            output_frames = [
                ("对账摘要", summary_frame),
                ("自动匹配", result.matched),
                ("金额差异", result.amount_difference),
                ("日期差异", result.date_difference),
                ("待人工确认", result.review),
                ("左侧独有", result.left_only),
                ("右侧独有", result.right_only),
                ("重复键隔离", result.duplicates),
            ]
            produced: list[str] = []
            table_ids: dict[str, str] = {}
            for label, frame in output_frames:
                table_id = SESSION.add_table(f"{output_name}_{label}", frame, source="高级对账")
                produced.append(table_id)
                table_ids[label] = table_id

            def add_frame_reviews(frame: pd.DataFrame, category: str, table_id: str, limit: int) -> None:
                items: list[dict[str, Any]] = []
                for _, row in frame.head(limit).iterrows():
                    reason = row.get("match_reason", row.get("unmatched_reason", row.get("duplicate_reason", "")))
                    left_position = row.get("left_row_position", row.get("source_row_position", ""))
                    right_position = row.get("right_row_position", "")
                    group = row.get("candidate_group_id", "")
                    record_key = (
                        str(group) if str(group or "").strip() else f"左行 {left_position} / 右行 {right_position}"
                    )
                    items.append(
                        {
                            "title": f"{category}：{record_key}",
                            "detail": str(reason or "请核对原始记录后确认处理口径"),
                            "reason": str(reason or ""),
                            "record_key": record_key,
                            "original": row.get("left_amount_decimal", row.get("left__金额", "")),
                            "candidate": row.get("right_amount_decimal", row.get("right__金额", "")),
                            "score": row.get("match_score"),
                            "evidence": {
                                "左侧行号": left_position,
                                "右侧行号": right_position,
                                "候选组": group,
                                "匹配类型": row.get("match_type", ""),
                                "金额差": row.get("amount_difference_decimal", ""),
                                "日期差天数": row.get("date_difference_days", ""),
                            },
                        }
                    )
                SESSION.add_review_items(category, "高级对账", items, table_id=table_id, limit=limit)

            add_frame_reviews(result.review, "对账候选", table_ids["待人工确认"], 300)
            add_frame_reviews(result.amount_difference, "金额差异", table_ids["金额差异"], 150)
            add_frame_reviews(result.date_difference, "日期差异", table_ids["日期差异"], 150)
            add_frame_reviews(result.duplicates, "重复键", table_ids["重复键隔离"], 100)
            add_frame_reviews(result.left_only, "左侧独有", table_ids["左侧独有"], 100)
            add_frame_reviews(result.right_only, "右侧独有", table_ids["右侧独有"], 100)

            matched_count = int(result.summary.get("matched_count", len(result.matched)))
            left_only_count = int(result.summary.get("left_only_count", len(result.left_only)))
            right_only_count = int(result.summary.get("right_only_count", len(result.right_only)))
            difference_count = len(result.amount_difference) + len(result.date_difference)
            SESSION.record(
                "高级容差对账",
                (
                    f"{left.name} ↔ {right.name}；自动匹配 {matched_count}；"
                    f"金额/日期差异 {difference_count}；待确认候选组 "
                    f"{result.summary.get('review_group_count', 0)}；未匹配 "
                    f"{left_only_count + right_only_count}"
                ),
                inputs=[left.name, right.name],
                produced=produced,
                before_rows=len(left.frame) + len(right.frame),
                after_rows=sum(len(frame) for _, frame in output_frames),
            )
        return {
            "message": (
                f"高级对账完成：自动匹配 {matched_count:,}，差异 {difference_count:,}，"
                f"左右独有 {left_only_count + right_only_count:,}；已生成完整证据表"
            ),
            "summary": _json_value(dict(result.summary)),
            "matched_count": matched_count,
            "left_only_count": left_only_count,
            "right_only_count": right_only_count,
            "difference_count": difference_count,
            "review_count": len(result.review),
            "table_ids": table_ids,
            "details": [
                {"label": label, "rows": len(frame), "table_id": table_ids[label]} for label, frame in output_frames
            ],
        }

    def _clean(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            dedupe_keep: str | bool = payload.get("dedupe_keep", "first")
            if dedupe_keep == "none":
                dedupe_keep = False
            config = CleaningConfig(
                trim_whitespace=bool(payload.get("trim_text", True)),
                normalize_blank_strings=bool(payload.get("trim_text", True)),
                drop_empty_rows=bool(payload.get("drop_empty", True)),
                drop_empty_columns=bool(payload.get("drop_empty", True)),
                drop_duplicates=bool(payload.get("drop_duplicates", False)),
                duplicate_subset=tuple(payload.get("dedupe_columns") or ()) or None,
                keep_duplicate=dedupe_keep,
                infer_types=bool(payload.get("infer_types", False)),
                missing_strategy="keep",
            )
            result, report = smart_clean(source.frame, config, table_name=source.name)
            if payload.get("fill_missing"):
                method = payload.get("fill_method", "value")
                if method == "ffill":
                    result = result.ffill()
                elif method == "bfill":
                    result = result.bfill()
                else:
                    result = result.fillna(payload.get("fill_value", ""))
            if payload.get("normalize_columns"):
                columns, warnings = _unique_columns([re.sub(r"\s+", " ", str(c).strip()) for c in result.columns])
                result.columns = columns
                SESSION.import_warnings.extend(warnings)
            output_name = _safe_table_name(payload.get("output_name"), fallback="清洗结果")
            new_id = SESSION.add_table(output_name, result, source=source.name)
            detail = (
                f"去空行 {report.empty_rows_removed}；去重复 {report.duplicate_rows_removed}；"
                f"空值 {report.missing_cells_before}→{int(result.isna().sum().sum())}"
            )
            SESSION.record(
                "一键清洗",
                detail,
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
        return {"message": f"清洗完成，已生成“{SESSION.tables[new_id].name}”"}

    def _columns(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            selected = list(payload.get("columns") or []) or None
            rename_column = str(payload.get("rename_column") or "")
            rename_value = str(payload.get("rename_value") or "").strip()
            rename = {rename_column: rename_value} if rename_column and rename_value else None
            if selected is not None and rename_column and rename_column not in selected:
                raise ApiError("需要重命名的字段不在“保留字段”中，请重新选择")
            sort_column = str(payload.get("sort_column") or "") or None
            if selected is not None and sort_column and sort_column not in selected:
                raise ApiError("排序字段不在“保留字段”中，请重新选择")
            effective_sort_column = (
                rename_value if sort_column and rename and sort_column == rename_column else sort_column
            )
            output_name = _safe_table_name(payload.get("output_name"), fallback="字段整理结果")
            result = select_rename_sort(
                source.frame,
                columns=selected,
                rename=rename,
                sort_by=effective_sort_column,
                ascending=bool(payload.get("ascending", True)),
                table_name=source.name,
                output_name=output_name,
            )
            new_id = SESSION.add_table(output_name, result, source="字段整理")
            detail_parts = [f"保留 {len(result.columns)} 个字段"]
            if rename:
                detail_parts.append(f"{rename_column}→{rename_value}")
            if sort_column:
                detail_parts.append(f"按 {effective_sort_column}{'升序' if payload.get('ascending', True) else '降序'}")
            SESSION.record(
                "字段整理与排序",
                "；".join(detail_parts),
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
        return {"message": f"字段整理完成，已生成“{SESSION.tables[new_id].name}”"}

    def _replace(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            column = str(payload.get("column") or "")
            if column not in source.frame.columns:
                raise ApiError("目标字段不存在，请重新选择")
            find_text = str(payload.get("find") or "")
            replacement = str(payload.get("replace") or "")
            if find_text == "":
                raise ApiError("查找内容不能为空")
            mode = str(payload.get("mode") or "exact")
            case_sensitive = bool(payload.get("case_sensitive", False))
            result = source.frame.copy(deep=True)
            values = result[column].astype("string")
            if mode == "exact":
                if case_sensitive:
                    matched = values.eq(find_text).fillna(False)
                else:
                    matched = values.str.casefold().eq(find_text.casefold()).fillna(False)
                result.loc[matched, column] = replacement
            elif mode == "contains":
                flags = 0 if case_sensitive else re.IGNORECASE
                pattern = re.compile(re.escape(find_text), flags)
                matched = values.str.contains(pattern, na=False)
                result.loc[matched, column] = values.loc[matched].map(
                    lambda value: pattern.sub(replacement, str(value))
                )
            else:
                raise ApiError("匹配方式必须是完全一致或包含内容")
            output_name = _safe_table_name(payload.get("output_name"), fallback="查找替换结果")
            new_id = SESSION.add_table(output_name, result, source="查找替换")
            count = int(matched.sum())
            SESSION.record(
                "按列查找替换",
                f"字段 {column}；匹配 {count} 行；方式 {mode}",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
        return {"message": f"查找替换完成，共修改 {count} 行"}

    def _concat(self, payload: dict[str, Any]) -> dict[str, Any]:
        ids = list(payload.get("tables") or [])
        if len(ids) < 2:
            raise ApiError("请至少选择两张数据表")
        with SESSION.lock:
            sources = [SESSION.get(item) for item in ids]
            mapping = {entry.name: entry.frame for entry in sources}
            output_name = _safe_table_name(payload.get("output_name"), fallback="追加合并结果")
            result = concat_tables(
                mapping,
                join=payload.get("strategy", "outer"),
                source_column="来源表" if payload.get("add_source", True) else None,
                output_name=output_name,
            )
            new_id = SESSION.add_table(output_name, result, source="纵向追加")
            SESSION.record(
                "纵向追加",
                f"合并 {len(sources)} 张表，字段策略：{payload.get('strategy', 'outer')}",
                inputs=[entry.name for entry in sources],
                produced=[new_id],
                before_rows=sum(len(entry.frame) for entry in sources),
                after_rows=len(result),
            )
        return {"message": f"已追加 {len(sources)} 张表，共 {len(result):,} 行"}

    def _join(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            left = SESSION.get(payload.get("left", ""))
            right = SESSION.get(payload.get("right", ""))
            left_key = str(payload.get("left_key") or "")
            right_key = str(payload.get("right_key") or "")
            if left_key not in left.frame.columns or right_key not in right.frame.columns:
                raise ApiError("所选匹配字段不存在，请重新选择")
            left_duplicates = int(left.frame.duplicated(subset=[left_key], keep=False).sum())
            right_duplicates = int(right.frame.duplicated(subset=[right_key], keep=False).sum())
            if right_duplicates:
                raise ApiError(
                    f"右表字段“{right_key}”发现 {right_duplicates} 行重复键，连接会造成数据膨胀。"
                    "请先对右表按该字段去重后再匹配。"
                )
            output_name = _safe_table_name(payload.get("output_name"), fallback="字段匹配结果")
            validate = "many_to_one" if left_duplicates else "one_to_one"
            result = join_tables(
                left.frame,
                right.frame,
                left_on=left_key,
                right_on=right_key,
                how=payload.get("how", "left"),
                validate=validate,
                left_name=left.name,
                right_name=right.name,
                output_name=output_name,
            )
            new_id = SESSION.add_table(output_name, result, source="关键字段匹配")
            SESSION.record(
                "关键字段匹配",
                f"{left_key} ↔ {right_key}；{payload.get('how', 'left')}；左表重复键行 {left_duplicates}",
                inputs=[left.name, right.name],
                produced=[new_id],
                before_rows=len(left.frame),
                after_rows=len(result),
            )
        return {"message": f"匹配完成，生成 {len(result):,} 行"}

    def _compare(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            base = SESSION.get(payload.get("base", ""))
            target = SESSION.get(payload.get("target", ""))
            key = str(payload.get("key") or "")
            if key not in base.frame.columns or key not in target.frame.columns:
                raise ApiError("唯一标识字段必须同时存在于两张表")
            common = [c for c in base.frame.columns if c in target.frame.columns and c != key]
            compare_columns = list(payload.get("columns") or common)
            missing = [c for c in compare_columns if c not in common]
            if missing:
                raise ApiError(f"比较字段不是两表共同字段：{missing}")
            prefix = _safe_table_name(payload.get("output_name"), fallback="数据比对")
            base_dup = base.frame[base.frame.duplicated(key, keep=False)].copy()
            target_dup = target.frame[target.frame.duplicated(key, keep=False)].copy()
            base_unique = base.frame.drop_duplicates(key, keep="first").copy()
            target_unique = target.frame.drop_duplicates(key, keep="first").copy()
            base_keys = set(base_unique[key].dropna().tolist())
            target_keys = set(target_unique[key].dropna().tolist())
            added = target_unique[target_unique[key].isin(target_keys - base_keys)].copy()
            removed = base_unique[base_unique[key].isin(base_keys - target_keys)].copy()
            common_keys = base_keys & target_keys
            left_common = base_unique[base_unique[key].isin(common_keys)].set_index(key)
            right_common = target_unique[target_unique[key].isin(common_keys)].set_index(key)
            order = [item for item in right_common.index if item in left_common.index]
            left_common = left_common.loc[order]
            right_common = right_common.loc[order]
            if compare_columns:
                equal_matrix = left_common[compare_columns].eq(right_common[compare_columns]) | (
                    left_common[compare_columns].isna() & right_common[compare_columns].isna()
                )
                changed_mask = ~equal_matrix.all(axis=1)
            else:
                changed_mask = pd.Series(False, index=left_common.index)
            changed_rows: list[dict[str, Any]] = []
            for row_key in left_common.index[changed_mask]:
                row: dict[str, Any] = {key: row_key}
                changed_fields: list[str] = []
                for column in compare_columns:
                    old = left_common.at[row_key, column]
                    new = right_common.at[row_key, column]
                    is_equal = (pd.isna(old) and pd.isna(new)) or (not pd.isna(old) and not pd.isna(new) and old == new)
                    if not is_equal:
                        changed_fields.append(column)
                    row[f"{column}_旧值"] = old
                    row[f"{column}_新值"] = new
                row["变更字段"] = "、".join(changed_fields)
                changed_rows.append(row)
            modified = pd.DataFrame(changed_rows)
            unchanged_keys = list(left_common.index[~changed_mask])
            unchanged = target_unique[target_unique[key].isin(unchanged_keys)].copy()
            duplicates = pd.concat(
                [base_dup.assign(来源表=base.name), target_dup.assign(来源表=target.name)],
                ignore_index=True,
                sort=False,
            )
            outputs = {
                f"{prefix}_新增": added,
                f"{prefix}_删除": removed,
                f"{prefix}_修改": modified,
                f"{prefix}_未变化": unchanged,
                f"{prefix}_重复键": duplicates,
            }
            produced = [SESSION.add_table(name, frame, source="数据比对") for name, frame in outputs.items()]
            SESSION.record(
                "新旧数据比对",
                f"新增 {len(added)}；删除 {len(removed)}；修改 {len(modified)}；未变化 {len(unchanged)}；重复键行 {len(duplicates)}",
                inputs=[base.name, target.name],
                produced=produced,
                before_rows=len(base.frame),
                after_rows=len(target.frame),
            )
        return {"message": f"比对完成：新增 {len(added)}，删除 {len(removed)}，修改 {len(modified)}"}

    def _summary(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            groups = list(payload.get("group_by") or [])
            column = str(payload.get("column") or "")
            method = str(payload.get("method") or "count")
            if method not in {"count", "nunique", "sum", "mean", "max", "min"}:
                raise ApiError("不支持的统计方式")
            output_name = _safe_table_name(payload.get("output_name"), fallback="汇总结果")
            result, engine = group_summary_auto(source.frame, by=groups, aggregations={column: method})
            new_id = SESSION.add_table(output_name, result, source="分组汇总")
            SESSION.record(
                "分组汇总",
                f"按 {'、'.join(groups)} 分组，对 {column} 执行 {method}；计算引擎 {engine.engine}",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
        return {"message": f"汇总完成，共 {len(result):,} 个分组", "engine": engine.to_dict()}

    def _split(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            column = str(payload.get("column") or "")
            parts = split_dataframe(source.frame, by=column, table_name=source.name)
            base = _safe_filename(payload.get("output_name") or "拆分结果")
            mode = payload.get("mode", "sheets")
            if mode == "sheets":
                destination = SESSION.output_dir / f"{base}.xlsx"
                export_tables(parts, destination, operation_log=SESSION.operation_log(), overwrite=True)
            elif mode == "files":
                staging = SESSION.output_dir / f".{uuid.uuid4().hex}"
                staging.mkdir(parents=True)
                destination = SESSION.output_dir / f"{base}.zip"
                try:
                    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
                        for name, frame in parts.items():
                            part_path = staging / f"{_safe_filename(name)}.xlsx"
                            export_tables({name: frame}, part_path, include_log=False, overwrite=True)
                            archive.write(part_path, arcname=part_path.name)
                finally:
                    shutil.rmtree(staging, ignore_errors=True)
            else:
                raise ApiError("拆分交付方式无效")
            url = SESSION.register_download(destination)
            SESSION.record(
                "批量拆分",
                f"按 {column} 拆分为 {len(parts)} 份；模式 {mode}",
                inputs=[source.name],
                produced=[],
                before_rows=len(source.frame),
                after_rows=len(source.frame),
            )
        return {"message": f"已拆分为 {len(parts)} 份", "download_url": url}

    def _mask(self, payload: dict[str, Any]) -> dict[str, Any]:
        with SESSION.lock:
            source = SESSION.get(payload.get("table", ""))
            columns = list(payload.get("columns") or [])
            mode = str(payload.get("mode") or "partial")
            if mode == "id_card" or mode == "bank_card":
                mode = "id"
            if mode == "auto":
                rules: dict[str, str] = {}
                for column in columns:
                    label = str(column).lower()
                    if "手机" in label or "电话" in label or "phone" in label:
                        rules[column] = "phone"
                    elif "邮箱" in label or "email" in label:
                        rules[column] = "email"
                    elif "身份证" in label or "银行卡" in label or "证件" in label:
                        rules[column] = "id"
                    elif "姓名" in label or "name" in label:
                        rules[column] = "name"
                    else:
                        rules[column] = "partial"
                result = mask_columns(source.frame, rules, table_name=source.name)
            else:
                result = mask_columns(source.frame, columns, strategy=mode, table_name=source.name)
            output_name = _safe_table_name(payload.get("output_name"), fallback="脱敏结果")
            new_id = SESSION.add_table(output_name, result, source="数据脱敏")
            SESSION.record(
                "敏感信息脱敏",
                f"字段：{'、'.join(columns)}；策略：{payload.get('mode', 'auto')}",
                inputs=[source.name],
                produced=[new_id],
                before_rows=len(source.frame),
                after_rows=len(result),
            )
        return {"message": f"已生成脱敏副本“{SESSION.tables[new_id].name}”"}

    def _export(self, payload: dict[str, Any]) -> dict[str, Any]:
        ids = list(payload.get("tables") or [])
        if not ids:
            raise ApiError("请至少选择一张表导出")
        export_format = str(payload.get("format") or "xlsx")
        if export_format not in {"xlsx", "csv_zip"}:
            raise ApiError("导出格式无效")
        professional = payload.get("professional", False)
        if not isinstance(professional, bool):
            raise ApiError("专业交付参数必须是布尔值")
        if professional and export_format != "xlsx":
            raise ApiError("专业 V3 交付包仅支持 Excel 工作簿格式")
        include_summary = bool(payload.get("include_summary", True))
        export_mode = str(payload.get("export_mode") or "data").strip().lower()
        if export_mode not in {"data", "preserve"}:
            raise ApiError("导出模式必须是 data 或 preserve")
        if export_mode == "preserve" and export_format != "xlsx":
            raise ApiError("原文件保真模式仅支持 Excel 工作簿")
        with SESSION.lock:
            entries = [SESSION.get(item) for item in ids]
            tables = {entry.name: entry.frame.copy(deep=True) for entry in entries}
            if payload.get("safe_csv", True):
                tables = {name: _escape_spreadsheet_formulas(frame) for name, frame in tables.items()}
            review_counts = SESSION.review_payload()["counts"]
            prefix_tables: dict[str, pd.DataFrame] = {}
            if include_summary:
                summary_rows = [
                    {"项目": "任务编号", "内容": SESSION.task_id},
                    {"项目": "任务名称", "内容": SESSION.task_name},
                    {"项目": "导出时间", "内容": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
                    {"项目": "导出数据表", "内容": "；".join(entry.name for entry in entries)},
                    {"项目": "处理步骤数", "内容": len(SESSION.operations)},
                ]
                for index, operation in enumerate(SESSION.operations, start=1):
                    summary_rows.append(
                        {
                            "项目": f"步骤 {index}：{operation['name']}",
                            "内容": operation.get("detail", ""),
                        }
                    )
                prefix_tables["处理摘要"] = pd.DataFrame(summary_rows)
            if professional:
                blank_detail = _blank_cell_detail_frame(entries)
                blank_counts = blank_detail.groupby("来源数据表").size().to_dict() if not blank_detail.empty else {}
                acceptance_rows: list[dict[str, Any]] = []
                for entry in entries:
                    missing_cells = int(blank_counts.get(entry.name, 0))
                    duplicate_rows = profile_dataframe(entry.frame).duplicate_row_count
                    acceptance_rows.append(
                        {
                            "数据表": entry.name,
                            "行数": len(entry.frame),
                            "列数": len(entry.frame.columns),
                            "空值单元格": missing_cells,
                            "重复组记录": duplicate_rows,
                            "验收状态": "需复核" if missing_cells or duplicate_rows else "基础检查通过",
                            "说明": "基础统计不替代客户业务口径确认",
                        }
                    )
                acceptance_rows.append(
                    {
                        "数据表": "人工核验中心",
                        "行数": review_counts["total"],
                        "列数": None,
                        "空值单元格": None,
                        "重复组记录": None,
                        "验收状态": ("存在待确认项" if review_counts["pending"] else "无待确认项"),
                        "说明": (
                            f"待确认 {review_counts['pending']}；已接受 {review_counts['accepted']}；"
                            f"已拒绝 {review_counts['rejected']}"
                        ),
                    }
                )
                prefix_tables["验收清单"] = pd.DataFrame(acceptance_rows)
                if not blank_detail.empty:
                    prefix_tables["空值清单"] = blank_detail
                long_text_detail = _long_text_detail_frame(entries)
                if not long_text_detail.empty:
                    prefix_tables["长文本明细"] = long_text_detail
            tables = {**prefix_tables, **tables}
            base = _safe_filename(payload.get("filename") or f"{SESSION.task_id}_处理结果")
            source_workbook: Path | None = None
            if export_mode == "preserve":
                source_candidates = sorted(
                    path for path in SESSION.upload_dir.iterdir() if path.suffix.lower() in {".xlsx", ".xlsm"}
                )
                if len(source_candidates) != 1:
                    raise ApiError("原文件保真模式要求当前任务恰好包含一个 Excel 源工作簿")
                source_workbook = source_candidates[0]
            suffix = (
                source_workbook.suffix.lower()
                if source_workbook is not None
                else (".xlsx" if export_format == "xlsx" else ".zip")
            )
            destination = SESSION.output_dir / f"{Path(base).stem}{suffix}"
            qa_tables = tables
            fidelity_inventory: dict[str, Any] | None = None
            if source_workbook is not None:
                fidelity_inventory = workbook_feature_inventory(source_workbook)
                existing_sheets = set(fidelity_inventory.get("worksheets") or [])
                replace_sheets: dict[str, str] = {}
                for entry in entries:
                    if not entry.original:
                        continue
                    inferred = entry.name.rsplit("__", 1)[-1]
                    if inferred in existing_sheets:
                        replace_sheets[entry.name] = inferred
                preserve_workbook_export(source_workbook, destination, tables, replace_sheets=replace_sheets)
                qa_tables = {replace_sheets.get(name, name): frame for name, frame in tables.items()}
            else:
                export_tables(tables, destination, include_log=False, overwrite=True)
            acceptance = verify_delivery(destination, qa_tables, allow_extra_tables=True)
            acceptance_path = destination.with_name(f"{destination.stem}_自动验收.json")
            write_acceptance_json(acceptance, acceptance_path)
            if acceptance.status != "passed":
                failed = [item.table for item in acceptance.tables if item.status != "passed"]
                raise ApiError(f"导出后自动验收失败：{'、'.join(failed)}；已阻止交付", 500)
            url = SESSION.register_download(destination)
            acceptance_url = SESSION.register_download(acceptance_path)
            SESSION.record(
                "导出交付包",
                (
                    f"格式 {suffix}；包含 {len(entries)} 张结果表；"
                    f"{'专业验收版' if professional else '标准版'}；"
                    f"{'原工作簿保真模式' if export_mode == 'preserve' else '数据交付模式'}；"
                    f"自动验收 {acceptance.checks_passed}/{acceptance.checks_total}；"
                    f"待人工确认 {review_counts['pending']}"
                ),
                inputs=[entry.name for entry in entries],
                produced=[],
                before_rows=sum(len(entry.frame) for entry in entries),
                after_rows=sum(len(entry.frame) for entry in entries),
            )
        return {
            "message": "交付文件已生成并通过自动验收，浏览器将开始下载",
            "download_url": url,
            "acceptance_url": acceptance_url,
            "acceptance": acceptance.to_dict(),
            "fidelity_inventory": fidelity_inventory,
        }

    def _undo(self, payload: dict[str, Any]) -> dict[str, Any]:
        del payload
        with SESSION.lock:
            SESSION.undo()
        return {"message": "已撤销上一步处理"}

    def _redo(self, payload: dict[str, Any]) -> dict[str, Any]:
        del payload
        with SESSION.lock:
            SESSION.redo()
        return {"message": "已重做上一步处理"}

    def _review_decision(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"ids", "decision", "note"})
        ids = payload.get("ids") or []
        if not isinstance(ids, list) or not ids:
            raise ApiError("请至少选择一条待核验记录")
        if len(ids) > 500:
            raise ApiError("单次最多处理 500 条核验记录")
        raw_decision = str(payload.get("decision") or "").strip().lower()
        decision = {
            "accept": "accepted",
            "accepted": "accepted",
            "reject": "rejected",
            "rejected": "rejected",
            "pending": "pending",
            "reset": "pending",
        }.get(raw_decision)
        if decision is None:
            raise ApiError("核验决定必须是接受、拒绝或恢复待确认")
        with SESSION.lock:
            changed = SESSION.decide_reviews([str(item) for item in ids], decision, str(payload.get("note") or ""))
            review_payload = SESSION.review_payload()
        return {
            "message": f"已更新 {changed} 条核验记录",
            **review_payload,
        }

    def _reset(self, payload: dict[str, Any]) -> dict[str, Any]:
        del payload
        SESSION.reset()
        # The browser must immediately switch its X-Task-ID header to the new
        # empty task.  Without this id, its next /api/state refresh restores
        # the previous durable task and makes the old tables appear again.
        return {
            "message": "任务数据已清空，新任务已创建",
            "task_id": SESSION.task_id,
            "task_name": SESSION.task_name,
            "new_task": True,
        }

    def _task_open(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"task_id"})
        task_id = str(payload.get("task_id") or "").strip()
        with SESSION.lock:
            SESSION.restore(task_id)
        return {"message": f"已恢复任务“{SESSION.task_name}”", "task_id": SESSION.task_id}

    def _task_delete(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"task_id"})
        task_id = str(payload.get("task_id") or "").strip()
        with SESSION.lock:
            if task_id == SESSION.task_id:
                raise ApiError("不能删除正在使用的任务；请先新建或恢复其他任务", 409)
            TASK_REPOSITORY.delete(task_id)
        return {"message": "任务及其本地文件已永久删除"}

    def _task_purge(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, set())
        with SESSION.lock:
            removed = [item for item in TASK_REPOSITORY.purge_expired() if item != SESSION.task_id]
        return {"message": f"已清理 {len(removed)} 个超过保留期的任务", "removed": removed}

    def _contract_generate(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table_id", "name", "contract_id", "strict_nulls", "allow_extra_columns"})
        table_id = str(payload.get("table_id") or SESSION.active_table or "")
        with SESSION.lock:
            entry = SESSION.get(table_id)
            contract_id = str(payload.get("contract_id") or "").strip() or None
            version = CONTRACT_STORE.next_version(contract_id) if contract_id else 1
            contract = infer_data_contract(
                entry.frame,
                name=str(payload.get("name") or f"{entry.name} 数据合同"),
                contract_id=contract_id,
                version=version,
                allow_extra_columns=bool(payload.get("allow_extra_columns", True)),
                strict_nulls=bool(payload.get("strict_nulls", False)),
            )
            CONTRACT_STORE.save(contract)
        return {"message": f"已生成并启用数据合同 v{contract.version}", "contract": contract.to_dict()}

    def _contract_validate(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"table_id", "contract_id", "version", "create_issue_table"})
        table_id = str(payload.get("table_id") or SESSION.active_table or "")
        contract_id = str(payload.get("contract_id") or "").strip()
        if not contract_id:
            raise ApiError("请选择数据合同")
        version = int(payload["version"]) if payload.get("version") not in (None, "") else None
        with SESSION.lock:
            entry = SESSION.get(table_id)
            contract = CONTRACT_STORE.load(contract_id, version)
            result = validate_contract(entry.frame, contract)
            issue_table_id = None
            if bool(payload.get("create_issue_table", True)) and result.issues:
                issue_table_id = SESSION.add_table(
                    f"{entry.name}_数据合同检查",
                    issues_frame(result),
                    source=f"数据合同 {contract.name} v{contract.version}",
                )
                SESSION.record(
                    "数据合同校验",
                    f"{contract.name} v{contract.version}：{'通过' if result.passed else '未通过'}",
                    inputs=[entry.id],
                    produced=[issue_table_id],
                    before_rows=len(entry.frame),
                    after_rows=len(result.issues),
                )
        return {
            "message": "数据合同校验通过" if result.passed else "检测到结构漂移或数据质量问题",
            "validation": result.to_dict(),
            "issue_table_id": issue_table_id,
        }

    def _lineage_export(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, set())
        destination = SESSION.output_dir / f"{SESSION.task_id}_处理血缘与审计证据.json"
        LINEAGE_STORE.export_evidence(SESSION.task_id, destination)
        return {"message": "处理血缘与审计证据已生成", "download_url": SESSION.register_download(destination)}

    def _job_submit(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"job_type", "payload", "max_attempts"})
        job_type = str(payload.get("job_type") or "")
        if job_type not in {"contract_validate", "lineage_export"}:
            raise ApiError("该后台任务类型未开放")
        job_payload = payload.get("payload") or {}
        if not isinstance(job_payload, dict):
            raise ApiError("后台任务参数必须是对象")
        job = JOB_ENGINE.submit(
            SESSION.task_id,
            job_type,
            job_payload,
            max_attempts=int(payload.get("max_attempts", 2)),
        )
        return {"message": "后台任务已进入队列", "job": job.to_dict()}

    def _job_cancel(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"job_id"})
        job = JOB_ENGINE.get(str(payload.get("job_id") or ""))
        if job.task_id != SESSION.task_id:
            raise ApiError("不能操作其他客户任务的后台作业", 403)
        return {"message": "取消请求已提交", "job": JOB_ENGINE.cancel(job.job_id).to_dict()}

    def _job_retry(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"job_id"})
        job = JOB_ENGINE.get(str(payload.get("job_id") or ""))
        if job.task_id != SESSION.task_id:
            raise ApiError("不能操作其他客户任务的后台作业", 403)
        return {"message": "后台任务已重新排队", "job": JOB_ENGINE.retry(job.job_id).to_dict()}

    def _database_profile_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(
            payload, {"name", "kind", "path", "connection_string", "dialect", "description", "profile_id"}
        )
        kind = str(payload.get("kind") or "sqlite")
        secret = (
            {"path": str(payload.get("path") or "")}
            if kind == "sqlite"
            else {"connection_string": str(payload.get("connection_string") or "")}
        )
        profile = DATABASE_CONNECTIONS.save(
            name=str(payload.get("name") or ""),
            kind=kind,
            secret=secret,
            dialect=str(payload.get("dialect") or ""),
            description=str(payload.get("description") or ""),
            profile_id=str(payload.get("profile_id") or "") or None,
        )
        return {"message": "数据库连接已加密保存", "profile": profile.to_dict()}

    def _database_profile_delete(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"profile_id"})
        removed = DATABASE_CONNECTIONS.delete(str(payload.get("profile_id") or ""))
        return {"message": "数据库连接已删除" if removed else "数据库连接不存在", "removed": removed}

    def _database_profile_test(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"profile_id"})
        return {"message": "数据库连接成功", "result": DATABASE_CONNECTIONS.test(str(payload.get("profile_id") or ""))}

    def _database_profile_schema(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"profile_id", "maximum_tables"})
        schema = DATABASE_CONNECTIONS.schema(
            str(payload.get("profile_id") or ""),
            maximum_tables=int(payload.get("maximum_tables", 200)),
        )
        return {"message": f"已读取 {len(schema)} 张表或视图的结构", "schema": schema}

    def _database_profile_query(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"profile_id", "sql", "max_rows", "output_name"})
        result = DATABASE_CONNECTIONS.query(
            str(payload.get("profile_id") or ""),
            str(payload.get("sql") or ""),
            max_rows=int(payload.get("max_rows", MAX_ROWS_PER_TABLE)),
        )
        with SESSION.lock:
            table_id = SESSION.add_table(
                str(payload.get("output_name") or "数据库查询结果"), result, source="数据库连接中心只读查询"
            )
            SESSION.record(
                "数据库只读查询",
                "连接凭据未写入任务文件，SQL 已通过只读校验",
                inputs=[],
                produced=[table_id],
                after_rows=len(result),
            )
        return {"message": f"只读查询完成，返回 {len(result):,} 行", "table_id": table_id}

    def _conversation_clear(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, set())
        ConversationStore(SESSION.task_dir / "conversation.json").clear()
        return {"message": "当前任务的连续追问上下文已清空"}

    def _ai_evaluation_run(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"confirm_cost", "maximum_scenarios"})
        if not bool(payload.get("confirm_cost")):
            raise ApiError("AI 回归测试会调用模型，请明确确认费用后执行", 409)
        config = _project_ai_config()
        if not config.get("configured"):
            raise ApiError("项目 DeepSeek API 尚未配置")
        scenarios = AI_SCENARIO_STORE.list()[: max(1, min(int(payload.get("maximum_scenarios", 10)), 50))]
        table_ids = list(SESSION.tables)[:AI_MAX_SELECTED_TABLES]

        def target(scenario: Any) -> dict[str, Any]:
            raw = self._ai_unified({"prompt": scenario.prompt, "table_ids": table_ids})
            plan = raw.get("plan") if isinstance(raw, dict) else None
            result = dict(raw) if isinstance(raw, dict) else {"message": str(raw)}
            if isinstance(plan, dict):
                result.setdefault("status", plan.get("status"))
                result.setdefault("steps", plan.get("steps") or [])
                result.setdefault("summary", plan.get("summary"))
            result.setdefault("status", "ready" if not result.get("error") else "unsupported")
            return result

        report = run_evaluation(scenarios, target, model=str(config.get("model") or "deepseek"))
        destination = SESSION.output_dir / f"{SESSION.task_id}_AI回归测试报告.json"
        destination.write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        return {
            "message": f"AI 回归测试完成：{report.passed} 项通过，{report.failed} 项失败",
            "report": report.to_dict(),
            "download_url": SESSION.register_download(destination),
        }

    def _order_quote(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"request", "table_count", "total_rows", "deadline_hours", "has_sample"})
        try:
            quote = quote_order(
                str(payload.get("request") or ""),
                table_count=int(payload.get("table_count", 1)),
                total_rows=int(payload.get("total_rows", 0)),
                deadline_hours=(
                    float(payload["deadline_hours"]) if payload.get("deadline_hours") not in (None, "") else None
                ),
                has_sample=bool(payload.get("has_sample", False)),
            )
        except (TypeError, ValueError) as exc:
            raise ApiError(f"订单评估参数无效：{exc}") from exc
        return {"message": "接单评估已完成", "quote": quote.to_dict()}

    def _database_query(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"database_file", "sql", "output_name"})
        filename = _safe_filename(str(payload.get("database_file") or ""))
        database = (SESSION.upload_dir / filename).resolve()
        if SESSION.upload_dir.resolve() not in database.parents or database.suffix.lower() not in {
            ".db",
            ".sqlite",
            ".sqlite3",
        }:
            raise ApiError("只允许查询当前任务上传的 SQLite 数据库")
        try:
            result = query_sqlite_read_only(database, str(payload.get("sql") or ""))
        except (FileNotFoundError, RuntimeError, ValueError, sqlite3.Error) as exc:
            raise ApiError(f"只读数据库查询失败：{exc}") from exc
        with SESSION.lock:
            name = _validated_output_name(payload.get("output_name"), fallback="数据库查询结果")
            table_id = SESSION.add_table(name, result, source="SQLite只读查询")
            SESSION.record(
                "数据库只读查询",
                "已在只读事务中执行 SELECT/WITH 查询",
                inputs=[filename],
                produced=[table_id],
                before_rows=None,
                after_rows=len(result),
            )
        return {"message": f"只读查询完成，返回 {len(result):,} 行", "table_id": table_id}

    def _vba_bundle(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"code", "module_name", "entry_macro", "filename"})
        filename = _safe_filename(payload.get("filename") or "VBA安全交付包.zip")
        if not filename.lower().endswith(".zip"):
            filename += ".zip"
        destination = SESSION.output_dir / filename
        try:
            result = build_vba_bundle(
                str(payload.get("code") or ""),
                destination,
                module_name=str(payload.get("module_name") or "BiaogeAutomation"),
                entry_macro=(str(payload["entry_macro"]) if payload.get("entry_macro") else None),
            )
        except (OSError, ValueError) as exc:
            raise ApiError(f"VBA交付包生成失败：{exc}") from exc
        with SESSION.lock:
            url = SESSION.register_download(destination)
            SESSION.record(
                "生成VBA安全交付包",
                "已执行危险指令扫描、哈希和回滚清单",
                inputs=[],
                produced=[],
                before_rows=None,
                after_rows=None,
            )
        return {"message": result.message, "download_url": url, "result": result.to_dict()}

    def _schedule(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(
            payload,
            {
                "action",
                "id",
                "name",
                "kind",
                "expression",
                "recipe_id",
                "task_id",
                "table_id",
                "output_name",
                "enabled",
            },
        )
        action = str(payload.get("action") or "list")
        try:
            if action == "add":
                schedule = SCHEDULER.add(
                    str(payload.get("name") or "定时Excel任务"),
                    str(payload.get("kind") or "daily"),
                    str(payload.get("expression") or "09:00"),
                    "recipe",
                    {
                        "recipe_id": str(payload.get("recipe_id") or ""),
                        "task_id": str(payload.get("task_id") or SESSION.task_id),
                        "table_id": str(payload.get("table_id") or SESSION.active_table or ""),
                        "output_name": str(payload.get("output_name") or "定时任务结果"),
                    },
                )
                message = f"已创建计划任务“{schedule.name}”"
            elif action == "enable":
                SCHEDULER.set_enabled(str(payload.get("id") or ""), bool(payload.get("enabled", True)))
                message = "计划任务状态已更新"
            elif action == "delete":
                SCHEDULER.delete(str(payload.get("id") or ""))
                message = "计划任务已删除"
            elif action == "run_due":
                count = SCHEDULER.run_due()
                message = f"已检查并运行 {count} 个到期任务"
            elif action == "list":
                message = "计划任务已加载"
            else:
                raise ValueError("未知计划任务操作")
        except (KeyError, TypeError, ValueError) as exc:
            raise ApiError(f"计划任务操作失败：{exc}") from exc
        return {"message": message, "schedules": [item.to_dict() for item in SCHEDULER.list()]}

    def _chart_history(self, payload: dict[str, Any]) -> dict[str, Any]:
        _validate_payload_keys(payload, {"action"})
        action = str(payload.get("action") or "list")
        with SESSION.lock:
            if action == "undo":
                if len(SESSION.chart_history) <= 1:
                    raise ApiError("没有更早的图表版本")
                SESSION.chart_redo_stack.append(SESSION.chart_history.pop())
            elif action == "redo":
                if not SESSION.chart_redo_stack:
                    raise ApiError("没有可恢复的图表版本")
                SESSION.chart_history.append(SESSION.chart_redo_stack.pop())
            elif action == "clear":
                SESSION.chart_history.clear()
                SESSION.chart_redo_stack.clear()
            elif action != "list":
                raise ApiError("未知图表历史操作")
            current = SESSION.chart_history[-1] if SESSION.chart_history else None
        return {
            "message": "图表历史已更新",
            "current": current,
            "history_count": len(SESSION.chart_history),
            "can_undo": len(SESSION.chart_history) > 1,
            "can_redo": bool(SESSION.chart_redo_stack),
        }


def _job_contract_validate(context: Any, payload: Any) -> dict[str, Any]:
    """Validate a durable task table without relying on HTTP request context."""

    session = AppSession(context.task_id)
    try:
        context.update(15, "正在读取任务数据")
        table_id = str(payload.get("table_id") or session.active_table or "")
        entry = session.get(table_id)
        contract_id = str(payload.get("contract_id") or "")
        version = int(payload["version"]) if payload.get("version") not in (None, "") else None
        contract = CONTRACT_STORE.load(contract_id, version)
        context.update(45, "正在检查结构漂移与业务规则")
        result = validate_contract(entry.frame, contract)
        destination = (
            session.output_dir / f"{session.task_id}_数据合同校验_{contract.contract_id}_v{contract.version}.json"
        )
        destination.write_text(json.dumps(result.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        context.update(90, "正在写入审计结果")
        LINEAGE_STORE.append_completed(
            task_id=session.task_id,
            job_name="后台数据合同校验",
            inputs=[dataset_metadata(entry.name, entry.frame, source=f"input:{entry.id}")],
            outputs=[],
            parameters={"contract_id": contract.contract_id, "version": contract.version, "passed": result.passed},
        )
        return {"passed": result.passed, "validation": result.to_dict(), "file": destination.name}
    finally:
        session.close()


def _job_lineage_export(context: Any, payload: Any) -> dict[str, Any]:
    del payload
    session = AppSession(context.task_id)
    try:
        context.update(35, "正在汇总操作证据")
        destination = session.output_dir / f"{session.task_id}_处理血缘与审计证据.json"
        LINEAGE_STORE.export_evidence(session.task_id, destination)
        context.update(90, "审计证据已写入交付目录")
        return {"file": destination.name, "run_count": LINEAGE_STORE.evidence(session.task_id)["run_count"]}
    finally:
        session.close()


JOB_ENGINE.register("contract_validate", _job_contract_validate)
JOB_ENGINE.register("lineage_export", _job_lineage_export)


def _urlquote(value: str) -> str:
    from urllib.parse import quote

    return quote(value, safe="")


def find_available_port(preferred: int) -> int:
    for port in range(preferred, preferred + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("无法找到可用本地端口")


def open_local_url(url: str) -> None:
    """Open the local UI with the most reliable Windows mechanism available."""

    try:
        if os.name == "nt":
            os.startfile(url)  # type: ignore[attr-defined]
            return
    except OSError:
        pass
    webbrowser.open(url)


def run_server(*, port: int = 8501, open_browser: bool = True) -> None:
    selected_port = find_available_port(port)
    server = ThreadingHTTPServer(("127.0.0.1", selected_port), ToolboxHandler)
    TASK_REPOSITORY.purge_expired()
    SCHEDULER.start(poll_seconds=30)
    JOB_ENGINE.start()
    url = f"http://127.0.0.1:{selected_port}"
    print("\n表格快处已启动")
    print(f"本地地址：{url}")
    print(f"数据只在本机按任务隔离保存；默认保留 {TASK_RETENTION_DAYS} 天。\n")
    if open_browser:
        threading.Timer(0.8, lambda: open_local_url(url)).start()
    try:
        server.serve_forever(poll_interval=0.25)
    except KeyboardInterrupt:
        print("\n正在关闭表格快处…")
    finally:
        JOB_ENGINE.stop()
        SCHEDULER.stop()
        server.server_close()
        SESSION.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="表格快处本地 Excel 数据处理工作台")
    parser.add_argument("--port", type=int, default=8501, help="首选本地端口")
    parser.add_argument("--no-browser", action="store_true", help="启动时不自动打开浏览器")
    args = parser.parse_args()
    run_server(port=args.port, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
