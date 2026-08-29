"""Deterministic Power BI/Fabric delivery and unattended publishing.

DeepSeek may describe the business intent, but this module owns every executable
artifact.  It builds a self-contained semantic model (embedded Power Query data),
PBIR report pages, a validation report and, when a Microsoft Entra service
principal is configured, publishes both items through the Fabric REST API.
"""

from __future__ import annotations

from dataclasses import dataclass
import base64
import csv
from datetime import datetime
import hashlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import time
from typing import Any, Callable, Mapping
from urllib import error as urllib_error
from urllib import parse as urllib_parse
from urllib import request as urllib_request
import uuid
import zipfile
import zlib

import pandas as pd


FABRIC_API_ROOT = "https://api.fabric.microsoft.com/v1"
FABRIC_TOKEN_SCOPE = "https://api.fabric.microsoft.com/.default"
POWER_BI_ENV_KEYS = (
    "POWER_BI_TENANT_ID",
    "POWER_BI_CLIENT_ID",
    "POWER_BI_CLIENT_SECRET",
    "POWER_BI_WORKSPACE_ID",
)
_GUID = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
    re.IGNORECASE,
)


class PowerBIAutomationError(RuntimeError):
    """A safe, user-facing Power BI automation failure."""


@dataclass(frozen=True)
class PowerBIConfig:
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    workspace_id: str = ""

    @property
    def missing(self) -> list[str]:
        values = {
            "POWER_BI_TENANT_ID": self.tenant_id,
            "POWER_BI_CLIENT_ID": self.client_id,
            "POWER_BI_CLIENT_SECRET": self.client_secret,
            "POWER_BI_WORKSPACE_ID": self.workspace_id,
        }
        return [name for name, value in values.items() if not value]

    @property
    def configured(self) -> bool:
        return not self.missing

    @classmethod
    def from_environment(
        cls,
        *,
        env: Mapping[str, str] | None = None,
        env_file: Path | None = None,
    ) -> "PowerBIConfig":
        values: dict[str, str] = {}
        if env_file and env_file.exists():
            try:
                for raw in env_file.read_text(encoding="utf-8-sig").splitlines():
                    line = raw.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    name, value = line.split("=", 1)
                    if name.strip() in POWER_BI_ENV_KEYS:
                        values[name.strip()] = value.strip().strip('"').strip("'")
            except OSError:
                values = {}
        source = os.environ if env is None else env
        for name in POWER_BI_ENV_KEYS:
            external = str(source.get(name, "")).strip()
            if external:
                values[name] = external
        config = cls(
            tenant_id=values.get("POWER_BI_TENANT_ID", ""),
            client_id=values.get("POWER_BI_CLIENT_ID", ""),
            client_secret=values.get("POWER_BI_CLIENT_SECRET", ""),
            workspace_id=values.get("POWER_BI_WORKSPACE_ID", ""),
        )
        for label, value in (
            ("tenant_id", config.tenant_id),
            ("client_id", config.client_id),
            ("workspace_id", config.workspace_id),
        ):
            if value and not _GUID.fullmatch(value):
                raise PowerBIAutomationError(f"{label} 必须是 Microsoft Entra GUID")
        if len(config.client_secret) > 4096 or any(ch in config.client_secret for ch in "\r\n"):
            raise PowerBIAutomationError("POWER_BI_CLIENT_SECRET 格式无效")
        return config

    def public_status(self) -> dict[str, Any]:
        return {
            "configured": self.configured,
            "missing": self.missing,
            "workspace_id": self.workspace_id if self.workspace_id else None,
        }


def fallback_power_bi_brief(prompt: str) -> dict[str, Any]:
    """Return a stable brief if the model emits an incompatible JSON shape."""

    return {
        "category": "power_bi",
        "normalized_request": str(prompt).strip()[:800],
        "scope": "将当前销售数据转换为可发布的 Power BI 星型模型、DAX 指标和三页交互报表，并执行本地自动验收。",
        "clarification_questions": [],
        "deliverables": [
            "自包含语义模型（FactSales 与五张维度表）",
            "Power Query M、核心 DAX、PBIR 三页报表定义",
            "Fabric REST 发布载荷、校验报告和自动化测试记录",
        ],
        "implementation_steps": [
            "识别字段并清洗订单主键、日期、金额、数量和成本",
            "生成 DimDate、DimCustomer、DimRegion、DimChannel、DimProduct",
            "建立一对多关系并创建核心 DAX 度量值",
            "生成管理概览、产品渠道、客户明细三页 PBIR 报表",
            "检查主外键、视觉字段角色、页面边界和发布载荷",
            "存在服务主体配置时发布语义模型和报表并回读验证",
        ],
        "artifacts": [],
        "test_checklist": [
            "维度主键全部唯一且事实表外键无孤儿记录",
            "总销售额、总成本、总利润与利润率可计算",
            "全部页面和视觉对象引用存在的字段或度量值",
            "发布请求不包含 DeepSeek 或 Microsoft 客户端密钥",
        ],
        "risks": ["首次发布仍需由租户管理员创建服务主体并授予目标工作区 Contributor 权限。"],
        "human_approval_points": [],
    }


_COLUMN_CANDIDATES: dict[str, tuple[str, ...]] = {
    "order": ("订单编号", "订单号", "orderid", "ordernumber", "order_no", "id"),
    "date": ("日期", "订单日期", "date", "orderdate", "createdat"),
    "customer": ("客户", "客户名称", "customer", "customername"),
    "region": ("地区", "区域", "region", "area"),
    "channel": ("渠道", "销售渠道", "channel"),
    "product": ("产品", "产品名称", "product", "productname"),
    "sales": ("订单金额", "销售额", "金额", "salesamount", "sales", "revenue"),
    "quantity": ("数量", "销量", "quantity", "qty", "units"),
    "cost": ("成本", "成本金额", "costamount", "cost"),
}


def _normalise_column_name(value: Any) -> str:
    return re.sub(r"[\s_\-./\\（）()]+", "", str(value)).casefold()


def _find_column(frame: pd.DataFrame, role: str) -> str:
    lookup = {_normalise_column_name(column): str(column) for column in frame.columns}
    for candidate in _COLUMN_CANDIDATES[role]:
        key = _normalise_column_name(candidate)
        if key in lookup:
            return lookup[key]
    raise PowerBIAutomationError(
        f"无法识别{role}字段；可识别字段示例：{'、'.join(_COLUMN_CANDIDATES[role][:4])}"
    )


def _clean_text(series: pd.Series) -> pd.Series:
    result = series.astype("string").str.replace(r"[\r\n\t]+", " ", regex=True).str.strip()
    return result.mask(result.isna() | result.eq(""), "(未知)")


def _star_tables(frame: pd.DataFrame) -> tuple[dict[str, pd.DataFrame], dict[str, Any]]:
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        raise PowerBIAutomationError("Power BI 自动化至少需要一行数据")
    columns = {role: _find_column(frame, role) for role in _COLUMN_CANDIDATES}
    source = pd.DataFrame(
        {
            "OrderNumber": _clean_text(frame[columns["order"]]),
            "Date": pd.to_datetime(frame[columns["date"]], errors="coerce"),
            "Customer": _clean_text(frame[columns["customer"]]),
            "Region": _clean_text(frame[columns["region"]]),
            "Channel": _clean_text(frame[columns["channel"]]),
            "Product": _clean_text(frame[columns["product"]]),
            "SalesAmount": pd.to_numeric(frame[columns["sales"]], errors="coerce"),
            "Quantity": pd.to_numeric(frame[columns["quantity"]], errors="coerce"),
            "CostAmount": pd.to_numeric(frame[columns["cost"]], errors="coerce"),
        }
    )
    invalid_dates = int(source["Date"].isna().sum())
    if invalid_dates:
        raise PowerBIAutomationError(f"日期字段有 {invalid_dates} 行无法解析，不能建立 DimDate")
    source = source.drop_duplicates(subset=["OrderNumber"], keep="first").reset_index(drop=True)
    source["SalesAmount"] = source["SalesAmount"].fillna(0.0).astype(float)
    source["CostAmount"] = source["CostAmount"].fillna(0.0).astype(float)
    source["Quantity"] = source["Quantity"].fillna(0).round().astype("int64")

    dim_specs = (
        ("Customer", "CustomerKey", "DimCustomer"),
        ("Region", "RegionKey", "DimRegion"),
        ("Channel", "ChannelKey", "DimChannel"),
        ("Product", "ProductKey", "DimProduct"),
    )
    tables: dict[str, pd.DataFrame] = {}
    for value_column, key_column, table_name in dim_specs:
        values = sorted(source[value_column].astype(str).unique().tolist())
        dimension = pd.DataFrame({key_column: range(1, len(values) + 1), value_column: values})
        tables[table_name] = dimension
        key_map = dict(zip(dimension[value_column], dimension[key_column]))
        source[key_column] = source[value_column].map(key_map).astype("int64")

    min_date = source["Date"].min().normalize()
    max_date = source["Date"].max().normalize()
    dates = pd.date_range(min_date, max_date, freq="D")
    dim_date = pd.DataFrame({"Date": dates})
    dim_date["DateKey"] = dim_date["Date"].dt.strftime("%Y%m%d").astype("int64")
    dim_date["Year"] = dim_date["Date"].dt.year.astype("int64")
    dim_date["Quarter"] = "Q" + dim_date["Date"].dt.quarter.astype(str)
    dim_date["MonthNumber"] = dim_date["Date"].dt.month.astype("int64")
    dim_date["Month"] = dim_date["Date"].dt.strftime("%m月")
    dim_date["YearMonth"] = dim_date["Date"].dt.strftime("%Y-%m")
    tables["DimDate"] = dim_date[
        ["DateKey", "Date", "Year", "Quarter", "MonthNumber", "Month", "YearMonth"]
    ]
    source["DateKey"] = source["Date"].dt.strftime("%Y%m%d").astype("int64")
    source["ProfitAmount"] = source["SalesAmount"] - source["CostAmount"]
    tables["FactSales"] = source[
        [
            "OrderNumber",
            "DateKey",
            "CustomerKey",
            "RegionKey",
            "ChannelKey",
            "ProductKey",
            "SalesAmount",
            "Quantity",
            "CostAmount",
            "ProfitAmount",
        ]
    ]
    return tables, {"source_columns": columns, "source_rows": len(frame), "fact_rows": len(source)}


def _json_write(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _csv_bytes(frame: pd.DataFrame) -> bytes:
    buffer = io.StringIO(newline="")
    frame.to_csv(buffer, index=False, lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
    return buffer.getvalue().encode("utf-8-sig")


def _embedded_m(frame: pd.DataFrame) -> list[str]:
    compressed = zlib.compress(_csv_bytes(frame), level=9)
    encoded = base64.b64encode(compressed).decode("ascii")
    type_map: dict[str, str] = {}
    for column in frame.columns:
        dtype = frame[column].dtype
        if pd.api.types.is_datetime64_any_dtype(dtype):
            type_map[str(column)] = "type date"
        elif pd.api.types.is_integer_dtype(dtype):
            type_map[str(column)] = "Int64.Type"
        elif pd.api.types.is_numeric_dtype(dtype):
            type_map[str(column)] = "type number"
        else:
            type_map[str(column)] = "type text"
    transforms = ", ".join(f'{{"{name}", {kind}}}' for name, kind in type_map.items())
    return [
        "let",
        f'    BinaryData = Binary.Decompress(Binary.FromText("{encoded}", BinaryEncoding.Base64), Compression.Deflate),',
        f'    Source = Csv.Document(BinaryData, [Delimiter=",", Columns={len(frame.columns)}, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),',
        "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),",
        f"    Typed = Table.TransformColumnTypes(PromotedHeaders, {{{transforms}}})",
        "in",
        "    Typed",
    ]


def _model_data_type(series: pd.Series) -> str:
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return "dateTime"
    if pd.api.types.is_integer_dtype(series.dtype):
        return "int64"
    if pd.api.types.is_numeric_dtype(series.dtype):
        return "double"
    return "string"


_MEASURES = [
    ("Total Sales", "SUM('FactSales'[SalesAmount])", "¥#,0.00"),
    ("Total Cost", "SUM('FactSales'[CostAmount])", "¥#,0.00"),
    ("Total Profit", "[Total Sales] - [Total Cost]", "¥#,0.00"),
    ("Profit Margin", "DIVIDE([Total Profit], [Total Sales])", "0.00%"),
    ("Order Count", "DISTINCTCOUNT('FactSales'[OrderNumber])", "#,0"),
    ("Units Sold", "SUM('FactSales'[Quantity])", "#,0"),
    ("Average Order Value", "DIVIDE([Total Sales], [Order Count])", "¥#,0.00"),
    ("Sales YTD", "TOTALYTD([Total Sales], 'DimDate'[Date])", "¥#,0.00"),
    ("Sales LY", "CALCULATE([Total Sales], SAMEPERIODLASTYEAR('DimDate'[Date]))", "¥#,0.00"),
    ("Sales YoY %", "DIVIDE([Total Sales] - [Sales LY], [Sales LY])", "0.00%"),
]


def _semantic_model(tables: Mapping[str, pd.DataFrame]) -> dict[str, Any]:
    model_tables: list[dict[str, Any]] = []
    for table_name, frame in tables.items():
        table: dict[str, Any] = {
            "name": table_name,
            "columns": [
                {
                    "name": str(column),
                    "dataType": _model_data_type(frame[column]),
                    "sourceColumn": str(column),
                }
                for column in frame.columns
            ],
            "partitions": [
                {
                    "name": table_name,
                    "mode": "import",
                    "source": {"type": "m", "expression": _embedded_m(frame)},
                }
            ],
        }
        if table_name == "FactSales":
            table["measures"] = [
                {"name": name, "expression": expression, "formatString": format_string}
                for name, expression, format_string in _MEASURES
            ]
        model_tables.append(table)
    relationships = []
    for dimension, key in (
        ("DimDate", "DateKey"),
        ("DimCustomer", "CustomerKey"),
        ("DimRegion", "RegionKey"),
        ("DimChannel", "ChannelKey"),
        ("DimProduct", "ProductKey"),
    ):
        relationships.append(
            {
                "name": str(uuid.uuid5(uuid.NAMESPACE_URL, f"biaoge:{dimension}:{key}")),
                "fromTable": "FactSales",
                "fromColumn": key,
                "toTable": dimension,
                "toColumn": key,
                "crossFilteringBehavior": "oneDirection",
            }
        )
    return {
        "compatibilityLevel": 1601,
        "model": {
            "culture": "zh-CN",
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "zh-CN",
            "tables": model_tables,
            "relationships": relationships,
            "annotations": [{"name": "PBI_QueryOrder", "value": json.dumps(list(tables), ensure_ascii=False)}],
        },
    }


def _visual_id(seed: str) -> str:
    return hashlib.sha1(seed.encode("utf-8")).hexdigest()[:20]


def _page_id(seed: str) -> str:
    return "ReportSection" + hashlib.sha1(seed.encode("utf-8")).hexdigest()[:24]


def _column_projection(table: str, column: str) -> dict[str, Any]:
    return {
        "field": {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}},
        "queryRef": f"{table}.{column}",
        "nativeQueryRef": column,
    }


def _measure_projection(measure: str) -> dict[str, Any]:
    return {
        "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "FactSales"}}, "Property": measure}},
        "queryRef": f"FactSales.{measure}",
        "nativeQueryRef": measure,
    }


def _visual(
    *,
    seed: str,
    visual_type: str,
    roles: Mapping[str, list[dict[str, Any]]],
    x: int,
    y: int,
    width: int,
    height: int,
    z: int,
    objects: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    visual: dict[str, Any] = {
        "visualType": visual_type,
        "query": {"queryState": {role: {"projections": fields} for role, fields in roles.items()}},
    }
    if objects:
        visual["objects"] = dict(objects)
    return {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json",
        "name": _visual_id(seed),
        "position": {"x": x, "y": y, "z": z, "height": height, "width": width, "tabOrder": z},
        "visual": visual,
    }


def _slicer(seed: str, table: str, column: str, x: int, y: int, width: int, z: int) -> dict[str, Any]:
    return _visual(
        seed=seed,
        visual_type="slicer",
        roles={"Values": [_column_projection(table, column)]},
        x=x,
        y=y,
        width=width,
        height=80,
        z=z,
        objects={
            "data": [{"properties": {"mode": {"expr": {"Literal": {"Value": "'Dropdown'"}}}}}],
            "header": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}, "text": {"expr": {"Literal": {"Value": f"'{column}'"}}}}}],
        },
    )


def _report_pages() -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []

    overview = [
        _visual(seed="overview-kpis", visual_type="cardVisual", roles={"Data": [_measure_projection(item) for item in ("Total Sales", "Total Profit", "Profit Margin", "Order Count")]}, x=20, y=20, width=900, height=120, z=1000),
        _slicer("overview-year", "DimDate", "Year", 940, 20, 150, 2000),
        _slicer("overview-region", "DimRegion", "Region", 1110, 20, 150, 3000),
        _visual(seed="overview-trend", visual_type="lineChart", roles={"Category": [_column_projection("DimDate", "YearMonth")], "Y": [_measure_projection("Total Sales"), _measure_projection("Total Profit")]}, x=20, y=165, width=760, height=520, z=4000),
        _visual(seed="overview-region-chart", visual_type="clusteredColumnChart", roles={"Category": [_column_projection("DimRegion", "Region")], "Y": [_measure_projection("Total Sales")]}, x=800, y=165, width=460, height=520, z=5000),
    ]
    pages.append({"display_name": "管理概览", "visuals": overview})

    product = [
        _slicer("product-year", "DimDate", "Year", 20, 20, 160, 1000),
        _slicer("product-channel", "DimChannel", "Channel", 200, 20, 200, 2000),
        _visual(seed="product-sales", visual_type="clusteredBarChart", roles={"Category": [_column_projection("DimProduct", "Product")], "Y": [_measure_projection("Total Sales"), _measure_projection("Total Profit")]}, x=20, y=125, width=610, height=560, z=3000),
        _visual(seed="channel-sales", visual_type="clusteredColumnChart", roles={"Category": [_column_projection("DimChannel", "Channel")], "Y": [_measure_projection("Total Sales")]}, x=650, y=125, width=610, height=270, z=4000),
        _visual(seed="channel-margin", visual_type="lineChart", roles={"Category": [_column_projection("DimChannel", "Channel")], "Y": [_measure_projection("Profit Margin")]}, x=650, y=415, width=610, height=270, z=5000),
    ]
    pages.append({"display_name": "产品与渠道", "visuals": product})

    detail_values = [
        _column_projection("DimCustomer", "Customer"),
        _measure_projection("Total Sales"),
        _measure_projection("Total Profit"),
        _measure_projection("Order Count"),
        _measure_projection("Average Order Value"),
    ]
    customer = [
        _slicer("customer-region", "DimRegion", "Region", 20, 20, 180, 1000),
        _slicer("customer-product", "DimProduct", "Product", 220, 20, 220, 2000),
        _visual(seed="customer-table", visual_type="tableEx", roles={"Values": detail_values}, x=20, y=125, width=1240, height=560, z=3000, objects={"columnHeaders": [{"properties": {"columnAdjustment": {"expr": {"Literal": {"Value": "'growToFit'"}}}, "autoSizeColumnWidth": {"expr": {"Literal": {"Value": "true"}}}}}]}),
    ]
    pages.append({"display_name": "客户明细", "visuals": customer})
    return pages


def _write_report_definition(root: Path, *, model_folder_name: str) -> dict[str, Any]:
    definition = root / "definition"
    pages_root = definition / "pages"
    pages = _report_pages()
    page_ids: list[str] = []
    visual_count = 0
    for page_number, page in enumerate(pages, start=1):
        page_id = _page_id(f"{page_number}:{page['display_name']}")
        page_ids.append(page_id)
        page_root = pages_root / page_id
        _json_write(
            page_root / "page.json",
            {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
                "name": page_id,
                "displayName": page["display_name"],
                "displayOption": "FitToPage",
                "height": 720,
                "width": 1280,
            },
        )
        for visual in page["visuals"]:
            _json_write(page_root / "visuals" / visual["name"] / "visual.json", visual)
            visual_count += 1
    _json_write(
        pages_root / "pages.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
            "pageOrder": page_ids,
            "activePageName": page_ids[0],
        },
    )
    _json_write(
        definition / "version.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
            "version": "2.0.0",
        },
    )
    _json_write(
        definition / "report.json",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.1.0/schema.json",
            "themeCollection": {"baseTheme": {"name": "CY25SU12", "reportVersionAtImport": {"visual": "2.5.0", "report": "3.1.0", "page": "2.3.0"}, "type": "SharedResources"}},
            "layoutOptimization": "None",
            "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True},
        },
    )
    _json_write(
        root / "definition.pbir",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
            "version": "4.0",
            "datasetReference": {"byPath": {"path": f"../{model_folder_name}"}},
        },
    )
    return {"page_ids": page_ids, "pages": pages, "visual_count": visual_count}


def _validate_bundle(
    *,
    tables: Mapping[str, pd.DataFrame],
    model: Mapping[str, Any],
    report_root: Path,
    report_meta: Mapping[str, Any],
) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []

    def check(name: str, passed: bool, detail: str) -> None:
        checks.append({"name": name, "passed": bool(passed), "detail": detail})

    fact = tables["FactSales"]
    for dimension, key in (
        ("DimDate", "DateKey"),
        ("DimCustomer", "CustomerKey"),
        ("DimRegion", "RegionKey"),
        ("DimChannel", "ChannelKey"),
        ("DimProduct", "ProductKey"),
    ):
        dim = tables[dimension]
        unique = bool(dim[key].is_unique and dim[key].notna().all())
        orphan = set(fact[key].dropna().tolist()) - set(dim[key].dropna().tolist())
        check(f"{dimension} 主键唯一", unique, f"{len(dim)} 个唯一键")
        check(f"FactSales.{key} 无孤儿", not orphan, f"孤儿键 {len(orphan)} 个")

    model_tables = {item["name"]: item for item in model["model"]["tables"]}
    measure_names = {item["name"] for item in model_tables["FactSales"].get("measures", [])}
    check("核心 DAX 齐全", {"Total Sales", "Total Profit", "Profit Margin", "Sales YoY %"} <= measure_names, f"共 {len(measure_names)} 个度量值")
    check("星型关系齐全", len(model["model"].get("relationships", [])) == 5, "五条单向多对一关系")

    known_columns = {name: set(map(str, frame.columns)) for name, frame in tables.items()}
    required_roles = {
        "cardVisual": {"Data"},
        "lineChart": {"Category", "Y"},
        "clusteredColumnChart": {"Category", "Y"},
        "clusteredBarChart": {"Category", "Y"},
        "slicer": {"Values"},
        "tableEx": {"Values"},
    }
    visual_errors: list[str] = []
    for path in report_root.glob("definition/pages/*/visuals/*/visual.json"):
        payload = json.loads(path.read_text(encoding="utf-8"))
        position = payload["position"]
        if position["x"] + position["width"] > 1280 or position["y"] + position["height"] > 720:
            visual_errors.append(f"{payload['name']}:超出页面")
        visual = payload["visual"]
        roles = visual.get("query", {}).get("queryState", {})
        if not required_roles.get(visual["visualType"], set()) <= set(roles):
            visual_errors.append(f"{payload['name']}:字段角色不完整")
        for role in roles.values():
            for projection in role.get("projections", []):
                field = projection.get("field", {})
                if "Column" in field:
                    node = field["Column"]
                    table = node["Expression"]["SourceRef"]["Entity"]
                    if node["Property"] not in known_columns.get(table, set()):
                        visual_errors.append(f"{payload['name']}:字段不存在")
                elif "Measure" in field:
                    if field["Measure"]["Property"] not in measure_names:
                        visual_errors.append(f"{payload['name']}:度量值不存在")
    check("PBIR 视觉字段和角色有效", not visual_errors, "；".join(visual_errors) or f"{report_meta['visual_count']} 个视觉对象全部通过")
    check("PBIR 页面完整", len(report_meta["page_ids"]) == 3, "管理概览、产品与渠道、客户明细")
    passed = all(item["passed"] for item in checks)
    return {
        "passed": passed,
        "checked_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "summary": {"checks": len(checks), "passed": sum(item["passed"] for item in checks), "failed": sum(not item["passed"] for item in checks)},
        "checks": checks,
    }


def build_power_bi_bundle(
    frame: pd.DataFrame,
    output_dir: Path,
    *,
    task_id: str,
    source_name: str,
    engineering_brief: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Build and validate an isolated, self-contained Power BI delivery ZIP."""

    safe_task = re.sub(r"[^A-Za-z0-9_-]+", "_", str(task_id))[:80] or "task"
    package_name = f"PowerBI_自动化交付_{safe_task}"
    package_root = output_dir / package_name
    if package_root.exists():
        shutil.rmtree(package_root)
    package_root.mkdir(parents=True, exist_ok=False)
    tables, mapping = _star_tables(frame)
    data_root = package_root / "data"
    for name, table in tables.items():
        data_root.mkdir(parents=True, exist_ok=True)
        (data_root / f"{name}.csv").write_bytes(_csv_bytes(table))

    project_name = "SalesAutomation"
    model_root = package_root / f"{project_name}.SemanticModel"
    report_root = package_root / f"{project_name}.Report"
    model = _semantic_model(tables)
    _json_write(model_root / "model.bim", model)
    _json_write(
        model_root / "definition.pbism",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.2",
            "settings": {"qnaEnabled": True},
        },
    )
    report_meta = _write_report_definition(report_root, model_folder_name=model_root.name)
    platform_schema = "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json"
    _json_write(
        model_root / ".platform",
        {
            "$schema": platform_schema,
            "metadata": {"type": "SemanticModel", "displayName": f"{project_name} 模型"},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"biaoge:{safe_task}:model"))},
        },
    )
    _json_write(
        report_root / ".platform",
        {
            "$schema": platform_schema,
            "metadata": {"type": "Report", "displayName": f"{project_name} 报表"},
            "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, f"biaoge:{safe_task}:report"))},
        },
    )
    _json_write(
        package_root / f"{project_name}.pbip",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
            "version": "1.0",
            "artifacts": [{"report": {"path": report_root.name}}],
            "settings": {"enableAutoRecovery": True},
        },
    )

    dax_text = "\n\n".join(f"{name} =\n{expression}\n-- Format: {fmt}" for name, expression, fmt in _MEASURES)
    (package_root / "DAX_核心指标.dax").write_text(dax_text, encoding="utf-8")
    pq_root = package_root / "PowerQuery"
    for name, table in tables.items():
        pq_root.mkdir(parents=True, exist_ok=True)
        (pq_root / f"{name}.m").write_text("\n".join(_embedded_m(table)), encoding="utf-8")

    model_spec = {
        "schema_version": 1,
        "source_name": source_name,
        "mapping": mapping,
        "tables": {name: {"rows": len(table), "columns": list(map(str, table.columns))} for name, table in tables.items()},
        "relationships": model["model"]["relationships"],
        "measures": [{"name": name, "expression": expression, "format": fmt} for name, expression, fmt in _MEASURES],
    }
    report_spec = {
        "canvas": {"width": 1280, "height": 720},
        "pages": [
            {"name": page["display_name"], "visual_types": [item["visual"]["visualType"] for item in page["visuals"]], "visual_count": len(page["visuals"])}
            for page in report_meta["pages"]
        ],
        "slicers": ["DimDate.Year", "DimRegion.Region", "DimChannel.Channel", "DimProduct.Product"],
    }
    _json_write(package_root / "model_spec.json", model_spec)
    _json_write(package_root / "report_spec.json", report_spec)
    if engineering_brief:
        _json_write(package_root / "AI_需求说明.json", dict(engineering_brief))

    validation = _validate_bundle(tables=tables, model=model, report_root=report_root, report_meta=report_meta)
    _json_write(package_root / "validation_report.json", validation)
    if not validation["passed"]:
        raise PowerBIAutomationError("Power BI 交付包本地自动校验失败")

    env_example = "\n".join(
        [
            "# 仅保存在本机 excel_data_toolbox/.env；不要上传或打包真实密钥",
            "POWER_BI_TENANT_ID=00000000-0000-0000-0000-000000000000",
            "POWER_BI_CLIENT_ID=00000000-0000-0000-0000-000000000000",
            "POWER_BI_CLIENT_SECRET=在此填写服务主体密钥",
            "POWER_BI_WORKSPACE_ID=00000000-0000-0000-0000-000000000000",
            "",
        ]
    )
    (package_root / "PowerBI_发布配置.example.env").write_text(env_example, encoding="utf-8")
    readme = f"""# Power BI 全自动交付包

- 任务：{task_id}
- 数据源：{source_name}
- 事实表：{len(tables['FactSales']):,} 行
- 维度表：DimDate、DimCustomer、DimRegion、DimChannel、DimProduct
- 度量值：{len(_MEASURES)} 个
- 报表：3 页、{report_meta['visual_count']} 个视觉对象
- 本地校验：{validation['summary']['passed']}/{validation['summary']['checks']} 通过

该包的数据已嵌入 Power Query M，不依赖本机 CSV 路径。程序检测到完整的
Microsoft Entra 服务主体配置后，会自动调用 Fabric REST API 创建语义模型和
PBIR 报表并回读验证。首次必须由租户管理员合法创建服务主体、允许 Fabric API，
并把服务主体加入目标工作区 Contributor；任何 AI 都不能绕过这一步。
"""
    (package_root / "README.md").write_text(readme, encoding="utf-8")

    zip_path = output_dir / f"{package_name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(package_root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(package_root.parent))
    return {
        "package_root": package_root,
        "zip_path": zip_path,
        "model_root": model_root,
        "report_root": report_root,
        "validation": validation,
        "model_spec": model_spec,
        "report_spec": report_spec,
    }


def _definition_parts(root: Path, *, semantic_model_id: str | None = None) -> list[dict[str, str]]:
    parts: list[dict[str, str]] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file() or any(part.startswith(".") for part in path.relative_to(root).parts):
            continue
        relative = path.relative_to(root).as_posix()
        raw = path.read_bytes()
        if relative == "definition.pbir" and semantic_model_id:
            payload = json.loads(raw.decode("utf-8"))
            payload["datasetReference"] = {"byConnection": {"connectionString": f"semanticmodelid={semantic_model_id}"}}
            raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        parts.append({"path": relative, "payload": base64.b64encode(raw).decode("ascii"), "payloadType": "InlineBase64"})
    return parts


@dataclass
class HttpResult:
    status: int
    headers: dict[str, str]
    payload: Any


def _default_transport(method: str, url: str, headers: Mapping[str, str], body: bytes | None, timeout: int) -> HttpResult:
    request = urllib_request.Request(url, data=body, method=method, headers=dict(headers))
    try:
        with urllib_request.urlopen(request, timeout=timeout) as response:
            raw = response.read(5_000_000)
            payload = json.loads(raw.decode("utf-8")) if raw else None
            return HttpResult(int(response.status), {str(k).lower(): str(v) for k, v in response.headers.items()}, payload)
    except urllib_error.HTTPError as exc:
        raw = exc.read(1_000_000)
        try:
            payload = json.loads(raw.decode("utf-8")) if raw else None
        except (UnicodeDecodeError, json.JSONDecodeError):
            payload = {"message": "Microsoft Fabric 返回了非 JSON 错误"}
        return HttpResult(int(exc.code), {str(k).lower(): str(v) for k, v in exc.headers.items()}, payload)
    except (urllib_error.URLError, TimeoutError, OSError) as exc:
        raise PowerBIAutomationError(f"无法连接 Microsoft Fabric：{getattr(exc, 'reason', exc)}") from None


class FabricPublisher:
    """Small Fabric REST client with injectable transport for full mock testing."""

    def __init__(
        self,
        config: PowerBIConfig,
        *,
        transport: Callable[[str, str, Mapping[str, str], bytes | None, int], HttpResult] = _default_transport,
        timeout_seconds: int = 60,
        max_polls: int = 40,
    ) -> None:
        if not config.configured:
            raise PowerBIAutomationError(f"Power BI 发布配置不完整：{', '.join(config.missing)}")
        self.config = config
        self.transport = transport
        self.timeout_seconds = max(10, min(int(timeout_seconds), 180))
        self.max_polls = max(1, min(int(max_polls), 120))
        self._token = ""

    def _token_value(self) -> str:
        if self._token:
            return self._token
        body = urllib_parse.urlencode(
            {
                "client_id": self.config.client_id,
                "client_secret": self.config.client_secret,
                "grant_type": "client_credentials",
                "scope": FABRIC_TOKEN_SCOPE,
            }
        ).encode("ascii")
        url = f"https://login.microsoftonline.com/{self.config.tenant_id}/oauth2/v2.0/token"
        response = self.transport("POST", url, {"Content-Type": "application/x-www-form-urlencoded"}, body, self.timeout_seconds)
        token = response.payload.get("access_token") if isinstance(response.payload, dict) else None
        if response.status != 200 or not isinstance(token, str) or not token:
            raise PowerBIAutomationError("Microsoft Entra 服务主体认证失败，请检查租户、客户端和密钥权限")
        self._token = token
        return token

    def _request(self, method: str, url: str, payload: Any = None) -> HttpResult:
        headers = {"Authorization": f"Bearer {self._token_value()}", "Content-Type": "application/json"}
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8") if payload is not None else None
        response = self.transport(method, url, headers, body, self.timeout_seconds)
        if response.status >= 400:
            message = ""
            if isinstance(response.payload, dict):
                message = str(response.payload.get("message") or response.payload.get("errorCode") or "")
            raise PowerBIAutomationError(f"Microsoft Fabric 请求失败（HTTP {response.status}）{f'：{message}' if message else ''}")
        return response

    def _complete_create(self, response: HttpResult) -> dict[str, Any]:
        if response.status == 201 and isinstance(response.payload, dict) and response.payload.get("id"):
            return response.payload
        if response.status != 202:
            raise PowerBIAutomationError(f"Fabric 创建项目返回了意外状态 HTTP {response.status}")
        location = response.headers.get("location")
        operation_id = response.headers.get("x-ms-operation-id")
        if not location and operation_id:
            location = f"{FABRIC_API_ROOT}/operations/{operation_id}"
        if not location:
            raise PowerBIAutomationError("Fabric 长任务未返回 Location")
        for _ in range(self.max_polls):
            poll = self._request("GET", location)
            payload = poll.payload if isinstance(poll.payload, dict) else {}
            status = str(payload.get("status") or "").casefold()
            if payload.get("id") and not status:
                return payload
            if status == "failed":
                raise PowerBIAutomationError("Fabric 长任务执行失败")
            if status == "succeeded":
                result_url = poll.headers.get("location")
                if not result_url and operation_id:
                    result_url = f"{FABRIC_API_ROOT}/operations/{operation_id}/result"
                if result_url:
                    result = self._request("GET", result_url)
                    if isinstance(result.payload, dict) and result.payload.get("id"):
                        return result.payload
                if payload.get("id"):
                    return payload
                raise PowerBIAutomationError("Fabric 长任务成功但未返回项目 ID")
            retry = min(max(int(poll.headers.get("retry-after", "1") or "1"), 0), 10)
            if retry:
                time.sleep(retry)
        raise PowerBIAutomationError("Fabric 长任务等待超时")

    def publish(self, *, model_root: Path, report_root: Path, display_name: str) -> dict[str, Any]:
        workspace = self.config.workspace_id
        semantic_response = self._request(
            "POST",
            f"{FABRIC_API_ROOT}/workspaces/{workspace}/semanticModels",
            {
                "displayName": f"{display_name} 模型",
                "description": "由表格快处 AI 自动创建并完成本地校验",
                "definition": {"parts": _definition_parts(model_root)},
            },
        )
        semantic = self._complete_create(semantic_response)
        semantic_id = str(semantic["id"])
        report_response = self._request(
            "POST",
            f"{FABRIC_API_ROOT}/workspaces/{workspace}/reports",
            {
                "displayName": f"{display_name} 报表",
                "description": "由表格快处 AI 自动创建并完成本地校验",
                "definition": {"parts": _definition_parts(report_root, semantic_model_id=semantic_id)},
            },
        )
        report = self._complete_create(report_response)
        report_id = str(report["id"])
        verify_model = self._request("GET", f"{FABRIC_API_ROOT}/workspaces/{workspace}/items/{semantic_id}")
        verify_report = self._request("GET", f"{FABRIC_API_ROOT}/workspaces/{workspace}/items/{report_id}")
        if not isinstance(verify_model.payload, dict) or str(verify_model.payload.get("id")) != semantic_id:
            raise PowerBIAutomationError("语义模型已创建但回读验证失败")
        if not isinstance(verify_report.payload, dict) or str(verify_report.payload.get("id")) != report_id:
            raise PowerBIAutomationError("报表已创建但回读验证失败")
        return {
            "status": "published",
            "workspace_id": workspace,
            "semantic_model_id": semantic_id,
            "report_id": report_id,
            "report_url": f"https://app.powerbi.com/groups/{workspace}/reports/{report_id}",
            "verified": True,
        }


def publish_bundle_if_configured(
    bundle: Mapping[str, Any],
    *,
    config: PowerBIConfig,
    display_name: str,
    publisher_factory: Callable[[PowerBIConfig], FabricPublisher] = FabricPublisher,
) -> dict[str, Any]:
    if not config.configured:
        return {
            "status": "credentials_required",
            "published": False,
            "message": "自动化交付包已生成并通过校验；未检测到微软服务主体配置，因此没有向外部工作区发布。",
            "missing": config.missing,
        }
    try:
        result = publisher_factory(config).publish(
            model_root=Path(bundle["model_root"]),
            report_root=Path(bundle["report_root"]),
            display_name=display_name,
        )
        return {**result, "published": True, "message": "语义模型和报表已自动发布并回读验证。"}
    except PowerBIAutomationError as exc:
        return {
            "status": "publish_failed",
            "published": False,
            "message": str(exc),
            "missing": [],
        }
