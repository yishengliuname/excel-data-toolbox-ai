"""Preflight diagnostics for a customer-facing workstation."""

from __future__ import annotations

import importlib
import json
import os
from pathlib import Path
import platform
import shutil
import socket
import sys


PROJECT = Path(__file__).resolve().parents[1]
PACKAGE_PARENT = PROJECT.parent
if str(PACKAGE_PARENT) not in sys.path:
    sys.path.insert(0, str(PACKAGE_PARENT))


def module_status(name: str) -> dict[str, object]:
    try:
        module = importlib.import_module(name)
        return {"available": True, "version": getattr(module, "__version__", "unknown")}
    except Exception as exc:
        return {"available": False, "error": f"{type(exc).__name__}: {exc}"}


def main() -> int:
    modules = {
        name: module_status(name)
        for name in ("pandas", "numpy", "openpyxl", "duckdb", "pdfplumber", "PIL", "pytesseract", "pyodbc")
    }
    required_ok = all(modules[name]["available"] for name in ("pandas", "numpy", "openpyxl"))
    try:
        from excel_data_toolbox.advanced_automation import (
            local_tessdata_dir,
            tesseract_executable,
            tesseract_languages,
        )

        ocr_executable = tesseract_executable()
        ocr_tessdata = local_tessdata_dir()
        ocr_languages = sorted(tesseract_languages(ocr_executable, ocr_tessdata)) if ocr_executable else []
    except Exception:
        ocr_executable = None
        ocr_tessdata = None
        ocr_languages = []
    ocr_ready = bool(
        modules["PIL"]["available"]
        and modules["pytesseract"]["available"]
        and ocr_executable
        and {"chi_sim", "eng"}.issubset(set(ocr_languages))
    )
    report = {
        "status": "ready" if required_ok else "blocked",
        "python": sys.version,
        "platform": platform.platform(),
        "project": str(PROJECT),
        "writable": os.access(PROJECT, os.W_OK),
        "free_space_bytes": shutil.disk_usage(PROJECT).free,
        "hostname": socket.gethostname(),
        "modules": modules,
        "features": {
            "core_excel": required_ok,
            "large_data": bool(modules["duckdb"]["available"]),
            "pdf_tables": bool(modules["pdfplumber"]["available"]),
            "image_ocr": ocr_ready,
            "tesseract_executable": str(ocr_executable) if ocr_executable else None,
            "tessdata_directory": str(ocr_tessdata) if ocr_tessdata else None,
            "ocr_languages": ocr_languages,
            "odbc": bool(modules["pyodbc"]["available"]),
            "power_bi_local_bundle": required_ok,
        },
        "notes": [
            "图片 OCR 只有在引擎及 chi_sim/eng 语言包均可用时才标记为可用。",
            "Power BI 云发布需要客户自己的 Microsoft Entra/Fabric 租户授权。",
            "Office 宏执行受客户信任中心策略控制，程序不会绕过。",
        ],
    }
    output = PROJECT / "outputs" / "health_check.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if required_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
