"""Controlled adapters for VBA, databases and OCR/document ingestion."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import ctypes
import hashlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from typing import Any, Iterable, Mapping
import zipfile

import pandas as pd

from .large_data import validate_read_only_sql


VBA_FORBIDDEN = re.compile(
    r"\b(shell|kill|createobject|getobject|wscript|powershell|cmd\.exe|filesystemobject|"
    r"winhttp|xmlhttp|regwrite|open\s+.+\s+for\s+(output|append|binary)|"
    r"declare\s+(ptrsafe\s+)?(sub|function)|vbproject|vbcomponents)\b",
    re.IGNORECASE,
)


def _windows_compatible_path(path: Path) -> Path:
    """Use an ASCII 8.3 path when a legacy Windows CLI rejects Unicode."""
    if os.name != "nt" or not path.exists() or str(path).isascii():
        return path
    try:
        get_short_path = ctypes.windll.kernel32.GetShortPathNameW
        needed = get_short_path(str(path), None, 0)
        if not needed:
            return path
        buffer = ctypes.create_unicode_buffer(needed + 1)
        written = get_short_path(str(path), buffer, len(buffer))
        if written and buffer.value:
            return Path(buffer.value)
    except (AttributeError, OSError, ValueError):
        pass
    return path


def _portable_tesseract_runtime() -> Path | None:
    """Stage the bundled OCR engine in a Windows CLI-safe local directory."""
    if not getattr(sys, "frozen", False):
        return None
    embedded = Path(sys.executable).resolve().parent / "tesseract"
    if not (embedded / "tesseract.exe").is_file():
        return None
    local_base = os.environ.get("LOCALAPPDATA", "").strip()
    if not local_base:
        return embedded
    runtime = Path(local_base) / "BiaogeKuaichu" / "ocr_runtime"
    required = [runtime / "tesseract.exe", runtime / "tessdata" / "chi_sim.traineddata"]
    if not all(item.is_file() for item in required):
        runtime.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(embedded, runtime, dirs_exist_ok=True)
    return _windows_compatible_path(runtime)


def local_tessdata_dir() -> Path | None:
    portable = _portable_tesseract_runtime()
    candidates = [
        os.environ.get("BIAOGE_TESSDATA", ""),
        str(portable / "tessdata") if portable else "",
        str(Path(os.environ.get("LOCALAPPDATA", "")) / "BiaogeKuaichu" / "tessdata")
        if os.environ.get("LOCALAPPDATA") else "",
        str(Path(__file__).resolve().parent / "user_data" / "tessdata"),
    ]
    for candidate in candidates:
        path = Path(candidate) if candidate else None
        if path is not None and path.is_dir() and any(path.glob("*.traineddata")):
            return _windows_compatible_path(path)
    return None


def tesseract_executable() -> Path | None:
    portable = _portable_tesseract_runtime()
    candidates = [
        os.environ.get("TESSERACT_CMD", ""),
        str(portable / "tesseract.exe") if portable else "",
        shutil.which("tesseract.exe") or shutil.which("tesseract") or "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return _windows_compatible_path(Path(candidate))
    return None


def tesseract_languages(executable: Path, tessdata: Path | None) -> set[str]:
    """Return installed languages without relying on pytesseract's Windows parser."""
    command = [str(executable)]
    if tessdata is not None:
        command.extend(["--tessdata-dir", str(tessdata)])
    command.append("--list-langs")
    completed = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
    )
    if completed.returncode != 0:
        return set()
    return {
        line.strip()
        for line in completed.stdout.splitlines()
        if line.strip() and not line.lower().startswith("list of available")
    }


@dataclass(frozen=True)
class AutomationResult:
    status: str
    kind: str
    message: str
    artifacts: tuple[str, ...]
    checks: tuple[str, ...]
    warnings: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def validate_vba_module(code: str) -> str:
    text = str(code).strip()
    if not text or len(text) > 100_000:
        raise ValueError("VBA 模块为空或过长")
    if VBA_FORBIDDEN.search(text):
        raise ValueError("VBA 模块包含文件、系统、网络、外部进程或自修改操作")
    if not re.search(r"\b(Sub|Function)\s+[A-Za-z_][A-Za-z0-9_]*", text, re.IGNORECASE):
        raise ValueError("VBA 模块没有可识别的 Sub/Function")
    if re.search(r"\bOn\s+Error\s+Resume\s+Next\b", text, re.IGNORECASE):
        raise ValueError("VBA 模块不得全局忽略错误")
    return text


def build_vba_bundle(
    code: str,
    destination: str | Path,
    *,
    module_name: str = "BiaogeAutomation",
    entry_macro: str | None = None,
) -> AutomationResult:
    safe_code = validate_vba_module(code)
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,39}", module_name):
        raise ValueError("VBA 模块名无效")
    if entry_macro and not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,79}", entry_macro):
        raise ValueError("入口宏名称无效")
    target = Path(destination)
    if target.suffix.lower() != ".zip":
        raise ValueError("VBA 交付包必须是 .zip")
    target.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(safe_code.encode("utf-8")).hexdigest()
    manifest = {
        "schema_version": 1,
        "module": f"{module_name}.bas",
        "entry_macro": entry_macro,
        "sha256": digest,
        "security": "已通过本地静态危险指令扫描；仍须在客户工作簿副本中测试",
    }
    readme = (
        "VBA安全交付包\n\n"
        "1. 只在客户工作簿副本中导入 .bas；2. 对比执行前后工作表数量、行数与关键合计；"
        "3. Office 宏策略、信任中心和签名由文件所有者控制，程序不会绕过；"
        "4. 模块哈希见 manifest.json，修改后必须重新扫描。\n"
    )
    checklist = "\n".join([
        "[ ] 使用备份副本", "[ ] 宏签名/来源已确认", "[ ] 关键数据合计已记录",
        "[ ] 执行后无新增外部链接", "[ ] 执行结果与人工样例一致", "[ ] 原文件可随时回滚",
    ])
    handle, temp_name = tempfile.mkstemp(prefix=".vba_", suffix=".zip", dir=target.parent)
    os.close(handle)
    temp = Path(temp_name)
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(f"{module_name}.bas", safe_code.encode("utf-8-sig"))
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            archive.writestr("使用说明.txt", readme.encode("utf-8-sig"))
            archive.writestr("测试与回滚清单.txt", checklist.encode("utf-8-sig"))
        os.replace(temp, target)
    finally:
        temp.unlink(missing_ok=True)
    return AutomationResult(
        "ready", "vba_bundle", "VBA模块已扫描并生成可审计交付包", (str(target),),
        ("危险指令扫描通过", "模块SHA-256已写入清单", "测试与回滚清单已生成"),
        ("Office 安全设置与客户授权不能被程序绕过",),
    )


def query_sqlite_read_only(database: str | Path, sql: str, *, max_rows: int = 1_000_000) -> pd.DataFrame:
    path = Path(database).resolve()
    if not path.is_file():
        raise FileNotFoundError("数据库文件不存在")
    query = validate_read_only_sql(sql)
    connection = sqlite3.connect(f"file:{path.as_posix()}?mode=ro", uri=True, timeout=10)
    try:
        connection.execute("PRAGMA query_only=ON")
        frame = pd.read_sql_query(f"SELECT * FROM ({query}) LIMIT {int(max_rows) + 1}", connection)
        if len(frame) > max_rows:
            raise ValueError(f"数据库查询结果超过 {max_rows:,} 行")
        return frame
    finally:
        connection.close()


def query_odbc_read_only(connection_string: str, sql: str, *, max_rows: int = 1_000_000, timeout: int = 30) -> pd.DataFrame:
    query = validate_read_only_sql(sql)
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError("未安装 pyodbc，无法连接 SQL Server/MySQL/PostgreSQL ODBC 数据源") from None
    connection = pyodbc.connect(str(connection_string), autocommit=False, timeout=max(1, min(int(timeout), 300)))
    try:
        cursor = connection.cursor()
        try:
            cursor.execute("SET TRANSACTION READ ONLY")
        except Exception:
            pass
        frame = pd.read_sql_query(f"SELECT * FROM ({query}) AS result", connection)
        if len(frame) > max_rows:
            raise ValueError(f"数据库查询结果超过 {max_rows:,} 行")
        connection.rollback()
        return frame
    finally:
        connection.close()


def extract_pdf_tables(path: str | Path, *, max_pages: int = 200) -> dict[str, pd.DataFrame]:
    source = Path(path).resolve()
    if source.suffix.lower() != ".pdf" or not source.is_file():
        raise ValueError("请选择有效的 PDF 文件")
    try:
        import pdfplumber
    except ImportError:
        raise RuntimeError("未安装 PDF 表格提取组件 pdfplumber") from None
    tables: dict[str, pd.DataFrame] = {}
    with pdfplumber.open(source) as document:
        if len(document.pages) > max_pages:
            raise ValueError(f"PDF 超过 {max_pages} 页安全上限")
        for page_index, page in enumerate(document.pages, start=1):
            for table_index, raw in enumerate(page.extract_tables() or [], start=1):
                if not raw:
                    continue
                header = [str(value or f"未命名列_{idx + 1}") for idx, value in enumerate(raw[0])]
                tables[f"第{page_index}页_表{table_index}"] = pd.DataFrame(raw[1:], columns=header)
    if not tables:
        raise ValueError("PDF 中没有识别到结构化表格；扫描件请使用 OCR 模式")
    return tables


def extract_image_text(path: str | Path, *, language: str = "chi_sim+eng") -> pd.DataFrame:
    source = Path(path).resolve()
    if source.suffix.lower() not in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"} or not source.is_file():
        raise ValueError("请选择有效图片")
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise RuntimeError("未安装 OCR 组件 pytesseract/Pillow") from None
    executable = tesseract_executable()
    if executable is None:
        raise RuntimeError("未找到 Tesseract OCR 引擎")
    pytesseract.pytesseract.tesseract_cmd = str(executable)
    config = ""
    tessdata = local_tessdata_dir()
    if tessdata is not None:
        # pytesseract uses shlex to split this string.  On Windows, wrapping an
        # already space-free path in quotes can leave the quote character in
        # Tesseract's resolved language-file path (".../chi_sim.traineddata).
        # The application-owned fallback path is deliberately ASCII and has no
        # spaces, so pass it without quotes.  For an operator-supplied path with
        # spaces, use TESSDATA_PREFIX instead of relying on shell quoting.
        if " " in str(tessdata):
            os.environ["TESSDATA_PREFIX"] = str(tessdata)
        else:
            config = f"--tessdata-dir {tessdata}"
    requested = [item for item in language.split("+") if item]
    available = tesseract_languages(executable, tessdata)
    missing = [item for item in requested if item not in available]
    if missing:
        raise RuntimeError(f"OCR 缺少语言包：{'、'.join(missing)}")
    data = pytesseract.image_to_data(
        Image.open(source), lang=language, config=config,
        output_type=pytesseract.Output.DATAFRAME,
    )
    useful = data.loc[data["text"].fillna("").astype(str).str.strip().ne(""), [
        "page_num", "block_num", "par_num", "line_num", "left", "top", "width", "height", "conf", "text"
    ]].reset_index(drop=True)
    if useful.empty:
        raise ValueError("图片中没有识别到文字")
    return useful


def document_capabilities() -> dict[str, bool]:
    capabilities = {}
    try:
        import pdfplumber  # noqa: F401
        capabilities["pdf_tables"] = True
    except ImportError:
        capabilities["pdf_tables"] = False
    try:
        import pytesseract  # noqa: F401
        from PIL import Image  # noqa: F401
        executable = tesseract_executable()
        languages = tesseract_languages(executable, local_tessdata_dir()) if executable else set()
        capabilities["image_ocr"] = bool(executable and {"chi_sim", "eng"}.issubset(languages))
    except ImportError:
        capabilities["image_ocr"] = False
    try:
        import pyodbc  # noqa: F401
        capabilities["odbc"] = True
    except ImportError:
        capabilities["odbc"] = False
    capabilities["sqlite"] = True
    capabilities["tesseract_engine"] = tesseract_executable() is not None
    return capabilities


__all__ = [
    "AutomationResult", "build_vba_bundle", "document_capabilities",
    "extract_image_text", "extract_pdf_tables", "query_odbc_read_only",
    "local_tessdata_dir", "query_sqlite_read_only", "tesseract_executable", "validate_vba_module",
]
