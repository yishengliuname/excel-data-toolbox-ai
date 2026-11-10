"""Start a packaged Windows app and verify the local demo/export path."""

from __future__ import annotations

import argparse
import io
import json
import os
from pathlib import Path
import socket
import subprocess
import tempfile
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from openpyxl import load_workbook


SALES_PROMPT = "按月份和地区分析订单金额，生成销售趋势和地区贡献的 Excel 报表。"


def free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def request_json(url: str, *, payload: dict | None = None, task_id: str = "") -> dict:
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    headers = {"Accept": "application/json"}
    if data is not None:
        headers["Content-Type"] = "application/json"
    if task_id:
        headers["X-Task-ID"] = task_id
    request = Request(url, data=data, headers=headers, method="POST" if data is not None else "GET")
    with urlopen(request, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def wait_for_health(base_url: str, timeout: float = 90) -> dict:
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            return request_json(urljoin(base_url, "/health"))
        except (URLError, TimeoutError, ConnectionError) as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(f"packaged server did not become healthy: {last_error}")


def download_workbook(base_url: str, result: dict, task_id: str) -> bytes:
    url = result.get("download_url", "")
    if not url.startswith("/download/"):
        raise RuntimeError(f"Missing workbook download: {result}")
    request = Request(urljoin(base_url, url), headers={"X-Task-ID": task_id})
    with urlopen(request, timeout=20) as response:
        return response.read()


def check_onboarding(base_url: str) -> bytes:
    demo = request_json(urljoin(base_url, "/api/demo"), payload={})
    task_id = str(demo["task_id"])
    state = request_json(urljoin(base_url, "/api/state"), task_id=task_id)
    tables = state.get("tables", [])
    if len(tables) != 1 or not tables[0]["original"] or tables[0]["rows"] <= 0:
        raise RuntimeError("Sales demo is not available as a selectable input")
    job = request_json(
        urljoin(base_url, "/api/ai-run"),
        payload={"prompt": SALES_PROMPT, "table_ids": [demo["table_id"]]},
        task_id=task_id,
    )["job"]
    deadline = time.monotonic() + 60
    while job["status"] in {"queued", "running", "retrying"} and time.monotonic() < deadline:
        time.sleep(0.2)
        jobs = request_json(urljoin(base_url, "/api/jobs"), task_id=task_id)["jobs"]
        job = next(item for item in jobs if item["job_id"] == job["job_id"])
    if job["status"] != "completed":
        raise RuntimeError(f"Sales demo did not finish: {job}")
    with io.BytesIO(download_workbook(base_url, job["result"], task_id)) as stream:
        workbook = load_workbook(stream, read_only=True)
        try:
            if len(workbook.sheetnames) < 2:
                raise RuntimeError(f"Sales demo workbook is incomplete: {workbook.sheetnames}")
        finally:
            workbook.close()

    demo = request_json(
        urljoin(base_url, "/api/demo"), payload={"scenario": "reconciliation"}, task_id=task_id
    )
    if demo["task_id"] == task_id:
        raise RuntimeError("Reconciliation demo did not create a separate task")
    expected = {
        "matched_count": 2,
        "review_group_count": 1,
        "left_only_count": 1,
        "right_only_count": 1,
        "duplicate_rows_count": 3,
    }
    actual = {key: demo["summary"].get(key) for key in expected}
    if actual != expected or demo["acceptance"]["status"] != "passed":
        raise RuntimeError(f"Unexpected reconciliation results: {actual}")
    contents = download_workbook(base_url, demo, demo["task_id"])
    workbook = load_workbook(io.BytesIO(contents), read_only=True)
    try:
        required = {"订单应结表", "平台实结表", "演示对账_对账摘要", "演示对账_待人工确认", "演示对账_重复键隔离"}
        if not required.issubset(workbook.sheetnames):
            raise RuntimeError(f"Reconciliation workbook is incomplete: {workbook.sheetnames}")
        rows = list(workbook["演示对账_待人工确认"].values)
        header_index = next(index for index, row in enumerate(rows) if "amount_difference_decimal" in row)
        amount_index = rows[header_index].index("amount_difference_decimal")
        if rows[header_index + 1][amount_index] != -10:
            raise RuntimeError("Synthetic order DEMO-002 must show a 10-unit shortfall")
    finally:
        workbook.close()
    return contents


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-dir", type=Path, required=True)
    parser.add_argument("--expected-version", required=True)
    parser.add_argument("--sample-output", type=Path, help="Write the verified synthetic workbook outside the app ZIP")
    args = parser.parse_args()

    app_dir = args.app_dir.resolve()
    executable = app_dir / "BiaogeKuaichuAI.exe"
    if not executable.is_file():
        raise SystemExit(f"missing packaged executable: {executable}")

    port = free_port()
    base_url = f"http://127.0.0.1:{port}"
    with tempfile.TemporaryDirectory(prefix="excel-toolbox-release-smoke-") as temp_name:
        temp_root = Path(temp_name)
        application_log = temp_root / "packaged-application.log"
        environment = os.environ.copy()
        environment["BIAOGE_USER_DATA"] = str(temp_root / "user-data")
        environment.pop("DEEPSEEK_API_KEY", None)
        with application_log.open("wb") as log_handle:
            process = subprocess.Popen(
                [str(executable), "--no-browser", "--port", str(port)],
                cwd=app_dir,
                env=environment,
                stdout=log_handle,
                stderr=subprocess.STDOUT,
            )
            try:
                health = wait_for_health(base_url)
                if health.get("status") != "ok" or health.get("version") != args.expected_version:
                    raise RuntimeError(f"unexpected health response: {health}")

                with urlopen(urljoin(base_url, "/"), timeout=20) as response:
                    page = response.read().decode("utf-8")
                if any(label not in page for label in ("表格快处", "加载演示数据", "试用电商订单对账")):
                    raise RuntimeError("packaged web UI did not load the expected application shell")
                with urlopen(urljoin(base_url, "/static/app.js"), timeout=20) as response:
                    if response.headers.get_content_type() != "text/javascript" or b"reconcileDemoBtn" not in response.read():
                        raise RuntimeError("packaged web UI did not load its demo controls")
                for path in ("/static/__init__.py", "/static/%2e%2e/domain_packs.json"):
                    try:
                        with urlopen(urljoin(base_url, path), timeout=20):
                            raise RuntimeError(f"Static asset allowlist exposed {path}")
                    except HTTPError as exc:
                        if exc.code != 404:
                            raise

                sample = check_onboarding(base_url)
                if args.sample_output:
                    args.sample_output.write_bytes(sample)
            except Exception:
                log_handle.flush()
                print("--- packaged application log ---")
                print(application_log.read_text(encoding="utf-8", errors="replace"))
                raise
            finally:
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=10)

    print("Windows package smoke test passed: health, UI, sales command, reconciliation, Excel export and reopen")


if __name__ == "__main__":
    main()
