from __future__ import annotations

import json
from urllib import error as urllib_error

import pandas as pd
import pytest

import excel_data_toolbox.nl_agent as nl
import excel_data_toolbox.server as server_module
from excel_data_toolbox.nl_agent import DeepSeekAPIError, DeepSeekClient
from excel_data_toolbox.server import ToolboxHandler, _chart_payload


def _visual_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "月份": ["2026-01", "2026-01", "2026-02", "2026-02", "2026-03", "2026-03"],
            "区域": ["华东", "华南", "华东", "华南", "华东", "华南"],
            "销售额": [100.0, 80.0, 130.0, 95.0, 160.0, 110.0],
            "广告费": [20.0, 18.0, 25.0, 20.0, 30.0, 24.0],
            "任务": ["调研", "设计", "开发", "测试", "上线", "复盘"],
            "开始": pd.date_range("2026-01-01", periods=6, freq="7D"),
            "结束": pd.date_range("2026-01-05", periods=6, freq="7D"),
            "进度": [100, 100, 80, 60, 20, 0],
        }
    )


def _base_payload(chart_type: str) -> dict[str, object]:
    return {
        "chart_type": chart_type,
        "dimension": "月份",
        "measure": "销售额",
        "aggregation": "sum",
        "top_n": 20,
        "date_grain": "auto",
        "style_3d": False,
    }


@pytest.mark.parametrize("chart_type", ["grouped_bar", "stacked_bar", "radar", "heatmap"])
def test_multi_series_visuals_return_aligned_series(chart_type: str) -> None:
    payload = _base_payload(chart_type)
    payload["series"] = "区域"
    payload["date_grain"] = "month"

    result = _chart_payload(_visual_frame(), payload)

    assert result["chart_type"] == chart_type
    assert result["labels"] == ["2026-01", "2026-02", "2026-03"]
    assert {item["name"] for item in result["series"]} == {"华东", "华南"}
    assert all(len(item["values"]) == 3 for item in result["series"])


@pytest.mark.parametrize(
    "chart_type",
    ["bar", "horizontal_bar", "line", "area", "pie", "funnel", "waterfall", "treemap"],
)
def test_single_series_visual_catalog_returns_aligned_values(chart_type: str) -> None:
    result = _chart_payload(_visual_frame(), _base_payload(chart_type))

    assert result["chart_type"] == chart_type
    assert len(result["labels"]) == len(result["values"]) == 3


def test_scatter_visual_includes_linear_regression_and_r_squared() -> None:
    payload = _base_payload("scatter")
    payload["dimension"] = "广告费"

    result = _chart_payload(_visual_frame(), payload)

    assert result["trendline"] is not None
    assert 0 <= result["trendline"]["r_squared"] <= 1
    assert len(result["trendline"]["points"]) == 2
    assert "不代表因果关系" in result["summary"]


def test_box_and_gantt_visuals_cover_distribution_and_project_schedule() -> None:
    box_payload = _base_payload("box")
    box_payload["dimension"] = "区域"
    box = _chart_payload(_visual_frame(), box_payload)

    gantt_payload = _base_payload("gantt")
    gantt_payload.update(
        {
            "dimension": "任务",
            "start": "开始",
            "end": "结束",
            "progress": "进度",
        }
    )
    gantt = _chart_payload(_visual_frame(), gantt_payload)

    assert len(box["boxes"]) == 2
    assert all({"q1", "median", "q3", "outliers"} <= set(item) for item in box["boxes"])
    assert len(gantt["items"]) == 6
    assert gantt["items"][0]["duration_days"] == 5
    assert all(0 <= item["progress"] <= 100 for item in gantt["items"])


def test_funnel_and_waterfall_preserve_business_sequence() -> None:
    frame = pd.DataFrame(
        {
            "项目": ["收入", "退款", "成本", "其他收益"],
            "金额": [850000, -32000, -410000, 12000],
        }
    )

    waterfall = _chart_payload(
        frame,
        {
            "chart_type": "waterfall",
            "dimension": "项目",
            "measure": "金额",
            "aggregation": "sum",
            "top_n": 20,
            "date_grain": "auto",
            "style_3d": False,
        },
    )

    assert waterfall["labels"] == ["收入", "退款", "成本", "其他收益"]
    assert waterfall["values"] == [850000.0, -32000.0, -410000.0, 12000.0]
    assert "累计结果" in waterfall["summary"]


class _FakeResponse:
    def __enter__(self) -> "_FakeResponse":
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self, _limit: int) -> bytes:
        return json.dumps(
            {"choices": [{"finish_reason": "stop", "message": {"content": "OK"}}]}
        ).encode("utf-8")


def test_connection_check_sends_only_fixed_probe_and_never_returns_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    captured: dict[str, object] = {}

    def fake_urlopen(request: object, timeout: int) -> _FakeResponse:
        captured["body"] = request.data.decode("utf-8")  # type: ignore[attr-defined]
        captured["timeout"] = timeout
        return _FakeResponse()

    monkeypatch.setattr(nl.urllib_request, "urlopen", fake_urlopen)
    key = "sk-unit-test-not-a-real-secret"
    result = DeepSeekClient(key, timeout_seconds=20).check_connection()

    assert result["authenticated"] is True
    assert key not in json.dumps(result, ensure_ascii=False)
    assert key not in str(captured["body"])
    assert "表" not in str(captured["body"])


def test_connection_check_explains_firewall_permission_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def blocked(*_args: object, **_kwargs: object):
        raise urllib_error.URLError(PermissionError(13, "blocked"))

    monkeypatch.setattr(nl.urllib_request, "urlopen", blocked)

    with pytest.raises(DeepSeekAPIError, match="防火墙|套接字权限"):
        DeepSeekClient("sk-unit-test").check_connection()


def test_ai_diagnosis_handler_redacts_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    handler = object.__new__(ToolboxHandler)
    key = "sk-unit-test-handler-secret"

    def fake_check(_self: object) -> dict[str, object]:
        return {
            "status": "connected",
            "authenticated": True,
            "model": "deepseek-v4-flash",
            "message": "连接成功",
        }

    monkeypatch.setattr(server_module.DeepSeekClient, "check_connection", fake_check)
    result = handler._ai_diagnose({"api_key": key, "model": "deepseek-v4-flash"})

    assert result["authenticated"] is True
    assert key not in json.dumps(result, ensure_ascii=False)
