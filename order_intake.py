"""Deterministic freelance order triage, quote and acceptance contract."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import math
import re
from typing import Any, Iterable, Mapping


@dataclass(frozen=True)
class OrderQuote:
    capability: str
    complexity_score: int
    risk_level: str
    estimated_hours: tuple[float, float]
    suggested_price: tuple[int, int]
    delivery_days: tuple[int, int]
    detected_services: tuple[str, ...]
    clarification_questions: tuple[str, ...]
    acceptance_items: tuple[str, ...]
    customer_reply: str
    assumptions: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


SERVICE_RULES: tuple[tuple[str, str, int, float], ...] = (
    (r"清洗|去重|空值|格式|拆分|替换", "数据清洗", 1, 0.8),
    (r"合并|拼接|关联|匹配|lookup|vlookup|xlookup", "多表合并匹配", 2, 1.5),
    (r"对账|核销|容差|回款", "复杂对账", 4, 3.0),
    (r"统计|分析|同比|环比|预测|回归|异常", "统计分析", 3, 2.5),
    (r"图|可视化|看板|仪表盘", "可视化看板", 3, 2.0),
    (r"宏|vba", "VBA自动化", 5, 4.0),
    (r"power\s*bi|dax|power\s*query|pbip", "Power BI", 6, 6.0),
    (r"数据库|sql|mysql|postgres|sqlserver|sqlite", "数据库只读集成", 5, 5.0),
    (r"pdf|图片|截图|ocr|扫描", "OCR与文档取数", 5, 4.0),
    (r"定时|每天|每周|自动运行|批处理", "定时批处理", 4, 3.0),
)


HIGH_RISK = re.compile(r"破解|绕过密码|伪造|篡改流水|考试代做|刷单|虚假证明", re.IGNORECASE)


def quote_order(
    request: str,
    *,
    table_count: int = 1,
    total_rows: int = 0,
    deadline_hours: float | None = None,
    has_sample: bool = False,
) -> OrderQuote:
    text = str(request).strip()
    if not text:
        raise ValueError("订单需求不能为空")
    if len(text) > 20_000:
        raise ValueError("订单需求过长")
    if HIGH_RISK.search(text):
        return OrderQuote(
            capability="refuse", complexity_score=10, risk_level="blocked",
            estimated_hours=(0, 0), suggested_price=(0, 0), delivery_days=(0, 0),
            detected_services=(), clarification_questions=(), acceptance_items=(),
            customer_reply="您好，该需求涉及伪造、破解或绕过安全限制，本服务无法承接。",
            assumptions=(),
        )
    services: list[str] = []
    score = 1
    hours = 0.5
    for pattern, label, points, effort in SERVICE_RULES:
        if re.search(pattern, text, re.IGNORECASE):
            services.append(label)
            score += points
            hours += effort
    score += min(3, max(0, int(table_count) - 1))
    if total_rows > 1_000_000:
        score += 3
        hours += 4
    elif total_rows > 300_000:
        score += 2
        hours += 2
    elif total_rows > 50_000:
        score += 1
        hours += 1
    if deadline_hours is not None and deadline_hours < 24:
        score += 2
        hours *= 1.2
    score = min(score, 10)
    risk = "high" if score >= 8 else ("medium" if score >= 5 else "low")
    capability = "needs_review" if any(item in services for item in ("VBA自动化", "Power BI", "数据库只读集成", "OCR与文档取数")) else "supported"
    lower_hours = max(1.0, round(hours * 0.8, 1))
    upper_hours = max(lower_hours, round(hours * 1.5, 1))
    hourly = 80 if score <= 3 else (120 if score <= 6 else 180)
    rush_multiplier = 1.35 if deadline_hours is not None and deadline_hours < 24 else 1.0
    lower_price = int(math.ceil(lower_hours * hourly * rush_multiplier / 10) * 10)
    upper_price = int(math.ceil(upper_hours * hourly * rush_multiplier / 10) * 10)
    days = (1, 2) if upper_hours <= 6 else ((2, 4) if upper_hours <= 16 else (4, 8))
    questions: list[str] = []
    if not has_sample:
        questions.append("请提供脱敏样表、字段说明和期望结果示例。")
    if not re.search(r"输出|交付|结果|图|报表|文件", text):
        questions.append("最终需要交付Excel、CSV、图片、Power BI工程还是自动化脚本？")
    if "对账" in text and not re.search(r"容差|完全一致|日期", text):
        questions.append("对账键、金额容差、日期容差和一对多处理规则是什么？")
    acceptance = (
        "导出文件能够正常打开且无损坏提示",
        "行数、字段、空值和数值合计通过自动验收",
        "需求中约定的统计口径与图表逐项核对",
        "存在人工业务判断的记录单独列入待确认清单",
        "客户确认脱敏样例结果后再处理完整数据",
    )
    service_text = "、".join(services) if services else "基础Excel处理"
    reply = (
        f"您好，需求初步识别为：{service_text}。预计 {days[0]}–{days[1]} 天，"
        f"参考价格 ¥{lower_price}–¥{upper_price}。正式报价前请提供脱敏样表和期望结果；"
        "我会先给处理预览与验收口径，确认后执行，交付文件附自动验收报告。"
    )
    return OrderQuote(
        capability=capability,
        complexity_score=score,
        risk_level=risk,
        estimated_hours=(lower_hours, upper_hours),
        suggested_price=(lower_price, upper_price),
        delivery_days=days,
        detected_services=tuple(services or ["基础Excel处理"]),
        clarification_questions=tuple(questions),
        acceptance_items=acceptance,
        customer_reply=reply,
        assumptions=("价格为程序自动估算，最终以样表、数据量、截止时间和验收口径为准。",),
    )


def quote_frame(quotes: Iterable[OrderQuote]):
    import pandas as pd
    return pd.DataFrame(
        [
            {
                "能力结论": quote.capability,
                "复杂度": quote.complexity_score,
                "风险": quote.risk_level,
                "预计工时": f"{quote.estimated_hours[0]}~{quote.estimated_hours[1]}",
                "建议报价": f"¥{quote.suggested_price[0]}~¥{quote.suggested_price[1]}",
                "交付周期": f"{quote.delivery_days[0]}~{quote.delivery_days[1]}天",
                "服务": "、".join(quote.detected_services),
            }
            for quote in quotes
        ]
    )


__all__ = ["OrderQuote", "quote_frame", "quote_order"]
