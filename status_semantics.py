"""Conservative business-status classification.

Status fields are not safe boolean strings.  A substring rule such as
``contains('完成')`` accepts ``未完成`` and ``退款申请完成但尚未到账``.  This
module gives every domain the same negative-first decision order and keeps
unrecognised states visible for manual verification.
"""

from __future__ import annotations

import re
from typing import Callable

import pandas as pd


ORDER_SUCCESS = "NORMAL_SUCCESS"
ORDER_INVALID = "INVALID"
REFUND_CONFIRMED = "REFUND_CONFIRMED"
REFUND_PENDING = "REFUND_PENDING"
RECEIPT_CONFIRMED = "RECEIPT_CONFIRMED"
RECEIPT_PENDING = "RECEIPT_PENDING"
UNKNOWN = "UNKNOWN"

_SEPARATORS = re.compile(r"[\s_\-—/|,，。；;：:（）()【】\[\]]+")


def _normalise(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return _SEPARATORS.sub("", str(value)).casefold()


def _has(text: str, phrases: tuple[str, ...]) -> bool:
    return any(phrase in text for phrase in phrases)


def classify_order_status(value: object) -> str:
    """Classify sales/order status with negative states taking precedence."""

    text = _normalise(value)
    if not text:
        return UNKNOWN
    if _has(
        text,
        (
            "未完成",
            "未支付",
            "待支付",
            "待付款",
            "支付失败",
            "付款失败",
            "交易失败",
            "已取消",
            "取消",
            "已关闭",
            "关闭",
            "作废",
        ),
    ):
        return ORDER_INVALID
    if _has(text, ("退款未完成", "退款中", "退款申请", "已退款", "全额退款")):
        return ORDER_INVALID
    if _has(text, ("已完成", "交易完成", "订单完成", "已支付", "支付成功", "已结账", "结账完成")):
        return ORDER_SUCCESS
    if text in {"完成", "支付", "结账", "成功"}:
        return ORDER_SUCCESS
    return UNKNOWN


def classify_refund_status(value: object) -> str:
    """Classify after-sale status without treating workflow completion as cash."""

    text = _normalise(value)
    if not text:
        return UNKNOWN
    # Pending/negative cash evidence must win over words such as “完成”.
    if _has(
        text,
        (
            "未退款",
            "退款失败",
            "退款未完成",
            "尚未到账",
            "未到账",
            "申请中",
            "处理中",
            "待审核",
            "待处理",
            "退款申请",
        ),
    ):
        return REFUND_PENDING
    if _has(text, ("拒绝", "驳回", "已取消", "已关闭", "撤销")):
        return UNKNOWN
    if _has(text, ("退款成功", "退款完成", "已退款", "已到账")):
        return REFUND_CONFIRMED
    return UNKNOWN


def classify_receipt_status(value: object) -> str:
    """Classify procurement receipt status with cancellation first."""

    text = _normalise(value)
    if not text:
        return UNKNOWN
    if _has(text, ("未入库", "待入库", "待验收", "入库失败", "取消", "关闭", "作废")):
        return RECEIPT_PENDING
    if _has(text, ("已入库", "入库完成", "验收完成", "已验收")):
        return RECEIPT_CONFIRMED
    return UNKNOWN


def classify_status_series(series: pd.Series, classifier: Callable[[object], str]) -> pd.Series:
    """Vector-friendly wrapper that preserves the source index."""

    return series.map(classifier).astype("string")


__all__ = [
    "ORDER_INVALID",
    "ORDER_SUCCESS",
    "RECEIPT_CONFIRMED",
    "RECEIPT_PENDING",
    "REFUND_CONFIRMED",
    "REFUND_PENDING",
    "UNKNOWN",
    "classify_order_status",
    "classify_receipt_status",
    "classify_refund_status",
    "classify_status_series",
]
