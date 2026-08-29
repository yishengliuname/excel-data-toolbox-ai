from __future__ import annotations

from excel_data_toolbox.status_semantics import (
    ORDER_INVALID,
    ORDER_SUCCESS,
    RECEIPT_CONFIRMED,
    RECEIPT_PENDING,
    REFUND_CONFIRMED,
    REFUND_PENDING,
    UNKNOWN,
    classify_order_status,
    classify_receipt_status,
    classify_refund_status,
)


def test_order_status_negative_phrases_win_over_positive_substrings() -> None:
    assert classify_order_status("未完成") == ORDER_INVALID
    assert classify_order_status("支付失败") == ORDER_INVALID
    assert classify_order_status("未支付") == ORDER_INVALID
    assert classify_order_status("已取消（支付成功后撤销）") == ORDER_INVALID
    assert classify_order_status("已完成") == ORDER_SUCCESS
    assert classify_order_status("支付成功") == ORDER_SUCCESS
    assert classify_order_status("奇怪的新状态") == UNKNOWN


def test_refund_status_requires_cash_confirmation() -> None:
    assert classify_refund_status("退款未完成") == REFUND_PENDING
    assert classify_refund_status("退款申请完成但尚未到账") == REFUND_PENDING
    assert classify_refund_status("处理中") == REFUND_PENDING
    assert classify_refund_status("退款成功") == REFUND_CONFIRMED
    assert classify_refund_status("已退款") == REFUND_CONFIRMED
    assert classify_refund_status("退款申请已成功到账") == REFUND_CONFIRMED
    assert classify_refund_status("申请已撤销") == UNKNOWN


def test_receipt_status_does_not_accept_unreceived_goods() -> None:
    assert classify_receipt_status("未入库") == RECEIPT_PENDING
    assert classify_receipt_status("待验收") == RECEIPT_PENDING
    assert classify_receipt_status("入库失败") == RECEIPT_PENDING
    assert classify_receipt_status("已入库") == RECEIPT_CONFIRMED
    assert classify_receipt_status("未知") == UNKNOWN
