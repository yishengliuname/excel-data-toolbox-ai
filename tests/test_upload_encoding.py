from __future__ import annotations

from email import policy
from email.parser import BytesParser

from excel_data_toolbox.server import _decode_multipart_text


def test_browser_formdata_chinese_text_decodes_as_utf8_without_charset() -> None:
    boundary = "----biaoge-boundary"
    task_name = "Excel工具客户测试数据"
    body = (
        f"--{boundary}\r\n"
        'Content-Disposition: form-data; name="task_name"\r\n\r\n'
    ).encode("ascii") + task_name.encode("utf-8") + f"\r\n--{boundary}--\r\n".encode("ascii")
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: multipart/form-data; boundary={boundary}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii")
        + body
    )
    part = next(message.iter_parts())

    assert _decode_multipart_text(part) == task_name
