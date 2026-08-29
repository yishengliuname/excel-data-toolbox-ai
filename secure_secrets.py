"""Windows DPAPI-backed local secret storage.

Secrets are encrypted for the current Windows user and never returned by status
APIs.  Environment variables remain supported for unattended deployments.
"""

from __future__ import annotations

import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import tempfile
from typing import Any


class SecretStoreError(RuntimeError):
    pass


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob(data: bytes) -> tuple[_DataBlob, Any]:
    buffer = ctypes.create_string_buffer(data)
    return _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def _protect(data: bytes) -> bytes:
    if os.name != "nt":
        raise SecretStoreError("安全密钥库目前仅支持 Windows DPAPI")
    input_blob, input_buffer = _blob(data)
    output_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptProtectData(
        ctypes.byref(input_blob), "表格快处本地密钥", None, None, None, 0,
        ctypes.byref(output_blob),
    ):
        raise SecretStoreError(f"Windows DPAPI 加密失败：{ctypes.GetLastError()}")
    del input_buffer
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        kernel32.LocalFree(output_blob.pbData)


def _unprotect(data: bytes) -> bytes:
    if os.name != "nt":
        raise SecretStoreError("安全密钥库目前仅支持 Windows DPAPI")
    input_blob, input_buffer = _blob(data)
    output_blob = _DataBlob()
    description = wintypes.LPWSTR()
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    if not crypt32.CryptUnprotectData(
        ctypes.byref(input_blob), ctypes.byref(description), None, None, None, 0,
        ctypes.byref(output_blob),
    ):
        raise SecretStoreError(f"Windows DPAPI 解密失败：{ctypes.GetLastError()}")
    del input_buffer
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        if description:
            kernel32.LocalFree(description)
        kernel32.LocalFree(output_blob.pbData)


class SecureSecretStore:
    """Small encrypted key/value store scoped to this Windows user."""

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    @property
    def available(self) -> bool:
        return os.name == "nt"

    def _read_all(self) -> dict[str, str]:
        if not self.path.exists():
            return {}
        try:
            decoded = json.loads(_unprotect(self.path.read_bytes()).decode("utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, SecretStoreError) as exc:
            raise SecretStoreError("本地安全密钥库损坏或不属于当前 Windows 用户") from exc
        if not isinstance(decoded, dict) or any(not isinstance(k, str) or not isinstance(v, str) for k, v in decoded.items()):
            raise SecretStoreError("本地安全密钥库结构无效")
        return decoded

    def get(self, name: str) -> str:
        if not isinstance(name, str) or not name.strip():
            raise ValueError("密钥名称不能为空")
        return self._read_all().get(name.strip(), "")

    def set(self, name: str, value: str) -> None:
        key = str(name).strip()
        secret = str(value).strip()
        if not key or len(key) > 100:
            raise ValueError("密钥名称无效")
        if not secret or len(secret) > 4096 or any(character in secret for character in "\r\n\x00"):
            raise ValueError("密钥内容无效")
        values = self._read_all()
        values[key] = secret
        self.path.parent.mkdir(parents=True, exist_ok=True)
        encrypted = _protect(json.dumps(values, ensure_ascii=False).encode("utf-8"))
        handle, temp_name = tempfile.mkstemp(prefix=".secrets_", dir=self.path.parent)
        os.close(handle)
        temp = Path(temp_name)
        try:
            temp.write_bytes(encrypted)
            os.replace(temp, self.path)
        finally:
            temp.unlink(missing_ok=True)

    def delete(self, name: str) -> bool:
        values = self._read_all()
        removed = values.pop(str(name).strip(), None) is not None
        if not removed:
            return False
        if not values:
            self.path.unlink(missing_ok=True)
            return True
        self.path.write_bytes(_protect(json.dumps(values, ensure_ascii=False).encode("utf-8")))
        return True

    def status(self, names: list[str]) -> dict[str, bool]:
        values = self._read_all()
        return {name: bool(values.get(name)) for name in names}


__all__ = ["SecureSecretStore", "SecretStoreError"]
