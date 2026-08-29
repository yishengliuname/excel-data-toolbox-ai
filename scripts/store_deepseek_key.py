"""Migrate a DeepSeek key into Windows DPAPI without echoing it."""

from __future__ import annotations

import argparse
from getpass import getpass
import os
from pathlib import Path
import sys
import tempfile

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT.parent))

from excel_data_toolbox.secure_secrets import SecureSecretStore, SecretStoreError


def _migrate_project_env(store: SecureSecretStore) -> None:
    env_path = PROJECT / ".env"
    if not env_path.is_file():
        raise SystemExit("项目 .env 不存在，无法自动迁移。")
    lines = env_path.read_text(encoding="utf-8-sig").splitlines()
    secret = ""
    retained: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            key, value = stripped.split("=", 1)
            if key.strip() == "DEEPSEEK_API_KEY":
                secret = value.strip().strip('"').strip("'")
                continue
        retained.append(line)
    if not secret:
        raise SystemExit(".env 中没有可迁移的 DEEPSEEK_API_KEY。")
    inaccessible_backup: Path | None = None
    try:
        store.set("DEEPSEEK_API_KEY", secret)
    except SecretStoreError:
        # A vault copied from another Windows account cannot be decrypted by
        # design.  Replace it only while the same key is still present in the
        # source .env, then remove the unreadable encrypted duplicate.
        if not store.path.is_file():
            raise
        inaccessible_backup = store.path.with_suffix(".inaccessible.dpapi")
        os.replace(store.path, inaccessible_backup)
        store.set("DEEPSEEK_API_KEY", secret)
    secret = ""
    handle, temp_name = tempfile.mkstemp(prefix=".env_", dir=PROJECT)
    os.close(handle)
    temp = Path(temp_name)
    try:
        temp.write_text("\n".join(retained).rstrip() + "\n", encoding="utf-8")
        os.replace(temp, env_path)
    finally:
        temp.unlink(missing_ok=True)
    if inaccessible_backup is not None:
        inaccessible_backup.unlink(missing_ok=True)
    print("密钥已迁移到 Windows DPAPI；.env 明文密钥行已移除。")


def _restore_project_env(store: SecureSecretStore) -> None:
    """Recovery bridge for re-encrypting under a different Windows identity."""
    secret = store.get("DEEPSEEK_API_KEY")
    if not secret:
        raise SystemExit("当前 Windows 用户的保险箱中没有可恢复密钥。")
    env_path = PROJECT / ".env"
    lines = env_path.read_text(encoding="utf-8-sig").splitlines() if env_path.is_file() else []
    retained = [
        line for line in lines
        if not (line.strip() and not line.lstrip().startswith("#") and line.split("=", 1)[0].strip() == "DEEPSEEK_API_KEY")
    ]
    retained.append(f"DEEPSEEK_API_KEY={secret}")
    secret = ""
    handle, temp_name = tempfile.mkstemp(prefix=".env_", dir=PROJECT)
    os.close(handle)
    temp = Path(temp_name)
    try:
        temp.write_text("\n".join(retained).rstrip() + "\n", encoding="utf-8")
        os.replace(temp, env_path)
    finally:
        temp.unlink(missing_ok=True)
    print("密钥已临时恢复到 .env；请立即在目标 Windows 用户下再次执行 --migrate-env。")


def main() -> None:
    parser = argparse.ArgumentParser(description="安全保存 DeepSeek API Key")
    parser.add_argument("--migrate-env", action="store_true", help="从项目 .env 迁移并删除明文密钥行")
    parser.add_argument("--restore-env-from-vault", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    store = SecureSecretStore(PROJECT / "user_data" / "secrets.dpapi")
    if args.migrate_env:
        _migrate_project_env(store)
        return
    if args.restore_env_from_vault:
        _restore_project_env(store)
        return
    secret = getpass("请输入 DeepSeek API Key（输入不会显示）：").strip()
    store.set("DEEPSEEK_API_KEY", secret)
    secret = ""
    print("密钥已使用当前 Windows 用户的 DPAPI 加密保存。")


if __name__ == "__main__":
    main()
