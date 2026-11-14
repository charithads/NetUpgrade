"""Reading secrets saved by NetUpgrade versions before the vault (only used to migrate them).

Old values look like 'dpapi:<base64>' (Windows DPAPI, no entropy) or 'fernet:<token>' (Fernet with the key
file %LOCALAPPDATA%\\NetUpgrade\\profile.key). New secrets are never written in these formats.
"""
from __future__ import annotations

import base64
import os
from pathlib import Path

from . import winsec


def legacy_key_files():
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home())
    return [root / "NetUpgrade" / "profile.key", root / "IOSXEUpgradeManager" / "profile.key"]


def is_legacy_secret(value) -> bool:
    return isinstance(value, str) and value.split(":", 1)[0] in ("dpapi", "fernet") and ":" in value


def unprotect_legacy(value: str) -> str:
    if not value:
        return ""
    kind, _, body = value.partition(":")
    if kind == "dpapi":
        return winsec.dpapi_unprotect(base64.b64decode(body)).decode("utf-8")
    if kind == "fernet":
        from cryptography.fernet import Fernet, InvalidToken
        for key_file in legacy_key_files():
            try:
                key = key_file.read_bytes().strip()
            except OSError:
                continue
            try:
                return Fernet(key).decrypt(body.encode("ascii")).decode("utf-8")
            except (InvalidToken, ValueError):
                continue
        raise ValueError("The old Fernet key is not available.")
    raise ValueError("Unknown secret format.")


def retire_legacy_key_files(log=print):
    """After every legacy secret has been moved into the vault, the old Fernet key has no purpose left:
    overwrite and delete it so the redacted backups cannot be combined with it."""
    for key_file in legacy_key_files():
        if key_file.is_file():
            try:
                winsec.secure_delete(key_file)
                log(f"Storage migration: removed the old key file {key_file}.")
            except OSError as exc:
                log(f"Storage migration: could not remove the old key file {key_file} ({exc}).")
