"""NetUpgrade storage layer: the application database and the secret vault.

    from netupgrade_app import storage
    storage.repositories.load_devices("Site A")
    ref = storage.protect("p@ssw0rd", purpose="device-password")   # -> 'vault:<id>'
    storage.reveal(ref)                                             # -> 'p@ssw0rd'

Files (all in the per-user data folder, %LOCALAPPDATA%\\NetUpgrade unless NETUPGRADE_DATA_DIR is set):
    netupgrade.db   application state (SQLite, WAL)       - no secrets, only vault references
    vault.db        AES-256-GCM encrypted secrets          - key in Windows Credential Manager
Backups, reports, run logs and IOS images stay as files in the storage folder chosen at install.
"""
from __future__ import annotations

import os
import threading
from pathlib import Path

from . import repositories  # noqa: F401  (re-exported for callers)
from .database import Database
from .vault import (REF_PREFIX, SecretVault, VaultError, default_key_provider, is_ref, open_from_worker,  # noqa: F401
                    seal_for_worker)

DB_FILENAME = "netupgrade.db"
VAULT_FILENAME = "vault.db"

_lock = threading.Lock()
_state = {"data_dir": None, "db": None, "vault": None, "key_provider": None}


def data_dir() -> Path:
    """The per-user folder that holds the database and the vault."""
    if _state["data_dir"] is None:
        configured = os.environ.get("NETUPGRADE_DATA_DIR")
        if configured:
            base = Path(os.path.expandvars(os.path.expanduser(configured)))
        else:
            base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".local" / "share") / "NetUpgrade"
        base.mkdir(parents=True, exist_ok=True)
        _state["data_dir"] = base
    return _state["data_dir"]


def configure(data_dir=None, key_provider=None):
    """Point the storage layer somewhere else (tests, portable runs). Closes anything already open."""
    with _lock:
        if _state["db"] is not None:
            _state["db"].close()
        _state.update(db=None, vault=None, key_provider=key_provider,
                      data_dir=Path(data_dir) if data_dir is not None else None)
        if data_dir is not None:
            Path(data_dir).mkdir(parents=True, exist_ok=True)


def get_db() -> Database:
    if _state["db"] is None:
        with _lock:
            if _state["db"] is None:
                _state["db"] = Database(data_dir() / DB_FILENAME)
    return _state["db"]


def get_vault() -> SecretVault:
    if _state["vault"] is None:
        with _lock:
            if _state["vault"] is None:
                provider = _state["key_provider"] or default_key_provider(data_dir())
                _state["vault"] = SecretVault(data_dir() / VAULT_FILENAME, provider)
    return _state["vault"]


def protect(text: str, purpose: str = "secret") -> str:
    """Store a secret in the vault and return its reference ('' stays '')."""
    return get_vault().put(text, purpose) if text else ""


def reveal(ref: str) -> str:
    """Decrypt a vault reference ('' stays '')."""
    return get_vault().get(ref) if ref else ""
