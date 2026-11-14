"""The secret vault: every password, enable secret and API key NetUpgrade stores.

How it is protected
-------------------
* Each secret is encrypted with **AES-256-GCM** (authenticated encryption) under a random 256-bit
  master key, with a fresh 96-bit nonce per secret. The secret's id and purpose are bound in as
  associated data, so a ciphertext cannot be swapped onto another record without detection.
* The ciphertexts live in their own database file, ``vault.db``, separate from the application
  database. The application database only holds references (``vault:<id>``).
* The **master key never touches the disk in clear**. It is kept in **Windows Credential Manager**
  (target ``NetUpgrade/Vault/MasterKey/v1``, this user on this PC only), additionally wrapped with
  DPAPI and an application-specific entropy value. If Credential Manager cannot be used (rare: e.g. a
  locked-down service account) the DPAPI-wrapped key is kept in ``vault.key`` instead.
* Both files are restricted to the current Windows user and SYSTEM.
* Secrets are only decrypted in memory, at the moment they are needed (a login to a switch, sending an
  e-mail, calling the AI provider), and are never written to logs or returned to the browser.

Only the same Windows user on the same PC can open the vault. Moving the files to another PC or user
does not reveal anything; the passwords then have to be entered again.
"""
from __future__ import annotations

import hashlib
import os
import sqlite3
import threading
import time
import uuid
from pathlib import Path

from . import winsec

REF_PREFIX = "vault:"
KEY_TARGET = "NetUpgrade/Vault/MasterKey/v1"
KEY_ENTROPY = b"NetUpgrade vault master key v1"
JOB_ENTROPY = b"NetUpgrade worker job credentials v1"
_KEY_BYTES = 32


class VaultError(Exception):
    """A secret could not be stored or read (the message is safe to show; it never contains a secret)."""


# ----------------------------------------------------------------------------------------------- #
# Where the master key is kept
# ----------------------------------------------------------------------------------------------- #
class MemoryKeyProvider:
    """Tests only: the key lives in this object."""
    name = "memory"

    def __init__(self, key=None):
        self._key = key

    def load(self):
        return self._key

    def store(self, key):
        self._key = key


class CredentialManagerKeyProvider:
    """Windows Credential Manager, the key wrapped with DPAPI + application entropy."""
    name = "Windows Credential Manager"

    def __init__(self, target=KEY_TARGET):
        self.target = target

    def load(self):
        blob = winsec.cred_read(self.target)
        return None if blob is None else winsec.dpapi_unprotect(blob, KEY_ENTROPY)

    def store(self, key):
        winsec.cred_write(self.target, winsec.dpapi_protect(key, KEY_ENTROPY), username="NetUpgrade",
                          comment="NetUpgrade vault master key - do not delete (stored passwords become unreadable)")


class DpapiFileKeyProvider:
    """Fallback: the DPAPI-wrapped key in a user-only file."""
    name = "DPAPI-protected key file"

    def __init__(self, path):
        self.path = Path(path)

    def load(self):
        try:
            blob = self.path.read_bytes()
        except FileNotFoundError:
            return None
        return winsec.dpapi_unprotect(blob, KEY_ENTROPY)

    def store(self, key):
        _write_private(self.path, winsec.dpapi_protect(key, KEY_ENTROPY))


class PlainFileKeyProvider:
    """Development on non-Windows systems only (the key file is readable by this OS user)."""
    name = "key file (non-Windows development only)"

    def __init__(self, path):
        self.path = Path(path)

    def load(self):
        try:
            return self.path.read_bytes()
        except FileNotFoundError:
            return None

    def store(self, key):
        _write_private(self.path, key)


class ChainKeyProvider:
    """Try each provider in turn. A key is loaded from the first that has one; a new key is stored in the
    first that accepts it (Credential Manager first, the key file only if that fails)."""

    def __init__(self, providers):
        self.providers = list(providers)
        self.name = self.providers[0].name
        self.used = None

    def load(self):
        for p in self.providers:
            try:
                key = p.load()
            except OSError:
                continue
            if key:
                self.used = p
                self.name = p.name
                return key
        return None

    def store(self, key):
        errors = []
        for p in self.providers:
            try:
                p.store(key)
                if p.load() == key:                       # prove it round-trips before relying on it
                    self.used = p
                    self.name = p.name
                    return
            except OSError as exc:
                errors.append(f"{p.name}: {exc}")
        raise VaultError("Could not store the vault key (" + "; ".join(errors) + ").")


def default_key_provider(data_dir):
    choice = os.environ.get("NETUPGRADE_VAULT_KEY", "").lower()
    if choice == "memory":
        return MemoryKeyProvider()
    if winsec.IS_WINDOWS and choice != "file":
        return ChainKeyProvider([CredentialManagerKeyProvider(), DpapiFileKeyProvider(Path(data_dir) / "vault.key")])
    if winsec.IS_WINDOWS:
        return DpapiFileKeyProvider(Path(data_dir) / "vault.key")
    return PlainFileKeyProvider(Path(data_dir) / "vault.key")


def _write_private(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    fd = os.open(str(tmp), os.O_WRONLY | os.O_CREAT | os.O_TRUNC | getattr(os, "O_BINARY", 0), 0o600)
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, path)
    winsec.restrict_to_current_user(path)


# ----------------------------------------------------------------------------------------------- #
# The vault
# ----------------------------------------------------------------------------------------------- #
_SCHEMA = """
CREATE TABLE IF NOT EXISTS secrets (
    id          TEXT PRIMARY KEY,
    purpose     TEXT NOT NULL,
    key_id      TEXT NOT NULL,
    nonce       BLOB NOT NULL,
    ciphertext  BLOB NOT NULL,
    created_at  REAL NOT NULL
)
"""


def is_ref(value) -> bool:
    return isinstance(value, str) and value.startswith(REF_PREFIX) and len(value) > len(REF_PREFIX)


class SecretVault:
    def __init__(self, path, key_provider):
        self.path = Path(path)
        self.provider = key_provider
        self._lock = threading.RLock()
        self._key = None
        self._ready = False

    # ---- setup ------------------------------------------------------------------------- #
    def _db(self):
        new = not self.path.exists()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(self.path), timeout=30, isolation_level=None)
        conn.execute("PRAGMA secure_delete = ON")
        conn.execute("PRAGMA busy_timeout = 30000")
        if not self._ready:
            conn.execute(_SCHEMA)
            if new:
                winsec.restrict_to_current_user(self.path)
            self._ready = True
        return conn

    def _master_key(self, create=True):
        if self._key is not None:
            return self._key
        try:
            key = self.provider.load()
        except OSError as exc:
            raise VaultError(f"The vault key could not be read ({exc}).") from None
        if not key:
            if not create:
                return None
            key = os.urandom(_KEY_BYTES)
            try:
                self.provider.store(key)
            except OSError as exc:
                raise VaultError(f"The vault key could not be stored ({exc}).") from None
        if len(key) != _KEY_BYTES:
            raise VaultError("The vault key is damaged.")
        self._key = key
        return key

    @staticmethod
    def _key_id(key):
        return hashlib.sha256(b"netupgrade-key-id|" + key).hexdigest()[:16]

    @staticmethod
    def _aad(secret_id, purpose):
        return f"netupgrade-vault|v1|{secret_id}|{purpose}".encode("utf-8")

    def status(self):
        """For diagnostics: where the key is kept and how many secrets are stored (never the secrets)."""
        with self._lock:
            conn = self._db()
            try:
                count = conn.execute("SELECT COUNT(*) FROM secrets").fetchone()[0]
            finally:
                conn.close()
            key = self._master_key(create=False)
            return {"secrets": count, "key_store": getattr(self.provider, "name", "?"), "key_present": bool(key),
                    "path": str(self.path)}

    # ---- secrets ----------------------------------------------------------------------- #
    def put(self, plaintext: str, purpose: str = "secret") -> str:
        """Encrypt and store a secret; returns its reference ('vault:<id>'). '' stays ''."""
        if not plaintext:
            return ""
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        with self._lock:
            key = self._master_key()
            secret_id = uuid.uuid4().hex
            nonce = os.urandom(12)
            ct = AESGCM(key).encrypt(nonce, plaintext.encode("utf-8"), self._aad(secret_id, purpose))
            conn = self._db()
            try:
                conn.execute("INSERT INTO secrets(id, purpose, key_id, nonce, ciphertext, created_at) VALUES (?,?,?,?,?,?)",
                             (secret_id, purpose, self._key_id(key), nonce, ct, time.time()))
            finally:
                conn.close()
            return REF_PREFIX + secret_id

    def get(self, ref: str) -> str:
        """Decrypt a secret by its reference. Raises VaultError if it is missing or cannot be decrypted."""
        if not ref:
            return ""
        if not is_ref(ref):
            raise VaultError("Not a vault reference.")
        from cryptography.exceptions import InvalidTag
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        secret_id = ref[len(REF_PREFIX):]
        with self._lock:
            conn = self._db()
            try:
                row = conn.execute("SELECT purpose, key_id, nonce, ciphertext FROM secrets WHERE id = ?",
                                   (secret_id,)).fetchone()
            finally:
                conn.close()
            if row is None:
                raise VaultError("The stored password is missing from the vault. Enter it again.")
            purpose, key_id, nonce, ct = row
            key = self._master_key(create=False)
            if key is None or self._key_id(key) != key_id:
                raise VaultError("The stored password was saved under a vault key that is not available for this "
                                 "Windows user on this PC. Enter it again.")
            try:
                return AESGCM(key).decrypt(nonce, ct, self._aad(secret_id, purpose)).decode("utf-8")
            except InvalidTag:
                raise VaultError("The stored password failed its integrity check. Enter it again.") from None

    def delete(self, refs) -> int:
        ids = [r[len(REF_PREFIX):] for r in ([refs] if isinstance(refs, str) else refs) if is_ref(r)]
        if not ids:
            return 0
        with self._lock:
            conn = self._db()
            try:
                n = 0
                for i in range(0, len(ids), 500):
                    part = ids[i:i + 500]
                    n += conn.execute(f"DELETE FROM secrets WHERE id IN ({','.join('?' * len(part))})", part).rowcount
                return n
            finally:
                conn.close()

    def prune(self, keep_refs, grace_seconds=3600) -> int:
        """Delete secrets nothing refers to any more (older than the grace period, so a secret that was just
        created and is about to be saved is never removed)."""
        keep = {r[len(REF_PREFIX):] for r in keep_refs if is_ref(r)}
        with self._lock:
            conn = self._db()
            try:
                rows = conn.execute("SELECT id FROM secrets WHERE created_at < ?",
                                    (time.time() - grace_seconds,)).fetchall()
                drop = [r[0] for r in rows if r[0] not in keep]
            finally:
                conn.close()
        return self.delete([REF_PREFIX + i for i in drop])


# ----------------------------------------------------------------------------------------------- #
# Credentials handed to a background upgrade worker (same Windows user, separate process)
# ----------------------------------------------------------------------------------------------- #
def seal_for_worker(text: str) -> str:
    """Wrap a password for the short-lived job file a worker reads (DPAPI, this user, app entropy)."""
    if not text:
        return ""
    if not winsec.IS_WINDOWS:
        return "plain:" + text
    import base64
    return "dpapi-job:" + base64.b64encode(winsec.dpapi_protect(text.encode("utf-8"), JOB_ENTROPY)).decode("ascii")


def open_from_worker(value) -> str:
    if not value:
        return ""
    value = str(value)
    if value.startswith("dpapi-job:"):
        import base64
        return winsec.dpapi_unprotect(base64.b64decode(value[len("dpapi-job:"):]), JOB_ENTROPY).decode("utf-8")
    if value.startswith("plain:"):
        return value[len("plain:"):]
    return value                                         # job file written by an older version
