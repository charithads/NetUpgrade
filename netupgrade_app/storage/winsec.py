"""Thin ctypes wrappers around the Windows security APIs used by the storage layer.

* DPAPI (CryptProtectData / CryptUnprotectData): encrypts bytes so only the same Windows user on the
  same PC can decrypt them. An application-specific "entropy" value is mixed in, so another program
  running as the same user cannot decrypt the blob by simply calling DPAPI without it.
* Credential Manager (CredWriteW / CredReadW / CredDeleteW): the Windows store for secrets, protected by
  DPAPI with the user's logon credentials. NetUpgrade keeps its vault master key there.
* File permissions: restrict a file to the current user and SYSTEM (removes inherited access).

Every function raises OSError on failure and is only usable on Windows (``IS_WINDOWS``).
"""
from __future__ import annotations

import ctypes
import os
import subprocess

IS_WINDOWS = os.name == "nt"

CRYPTPROTECT_UI_FORBIDDEN = 0x1
CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2          # this user, this PC only - never roams to other machines
CRED_MAX_BLOB = 2560
ERROR_NOT_FOUND = 1168

if IS_WINDOWS:
    from ctypes import wintypes

    class _BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    class _CREDENTIALW(ctypes.Structure):
        _fields_ = [("Flags", wintypes.DWORD), ("Type", wintypes.DWORD), ("TargetName", wintypes.LPWSTR),
                    ("Comment", wintypes.LPWSTR), ("LastWritten", wintypes.FILETIME),
                    ("CredentialBlobSize", wintypes.DWORD), ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
                    ("Persist", wintypes.DWORD), ("AttributeCount", wintypes.DWORD), ("Attributes", ctypes.c_void_p),
                    ("TargetAlias", wintypes.LPWSTR), ("UserName", wintypes.LPWSTR)]

    _crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)

    for _fn in (_crypt32.CryptProtectData, _crypt32.CryptUnprotectData):
        _fn.argtypes = [ctypes.POINTER(_BLOB), ctypes.c_void_p, ctypes.POINTER(_BLOB), ctypes.c_void_p,
                        ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_BLOB)]
        _fn.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p
    _advapi32.CredWriteW.argtypes = [ctypes.POINTER(_CREDENTIALW), wintypes.DWORD]
    _advapi32.CredWriteW.restype = wintypes.BOOL
    _advapi32.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                    ctypes.POINTER(ctypes.POINTER(_CREDENTIALW))]
    _advapi32.CredReadW.restype = wintypes.BOOL
    _advapi32.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
    _advapi32.CredDeleteW.restype = wintypes.BOOL
    _advapi32.CredFree.argtypes = [ctypes.c_void_p]
    _advapi32.CredFree.restype = None


def _require_windows():
    if not IS_WINDOWS:
        raise OSError("This operation needs Windows.")


def _blob(data: bytes):
    buf = ctypes.create_string_buffer(data, len(data))
    return _BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _dpapi(data: bytes, entropy: bytes, protect: bool) -> bytes:
    _require_windows()
    blob_in, keep_in = _blob(data)
    ent, keep_ent = _blob(entropy) if entropy else (None, None)
    blob_out = _BLOB()
    fn = _crypt32.CryptProtectData if protect else _crypt32.CryptUnprotectData
    ok = fn(ctypes.byref(blob_in), None, ctypes.byref(ent) if ent is not None else None,
            None, None, CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out))
    del keep_in, keep_ent
    if not ok:
        raise OSError(f"DPAPI {'protect' if protect else 'unprotect'} failed (error {ctypes.get_last_error()})")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def dpapi_protect(data: bytes, entropy: bytes = b"") -> bytes:
    """Encrypt for the current Windows user on this PC."""
    return _dpapi(data, entropy, True)


def dpapi_unprotect(blob: bytes, entropy: bytes = b"") -> bytes:
    return _dpapi(blob, entropy, False)


def cred_write(target: str, secret: bytes, username: str = "NetUpgrade", comment: str = "") -> None:
    """Create or replace a generic credential for the current user on this PC."""
    _require_windows()
    if len(secret) > CRED_MAX_BLOB:
        raise OSError("Credential too large for Credential Manager.")
    buf = (ctypes.c_ubyte * len(secret)).from_buffer_copy(secret)
    cred = _CREDENTIALW()
    cred.Type = CRED_TYPE_GENERIC
    cred.TargetName = target
    cred.Comment = comment or None
    cred.CredentialBlobSize = len(secret)
    cred.CredentialBlob = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
    cred.Persist = CRED_PERSIST_LOCAL_MACHINE
    cred.UserName = username
    if not _advapi32.CredWriteW(ctypes.byref(cred), 0):
        raise OSError(f"CredWrite failed (error {ctypes.get_last_error()})")
    ctypes.memset(buf, 0, len(secret))


def cred_read(target: str) -> bytes | None:
    """The credential's secret bytes, or None when there is no such credential."""
    _require_windows()
    pcred = ctypes.POINTER(_CREDENTIALW)()
    if not _advapi32.CredReadW(target, CRED_TYPE_GENERIC, 0, ctypes.byref(pcred)):
        err = ctypes.get_last_error()
        if err == ERROR_NOT_FOUND:
            return None
        raise OSError(f"CredRead failed (error {err})")
    try:
        c = pcred.contents
        return ctypes.string_at(c.CredentialBlob, c.CredentialBlobSize)
    finally:
        _advapi32.CredFree(pcred)


def cred_delete(target: str) -> bool:
    _require_windows()
    if _advapi32.CredDeleteW(target, CRED_TYPE_GENERIC, 0):
        return True
    if ctypes.get_last_error() == ERROR_NOT_FOUND:
        return False
    raise OSError(f"CredDelete failed (error {ctypes.get_last_error()})")


def restrict_to_current_user(path) -> bool:
    """Best effort: give only the current user and SYSTEM access to ``path`` (no inherited entries).
    Access is granted first and inheritance removed only after that succeeded, so a failure can never
    lock the user out. Returns True when both steps worked."""
    if not IS_WINDOWS:
        try:
            os.chmod(path, 0o600)
            return True
        except OSError:
            return False
    user = os.environ.get("USERNAME")
    domain = os.environ.get("USERDOMAIN")
    if not user:
        return False
    account = f"{domain}\\{user}" if domain else user
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        grant = subprocess.run(["icacls", str(path), "/grant:r", f"{account}:(F)", "*S-1-5-18:(F)"],
                               capture_output=True, timeout=30, creationflags=flags)
        if grant.returncode != 0:
            return False
        strip = subprocess.run(["icacls", str(path), "/inheritance:r"],
                               capture_output=True, timeout=30, creationflags=flags)
        return strip.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def secure_delete(path) -> None:
    """Overwrite a file with random bytes, flush it to disk, then delete it. On SSDs and some file
    systems the old blocks may survive; this still removes the content from the normal file view."""
    try:
        size = os.path.getsize(path)
        with open(path, "r+b", buffering=0) as fh:
            remaining = size
            while remaining > 0:
                chunk = min(remaining, 1024 * 1024)
                fh.write(os.urandom(chunk))
                remaining -= chunk
            fh.flush()
            os.fsync(fh.fileno())
    except OSError:
        pass
    os.remove(path)
