"""Wait for NetUpgrade to exit, install a verified update, then restart it.

This module is packaged as NetUpgradeUpdater.exe beside NetUpgrade.exe by Inno Setup.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import hmac
import os
import re
import subprocess
import sys
import time
from pathlib import Path

APP_NAME = "NetUpgrade"
INSTALL_RELATIVE = Path("Programs") / APP_NAME / f"{APP_NAME}.exe"
UPDATE_RELATIVE = Path(APP_NAME) / "Updates"
_INSTALLER_RE = re.compile(r"^NetUpgrade-Setup-v\d+\.\d+\.\d+\.exe$")
_SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")
WAIT_TIMEOUT_SECONDS = 180
INSTALL_TIMEOUT_SECONDS = 600


def expected_paths(local_app_data: str | os.PathLike) -> tuple[Path, Path]:
    root = Path(local_app_data)
    return root / INSTALL_RELATIVE, root / UPDATE_RELATIVE


def validate_request(pid: int, installer: str | os.PathLike, sha256: str,
                    app_path: str | os.PathLike, local_app_data: str | os.PathLike | None = None):
    """Validate all paths and arguments before running the installer."""
    if pid <= 0 or pid == os.getpid():
        raise ValueError("Invalid NetUpgrade process id.")
    local_app_data = local_app_data or os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise ValueError("Windows local application data path is unavailable.")
    expected_app, updates_dir = expected_paths(local_app_data)
    raw_app = Path(app_path)
    raw_installer = Path(installer)
    if raw_installer.is_symlink():
        raise ValueError("The downloaded installer is missing or unsafe.")
    supplied_app = raw_app.resolve()
    supplied_installer = raw_installer.resolve()
    if os.path.normcase(str(supplied_app)) != os.path.normcase(str(expected_app.resolve())):
        raise ValueError("The installed NetUpgrade path did not match the expected location.")
    if os.path.normcase(str(supplied_installer.parent)) != os.path.normcase(str(updates_dir.resolve())):
        raise ValueError("The installer must be inside NetUpgrade's update folder.")
    if not _INSTALLER_RE.fullmatch(supplied_installer.name):
        raise ValueError("The update file is not a NetUpgrade setup package.")
    if not _SHA256_RE.fullmatch(str(sha256 or "")):
        raise ValueError("The installer SHA-256 digest is invalid.")
    if not supplied_installer.is_file() or supplied_installer.is_symlink():
        raise ValueError("The downloaded installer is missing or unsafe.")
    return supplied_installer, supplied_app, str(sha256).lower()


def file_sha256(path: str | os.PathLike) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def wait_for_process_exit(pid: int, timeout: float = WAIT_TIMEOUT_SECONDS) -> bool:
    """Wait on a Windows process handle; return False on timeout or wait failure."""
    if os.name != "nt":
        raise OSError("The NetUpgrade updater can only run on Windows.")
    synchronize = 0x00100000
    wait_object_0, wait_timeout = 0, 258
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_ulong)
    kernel.WaitForSingleObject.restype = ctypes.c_ulong
    kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
    kernel.CloseHandle.restype = ctypes.c_int
    handle = kernel.OpenProcess(synchronize, False, pid)
    if not handle:
        import psutil
        return not psutil.pid_exists(pid)
    try:
        result = kernel.WaitForSingleObject(handle, max(0, int(timeout * 1000)))
        return result == wait_object_0 and result != wait_timeout
    finally:
        kernel.CloseHandle(handle)


def _show_error(message: str) -> None:
    if os.name == "nt":
        ctypes.windll.user32.MessageBoxW(None, message, APP_NAME, 0x10)


def _log(message: str) -> None:
    root = os.environ.get("LOCALAPPDATA")
    if not root:
        return
    try:
        path = Path(root) / APP_NAME / "update.log"
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as output:
            output.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


def apply_update(pid: int, installer: str, sha256: str, app_path: str,
                 wait_timeout: float = WAIT_TIMEOUT_SECONDS) -> int:
    installer_path, app_exe, expected_digest = validate_request(pid, installer, sha256, app_path)
    if not hmac.compare_digest(file_sha256(installer_path), expected_digest):
        raise ValueError("The installer changed after download and failed SHA-256 verification.")
    _log(f"Waiting for NetUpgrade process {pid} to close.")
    if not wait_for_process_exit(pid, wait_timeout):
        raise TimeoutError("NetUpgrade did not close in time. The update was cancelled.")
    _log(f"Installing {installer_path.name}.")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    result = subprocess.run(
        [str(installer_path), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART",
         "/SP-"],
        cwd=str(installer_path.parent), creationflags=flags, timeout=INSTALL_TIMEOUT_SECONDS,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"The installer stopped with exit code {result.returncode}.")
    if not app_exe.is_file():
        raise RuntimeError("The installer finished but NetUpgrade.exe was not found.")
    _log("Installation complete; starting NetUpgrade.")
    subprocess.Popen([str(app_exe)], cwd=str(app_exe.parent),
                     creationflags=flags, close_fds=True)
    return 0


def relaunch_from_updates(args) -> bool:
    """Run a private copy outside the app folder so setup can replace installed files."""
    if os.name != "nt":
        return False
    local = os.environ.get("LOCALAPPDATA")
    if not local:
        raise ValueError("Windows local application data path is unavailable.")
    updates_dir = expected_paths(local)[1]
    if Path(sys.executable).resolve().parent == updates_dir.resolve():
        return False
    updates_dir.mkdir(parents=True, exist_ok=True)
    staged = updates_dir / f"NetUpgradeUpdater-{os.getpid()}.exe"
    with open(sys.executable, "rb") as source, staged.open("wb") as dest:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            dest.write(chunk)
        dest.flush()
        os.fsync(dest.fileno())
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen([str(staged), "--pid", str(args.pid), "--installer", args.installer,
                      "--sha256", args.sha256, "--app", args.app],
                     cwd=str(updates_dir), creationflags=flags, close_fds=True)
    return True


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--installer", required=True)
    parser.add_argument("--sha256", required=True)
    parser.add_argument("--app", required=True)
    args = parser.parse_args(argv)
    try:
        if relaunch_from_updates(args):
            return 0
        return apply_update(args.pid, args.installer, args.sha256, args.app)
    except Exception as exc:
        _log(f"Update failed: {type(exc).__name__}: {exc}")
        _show_error(f"NetUpgrade could not finish the update.\n\n{exc}\n\n"
                    f"The current version is still installed. See the update log in "
                    f"%LOCALAPPDATA%\\NetUpgrade\\update.log.")
        return 1


if __name__ == "__main__":
    sys.exit(main())

