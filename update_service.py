"""Safe release discovery and installer download helpers for NetUpgrade."""

from __future__ import annotations

import hashlib
import hmac
import os
import re
import tempfile
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

APP_NAME = "NetUpgrade"
MAX_INSTALLER_BYTES = 512 * 1024 * 1024
_VERSION_RE = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")
_SHA256_RE = re.compile(r"^sha256:([0-9a-fA-F]{64})$")


def parse_version(value: str) -> tuple[int, int, int] | None:
    """Parse a stable x.y.z release tag; reject prereleases and malformed tags."""
    match = _VERSION_RE.fullmatch(str(value or "").strip())
    return tuple(int(part) for part in match.groups()) if match else None


def is_newer_release(release: dict, current_version: str) -> bool:
    latest = parse_version(release.get("tag_name", ""))
    current = parse_version(current_version)
    return latest is not None and current is not None and latest > current


def installer_asset_name(tag: str) -> str:
    """Return the exact installer asset name expected for a release tag."""
    version = parse_version(tag)
    if version is None:
        raise ValueError("GitHub returned an invalid release version.")
    return f"{APP_NAME}-Setup-v{'.'.join(map(str, version))}.exe"


def pick_installer_asset(release: dict) -> dict | None:
    """Choose only this release's Windows installer, with GitHub SHA-256 metadata."""
    tag = release.get("tag_name", "")
    try:
        expected_name = installer_asset_name(tag)
    except ValueError:
        return None
    for asset in release.get("assets") or []:
        if asset.get("name") != expected_name:
            continue
        url = urlparse(str(asset.get("browser_download_url") or ""))
        digest = _SHA256_RE.fullmatch(str(asset.get("digest") or ""))
        size = asset.get("size")
        if (url.scheme == "https" and url.hostname == "github.com" and url.path.startswith("/")
                and digest and isinstance(size, int) and 0 < size <= MAX_INSTALLER_BYTES):
            return {**asset, "sha256": digest.group(1).lower()}
    return None


def download_verified_installer(asset: dict, directory: str | os.PathLike,
                                progress=None, opener=urlopen) -> Path:
    """Download to a temporary file, verify size and SHA-256, then atomically publish it."""
    expected_name = str(asset.get("name") or "")
    if not re.fullmatch(r"NetUpgrade-Setup-v\d+\.\d+\.\d+\.exe", expected_name):
        raise ValueError("GitHub returned an unexpected installer filename.")
    url = urlparse(str(asset.get("browser_download_url") or ""))
    if url.scheme != "https" or url.hostname != "github.com":
        raise ValueError("GitHub returned an unsafe installer download URL.")
    expected_digest = str(asset.get("sha256") or "").lower()
    if not re.fullmatch(r"[0-9a-f]{64}", expected_digest):
        raise ValueError("GitHub did not provide a valid SHA-256 digest for the installer.")
    expected_size = asset.get("size")
    if not isinstance(expected_size, int) or not 0 < expected_size <= MAX_INSTALLER_BYTES:
        raise ValueError("GitHub returned an invalid installer size.")

    folder = Path(directory)
    folder.mkdir(parents=True, exist_ok=True)
    fd, partial_name = tempfile.mkstemp(prefix="NetUpgrade-", suffix=".part", dir=folder)
    partial = Path(partial_name)
    destination = folder / expected_name
    digest = hashlib.sha256()
    received = 0
    try:
        request = Request(url.geturl(), headers={"User-Agent": "NetUpgrade-Updater"})
        with os.fdopen(fd, "wb") as output, opener(request, timeout=30) as response:
            while True:
                chunk = response.read(256 * 1024)
                if not chunk:
                    break
                received += len(chunk)
                if received > expected_size or received > MAX_INSTALLER_BYTES:
                    raise ValueError("The installer download exceeded the size published by GitHub.")
                output.write(chunk)
                digest.update(chunk)
                if progress:
                    progress(received, expected_size)
            output.flush()
            os.fsync(output.fileno())
        if received != expected_size:
            raise ValueError("The installer download was incomplete.")
        if not hmac.compare_digest(digest.hexdigest(), expected_digest):
            raise ValueError("The downloaded installer failed GitHub's SHA-256 check.")
        os.replace(partial, destination)
        return destination
    except Exception:
        try:
            partial.unlink(missing_ok=True)
        except OSError:
            pass
        raise

