"""One-time import of the JSON files older NetUpgrade versions wrote.

For every legacy file that exists:
  1. its content is imported into the database (only when the database has nothing of that kind yet, so
     a stray old file can never overwrite newer data);
  2. secrets inside it (DPAPI- or Fernet-encrypted passwords and keys) are decrypted with the old method
     and re-encrypted into the vault;
  3. a backup copy goes to  <data dir>/legacy-json-backup/<timestamp>/...  - for files that held secrets the
     copy is REDACTED (the secret values are replaced by a marker), so no password stays on disk;
  4. the original is overwritten with random bytes and deleted.
A file that cannot be imported is left exactly where it is and the problem is reported.
"""
from __future__ import annotations

import json
import shutil
import time
from pathlib import Path

from . import data_dir, protect, repositories as repo
from .legacy import unprotect_legacy
from .winsec import secure_delete

REDACTED = "[moved to the NetUpgrade vault]"
_RUN_STAMP = time.strftime("%Y%m%d_%H%M%S")


def _read(path):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _backup_root():
    return data_dir() / "legacy-json-backup" / _RUN_STAMP


def _retire(path, backup_rel, redacted=None, keep_original_copy=False):
    """Back up (redacted when given) and securely delete a migrated legacy file."""
    dest = _backup_root() / backup_rel
    dest.parent.mkdir(parents=True, exist_ok=True)
    if redacted is not None and not keep_original_copy:
        dest.write_text(json.dumps(redacted, indent=2), encoding="utf-8")
    else:
        shutil.copy2(path, dest)
    secure_delete(path)
    for leftover in (Path(str(path) + ".tmp"), Path(path).with_suffix(".json.tmp")):
        if leftover.exists():
            try:
                secure_delete(leftover)
            except OSError:
                pass


class _Secrets:
    """Re-encrypts legacy secrets into the vault, remembering whether any could not be read."""

    def __init__(self, legacy_unprotect, purpose):
        self.legacy_unprotect = legacy_unprotect
        self.purpose = purpose
        self.failed = 0

    def convert(self, value):
        if not value or not isinstance(value, str):
            return ""
        try:
            plain = self.legacy_unprotect(value)
        except Exception:  # noqa: BLE001 - another user's DPAPI blob, missing Fernet key, damaged value
            self.failed += 1
            return ""
        return protect(plain, self.purpose) if plain else ""


def _run(name, path, importer, backup_rel, log, results):
    """Import one legacy file; never raises."""
    path = Path(path)
    if not path.is_file():
        return
    try:
        data = _read(path)
    except (OSError, ValueError) as exc:
        log(f"Storage migration: '{path.name}' could not be read ({type(exc).__name__}) - left in place.")
        results.append((name, "unreadable"))
        return
    try:
        outcome = importer(data) or {}
    except Exception as exc:  # noqa: BLE001
        log(f"Storage migration: '{path}' could not be imported ({type(exc).__name__}: {exc}) - left in place.")
        results.append((name, "failed"))
        return
    try:
        _retire(path, backup_rel, outcome.get("redacted"), keep_original_copy=outcome.get("keep_original", False))
        results.append((name, outcome.get("status", "imported")))
    except OSError as exc:
        log(f"Storage migration: '{path}' was imported but could not be removed ({exc}).")
        results.append((name, "imported, file kept"))


# ----------------------------------------------------------------------------------------------- #
# Global files (next to the program)
# ----------------------------------------------------------------------------------------------- #
def migrate_global(program_dir, legacy_unprotect=unprotect_legacy, log=print):
    program_dir = Path(program_dir)
    results = []

    def settings(data):
        if not repo.has_setting("app", "settings") and isinstance(data, dict):
            repo.set_setting("app", "settings", data)
        return {}

    def timeouts(data):
        if not repo.load_timeout_profiles()["profiles"] and isinstance(data, dict):
            repo.save_timeout_profiles({"profiles": data.get("profiles") or {}, "last_applied": data.get("last_applied") or ""})
        return {}

    def diag_tasks(data):
        tasks = data.get("tasks") if isinstance(data, dict) else None
        if not repo.load_diag_tasks() and isinstance(tasks, list):
            repo.save_diag_tasks([t for t in tasks if isinstance(t, dict) and isinstance(t.get("id"), str)])
        return {}

    def secret_config(scope, field, purpose):
        def importer(data):
            if not isinstance(data, dict):
                return {}
            redacted = dict(data)
            if field in redacted and redacted[field]:
                redacted[field] = REDACTED
            if repo.has_setting(scope, "config"):
                return {"redacted": redacted, "status": "skipped (already in database)"}
            conv = _Secrets(legacy_unprotect, purpose)
            cfg = dict(data)
            cfg[field] = conv.convert(data.get(field))
            repo.set_setting(scope, "config", cfg)
            if conv.failed:
                log(f"Storage migration: the saved {scope} secret could not be decrypted for this Windows user - "
                    "enter it again in Settings.")
                return {"redacted": redacted, "keep_original": True, "status": "imported, secret must be re-entered"}
            return {"redacted": redacted}
        return importer

    _run("settings", program_dir / "settings.json", settings, "global/settings.json", log, results)
    _run("command_profiles", program_dir / "command_profiles.json", timeouts, "global/command_profiles.json", log, results)
    _run("diag_tasks", program_dir / "diag_tasks.json", diag_tasks, "global/diag_tasks.json", log, results)
    _run("email_settings", program_dir / "email_settings.json", secret_config("email", "password_enc", "smtp-password"),
         "global/email_settings.json", log, results)
    _run("ai_settings", program_dir / "ai_settings.json", secret_config("ai", "api_key_enc", "ai-api-key"),
         "global/ai_settings.json", log, results)
    _run("device_versions", program_dir / "device_versions.json", lambda data: {"status": "backed up (no site)"},
         "global/device_versions.json", log, results)
    return results


# ----------------------------------------------------------------------------------------------- #
# Per-site files (in each site's folder)
# ----------------------------------------------------------------------------------------------- #
SITE_FILES = ("devices.json", "upgrade_history.json", "device_versions.json", "credential_profiles.json",
              "schedules.json", "notifications.json", "topology.json", "diagnosis.json", "upgrade_settings.json")


def site_has_legacy_files(site_folder):
    folder = Path(site_folder)
    return any((folder / f).is_file() for f in SITE_FILES)


def migrate_site(site, site_folder, legacy_unprotect=unprotect_legacy, log=print):
    folder = Path(site_folder)
    results = []
    if not site_has_legacy_files(folder):
        return results
    repo.site_id(site)
    rel = f"sites/{site}"

    def devices(data):
        if not repo.site_has(site, "devices") and isinstance(data, list):
            repo.save_devices(site, [d for d in data if isinstance(d, dict) and d.get("host")])
        return {}

    def history(data):
        if not repo.site_has(site, "history") and isinstance(data, list):
            repo.replace_history(site, data)
        return {}

    def versions(data):
        if not repo.site_has(site, "versions") and isinstance(data, dict):
            for ip, entry in data.items():
                if isinstance(entry, dict):
                    repo.save_known_version(site, ip, entry)
        return {}

    def profiles(data):
        if not isinstance(data, dict):
            return {}
        raw = data.get("profiles") or {}
        redacted = {"profiles": {n: dict(r, password=REDACTED if r.get("password") else "",
                                         secret=REDACTED if r.get("secret") else "")
                                 for n, r in raw.items() if isinstance(r, dict)},
                    "assignments": data.get("assignments") or {}}
        if repo.site_has(site, "profiles"):
            return {"redacted": redacted, "status": "skipped (already in database)"}
        conv = _Secrets(legacy_unprotect, "device-credential")
        out = {}
        for name, r in raw.items():
            if not (isinstance(r, dict) and isinstance(r.get("username"), str)):
                continue
            proto = "telnet" if r.get("protocol") == "telnet" else "ssh"
            try:
                port = int(r.get("port") or (23 if proto == "telnet" else 22))
            except (TypeError, ValueError):
                port = 23 if proto == "telnet" else 22
            out[str(name)] = {"username": r["username"], "password": conv.convert(r.get("password")),
                              "secret": conv.convert(r.get("secret")), "protocol": proto,
                              "port": port if 1 <= port <= 65535 else (23 if proto == "telnet" else 22)}
        assign = {h: p for h, p in (data.get("assignments") or {}).items() if isinstance(h, str) and p in out}
        repo.save_profiles(site, out, assign)
        if conv.failed:
            log(f"Storage migration: {conv.failed} saved password(s) of site '{site}' could not be decrypted for this "
                "Windows user - open Home > Credential Profiles and save those profiles again.")
            return {"redacted": redacted, "keep_original": True, "status": "imported, some passwords must be re-entered"}
        return {"redacted": redacted}

    def schedules(data):
        if not repo.site_has(site, "schedules") and isinstance(data, dict):
            repo.save_schedules(site, {h: e for h, e in data.items() if isinstance(e, dict) and
                                       isinstance(e.get("run_at"), (int, float))})
        return {}

    def doc(kind):
        def importer(data):
            if not repo.site_has(site, kind):
                repo.set_site_doc(site, kind, data)
            return {}
        return importer

    _run("devices", folder / "devices.json", devices, f"{rel}/devices.json", log, results)
    _run("history", folder / "upgrade_history.json", history, f"{rel}/upgrade_history.json", log, results)
    _run("versions", folder / "device_versions.json", versions, f"{rel}/device_versions.json", log, results)
    _run("profiles", folder / "credential_profiles.json", profiles, f"{rel}/credential_profiles.json", log, results)
    _run("schedules", folder / "schedules.json", schedules, f"{rel}/schedules.json", log, results)
    _run("notes", folder / "notifications.json", doc("notes"), f"{rel}/notifications.json", log, results)
    _run("topology", folder / "topology.json", doc("topology"), f"{rel}/topology.json", log, results)
    _run("diagnosis", folder / "diagnosis.json", doc("diagnosis"), f"{rel}/diagnosis.json", log, results)
    _run("upgrade_settings", folder / "upgrade_settings.json", doc("upgrade_settings"),
         f"{rel}/upgrade_settings.json", log, results)
    return results


# ----------------------------------------------------------------------------------------------- #
# The Windows launcher's settings (dashboard accounts, network options, HTTPS certificate details)
# ----------------------------------------------------------------------------------------------- #
def migrate_launcher(settings_file, cert_meta_file=None, legacy_unprotect=unprotect_legacy, log=print):
    results = []

    def settings(data):
        if not isinstance(data, dict):
            return {}
        data = dict(data)
        email = data.pop("email", None)               # old launcher e-mail settings -> Settings > Notifications
        if isinstance(email, dict) and email.get("host") and not repo.has_setting("email", "config"):
            conv = _Secrets(legacy_unprotect, "smtp-password")
            repo.set_setting("email", "config", dict(email, password_enc=conv.convert(email.get("password_enc"))))
        if repo.load_launcher_settings() is None:
            repo.save_launcher_settings(data)
        return {"redacted": _redact_launcher(dict(data, email=email) if email else data)}

    def cert_meta(data):
        if not repo.has_setting("launcher", "tls_meta") and isinstance(data, dict):
            repo.set_setting("launcher", "tls_meta", data)
        return {}

    _run("launcher_settings", settings_file, settings, "launcher/settings.json", log, results)
    if cert_meta_file:
        _run("tls_meta", cert_meta_file, cert_meta, "launcher/cert.json", log, results)
    return results


def migrate_everything(program_dir, sites_root, site_names, log=print):
    """Startup: import every legacy JSON file still present (global, then each site), then tidy up the vault.
    Returns a short summary for the console."""
    results = migrate_global(program_dir, log=log)
    for name in site_names:
        results += [(f"{name}/{n}", s) for n, s in migrate_site(name, Path(sites_root) / name, log=log)]
    imported = [r for r in results if r[1].startswith("imported")]
    problems = [r for r in results if r[1] in ("failed", "unreadable") or "re-entered" in r[1]]
    if imported and not problems:
        from .legacy import retire_legacy_key_files
        retire_legacy_key_files(log)
    try:
        repo.prune_vault()
    except Exception as exc:  # noqa: BLE001 - housekeeping only
        log(f"Vault housekeeping skipped ({type(exc).__name__}).")
    if imported:
        log(f"Storage migration: {len(imported)} JSON file(s) moved into the database "
            f"(backups in {_backup_root()}).")
    return results


def _redact_launcher(data):
    out = dict(data)
    if isinstance(out.get("admin"), dict):
        out["admin"] = {k: (REDACTED if k in ("salt", "hash") else v) for k, v in out["admin"].items()}
    if isinstance(out.get("users"), dict):
        out["users"] = {n: {k: (REDACTED if k in ("salt", "hash") else v) for k, v in r.items()}
                        for n, r in out["users"].items() if isinstance(r, dict)}
    email = out.get("email")
    if isinstance(email, dict) and email.get("password_enc"):
        out["email"] = dict(email, password_enc=REDACTED)
    return out
