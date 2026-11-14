"""Typed read/write functions over the database - what the application calls instead of JSON files.

Every function takes plain Python data and returns plain Python data in the same shapes the application
used with its JSON files, so callers did not need to change their logic. Secrets never pass through here
in clear: records only carry vault references ('vault:<id>'), and the functions that replace a record
delete the vault entries the old record pointed to.
"""
from __future__ import annotations

import re

from .database import dumps, loads
from .vault import is_ref

_REF_RE = re.compile(r"vault:[0-9a-f]{32}")

HISTORY_MAX = 2000


def _db():
    from . import get_db
    return get_db()


def _vault():
    from . import get_vault
    return get_vault()


def _drop_refs(refs):
    refs = [r for r in refs if is_ref(r)]
    if refs:
        _vault().delete(refs)


# ----------------------------------------------------------------------------------------------- #
# Key/value settings (global settings, e-mail, AI, launcher, ...)
# ----------------------------------------------------------------------------------------------- #
def get_setting(scope, key, default=None):
    row = _db().one("SELECT value FROM app_settings WHERE scope = ? AND key = ?", (scope, key))
    return default if row is None else loads(row["value"], default)


def has_setting(scope, key):
    return _db().one("SELECT 1 FROM app_settings WHERE scope = ? AND key = ?", (scope, key)) is not None


def set_setting(scope, key, value):
    """Store a value; vault references the previous value held and the new one does not are deleted."""
    text = dumps(value)
    with _db().transaction() as c:
        row = c.execute("SELECT value FROM app_settings WHERE scope = ? AND key = ?", (scope, key)).fetchone()
        old_refs = set(_REF_RE.findall(row["value"])) if row else set()
        c.execute("INSERT INTO app_settings(scope, key, value, updated_at) VALUES (?, ?, ?, "
                  "strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')) "
                  "ON CONFLICT(scope, key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at",
                  (scope, key, text))
    _drop_refs(old_refs - set(_REF_RE.findall(text)))


def delete_setting(scope, key):
    with _db().transaction() as c:
        row = c.execute("SELECT value FROM app_settings WHERE scope = ? AND key = ?", (scope, key)).fetchone()
        c.execute("DELETE FROM app_settings WHERE scope = ? AND key = ?", (scope, key))
    if row:
        _drop_refs(_REF_RE.findall(row["value"]))


# ----------------------------------------------------------------------------------------------- #
# Command duration (timeout) profiles
# ----------------------------------------------------------------------------------------------- #
def load_timeout_profiles():
    profiles = {}
    for r in _db().all("SELECT * FROM timeout_profiles ORDER BY name COLLATE NOCASE"):
        profiles[r["name"]] = {"values": loads(r["timeouts"], {}), "note": r["note"], "saved_at": r["saved_at"],
                               "saved_by": r["saved_by"], "site": r["site"], "source": r["source"], "switch": r["switch"]}
    return {"profiles": profiles, "last_applied": get_setting("app", "timeout_profile_last_applied", "")}


def save_timeout_profiles(store):
    with _db().transaction() as c:
        c.execute("DELETE FROM timeout_profiles")
        for name, rec in (store.get("profiles") or {}).items():
            c.execute("INSERT INTO timeout_profiles(name, timeouts, note, saved_at, saved_by, site, source, switch) "
                      "VALUES (?,?,?,?,?,?,?,?)",
                      (name, dumps(rec.get("values") or {}), rec.get("note") or "", rec.get("saved_at") or "",
                       rec.get("saved_by") or "", rec.get("site") or "", rec.get("source") or "manual",
                       rec.get("switch") or ""))
        set_setting("app", "timeout_profile_last_applied", store.get("last_applied") or "")


# ----------------------------------------------------------------------------------------------- #
# Saved Diagnosis tasks
# ----------------------------------------------------------------------------------------------- #
def load_diag_tasks():
    return [t for t in (loads(r["data"]) for r in _db().all("SELECT data FROM diag_tasks ORDER BY position"))
            if isinstance(t, dict)]


def save_diag_tasks(tasks):
    with _db().transaction() as c:
        c.execute("DELETE FROM diag_tasks")
        for i, t in enumerate(tasks):
            c.execute("INSERT OR REPLACE INTO diag_tasks(id, position, data) VALUES (?,?,?)", (str(t.get("id")), i, dumps(t)))


# ----------------------------------------------------------------------------------------------- #
# Sites
# ----------------------------------------------------------------------------------------------- #
def site_id(name, create=True):
    row = _db().one("SELECT id FROM sites WHERE name = ?", (name,))
    if row:
        return row["id"]
    if not create:
        return None
    with _db().transaction() as c:
        c.execute("INSERT OR IGNORE INTO sites(name) VALUES (?)", (name,))
        return c.execute("SELECT id FROM sites WHERE name = ?", (name,)).fetchone()["id"]


def site_exists(name):
    return site_id(name, create=False) is not None


def rename_site(old, new):
    if old.lower() != new.lower() and site_exists(new) and site_exists(old):
        delete_site(new)                      # a stale row whose folder no longer exists
    with _db().transaction() as c:
        c.execute("UPDATE sites SET name = ? WHERE name = ?", (new, old))


def delete_site(name):
    """Delete a site and everything stored for it, including its credential profiles' secrets."""
    sid = site_id(name, create=False)
    if sid is None:
        return
    refs = _profile_refs(sid)
    with _db().transaction() as c:
        c.execute("DELETE FROM profile_assignments WHERE site_id = ?", (sid,))
        c.execute("DELETE FROM sites WHERE id = ?", (sid,))
    _drop_refs(refs)


def site_has(name, kind):
    """True when the database already holds data of this kind for the site (used by the migration)."""
    sid = site_id(name, create=False)
    if sid is None:
        return False
    table = {"devices": "devices", "profiles": "credential_profiles", "schedules": "schedules",
             "history": "upgrade_history", "versions": "known_versions"}.get(kind)
    if table:
        return _db().one(f"SELECT 1 FROM {table} WHERE site_id = ? LIMIT 1", (sid,)) is not None
    return _db().one("SELECT 1 FROM site_documents WHERE site_id = ? AND kind = ?", (sid, kind)) is not None


# ----------------------------------------------------------------------------------------------- #
# Devices
# ----------------------------------------------------------------------------------------------- #
_DEVICE_COLUMNS = ("host", "label", "type", "platform")


def load_devices(site):
    sid = site_id(site, create=False)
    if sid is None:
        return []
    out = []
    for r in _db().all("SELECT * FROM devices WHERE site_id = ? ORDER BY position", (sid,)):
        d = loads(r["extra"], {}) or {}
        d.update(host=r["host"], label=r["label"], type=r["type"])
        if r["platform"]:
            d["platform"] = r["platform"]
        out.append(d)
    return out


def save_devices(site, hosts):
    sid = site_id(site)
    with _db().transaction() as c:
        c.execute("DELETE FROM devices WHERE site_id = ?", (sid,))
        for i, h in enumerate(hosts):
            extra = {k: v for k, v in h.items() if k not in _DEVICE_COLUMNS}
            c.execute("INSERT OR REPLACE INTO devices(site_id, host, position, label, type, platform, extra) "
                      "VALUES (?,?,?,?,?,?,?)",
                      (sid, str(h["host"]), i, str(h.get("label") or ""), str(h.get("type") or "cisco-xe"),
                       str(h.get("platform") or ""), dumps(extra)))


# ----------------------------------------------------------------------------------------------- #
# Credential profiles and which device uses which (the passwords are vault references)
# ----------------------------------------------------------------------------------------------- #
def _profile_refs(sid):
    refs = set()
    for r in _db().all("SELECT password_ref, secret_ref FROM credential_profiles WHERE site_id = ?", (sid,)):
        refs.update(x for x in (r["password_ref"], r["secret_ref"]) if x)
    return refs


def load_profiles(site):
    """(profiles, assignments): {name: {username, password, secret, protocol, port}} with password/secret as
    vault references, and {host: profile name}."""
    sid = site_id(site, create=False)
    if sid is None:
        return {}, {}
    profiles = {}
    for r in _db().all("SELECT * FROM credential_profiles WHERE site_id = ? ORDER BY name COLLATE NOCASE", (sid,)):
        profiles[r["name"]] = {"username": r["username"], "password": r["password_ref"], "secret": r["secret_ref"],
                               "protocol": r["protocol"], "port": r["port"]}
    assign = {r["host"]: r["profile_name"]
              for r in _db().all("SELECT host, profile_name FROM profile_assignments WHERE site_id = ?", (sid,))}
    return profiles, assign


def save_profiles(site, profiles, assignments):
    sid = site_id(site)
    old_refs = _profile_refs(sid)
    new_refs = set()
    with _db().transaction() as c:
        c.execute("DELETE FROM profile_assignments WHERE site_id = ?", (sid,))
        c.execute("DELETE FROM credential_profiles WHERE site_id = ?", (sid,))
        for name, rec in profiles.items():
            pw, sec = rec.get("password") or "", rec.get("secret") or ""
            for ref in (pw, sec):
                if ref and not is_ref(ref):
                    raise ValueError("Credential profile secrets must be stored in the vault.")
            new_refs.update(x for x in (pw, sec) if x)
            proto = "telnet" if rec.get("protocol") == "telnet" else "ssh"
            c.execute("INSERT INTO credential_profiles(site_id, name, username, password_ref, secret_ref, protocol, port) "
                      "VALUES (?,?,?,?,?,?,?)",
                      (sid, name, rec.get("username") or "", pw, sec, proto, int(rec.get("port") or (23 if proto == "telnet" else 22))))
        for host, pname in assignments.items():
            if pname in profiles:
                c.execute("INSERT INTO profile_assignments(site_id, host, profile_name) VALUES (?,?,?)", (sid, host, pname))
    _drop_refs(old_refs - new_refs)


# ----------------------------------------------------------------------------------------------- #
# Scheduled upgrades
# ----------------------------------------------------------------------------------------------- #
def load_schedules(site):
    sid = site_id(site, create=False)
    if sid is None:
        return {}
    return {r["host"]: loads(r["data"], {}) for r in _db().all("SELECT host, data FROM schedules WHERE site_id = ?", (sid,))}


def save_schedules(site, schedule):
    sid = site_id(site)
    with _db().transaction() as c:
        c.execute("DELETE FROM schedules WHERE site_id = ?", (sid,))
        for host, entry in schedule.items():
            c.execute("INSERT INTO schedules(site_id, host, run_at, data) VALUES (?,?,?,?)",
                      (sid, host, float(entry.get("run_at") or 0), dumps(entry)))


# ----------------------------------------------------------------------------------------------- #
# Upgrade run history
# ----------------------------------------------------------------------------------------------- #
def load_history(site):
    sid = site_id(site, create=False)
    if sid is None:
        return []
    return [loads(r["data"], {}) for r in _db().all("SELECT data FROM upgrade_history WHERE site_id = ? ORDER BY id", (sid,))]


def append_history(site, record, keep=HISTORY_MAX):
    sid = site_id(site)
    with _db().transaction() as c:
        c.execute("INSERT INTO upgrade_history(site_id, ip_address, status, finished, data) VALUES (?,?,?,?,?)",
                  (sid, str(record.get("ip_address") or ""), str(record.get("status") or ""),
                   str(record.get("finished") or ""), dumps(record)))
        c.execute("DELETE FROM upgrade_history WHERE site_id = ? AND id NOT IN "
                  "(SELECT id FROM upgrade_history WHERE site_id = ? ORDER BY id DESC LIMIT ?)", (sid, sid, keep))


def replace_history(site, records, keep=HISTORY_MAX):
    sid = site_id(site)
    with _db().transaction() as c:
        c.execute("DELETE FROM upgrade_history WHERE site_id = ?", (sid,))
        for record in list(records)[-keep:]:
            if isinstance(record, dict):
                c.execute("INSERT INTO upgrade_history(site_id, ip_address, status, finished, data) VALUES (?,?,?,?,?)",
                          (sid, str(record.get("ip_address") or ""), str(record.get("status") or ""),
                           str(record.get("finished") or ""), dumps(record)))


def history_revision(site):
    """Changes whenever the site's history changes (used to cache values computed from it)."""
    sid = site_id(site, create=False)
    if sid is None:
        return None
    row = _db().one("SELECT MAX(id) AS m, COUNT(*) AS n FROM upgrade_history WHERE site_id = ?", (sid,))
    return (row["m"], row["n"])


# ----------------------------------------------------------------------------------------------- #
# Last known software version per device
# ----------------------------------------------------------------------------------------------- #
def load_known_versions(site):
    sid = site_id(site, create=False)
    if sid is None:
        return {}
    out = {}
    for r in _db().all("SELECT * FROM known_versions WHERE site_id = ?", (sid,)):
        e = {"hostname": r["hostname"], "version": r["version"], "updated_at": r["updated_at"]}
        if r["model"]:
            e["model"] = r["model"]
        boot = loads(r["boot"]) if r["boot"] else None
        if isinstance(boot, dict) and boot:
            e["boot"] = boot
        out[r["ip"]] = e
    return out


def save_known_version(site, ip, entry):
    sid = site_id(site)
    boot = entry.get("boot")
    with _db().transaction() as c:
        c.execute("INSERT OR REPLACE INTO known_versions(site_id, ip, hostname, version, model, boot, updated_at) "
                  "VALUES (?,?,?,?,?,?,?)",
                  (sid, ip, entry.get("hostname") or ip, entry.get("version") or "", entry.get("model") or "",
                   dumps(boot) if isinstance(boot, dict) and boot else "", entry.get("updated_at") or ""))


# ----------------------------------------------------------------------------------------------- #
# Small per-site documents: notifications, topology, diagnosis commands, upgrade settings
# ----------------------------------------------------------------------------------------------- #
def get_site_doc(site, kind, default=None):
    sid = site_id(site, create=False)
    if sid is None:
        return default
    row = _db().one("SELECT data FROM site_documents WHERE site_id = ? AND kind = ?", (sid, kind))
    return default if row is None else loads(row["data"], default)


def set_site_doc(site, kind, value):
    sid = site_id(site)
    with _db().transaction() as c:
        c.execute("INSERT INTO site_documents(site_id, kind, data, updated_at) VALUES (?, ?, ?, "
                  "strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')) "
                  "ON CONFLICT(site_id, kind) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at",
                  (sid, kind, dumps(value)))


# ----------------------------------------------------------------------------------------------- #
# Launcher (Windows application) settings and dashboard accounts
# ----------------------------------------------------------------------------------------------- #
def load_launcher_settings():
    """The launcher's settings dict, with 'users' and 'admin' rebuilt from the web_users table."""
    data = get_setting("launcher", "settings", None)
    if not isinstance(data, dict):
        return None
    users, admin = {}, None
    for r in _db().all("SELECT * FROM web_users ORDER BY name COLLATE NOCASE"):
        rec = loads(r["extra"], {}) or {}
        rec.update(salt=r["salt"], hash=r["hash"])
        if r["is_admin"]:
            admin = dict(rec, name=r["name"])
        else:
            rec["role"] = r["role"]
            users[r["name"]] = rec
    data["users"] = users
    if admin:
        data["admin"] = admin
    else:
        data.pop("admin", None)
    return data


def save_launcher_settings(settings):
    data = {k: v for k, v in settings.items() if k not in ("users", "admin")}
    with _db().transaction() as c:
        c.execute("DELETE FROM web_users")
        admin = settings.get("admin")
        if isinstance(admin, dict) and admin.get("name"):
            extra = {k: v for k, v in admin.items() if k not in ("name", "salt", "hash", "role")}
            c.execute("INSERT INTO web_users(name, salt, hash, role, is_admin, extra) VALUES (?,?,?,?,1,?)",
                      (admin["name"], admin["salt"], admin["hash"], "administrator", dumps(extra)))
        for name, rec in (settings.get("users") or {}).items():
            if isinstance(admin, dict) and name.lower() == str(admin.get("name", "")).lower():
                continue
            extra = {k: v for k, v in rec.items() if k not in ("salt", "hash", "role", "admin")}
            c.execute("INSERT OR REPLACE INTO web_users(name, salt, hash, role, is_admin, extra) VALUES (?,?,?,?,0,?)",
                      (name, rec["salt"], rec["hash"], rec.get("role") or "viewer", dumps(extra)))
        set_setting("launcher", "settings", data)


# ----------------------------------------------------------------------------------------------- #
# Vault housekeeping
# ----------------------------------------------------------------------------------------------- #
def referenced_vault_refs():
    """Every vault reference the database still points to."""
    refs = set()
    for r in _db().all("SELECT password_ref, secret_ref FROM credential_profiles"):
        refs.update(x for x in (r["password_ref"], r["secret_ref"]) if x)
    for r in _db().all("SELECT value FROM app_settings"):
        refs.update(_REF_RE.findall(r["value"]))
    for r in _db().all("SELECT data FROM site_documents"):
        refs.update(_REF_RE.findall(r["data"]))
    return refs


def prune_vault(grace_seconds=3600):
    return _vault().prune(referenced_vault_refs(), grace_seconds)


def migration_done(name):
    return _db().one("SELECT 1 FROM migrations_done WHERE name = ?", (name,)) is not None


def mark_migration_done(name, detail=""):
    _db().run("INSERT OR REPLACE INTO migrations_done(name, detail) VALUES (?, ?)", (name, detail[:2000]))
