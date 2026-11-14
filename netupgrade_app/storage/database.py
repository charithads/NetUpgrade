"""SQLite database for NetUpgrade's application state (no secrets - those live in the vault).

* One file, ``netupgrade.db``, in the per-user data folder (%LOCALAPPDATA%\\NetUpgrade).
* WAL journal, so the dashboard, the Windows launcher and background workers can use it at once.
* ``secure_delete`` is on: deleted rows are overwritten with zeros inside the file.
* Foreign keys are enforced; deleting a site removes everything that belongs to it.
* The schema is versioned (``PRAGMA user_version``) and upgraded in place by ``MIGRATIONS``.
"""
from __future__ import annotations

import json
import sqlite3
import threading
from contextlib import contextmanager
from pathlib import Path

SCHEMA_VERSION = 1

MIGRATIONS = {
    1: """
    CREATE TABLE app_settings (
        scope       TEXT NOT NULL,
        key         TEXT NOT NULL,
        value       TEXT NOT NULL,                 -- JSON; secrets only ever appear as 'vault:<id>' references
        updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
        PRIMARY KEY (scope, key)
    );

    CREATE TABLE web_users (                       -- dashboard accounts managed by the Windows launcher
        name        TEXT PRIMARY KEY COLLATE NOCASE,
        salt        TEXT NOT NULL,
        hash        TEXT NOT NULL,                 -- PBKDF2-HMAC-SHA256, 200 000 rounds (one-way, not reversible)
        role        TEXT NOT NULL DEFAULT 'viewer',
        is_admin    INTEGER NOT NULL DEFAULT 0,    -- the launcher's own administrator account
        extra       TEXT NOT NULL DEFAULT '{}'
    );
    CREATE UNIQUE INDEX ux_web_users_admin ON web_users(is_admin) WHERE is_admin = 1;

    CREATE TABLE timeout_profiles (
        name        TEXT PRIMARY KEY COLLATE NOCASE,
        timeouts    TEXT NOT NULL,                 -- JSON {setting key: seconds}
        note        TEXT NOT NULL DEFAULT '',
        saved_at    TEXT NOT NULL DEFAULT '',
        saved_by    TEXT NOT NULL DEFAULT '',
        site        TEXT NOT NULL DEFAULT '',
        source      TEXT NOT NULL DEFAULT 'manual',
        switch      TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE diag_tasks (
        id          TEXT PRIMARY KEY,
        position    INTEGER NOT NULL,
        data        TEXT NOT NULL                  -- JSON task record
    );

    CREATE TABLE sites (
        id          INTEGER PRIMARY KEY,
        name        TEXT NOT NULL UNIQUE COLLATE NOCASE,
        created_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime'))
    );

    CREATE TABLE devices (
        site_id     INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        host        TEXT NOT NULL,
        position    INTEGER NOT NULL,
        label       TEXT NOT NULL DEFAULT '',
        type        TEXT NOT NULL DEFAULT 'cisco-xe',
        platform    TEXT NOT NULL DEFAULT '',
        extra       TEXT NOT NULL DEFAULT '{}',
        PRIMARY KEY (site_id, host)
    );

    CREATE TABLE credential_profiles (
        site_id      INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        name         TEXT NOT NULL,
        username     TEXT NOT NULL,
        password_ref TEXT NOT NULL DEFAULT '',     -- 'vault:<id>' - the password itself is only in the vault
        secret_ref   TEXT NOT NULL DEFAULT '',     -- enable secret, same
        protocol     TEXT NOT NULL DEFAULT 'ssh' CHECK (protocol IN ('ssh', 'telnet')),
        port         INTEGER NOT NULL DEFAULT 22 CHECK (port BETWEEN 1 AND 65535),
        PRIMARY KEY (site_id, name)
    );

    CREATE TABLE profile_assignments (
        site_id      INTEGER NOT NULL,
        host         TEXT NOT NULL,
        profile_name TEXT NOT NULL,
        PRIMARY KEY (site_id, host),
        FOREIGN KEY (site_id, profile_name) REFERENCES credential_profiles(site_id, name)
            ON DELETE CASCADE ON UPDATE CASCADE
    );

    CREATE TABLE schedules (
        site_id     INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        host        TEXT NOT NULL,
        run_at      REAL NOT NULL,
        data        TEXT NOT NULL,                 -- JSON schedule entry (never holds credentials)
        PRIMARY KEY (site_id, host)
    );

    CREATE TABLE upgrade_history (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        site_id     INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        ip_address  TEXT NOT NULL DEFAULT '',
        status      TEXT NOT NULL DEFAULT '',
        finished    TEXT NOT NULL DEFAULT '',
        data        TEXT NOT NULL                  -- JSON run record
    );
    CREATE INDEX ix_upgrade_history_site ON upgrade_history(site_id, id);

    CREATE TABLE known_versions (
        site_id     INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        ip          TEXT NOT NULL,
        hostname    TEXT NOT NULL DEFAULT '',
        version     TEXT NOT NULL DEFAULT '',
        model       TEXT NOT NULL DEFAULT '',
        boot        TEXT NOT NULL DEFAULT '',      -- JSON or ''
        updated_at  TEXT NOT NULL DEFAULT '',
        PRIMARY KEY (site_id, ip)
    );

    CREATE TABLE site_documents (                  -- small per-site state: notifications, topology, diagnosis, ...
        site_id     INTEGER NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
        kind        TEXT NOT NULL,
        data        TEXT NOT NULL,
        updated_at  TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
        PRIMARY KEY (site_id, kind)
    );

    CREATE TABLE migrations_done (                 -- one-time imports of the old JSON files
        name        TEXT PRIMARY KEY,
        done_at     TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
        detail      TEXT NOT NULL DEFAULT ''
    );
    """,
}


class Database:
    """Thread-safe access to one SQLite file: one connection per thread, explicit transactions."""

    def __init__(self, path):
        self.path = Path(path)
        self._local = threading.local()
        self._init_lock = threading.Lock()
        self._ready = False

    # ---- connections ------------------------------------------------------------------- #
    def _connect(self):
        conn = sqlite3.connect(str(self.path), timeout=30, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA busy_timeout = 30000")
        conn.execute("PRAGMA secure_delete = ON")
        conn.execute("PRAGMA synchronous = NORMAL")
        return conn

    def conn(self):
        self._ensure_ready()
        c = getattr(self._local, "conn", None)
        if c is None:
            c = self._local.conn = self._connect()
            self._local.depth = 0
        return c

    def close(self):
        c = getattr(self._local, "conn", None)
        if c is not None:
            c.close()
            self._local.conn = None

    def _ensure_ready(self):
        if self._ready:
            return
        with self._init_lock:
            if self._ready:
                return
            new = not self.path.exists()
            self.path.parent.mkdir(parents=True, exist_ok=True)
            c = self._connect()
            try:
                c.execute("PRAGMA journal_mode = WAL")
                self._migrate(c)
            finally:
                c.close()
            if new:
                from .winsec import restrict_to_current_user
                restrict_to_current_user(self.path)
            self._ready = True

    @staticmethod
    def _migrate(c):
        current = c.execute("PRAGMA user_version").fetchone()[0]
        if current > SCHEMA_VERSION:
            raise RuntimeError(f"The NetUpgrade database was created by a newer version (schema {current}). "
                               "Install the newer NetUpgrade or restore the previous database.")
        for version in range(current + 1, SCHEMA_VERSION + 1):
            c.execute("BEGIN IMMEDIATE")
            try:
                for statement in _split_sql(MIGRATIONS[version]):
                    c.execute(statement)
                c.execute(f"PRAGMA user_version = {version}")
                c.execute("COMMIT")
            except BaseException:
                c.execute("ROLLBACK")
                raise

    # ---- transactions and queries ------------------------------------------------------ #
    @contextmanager
    def transaction(self):
        """``with db.transaction() as c:`` - nested calls join the outer transaction."""
        c = self.conn()
        if self._local.depth == 0:
            c.execute("BEGIN IMMEDIATE")
        self._local.depth += 1
        try:
            yield c
        except BaseException:
            self._local.depth -= 1
            if self._local.depth == 0:
                c.execute("ROLLBACK")
            raise
        else:
            self._local.depth -= 1
            if self._local.depth == 0:
                c.execute("COMMIT")

    def all(self, sql, params=()):
        return self.conn().execute(sql, params).fetchall()

    def one(self, sql, params=()):
        return self.conn().execute(sql, params).fetchone()

    def run(self, sql, params=()):
        with self.transaction() as c:
            return c.execute(sql, params)


def _split_sql(script):
    """Split a migration script into statements (none of ours contain ';' inside strings)."""
    lines = [ln.split("--", 1)[0] for ln in script.splitlines()]
    return [s.strip() for s in "\n".join(lines).split(";") if s.strip()]


def dumps(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def loads(text, default=None):
    try:
        return json.loads(text)
    except (TypeError, ValueError):
        return default
