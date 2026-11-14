"""Tests for the database, the secret vault and the one-time JSON migration."""
import base64
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
from netupgrade_app import storage
from netupgrade_app.storage import migrate, repositories as repo, vault as vault_mod, winsec


class StorageTestCase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.key = vault_mod.MemoryKeyProvider()
        storage.configure(Path(self.temp.name) / "data", self.key)
        self.addCleanup(storage.configure, None, None)


class VaultTests(StorageTestCase):
    def test_round_trip_and_reference_format(self):
        ref = storage.protect("S3cret!", "device-password")
        self.assertTrue(ref.startswith("vault:"))
        self.assertNotIn("S3cret", ref)
        self.assertEqual(storage.reveal(ref), "S3cret!")
        self.assertEqual(storage.protect(""), "")
        self.assertEqual(storage.reveal(""), "")

    def test_plaintext_never_written_to_disk(self):
        storage.protect("UniquePlaintextMarker42", "device-password")
        for f in Path(self.temp.name).rglob("*"):
            if f.is_file():
                self.assertNotIn(b"UniquePlaintextMarker42", f.read_bytes(), f.name)

    def test_tampered_or_swapped_ciphertext_is_rejected(self):
        a = storage.protect("first", "p")
        b = storage.protect("second", "p")
        db = Path(self.temp.name) / "data" / storage.VAULT_FILENAME
        conn = sqlite3.connect(str(db))
        ct_b = conn.execute("SELECT ciphertext, nonce FROM secrets WHERE id = ?", (b[6:],)).fetchone()
        conn.execute("UPDATE secrets SET ciphertext = ?, nonce = ? WHERE id = ?", (ct_b[0], ct_b[1], a[6:]))
        conn.commit()
        conn.close()
        with self.assertRaises(storage.VaultError):        # bound to its id: cannot be moved to another record
            storage.reveal(a)

    def test_lost_key_gives_a_clear_error(self):
        ref = storage.protect("x", "p")
        storage.configure(Path(self.temp.name) / "data", vault_mod.MemoryKeyProvider())   # a different key
        with self.assertRaisesRegex(storage.VaultError, "Enter it again"):
            storage.reveal(ref)

    def test_prune_keeps_referenced_and_recent_secrets(self):
        keep = storage.protect("keep", "p")
        drop = storage.protect("drop", "p")
        self.assertEqual(storage.get_vault().prune({keep}, grace_seconds=0), 1)
        self.assertEqual(storage.reveal(keep), "keep")
        with self.assertRaises(storage.VaultError):
            storage.reveal(drop)
        fresh = storage.protect("fresh", "p")
        storage.get_vault().prune(set(), grace_seconds=3600)
        self.assertEqual(storage.reveal(fresh), "fresh")

    def test_worker_handoff_is_sealed(self):
        secret = "Sw1tch-handoff-P@ss-0451"            # long enough never to appear in ciphertext by chance
        sealed = storage.seal_for_worker(secret)
        if winsec.IS_WINDOWS:
            self.assertTrue(sealed.startswith("dpapi-job:"))
            blob = base64.b64decode(sealed[len("dpapi-job:"):])
            self.assertNotIn(secret, sealed)
            self.assertNotIn(secret.encode("utf-8"), blob)
            self.assertNotIn(secret.encode("utf-16-le"), blob)
        self.assertEqual(storage.open_from_worker(sealed), secret)
        self.assertEqual(storage.open_from_worker("legacy-plain"), "legacy-plain")


class RepositoryTests(StorageTestCase):
    def test_profiles_store_only_references_and_drop_replaced_secrets(self):
        old = storage.protect("old", "p")
        repo.save_profiles("Site", {"P": {"username": "u", "password": old, "secret": "", "protocol": "ssh", "port": 22}},
                           {"10.0.0.1": "P"})
        new = storage.protect("new", "p")
        repo.save_profiles("Site", {"P": {"username": "u", "password": new, "secret": "", "protocol": "ssh", "port": 22}},
                           {"10.0.0.1": "P"})
        profiles, assign = repo.load_profiles("Site")
        self.assertEqual(profiles["P"]["password"], new)
        self.assertEqual(assign, {"10.0.0.1": "P"})
        with self.assertRaises(storage.VaultError):
            storage.reveal(old)
        with self.assertRaises(ValueError):
            repo.save_profiles("Site", {"P": {"username": "u", "password": "plaintext", "protocol": "ssh", "port": 22}}, {})

    def test_deleting_a_site_removes_its_data_and_secrets(self):
        ref = storage.protect("pw", "p")
        repo.save_devices("Gone", [{"host": "10.0.0.2", "label": "sw", "type": "cisco-xe"}])
        repo.save_profiles("Gone", {"P": {"username": "u", "password": ref, "protocol": "ssh", "port": 22}}, {})
        repo.append_history("Gone", {"ip_address": "10.0.0.2", "status": "completed"})
        repo.delete_site("Gone")
        self.assertEqual(repo.load_devices("Gone"), [])
        self.assertEqual(repo.load_history("Gone"), [])
        with self.assertRaises(storage.VaultError):
            storage.reveal(ref)

    def test_rename_keeps_everything(self):
        repo.save_devices("A", [{"host": "1.1.1.1", "label": "x", "type": "cisco-xe", "platform": "C9300", "note": 1}])
        repo.set_site_doc("A", "topology", {"devices": {}, "roots": ["1.1.1.1"]})
        repo.rename_site("A", "B")
        self.assertEqual(repo.load_devices("B")[0]["note"], 1)
        self.assertEqual(repo.get_site_doc("B", "topology")["roots"], ["1.1.1.1"])

    def test_history_is_capped(self):
        for i in range(5):
            repo.append_history("H", {"n": i}, keep=3)
        self.assertEqual([r["n"] for r in repo.load_history("H")], [2, 3, 4])

    def test_setting_change_drops_the_old_secret(self):
        a = storage.protect("a", "smtp-password")
        repo.set_setting("email", "config", {"password_enc": a})
        b = storage.protect("b", "smtp-password")
        repo.set_setting("email", "config", {"password_enc": b})
        with self.assertRaises(storage.VaultError):
            storage.reveal(a)
        self.assertEqual(storage.reveal(repo.get_setting("email", "config")["password_enc"]), "b")

    def test_launcher_accounts_round_trip(self):
        settings = {"port": 443, "admin": {"name": "admin", "salt": "aa", "hash": "bb"},
                    "users": {"op": {"salt": "cc", "hash": "dd", "role": "operator"}}}
        repo.save_launcher_settings(settings)
        loaded = repo.load_launcher_settings()
        self.assertEqual(loaded["port"], 443)
        self.assertEqual(loaded["admin"]["name"], "admin")
        self.assertEqual(loaded["users"]["op"]["role"], "operator")


class MigrationTests(StorageTestCase):
    def test_site_files_are_imported_backed_up_redacted_and_removed(self):
        folder = Path(self.temp.name) / "sites" / "Legacy"
        folder.mkdir(parents=True)
        (folder / "devices.json").write_text(json.dumps([{"host": "10.1.1.1", "label": "sw1", "type": "cisco-xe"}]))
        (folder / "credential_profiles.json").write_text(json.dumps({
            "profiles": {"Core": {"username": "admin", "password": "dpapi:OLD", "secret": "", "protocol": "ssh", "port": 22}},
            "assignments": {"10.1.1.1": "Core"}}))
        (folder / "notifications.json").write_text(json.dumps([{"id": 1, "kind": "completed"}]))
        results = migrate.migrate_site("Legacy", folder, legacy_unprotect=lambda v: "DevicePass1")

        self.assertTrue(all(s == "imported" for _, s in results), results)
        self.assertFalse((folder / "devices.json").exists())
        self.assertFalse((folder / "credential_profiles.json").exists())
        profiles, assign = repo.load_profiles("Legacy")
        self.assertEqual(storage.reveal(profiles["Core"]["password"]), "DevicePass1")
        self.assertEqual(assign, {"10.1.1.1": "Core"})
        self.assertEqual(repo.get_site_doc("Legacy", "notes")[0]["kind"], "completed")
        backups = list((storage.data_dir() / "legacy-json-backup").rglob("credential_profiles.json"))
        self.assertEqual(len(backups), 1)
        text = backups[0].read_text(encoding="utf-8")
        self.assertNotIn("dpapi:OLD", text)
        self.assertIn(migrate.REDACTED, text)

    def test_existing_database_data_is_never_overwritten(self):
        repo.save_devices("Keep", [{"host": "2.2.2.2", "label": "new", "type": "cisco-xe"}])
        folder = Path(self.temp.name) / "Keep"
        folder.mkdir()
        (folder / "devices.json").write_text(json.dumps([{"host": "9.9.9.9", "label": "old", "type": "cisco-xe"}]))
        migrate.migrate_site("Keep", folder)
        self.assertEqual([d["host"] for d in repo.load_devices("Keep")], ["2.2.2.2"])

    def test_undecryptable_secret_keeps_profile_and_reports(self):
        folder = Path(self.temp.name) / "Other"
        folder.mkdir()
        (folder / "credential_profiles.json").write_text(json.dumps(
            {"profiles": {"P": {"username": "u", "password": "dpapi:FOREIGN", "protocol": "ssh"}}, "assignments": {}}))

        def fail(_value):
            raise OSError("not this user")
        logs = []
        results = migrate.migrate_site("Other", folder, legacy_unprotect=fail, log=logs.append)
        self.assertIn("re-entered", results[0][1])
        self.assertEqual(repo.load_profiles("Other")[0]["P"]["password"], "")
        self.assertTrue(any("could not be decrypted" in m for m in logs))

    def test_global_secret_settings_move_into_the_vault(self):
        prog = Path(self.temp.name) / "prog"
        prog.mkdir()
        (prog / "ai_settings.json").write_text(json.dumps({"enabled": True, "api_key_enc": "dpapi:K"}))
        (prog / "settings.json").write_text(json.dumps({"max_parallel": 3}))
        migrate.migrate_global(prog, legacy_unprotect=lambda v: "sk-test-key")
        cfg = repo.get_setting("ai", "config")
        self.assertTrue(cfg["api_key_enc"].startswith("vault:"))
        self.assertEqual(storage.reveal(cfg["api_key_enc"]), "sk-test-key")
        self.assertEqual(repo.get_setting("app", "settings")["max_parallel"], 3)
        self.assertFalse((prog / "ai_settings.json").exists())


if __name__ == "__main__":
    unittest.main()
