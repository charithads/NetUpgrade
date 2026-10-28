"""Regression tests for configurable concurrent switch upgrades."""
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import ios_xe_upgrade_manager as manager


class EventSink:
    def __init__(self):
        self.entries = []

    def add(self, *entry):
        self.entries.append(entry)


def bare_manager(site="test-site"):
    obj = object.__new__(manager.Manager)
    obj.site_name = site
    obj.lock = threading.RLock()
    obj.events = EventSink()
    obj.max_parallel = 1
    obj.stop_on_failure = True
    obj.settings = manager.settings_defaults()
    return obj


class ParallelSettingsTests(unittest.TestCase):
    def test_parallelism_and_stop_policy_are_persisted(self):
        with tempfile.TemporaryDirectory() as temp:
            settings_file = Path(temp) / "settings.json"
            obj = bare_manager()
            with patch.object(manager, "settings_path", return_value=settings_file):
                result = obj.update_settings(
                    {"max_parallel": "4", "stop_on_failure": False}, who="admin"
                )
                self.assertEqual(result["values"]["max_parallel"], 4)
                self.assertFalse(result["values"]["stop_on_failure"])

                loaded = bare_manager()
                loaded.settings = manager.load_global_settings()
                loaded._apply_settings()

            self.assertEqual(loaded.upgrade_settings_view(), {
                "max_parallel": 4,
                "stop_on_failure": False,
                "limit": manager.MAX_PARALLEL_LIMIT,
            })

    def test_parallelism_rejects_non_integer_and_out_of_range_values(self):
        obj = bare_manager()
        for value in (True, "invalid", 0, manager.MAX_PARALLEL_LIMIT + 1):
            with self.subTest(value=value), self.assertRaises(manager.ApiError):
                obj.update_settings({"max_parallel": value, "stop_on_failure": True})

    def test_only_administrator_can_change_rollout_settings(self):
        allowed, message = manager.role_allows(
            manager.ROLE_OPERATOR, "POST", "/api/settings")
        self.assertFalse(allowed)
        self.assertIn("change the settings", message)
        self.assertEqual(
            manager.role_allows(manager.ROLE_ADMIN, "POST", "/api/settings"),
            (True, ""),
        )


class ParallelQueueTests(unittest.TestCase):
    def test_batch_starts_configured_parallel_count_and_queues_remaining_devices(self):
        obj = bare_manager()
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        hosts = ["192.0.2.1", "192.0.2.2", "192.0.2.3"]
        obj.monitor = type("MonitorStub", (), {"running": True})()
        obj.max_parallel = 2
        obj.stop_on_failure = True
        obj.order = hosts
        obj.ops = {h: {"backup": {"state": "idle"},
                       "upgrade": {"state": "idle", "action": None}} for h in hosts}
        obj.schedule = {}
        obj.rollouts = {}
        obj.rollout_of = {}
        obj.runs_dir = Path(temp.name) / "Runs"
        started = []
        backups_started = []

        def start_backup(run_hosts):
            backups_started.extend(run_hosts)
            return {"message": "Backup started on all selected switches."}

        with patch.object(obj, "_validate_hosts", return_value=hosts), \
                patch.object(obj, "_creds_map", return_value={h: ("user", "secret") for h in hosts}), \
                patch.object(obj, "_check_image", return_value=("C:/images/image.bin", 100)), \
                patch.object(obj, "_poll_all"), patch.object(obj, "_check_upgradable"), \
                patch.object(obj, "_require_platform_known"), patch.object(obj, "_enforce_image_check"), \
                patch.object(obj, "_active_reload_count", return_value=0), \
                patch.object(obj, "_save_schedule"), \
                patch.object(obj, "start_backup", side_effect=start_backup), \
                patch.object(manager.threading, "Thread") as background_thread:
            result = obj.start_upgrade(hosts, "image.bin", started_by="admin")

        self.assertEqual(backups_started, hosts[:2])
        self.assertEqual(started, [])  # upgrade jobs wait for each backup to finish
        self.assertTrue(all(obj.ops[h]["upgrade"]["phase"] == "waiting_backup" for h in hosts[:2]))
        background_thread.return_value.start.assert_called_once()
        self.assertEqual(set(obj.schedule), {hosts[2]})
        self.assertTrue(obj.schedule[hosts[2]]["queued"])
        self.assertEqual(obj.schedule[hosts[2]]["rollout"], obj.rollout_of[hosts[2]])
        self.assertIn("1 more switch(es) are queued", result["message"])


if __name__ == "__main__":
    unittest.main()
