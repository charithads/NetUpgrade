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
    return obj


class ParallelSettingsTests(unittest.TestCase):
    def test_parallelism_and_stop_policy_are_persisted_per_site(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = Path(temp) / "upgrade_settings.json"
            obj = bare_manager()
            with patch.object(manager, "site_upgrade_settings_path", return_value=settings):
                result = obj.set_upgrade_settings("4", False, who="admin")
                self.assertEqual(result["max_parallel"], 4)
                self.assertFalse(result["stop_on_failure"])

                loaded = bare_manager()
                loaded._load_upgrade_settings("test-site")

            self.assertEqual(loaded.upgrade_settings_view(), {
                "max_parallel": 4, "stop_on_failure": False, "limit": manager.MAX_PARALLEL_LIMIT,
            })

    def test_parallelism_rejects_non_integer_and_out_of_range_values(self):
        obj = bare_manager()
        for value in (True, "invalid", 0, manager.MAX_PARALLEL_LIMIT + 1):
            with self.subTest(value=value), self.assertRaises(manager.ApiError):
                obj.set_upgrade_settings(value, True)

    def test_only_administrator_can_change_rollout_settings(self):
        allowed, message = manager.role_allows(
            manager.ROLE_OPERATOR, "POST", "/api/upgrade-settings")
        self.assertFalse(allowed)
        self.assertIn("upgrade settings", message)
        self.assertEqual(manager.role_allows(
            manager.ROLE_ADMIN, "POST", "/api/upgrade-settings"), (True, ""))


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

        with patch.object(obj, "_validate_hosts", return_value=hosts), \
                patch.object(obj, "_creds_map", return_value={h: ("user", "secret") for h in hosts}), \
                patch.object(obj, "_check_image", return_value=("C:/images/image.bin", 100)), \
                patch.object(obj, "_poll_all"), patch.object(obj, "_check_upgradable"), \
                patch.object(obj, "_require_platform_known"), patch.object(obj, "_enforce_image_check"), \
                patch.object(obj, "_active_reload_count", return_value=0), \
                patch.object(obj, "_save_schedule"), \
                patch.object(obj, "_launch_job", side_effect=lambda host, *args: started.append(host)):
            result = obj.start_upgrade(hosts, "image.bin", started_by="admin")

        self.assertEqual(started, hosts[:2])
        self.assertEqual(set(obj.schedule), {hosts[2]})
        self.assertTrue(obj.schedule[hosts[2]]["queued"])
        self.assertEqual(obj.schedule[hosts[2]]["rollout"], obj.rollout_of[hosts[2]])
        self.assertIn("1 more switch(es) are queued", result["message"])


if __name__ == "__main__":
    unittest.main()
