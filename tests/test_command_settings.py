"""Settings > Check & Backup Commands: selectable commands, at least 3 per list, and no command run twice."""
import tempfile
import unittest
from unittest.mock import patch

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager


class FakeConn(object):
    """Records every command sent; answers with an empty output (or a canned one)."""

    def __init__(self, answers=None):
        self.sent = []
        self.answers = answers or {}

    def send_command(self, command, *args, **kwargs):
        self.sent.append(command)
        return self.answers.get(command, "")


ITEMS = {i["key"]: i for i in manager.SETTINGS_ITEMS if i["type"] == "commands"}
CATALOG = {"commands_catalog": list(manager.SETTINGS_DEFAULT_CATALOG)}    # what the app saves with every list


def sel(**lists):
    """A saved selection as the app stores it (with the list of commands it offered)."""
    return dict(CATALOG, **lists)


class SelectionTests(unittest.TestCase):
    def tearDown(self):
        manager.apply_command_selection({})              # back to every command

    def test_defaults_select_every_command_once(self):
        values = manager.clean_settings({})
        self.assertEqual(len(values["check_commands"]), 23)
        self.assertEqual(len(set(values["check_commands"])), 23)
        self.assertIn("show mac address-table dynamic", values["check_commands"])
        self.assertIn("show authentication sessions", values["check_commands"])
        self.assertEqual(values["backup_commands"],
                         ["show running-config", "show vlan", "show version", "show boot", "show inventory"])

    def test_fewer_than_three_is_refused_on_save(self):
        with self.assertRaisesRegex(manager.ApiError, "Select at least 3"):
            manager.clean_settings({"check_commands": ["show version", "show boot"]}, strict=True)
        with self.assertRaisesRegex(manager.ApiError, "Select at least 3 backup commands"):
            manager.clean_settings({"backup_commands": ["show running-config"]}, strict=True)

    def test_locked_commands_are_always_kept_and_order_is_the_page_order(self):
        values = manager.clean_settings({"check_commands": ["show processes cpu", "show clock", "show vlan brief"],
                                         "backup_commands": ["show inventory", "show vlan"]}, strict=True)
        self.assertEqual(values["check_commands"], ["show version", "show clock", "show vlan brief", "show processes cpu"])
        self.assertEqual(values["backup_commands"], ["show running-config", "show vlan", "show inventory"])

    def test_unknown_command_is_refused_and_a_damaged_value_becomes_the_default(self):
        with self.assertRaisesRegex(manager.ApiError, "Unknown command"):
            manager.clean_settings({"check_commands": ["show version", "show boot", "reload"]}, strict=True)
        self.assertEqual(manager.clean_settings(sel(check_commands="show version"))["check_commands"],
                         ITEMS["check_commands"]["default"])
        self.assertEqual(manager.clean_settings(sel(check_commands=["show boot"]))["check_commands"],
                         ITEMS["check_commands"]["default"])

    def test_snapshot_and_health_follow_the_selection(self):
        manager.apply_command_selection(sel(check_commands=["show version", "show interfaces status", "show vlan brief",
                                                            "show processes cpu"]))
        conn = FakeConn()
        outputs, _snap = manager.read_snapshot(conn)
        manager.read_health(conn)
        self.assertEqual(conn.sent, ["show interfaces status", "show vlan brief", "show processes cpu"])
        self.assertEqual(sorted(k for k in outputs if not k.endswith("_error")), ["int_status", "vlan_brief"])
        self.assertNotIn("show ip interface brief", manager.snapshot_text(outputs, "x"))

    def test_diagnosis_checks_use_the_selection(self):
        manager.apply_command_selection(sel(check_commands=["show clock", "show module", "show switch"]))
        self.assertEqual(manager.selected_check_commands(), ["show clock", "show version", "show module", "show switch"])
        manager.apply_command_selection({})
        self.assertEqual(len(manager.selected_check_commands()), 23)


class NoDuplicateTests(unittest.TestCase):
    def test_once_session_runs_a_command_once(self):
        conn = FakeConn({"show switch": "1 Active"})
        once = manager.OnceSession(conn)
        self.assertEqual(once.send_command("show switch"), "1 Active")
        self.assertEqual(once.send_command("  SHOW   switch "), "1 Active")
        self.assertEqual(conn.sent, ["show switch"])

    def test_post_check_sends_every_command_only_once(self):
        conn = FakeConn()
        manager.evaluate_postcheck(conn, None, None)
        self.assertEqual(len(conn.sent), len(set(conn.sent)), conn.sent)
        self.assertIn("show switch", conn.sent)

    def test_additional_backup_commands_are_deduplicated(self):
        self.assertEqual(manager.Manager._parse_custom_commands("show clock\nSHOW  CLOCK\nshow logging\n\nshow clock"),
                         ["show clock", "show logging"])

    def test_backup_skips_turned_off_and_repeated_commands(self):
        conn = FakeConn({"show running-config": "hostname sw1", "show version": "Cisco IOS XE Software, Version 17.09.04a"})
        conn.enable = lambda: None
        conn.find_prompt = lambda: "sw1#"
        conn.__enter__ = lambda: conn
        conn.__exit__ = lambda *a: False

        class Ctx(object):
            def __enter__(self):
                return conn

            def __exit__(self, *exc):
                return False

        manager.apply_command_selection(sel(backup_commands=["show running-config", "show version", "show inventory"]))
        self.addCleanup(manager.apply_command_selection, {})
        with tempfile.TemporaryDirectory() as tmp, patch.object(manager, "ConnectHandler", lambda **kw: Ctx()):
            manager.do_backup("10.0.0.1", "cisco_xe", "u", "p", "", tmp, "20261123_120000",
                              ["show version", "show logging", "show running-config"])
        self.assertEqual(conn.sent, ["show running-config", "show version", "show inventory", "show logging"])


if __name__ == "__main__":
    unittest.main()
