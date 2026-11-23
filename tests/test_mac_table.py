"""Pre & post checks: the MAC addresses learnt on each access port must be learnt again after the upgrade."""
import unittest

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager

MAC_TABLE = """          Mac Address Table
-------------------------------------------

Vlan    Mac Address       Type        Ports
----    -----------       --------    -----
 All    0100.0ccc.cccc    STATIC      CPU
   1    00a3.d111.1100    DYNAMIC     Gi1/0/48
  10    0011.2233.4401    DYNAMIC     Gi1/0/1
  10    0011.2233.4402    DYNAMIC     Gi1/0/1
  20    0011.2233.4403    DYNAMIC     Gi2/0/5
  30    0011.2233.4404    STATIC      Gi1/0/7
  10    0011.2233.4405    DYNAMIC     Po1
  10    aabb.ccdd.eeff    DYNAMIC     Te1/1/1
Total Mac Addresses for this criterion: 8
"""
TRUNK = """Port        Mode             Encapsulation  Status        Native vlan
Te1/1/1     on               802.1q         trunking      1
Gi1/0/48    on               802.1q         trunking      1
"""


def snapshot(mac_text, trunk_text=TRUNK):
    return manager.parse_snapshot({"mac_table": mac_text, "int_trunk": trunk_text})


class ParseTests(unittest.TestCase):
    def test_only_dynamic_macs_on_access_ports_are_kept(self):
        macs = manager.parse_mac_table(MAC_TABLE, ["Te1/1/1", "Gi1/0/48"])
        self.assertEqual(macs, {"Gi1/0/1": {"0011.2233.4401": "10", "0011.2233.4402": "10"},
                                "Gi2/0/5": {"0011.2233.4403": "20"}})

    def test_not_a_mac_table(self):
        self.assertIsNone(manager.parse_mac_table(""))
        self.assertIsNone(manager.parse_mac_table("% Invalid input detected at '^' marker."))
        self.assertEqual(manager.parse_mac_table("Mac Address Table\nTotal Mac Addresses for this criterion: 0"), {})

    def test_snapshot_leaves_out_trunks_and_counts_in_the_summary(self):
        snap = snapshot(MAC_TABLE)
        self.assertTrue(snap["ok"]["mac_table"])
        self.assertEqual(sorted(snap["macs"]), ["Gi1/0/1", "Gi2/0/5"])
        self.assertIn("3 MAC address(es) learnt on 2 access port(s)", manager.summarize_snapshot(snap))


class CompareTests(unittest.TestCase):
    def test_same_devices_back_is_clean(self):
        self.assertEqual(manager.compare_snapshots(snapshot(MAC_TABLE), snapshot(MAC_TABLE)), [])

    def test_missing_device_is_reported_by_port(self):
        after = MAC_TABLE.replace("  10    0011.2233.4402    DYNAMIC     Gi1/0/1\n", "")
        issues = manager.compare_snapshots(snapshot(MAC_TABLE), snapshot(after))
        self.assertEqual(issues, ["1 MAC address(es) learnt before the upgrade are not learnt again: Gi1/0/1 (1: 0011.2233.4402)"])

    def test_device_on_another_port_is_reported(self):
        after = MAC_TABLE.replace("0011.2233.4403    DYNAMIC     Gi2/0/5", "0011.2233.4403    DYNAMIC     Gi2/0/9")
        issues = manager.compare_snapshots(snapshot(MAC_TABLE), snapshot(after))
        self.assertEqual(issues, ["1 MAC address(es) now learnt on another port: 0011.2233.4403 Gi2/0/5 -> Gi2/0/9"])

    def test_many_missing_macs_are_summarised(self):
        before = {f"Gi1/0/{p}": {f"0011.2233.{p:02d}{m:02d}": "10" for m in range(5)} for p in range(1, 21)}
        issue = manager.compare_mac_tables(before, {})[0]
        self.assertTrue(issue.startswith("100 MAC address(es) learnt before the upgrade are not learnt again: "))
        self.assertIn("(+12 more)", issue)
        self.assertIn(", ...)", issue)

    def test_snapshot_without_mac_data_is_not_compared(self):
        old = snapshot(MAC_TABLE)
        old["ok"].pop("mac_table")                   # a pre-upgrade snapshot saved by an older version
        old.pop("macs")
        self.assertEqual(manager.compare_snapshots(old, snapshot("")), [])

    def test_diagnosis_comparison_table_has_a_mac_row(self):
        rows = {r[0]: r for r in manager._check_metrics(snapshot(MAC_TABLE), "17.9.4a", 5, 50)}
        self.assertEqual(rows["MAC addresses (access ports)"][2], 3)
        self.assertEqual(manager._metric_status("MAC addresses (access ports)", 3, 2), "Worse")


class SettingsTests(unittest.TestCase):
    MAC = "show mac address-table dynamic"

    def tearDown(self):
        manager.apply_command_selection({})

    def test_the_command_is_in_the_pre_and_post_check_lists(self):
        self.assertIn(self.MAC, [c for _k, c in manager.SNAPSHOT_COMMANDS])
        self.assertIn(self.MAC, [c for _k, c in manager.MANUAL_CHECK_COMMANDS])
        groups = {g["group"]: g["commands"] for g in manager.CHECK_COMMAND_GROUPS}
        self.assertEqual(groups["MAC address table"], [self.MAC])
        self.assertNotIn(self.MAC, groups["Stack"])

    def test_a_list_saved_before_the_command_existed_gets_it_switched_on(self):
        old = [c for c in manager.clean_settings({})["check_commands"] if c not in (self.MAC, "show clock")]
        values = manager.clean_settings({"check_commands": old})       # 2.3.0 / 2.3.1 save: no catalog
        self.assertIn(self.MAC, values["check_commands"])
        self.assertNotIn("show clock", values["check_commands"])         # what the user turned off stays off

    def test_a_list_saved_with_the_command_turned_off_keeps_it_off(self):
        saved = manager.clean_settings({"check_commands": ["show version", "show boot", "show switch"]}, strict=True)
        self.assertNotIn(self.MAC, saved["check_commands"])
        again = manager.clean_settings(saved)                            # loaded back (catalog saved with it)
        self.assertNotIn(self.MAC, again["check_commands"])
        manager.apply_command_selection({k: saved[k] for k in ("check_commands", "backup_commands", "commands_catalog")})
        self.assertFalse(manager.check_command_on(self.MAC))             # the upgrade worker's view (job file)


if __name__ == "__main__":
    unittest.main()
