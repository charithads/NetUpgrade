"""Device drawer: every member of a stack is shown, not only the active switch."""
import tempfile
import unittest
from pathlib import Path

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager

SHOW_VERSION = """Cisco IOS XE Software, Version 17.09.04a
Cisco IOS Software [Cupertino], Catalyst L3 Switch Software (CAT9K_IOSXE), Version 17.9.4a, RELEASE SOFTWARE (fc3)
ROM: IOS-XE ROMMON
BOOTLDR: System Bootstrap, Version 17.9.1r, RELEASE SOFTWARE (P)

core-sw uptime is 3 weeks, 1 day, 4 hours, 12 minutes
Uptime for this control processor is 3 weeks, 1 day, 4 hours, 14 minutes
System returned to ROM by Reload Command
System image file is "flash:packages.conf"
Last reload reason: Reload Command

cisco C9300-48P (X86) processor with 1331521K/6147K bytes of memory.
Processor board ID FOC2233A1AA
Base Ethernet MAC Address          : 00:a3:d1:11:11:00
Motherboard Assembly Number        : 73-17958-06
Motherboard Serial Number          : FOC2233A1AA
Model Revision Number              : B0
Motherboard Revision Number        : C0
Model Number                       : C9300-48P
System Serial Number               : FCW2233L0AA
Top Assembly Part Number           : 68-100829-01
Top Assembly Revision Number       : B0
Version ID                         : V02
CLEI Code Number                   : INM1J10ARA


Switch Ports Model              SW Version        SW Image              Mode
------ ----- -----              ----------        ----------            ----
     1 62    C9300-48U          17.09.04a         CAT9K_IOSXE           INSTALL
*    2 62    C9300-48P          17.09.04a         CAT9K_IOSXE           INSTALL
     3 30    C9300-24T          17.09.04a         CAT9K_IOSXE           INSTALL


Switch 01
---------
Switch uptime                      : 3 weeks, 1 day, 4 hours, 11 minutes
Base Ethernet MAC Address          : 00:a3:d1:22:22:00
Motherboard Assembly Number        : 73-17958-06
Motherboard Serial Number          : FOC2233B2BB
Model Revision Number              : B0
Motherboard Revision Number        : C0
Model Number                       : C9300-48U
System Serial Number               : FCW2233L0BB
Top Assembly Part Number           : 68-100830-01
Version ID                         : V01

Switch 03
---------
Switch uptime                      : 2 days, 6 hours, 2 minutes
Base Ethernet MAC Address          : 00:a3:d1:33:33:00
Model Number                       : C9300-24T
System Serial Number               : FCW2233L0CC
Version ID                         : V03

Configuration register is 0x102
"""

SHOW_INVENTORY = """NAME: "c93xx Stack", DESCR: "c93xx Stack"
PID: C9300-48P         , VID: V02  , SN: FCW2233L0AA

NAME: "Switch 1", DESCR: "C9300-48U"
PID: C9300-48U         , VID: V01  , SN: FCW2233L0BB

NAME: "Switch 1 - Power Supply A", DESCR: "Switch 1 - Power Supply A"
PID: PWR-C1-1100WAC    , VID: V02  , SN: DCB2233P1AA

NAME: "Switch 2", DESCR: "C9300-48P"
PID: C9300-48P         , VID: V02  , SN: FCW2233L0AA

NAME: "Switch 2 - Power Supply A", DESCR: "Switch 2 - Power Supply A"
PID: PWR-C1-715WAC     , VID: V01  , SN: DCB2233P2AA

NAME: "Switch 2 FRU Uplink Module 1", DESCR: "8x10G Uplink Module"
PID: C9300-NM-8X       , VID: V02  , SN: FOC2233N2AA

NAME: "Switch 3", DESCR: "C9300-24T"
PID: C9300-24T         , VID: V03  , SN: FCW2233L0CC

NAME: "Switch 12", DESCR: "not a real member - must not be matched as switch 1"
PID: C9300-24T         , VID: V03  , SN: FCW2233L0ZZ
"""

SHOW_SWITCH = """Switch/Stack Mac Address : 00a3.d111.1100 - Local Mac Address
Mac persistency wait time: Indefinite
                                             H/W   Current
Switch#   Role    Mac Address     Priority Version  State
-------------------------------------------------------------------------------------
 1       Standby  00a3.d122.2200     14     V01     Ready
*2       Active   00a3.d111.1100     15     V02     Ready
 3       Member   00a3.d133.3300     1      V03     Ready
"""


def backup_file(folder, commands):
    text = manager.command_sections_text("Configuration backup - core-sw", commands)
    (Path(folder) / f"{manager.BACKUP_FILE_PREFIX}_20261123_090000.txt").write_text(text, encoding="utf-8")


class StackMemberTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        manager._FACTS_CACHE.clear()

    def facts(self, commands):
        backup_file(self.temp.name, commands)
        return manager.read_device_facts(self.temp.name)

    def test_every_member_gets_its_own_details(self):
        f = self.facts([("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY)])
        members = {m["switch"]: m for m in f["members"]}
        self.assertEqual(sorted(members), [1, 2, 3])
        self.assertEqual(members[1]["serial"], "FCW2233L0BB")
        self.assertEqual(members[1]["hardware_version"], "V01")
        self.assertEqual(members[1]["pid"], "C9300-48U")
        self.assertEqual(members[1]["mac"], "00:a3:d1:22:22:00")
        self.assertEqual(members[1]["uptime"], "3 weeks, 1 day, 4 hours, 11 minutes")
        self.assertEqual(members[1]["parts"], ["Power Supply A: PWR-C1-1100WAC (SN DCB2233P1AA)"])
        self.assertEqual(members[3]["serial"], "FCW2233L0CC")
        self.assertEqual(members[3]["uptime"], "2 days, 6 hours, 2 minutes")
        self.assertEqual(members[3]["ports"], 30)

    def test_active_member_uses_the_values_before_the_table(self):
        f = self.facts([("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY)])
        active = next(m for m in f["members"] if m["active"])
        self.assertEqual(active["switch"], 2)
        self.assertEqual(active["serial"], "FCW2233L0AA")
        self.assertEqual(active["mac"], "00:a3:d1:11:11:00")
        self.assertEqual(active["uptime"], "3 weeks, 1 day, 4 hours, 12 minutes")
        self.assertEqual(active["parts"], ["Power Supply A: PWR-C1-715WAC (SN DCB2233P2AA)",
                                           "FRU Uplink Module 1: C9300-NM-8X (SN FOC2233N2AA)"])
        # the device-level serial / hardware version are the ACTIVE switch's, not switch 1's
        self.assertEqual(f["serial"], "FCW2233L0AA")
        self.assertEqual(f["hardware_version"], "V02")

    def test_show_version_alone_still_gives_every_member(self):
        f = self.facts([("show version", SHOW_VERSION)])
        members = {m["switch"]: m for m in f["members"]}
        self.assertEqual(members[1]["serial"], "FCW2233L0BB")
        self.assertEqual(members[1]["hardware_version"], "V01")
        self.assertEqual(members[3]["mac"], "00:a3:d1:33:33:00")
        self.assertEqual(members[2]["serial"], "FCW2233L0AA")

    def test_roles_come_from_show_switch_when_it_was_backed_up(self):
        f = self.facts([("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY), ("show switch", SHOW_SWITCH)])
        roles = {m["switch"]: (m.get("role"), m.get("state")) for m in f["members"]}
        self.assertEqual(roles, {1: ("Standby", "Ready"), 2: ("Active", "Ready"), 3: ("Member", "Ready")})
        self.assertNotIn("switch_roles", f)

    def test_a_standalone_switch_is_unchanged(self):
        single = SHOW_VERSION.split("Switch 01")[0].replace(
            "     1 62    C9300-48U          17.09.04a         CAT9K_IOSXE           INSTALL\n", "").replace(
            "     3 30    C9300-24T          17.09.04a         CAT9K_IOSXE           INSTALL\n", "")
        f = self.facts([("show version", single)])
        self.assertEqual(len(f["members"]), 1)
        self.assertEqual(f["members"][0]["serial"], "FCW2233L0AA")
        self.assertEqual(f["serial"], "FCW2233L0AA")


class RefreshVersionTests(unittest.TestCase):
    """Refresh Version keeps what it read; the drawer uses it when it is newer than the last backup."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.backups = Path(self.temp.name) / "backups"
        self.runs = Path(self.temp.name) / "runs"
        self.backups.mkdir()
        manager._FACTS_CACHE.clear()

    def save_refresh(self, sections):
        class Owner(object):
            def label(self, host):
                return "core-sw"
        manager.Manager._save_refresh_facts(Owner(), self.runs, "10.0.0.1", sections)
        return manager.refresh_facts_path(self.runs, "10.0.0.1")

    def set_age(self, path, seconds_ago):
        import os
        import time
        t = time.time() - seconds_ago
        os.utime(path, (t, t))

    def test_refresh_newer_than_the_backup_is_shown(self):
        old = SHOW_VERSION.replace("FCW2233L0BB", "OLDSERIAL01")      # switch 1 was replaced since the backup
        backup_file(self.backups, [("show version", old), ("show inventory", SHOW_INVENTORY.replace("FCW2233L0BB", "OLDSERIAL01"))])
        path = self.save_refresh([("show version", SHOW_VERSION), ("show switch", SHOW_SWITCH), ("show inventory", SHOW_INVENTORY)])
        f = manager.read_device_facts(self.backups, path)
        self.assertEqual(f["read_from"]["kind"], "refresh")
        members = {m["switch"]: m for m in f["members"]}
        self.assertEqual(members[1]["serial"], "FCW2233L0BB")
        self.assertEqual(members[1]["role"], "Standby")

    def test_backup_newer_than_the_refresh_is_shown(self):
        path = self.save_refresh([("show version", SHOW_VERSION.replace("FCW2233L0BB", "OLDSERIAL01"))])
        self.set_age(path, 3 * 365 * 86400)                       # long before the backup's stamp
        backup_file(self.backups, [("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY)])
        f = manager.read_device_facts(self.backups, path)
        self.assertEqual(f["read_from"]["kind"], "backup")
        self.assertEqual({m["switch"]: m["serial"] for m in f["members"]}[1], "FCW2233L0BB")

    def test_a_command_missing_from_the_refresh_comes_from_the_backup(self):
        backup_file(self.backups, [("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY)])
        path = self.save_refresh([("show version", SHOW_VERSION)])   # show inventory failed during the refresh
        f = manager.read_device_facts(self.backups, path)
        self.assertEqual(f["read_from"]["kind"], "refresh")
        members = {m["switch"]: m for m in f["members"]}
        self.assertEqual(members[2]["parts"][0], "Power Supply A: PWR-C1-715WAC (SN DCB2233P2AA)")

    def test_refresh_alone_works_without_any_backup(self):
        path = self.save_refresh([("show version", SHOW_VERSION), ("show inventory", SHOW_INVENTORY)])
        f = manager.read_device_facts(self.temp.name + "/no-such-folder", path)
        self.assertEqual(len(f["members"]), 3)
        self.assertEqual(f["serial"], "FCW2233L0AA")
        self.assertEqual(manager.read_device_facts(self.temp.name + "/no-such-folder", None), {})


if __name__ == "__main__":
    unittest.main()
