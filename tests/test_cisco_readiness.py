"""Cisco pre-upgrade readiness checks (Catalyst 9000 upgrade guide command reference), shared by the
Readiness Check and the Diagnosis Pre & Post Checks."""
import unittest

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as m

SHOW_VERSION = """Cisco IOS XE Software, Version 17.09.04a
ROM: IOS-XE ROMMON
BOOTLDR: System Bootstrap, Version 17.9.1r[FC2], RELEASE SOFTWARE (P)
Installation mode is INSTALL

Switch Ports Model              SW Version        SW Image              Mode
------ ----- -----              ----------        ----------            ----
*    1 52    C9300-48P          17.09.04a         CAT9K_IOSXE           INSTALL
     2 52    C9300-48P          17.09.04a         CAT9K_IOSXE           INSTALL
"""

SHOW_BOOT = """---------------------------
Switch 1
---------------------------
Current Boot Variables:
BOOT variable = flash:packages.conf;

Boot Variables on next reload:
BOOT variable = flash:packages.conf;
Manual Boot = no
---------------------------
Switch 2
---------------------------
Current Boot Variables:
BOOT variable = flash:packages.conf;

Boot Variables on next reload:
BOOT variable = flash:packages.conf;
Manual Boot = no
"""

INSTALL_OK = """[ Switch 1 2 ] Installed Package(s) Information:
State (St): I - Inactive, U - Activated & Uncommitted,
            C - Activated & Committed, D - Deactivated & Uncommitted
--------------------------------------------------------------------------------
Type  St   Filename/Version
--------------------------------------------------------------------------------
IMG   I    17.06.05.0.6
IMG   C    17.09.04a.0.6

--------------------------------------------------------------------------------
Auto abort timer: inactive
--------------------------------------------------------------------------------
"""

REDUNDANCY_OK = """Redundant System Information :
------------------------------
       Available system uptime = 3 weeks, 2 days
Switchovers system experienced = 0
              Standby failures = 0
        Last switchover reason = none

                 Hardware Mode = Duplex
    Configured Redundancy Mode = sso
     Operating Redundancy Mode = sso
              Maintenance Mode = Disabled
                Communications = Up

Current Processor Information :
-------------------------------
               Active Location = slot 1
        Current Software state = ACTIVE

Peer Processor Information :
----------------------------
              Standby Location = slot 2
        Current Software state = STANDBY HOT
"""

MODULE_OK = """Switch  Ports    Model                Serial No.   MAC address     Hw Ver.       Sw Ver.
------  -----   ---------             -----------  --------------  -------       --------
 1       52     C9300-48P            FOC1111X1AA  aaaa.bbbb.0001  V02           17.09.04a
 2       52     C9300-48P            FOC2222X2BB  aaaa.bbbb.0002  V02           17.09.04a
"""

FILE_SYSTEMS_OK = """File Systems:

       Size(b)       Free(b)      Type  Flags  Prefixes
*  11250098176    8055596544      disk     rw   flash:
   11250098176    8100000000      disk     rw   flash-2:
             -             -    opaque     rw   system:
"""

PACKAGES_OK = """boot  rp 0 0   rp_boot       cat9k-rpboot.17.09.04a.SPA.pkg
iso   rp 0 0   rp_base       cat9k-rpbase.17.09.04a.SPA.pkg
iso   rp 0 0   guestshell    cat9k-guestshell.17.09.04a.SPA.pkg
"""


def outputs(**over):
    base = {"version": SHOW_VERSION, "boot": SHOW_BOOT, "install": INSTALL_OK, "redundancy": REDUNDANCY_OK,
            "module": MODULE_OK, "file_systems": FILE_SYSTEMS_OK, "packages_conf": PACKAGES_OK}
    base.update(over)
    return base


class CiscoReadinessTests(unittest.TestCase):
    def test_healthy_stack_has_no_problems(self):
        f = m.cisco_readiness_findings(outputs(), include_boot=True)
        self.assertEqual(f["problems"], [])
        self.assertEqual(f["warnings"], [])
        s = f["summary"]
        self.assertEqual(s["Stack software"], "2 member(s), all 17.09.04a INSTALL")
        self.assertEqual(s["Install state"], "committed 17.09.04a.0.6")
        self.assertIn("sso", s["Redundancy"])
        self.assertEqual(s["Modules"], "no faults reported")
        self.assertIn("writable", s["File systems"])
        self.assertIn("release 17.09.04a", s["packages.conf"])
        self.assertIn("17.9.1r", s["ROMMON"])
        self.assertTrue(any("(rollback points): 17.06.05.0.6" in n for n in f["notes"]))

    def test_uncommitted_install_and_abort_timer_are_problems(self):
        pending = INSTALL_OK.replace("IMG   C    17.09.04a.0.6", "IMG   U    17.12.04.0.1").replace(
            "Auto abort timer: inactive", "Auto abort timer: active on install_activate, time before rollback - 01:50:12")
        f = m.cisco_readiness_findings(outputs(install=pending))
        self.assertTrue(any("not finished" in p and "17.12.04.0.1" in p for p in f["problems"]))
        self.assertTrue(any("auto-abort timer" in p for p in f["problems"]))
        self.assertEqual(f["summary"]["Install state"], "NOT committed")

    def test_stack_version_and_mode_mismatch(self):
        bad = SHOW_VERSION.replace("     2 52    C9300-48P          17.09.04a         CAT9K_IOSXE           INSTALL",
                                   "     2 52    C9300-48P          17.06.05          CAT9K_IOSXE           BUNDLE")
        f = m.cisco_readiness_findings(outputs(version=bad))
        self.assertTrue(any("different software versions" in p for p in f["problems"]))
        self.assertTrue(any("different modes" in p for p in f["problems"]))

    def test_standby_not_hot_and_not_sso(self):
        red = REDUNDANCY_OK.replace("Operating Redundancy Mode = sso", "Operating Redundancy Mode = rpr").replace(
            "Current Software state = STANDBY HOT", "Current Software state = STANDBY COLD")
        f = m.cisco_readiness_findings(outputs(redundancy=red))
        self.assertTrue(any("not SSO" in w for w in f["warnings"]))
        self.assertTrue(any("STANDBY COLD" in w for w in f["warnings"]))

    def test_standalone_switch_is_fine(self):
        standalone = ("Hardware Mode = Simplex\nConfigured Redundancy Mode = Non-redundant\n"
                      "Operating Redundancy Mode = Non-redundant\nPeer (slot: unknown) state = DISABLED\n")
        version = SHOW_VERSION.replace("     2 52    C9300-48P          17.09.04a         CAT9K_IOSXE           INSTALL\n", "")
        f = m.cisco_readiness_findings(outputs(version=version, redundancy=standalone))
        self.assertEqual(f["warnings"], [])
        self.assertEqual(f["summary"]["Redundancy"], "standalone (no standby)")

    def test_faulty_module_read_only_flash_and_missing_packages_conf(self):
        mod = MODULE_OK.replace("17.09.04a\n", "17.09.04a  faulty\n", 1)
        fs = FILE_SYSTEMS_OK.replace("disk     rw   flash:", "disk     ro   flash:")
        f = m.cisco_readiness_findings(outputs(module=mod, file_systems=fs,
                                               packages_conf="%Error opening flash:packages.conf (No such file or directory)"))
        self.assertTrue(any("module(s) not healthy" in w for w in f["warnings"]))
        self.assertTrue(any("not writable: flash:" in p for p in f["problems"]))
        self.assertTrue(any("packages.conf is missing" in p for p in f["problems"]))

    def test_packages_conf_pointing_to_another_release(self):
        f = m.cisco_readiness_findings(outputs(packages_conf=PACKAGES_OK.replace("17.09.04a", "17.12.04")))
        self.assertTrue(any("points to 17.12.04 but 17.09.04a is running" in w for w in f["warnings"]))

    def test_bundle_mode_does_not_need_packages_conf(self):
        version = SHOW_VERSION.replace("INSTALL", "BUNDLE")
        f = m.cisco_readiness_findings(outputs(version=version, packages_conf=""))
        self.assertFalse(any("packages.conf" in p for p in f["problems"]))
        self.assertEqual(f["summary"]["packages.conf"], "not used (BUNDLE mode)")

    def test_unavailable_commands_never_fail(self):
        f = m.cisco_readiness_findings({"version": SHOW_VERSION, "install": "", "redundancy": "% Invalid input detected",
                                        "module": "", "file_systems": "", "packages_conf": PACKAGES_OK})
        self.assertEqual(f["problems"], [])
        self.assertTrue(any("show redundancy" in n for n in f["notes"]))

    def test_readiness_verdict_uses_cisco_problems(self):
        checks = {"reachable": True, "login": True, "space_ok": True, "space_level": "ok", "space_detail": "ok",
                  "boot": None, "cpu_pct": 5, "mem_free_pct": 50, "stack": "2 member(s), all Ready", "version": "17.09.04a",
                  "cisco": {"problems": ["an earlier install is not finished"], "warnings": [], "notes": [], "summary": {}}}
        state, detail = m.compute_readiness_verdict(checks)
        self.assertEqual(state, "not_ready")
        self.assertIn("not finished", detail)

    def test_diagnosis_precheck_runs_and_reports_the_cisco_commands(self):
        commands = [c for _k, c in m.MANUAL_CHECK_COMMANDS]
        for cmd in ("show install summary", "show redundancy", "show module", "show file systems",
                    "more flash:packages.conf | include pkg"):
            self.assertIn(cmd, commands)
        pending = INSTALL_OK.replace("IMG   C    17.09.04a.0.6", "IMG   U    17.12.04.0.1")
        a = m.analyze_manual_check(outputs(install=pending), "pre")
        self.assertEqual(a["state"], "attention")
        self.assertTrue(any("not finished" in p for p in a["problems"]))
        names = [r["check"] for r in a["rows"]]
        for n in ("Install state", "Redundancy", "File systems", "packages.conf", "Stack software"):
            self.assertIn(n, names)

    def test_postcheck_compares_cisco_rows(self):
        post = outputs(version=SHOW_VERSION.replace("17.09.04a", "17.12.04"),
                       packages_conf=PACKAGES_OK.replace("17.09.04a", "17.12.04"),
                       install=INSTALL_OK.replace("17.09.04a.0.6", "17.12.04.0.1"))
        a = m.analyze_manual_check(post, "post", outputs(), "17.12.04")
        rows = {r["check"]: r for r in a["rows"]}
        self.assertEqual(rows["Install state"]["status"], "Changed")
        self.assertEqual(rows["Redundancy"]["status"], "Same")
        self.assertEqual(a["problems"], [])


class ReadinessOverrideTests(unittest.TestCase):
    def manager(self, readiness):
        import threading
        obj = object.__new__(m.Manager)
        obj.lock = threading.RLock()
        obj.info = {"10.0.0.1": {"host": "10.0.0.1", "label": "sw1", "type": "cisco-xe"}}
        obj.ops = {"10.0.0.1": {"readiness": readiness}}
        obj.entries = []
        obj.events = type("Sink", (), {"add": lambda _s, *e: obj.entries.append(e)})()
        return obj

    def blocked(self):
        return {"state": "not_ready", "detail": "x",
                "checks": {"cisco": {"problems": ["an earlier install is not finished"], "warnings": [], "notes": []}}}

    def test_block_refuses_without_override(self):
        obj = self.manager(self.blocked())
        with self.assertRaisesRegex(m.ApiError, "Ignore readiness check blocks"):
            obj._enforce_readiness(["10.0.0.1"], False, "admin")

    def test_override_allows_and_is_logged(self):
        obj = self.manager(self.blocked())
        obj._enforce_readiness(["10.0.0.1"], True, "admin")
        self.assertTrue(any("ignored by admin" in e[2] and e[3] == "warn" for e in obj.entries))

    def test_one_option_replaces_the_other_two(self):
        html = m.DASHBOARD_HTML
        self.assertIn("Ignore readiness check blocks", html)
        for gone in ('id="skip-image-check"', 'id="ignore-space"', "Override readiness"):
            self.assertNotIn(gone, html)

    def test_ignore_readiness_also_ignores_image_and_flash_checks(self):
        from unittest.mock import patch
        obj = self.manager(self.blocked())
        seen = {}

        def capture(hosts, image, skip, user):
            seen["skip_image_check"] = skip
            raise m.ApiError("stop here")              # nothing else of the upgrade needs to run
        obj.monitor = type("Mon", (), {"running": True})()
        obj.site_name, obj.max_parallel, obj.schedule = "s", 1, {}
        with patch.object(obj, "_validate_hosts", return_value=["10.0.0.1"]), \
                patch.object(obj, "_creds_map", return_value={}), \
                patch.object(obj, "_check_image", return_value=("C:/img.bin", 1)), \
                patch.object(obj, "_poll_all"), patch.object(obj, "_check_upgradable"), \
                patch.object(obj, "_active_reload_count", return_value=0), \
                patch.object(obj, "_enforce_image_check", side_effect=capture):
            with self.assertRaisesRegex(m.ApiError, "stop here"):
                obj.start_upgrade(["10.0.0.1"], "img.bin", ignore_readiness=True)
        self.assertTrue(seen["skip_image_check"])

    def test_no_block_when_ready_warning_or_not_run(self):
        for rd in ({}, {"state": "ready"}, {"state": "warning", "checks": {"cisco": {"problems": []}}},
                   {"state": "not_ready", "checks": {}}):          # e.g. could not log in: not a Cisco block
            with self.subTest(rd=rd):
                obj = self.manager(rd)
                obj._enforce_readiness(["10.0.0.1"], False, "admin")
                self.assertEqual(obj.readiness_blocks("10.0.0.1"), [])

    def test_help_and_override_are_in_the_dashboard(self):
        html = m.DASHBOARD_HTML
        for marker in ('id="btn-help"', 'id="help-modal"', 'id="help-pane-upgrade"', 'id="ignore-ready"',
                       "install add file flash:&lt;image&gt; activate commit", "verify /md5 flash:&lt;image&gt; &lt;md5&gt;"):
            self.assertIn(marker, html)


if __name__ == "__main__":
    unittest.main()
