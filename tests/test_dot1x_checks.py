"""Pre & post checks: every 802.1X / MAB device must be authenticated again after the upgrade."""
import unittest

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager

SID = "0A0B0C0D000000{:02d}1A2B3C4D"
BEFORE = """Interface                MAC Address    Method  Domain  Status Fg  Session ID
--------------------------------------------------------------------------------------------
Gi1/0/1                  0011.2233.4401 dot1x   DATA    Auth        {}
Gi1/0/2                  0011.2233.4402 mab     VOICE   Auth    D   {}
Gi1/0/2                  0011.2233.4403 dot1x   DATA    Auth        {}
Gi1/0/3                  0011.2233.4404 dot1x   DATA    Unauth      {}
Gi2/0/5                  0011.2233.4405 mab     DATA    Authz Success {}

Session count = 5
""".format(*(SID.format(i) for i in range(5)))


def snap(text):
    return manager.parse_snapshot({"auth_sessions": text})


class ParseTests(unittest.TestCase):
    def test_flag_letters_are_not_part_of_the_status(self):
        rows = manager.parse_auth_sessions(BEFORE)
        self.assertEqual([r["status"] for r in rows], ["Auth", "Auth", "Auth", "Unauth", "Authz Success"])
        self.assertEqual([r["ok"] for r in rows], [True, True, True, False, True])

    def test_snapshot_keeps_the_sessions_and_says_who_is_not_authenticated(self):
        s = snap(BEFORE)
        self.assertTrue(s["ok"]["auth_sessions"])
        self.assertEqual(len(s["auth"]), 5)
        self.assertEqual(manager.auth_status_text(s),
                         "802.1X/MAB: 4 of 5 session(s) authenticated (2 dot1x, 2 mab); "
                         "not authenticated: Gi1/0/3 0011.2233.4404 (dot1x, Unauth)")
        self.assertIn("802.1X/MAB: 4 of 5", manager.summarize_snapshot(s))

    def test_no_sessions_and_unsupported_command(self):
        self.assertEqual(manager.auth_status_text(snap("No sessions currently exist")), "802.1X/MAB: no authentication sessions")
        unsupported = snap("% Invalid input detected at '^' marker.")
        self.assertFalse(unsupported["ok"]["auth_sessions"])
        self.assertIsNone(manager.auth_status_text(unsupported))


class CompareTests(unittest.TestCase):
    def test_everyone_back_is_clean(self):
        self.assertEqual(manager.compare_snapshots(snap(BEFORE), snap(BEFORE)), [])

    def test_device_not_authenticated_after_the_upgrade(self):
        after = BEFORE.replace("0011.2233.4401 dot1x   DATA    Auth  ", "0011.2233.4401 dot1x   DATA    Unauth")
        self.assertEqual(manager.compare_snapshots(snap(BEFORE), snap(after)),
                         ["1 device(s) authenticated before the upgrade are NOT authenticated now: "
                          "Gi1/0/1 0011.2233.4401 (dot1x, Unauth)"])

    def test_device_without_a_session_after_the_upgrade(self):
        after = "\n".join(line for line in BEFORE.splitlines() if "0011.2233.4405" not in line)
        self.assertEqual(manager.compare_snapshots(snap(BEFORE), snap(after)),
                         ["1 device(s) authenticated before the upgrade have no authentication session now: "
                          "Gi2/0/5 0011.2233.4405"])

    def test_a_new_failing_session_is_reported_but_an_old_failure_is_not(self):
        after = BEFORE.replace("Session count = 5",
                               "Gi1/0/9                  0011.2233.4499 dot1x   DATA    Unauth      " + SID.format(9))
        issues = manager.compare_snapshots(snap(BEFORE), snap(after))
        self.assertEqual(issues, ["1 new session(s) not authenticated: Gi1/0/9 0011.2233.4499 (dot1x, Unauth)"])

    def test_post_check_detail_always_lists_the_not_authenticated_ports(self):
        conn_out = {"show authentication sessions": BEFORE.replace("0011.2233.4401 dot1x   DATA    Auth  ",
                                                                   "0011.2233.4401 dot1x   DATA    Unauth")}

        class Conn(object):
            def send_command(self, command, *a, **k):
                return conn_out.get(command, "")

        pre = snap(BEFORE)
        manager.apply_command_selection({})
        result, _outputs, _snap = manager.evaluate_postcheck(Conn(), None, pre)
        self.assertEqual(result["state"], "warning")
        self.assertIn("NOT authenticated now: Gi1/0/1 0011.2233.4401", result["detail"])
        self.assertIn("802.1X/MAB: 3 of 5 session(s) authenticated", result["detail"])
        self.assertIn("Gi1/0/3 0011.2233.4404 (dot1x, Unauth)", result["detail"])
        self.assertEqual(result["checks"]["auth"].split(";")[0], "802.1X/MAB: 3 of 5 session(s) authenticated (1 dot1x, 2 mab)")

    def test_diagnosis_comparison_rows(self):
        rows = {r[0]: r for r in manager._check_metrics(snap(BEFORE), "17.9.4a", 5, 50)}
        self.assertEqual(rows["Authenticated sessions"][2], 4)
        self.assertEqual(rows["Not authenticated sessions"][2], 1)
        self.assertEqual(manager._metric_status("Not authenticated sessions", 1, 3), "Worse")
        self.assertEqual(manager._metric_status("Authenticated sessions", 4, 3), "Worse")


class SettingsTests(unittest.TestCase):
    CMD = "show authentication sessions"

    def tearDown(self):
        manager.apply_command_selection({})

    def test_it_has_its_own_group_and_runs_in_the_checks(self):
        groups = {g["group"]: g["commands"] for g in manager.CHECK_COMMAND_GROUPS}
        self.assertEqual(groups["802.1X / MAB authentication"], [self.CMD])
        self.assertNotIn(self.CMD, groups["Stack"])
        self.assertIn(self.CMD, manager.selected_check_commands())

    def test_lists_saved_by_2_3_2_get_it_switched_on(self):
        catalog_2_3_2 = [c for c in manager.SETTINGS_DEFAULT_CATALOG if c != self.CMD]
        saved = [c for c in manager.clean_settings({})["check_commands"] if c not in (self.CMD, "show clock")]
        values = manager.clean_settings({"check_commands": saved, "commands_catalog": catalog_2_3_2})
        self.assertIn(self.CMD, values["check_commands"])
        self.assertNotIn("show clock", values["check_commands"])


if __name__ == "__main__":
    unittest.main()
