"""Settings > About shows only stable releases, newest first, starting with the running version."""
import re
import unittest

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager


class ChangelogTests(unittest.TestCase):
    def test_only_stable_versions_are_listed(self):
        for entry in manager.CHANGELOG:
            self.assertRegex(entry["version"], r"^\d+\.\d+\.\d+$", "pre-release versions are not shown in About")

    def test_newest_first_and_current_version_on_top(self):
        keys = [tuple(int(p) for p in re.findall(r"\d+", e["version"])) for e in manager.CHANGELOG]
        self.assertEqual(keys, sorted(keys, reverse=True))
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(manager.CHANGELOG[0]["version"], manager.APP_VERSION)

    def test_every_entry_has_a_date_and_features(self):
        for entry in manager.CHANGELOG:
            self.assertRegex(entry["date"], r"^\d{4}-\d{2}-\d{2}$")
            self.assertTrue(entry["features"])


if __name__ == "__main__":
    unittest.main()
