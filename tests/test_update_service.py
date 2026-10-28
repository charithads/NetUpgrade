"""Tests for safe release selection, downloading, and updater handoff validation."""
import hashlib
import io
import tempfile
import unittest
from pathlib import Path

import update_service
import updater_helper


class UpdateServiceTests(unittest.TestCase):
    def release(self, name="NetUpgrade-Setup-v5.12.2.exe", **overrides):
        data = b"windows setup payload"
        asset = {
            "name": name,
            "browser_download_url": "https://github.com/charithads/NetUpgrade/releases/download/v5.12.2/" + name,
            "size": len(data),
            "digest": "sha256:" + hashlib.sha256(data).hexdigest(),
        }
        asset.update(overrides)
        return {"tag_name": "v5.12.2", "assets": [asset]}, asset, data

    def test_version_parser_only_accepts_stable_three_part_versions(self):
        self.assertEqual(update_service.parse_version("v5.12.2"), (5, 12, 2))
        self.assertIsNone(update_service.parse_version("v5.12.2-rc1"))
        self.assertIsNone(update_service.parse_version("latest"))

    def test_only_exact_versioned_installer_is_selected(self):
        release, _, _ = self.release()
        release["assets"].insert(0, {"name": "NetUpgrade.exe", "browser_download_url": "https://github.com/x"})
        selected = update_service.pick_installer_asset(release)
        self.assertEqual(selected["name"], "NetUpgrade-Setup-v5.12.2.exe")
        self.assertTrue(update_service.is_newer_release(release, "5.12.1"))
        self.assertFalse(update_service.is_newer_release(release, "5.12.2"))

    def test_installer_is_not_selected_without_github_digest_or_safe_url(self):
        for override in ({"digest": ""}, {"browser_download_url": "http://github.com/file"},
                         {"browser_download_url": "https://example.com/file"}):
            with self.subTest(override=override):
                release, _, _ = self.release(**override)
                self.assertIsNone(update_service.pick_installer_asset(release))

    def test_download_verifies_size_and_digest_and_uses_expected_name(self):
        release, asset, payload = self.release()
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]) as temp:
            asset = update_service.pick_installer_asset(release)
            result = update_service.download_verified_installer(
                asset, temp, opener=lambda request, timeout: io.BytesIO(payload))
            self.assertEqual(result.name, "NetUpgrade-Setup-v5.12.2.exe")
            self.assertEqual(result.read_bytes(), payload)

    def test_bad_digest_and_incomplete_download_are_rejected(self):
        release, _, payload = self.release()
        asset = update_service.pick_installer_asset(release)
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]) as temp:
            bad = dict(asset, sha256="0" * 64)
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                update_service.download_verified_installer(
                    bad, temp, opener=lambda request, timeout: io.BytesIO(payload))
            with self.assertRaisesRegex(ValueError, "incomplete"):
                update_service.download_verified_installer(
                    asset, temp, opener=lambda request, timeout: io.BytesIO(payload[:-1]))
            self.assertEqual(list(Path(temp).iterdir()), [])

    def test_updater_only_accepts_files_under_expected_user_paths(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]) as temp:
            app_path, update_dir = updater_helper.expected_paths(temp)
            update_dir.mkdir(parents=True)
            installer = update_dir / "NetUpgrade-Setup-v5.12.2.exe"
            installer.write_bytes(b"installer")
            normalized, app, digest = updater_helper.validate_request(
                1234, installer, "A" * 64, app_path, temp)
            self.assertEqual(normalized, installer.resolve())
            self.assertEqual(app, app_path.resolve())
            self.assertEqual(digest, "a" * 64)
            with self.assertRaisesRegex(ValueError, "expected location"):
                updater_helper.validate_request(1234, installer, "a" * 64,
                    Path(temp) / "other.exe", temp)

    def test_updater_rejects_unsafe_installer_name_and_digest(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1]) as temp:
            app_path, update_dir = updater_helper.expected_paths(temp)
            update_dir.mkdir(parents=True)
            wrong_name = update_dir / "setup.exe"
            wrong_name.write_bytes(b"x")
            with self.assertRaisesRegex(ValueError, "setup package"):
                updater_helper.validate_request(1000, wrong_name, "a" * 64, app_path, temp)
            right_name = update_dir / "NetUpgrade-Setup-v5.12.2.exe"
            right_name.write_bytes(b"x")
            with self.assertRaisesRegex(ValueError, "digest"):
                updater_helper.validate_request(1000, right_name, "invalid", app_path, temp)


if __name__ == "__main__":
    unittest.main()

