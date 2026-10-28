"""Regression tests for image integrity checks; no live switch is required."""
import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import ios_xe_upgrade_manager as manager


class ImageIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.image = self.root / "cat9k_iosxe.bin"
        self.payload = b"representative IOS-XE image bytes\x00\x01"
        self.image.write_bytes(self.payload)
        self.md5_sidecar = Path(str(self.image) + ".md5")
        self.sha256_sidecar = Path(str(self.image) + ".sha256")
        self.digest_md5 = hashlib.md5(self.payload).hexdigest()
        self.digest_sha256 = hashlib.sha256(self.payload).hexdigest()

    def test_md5_digest_is_verified_in_chunks(self):
        self.assertEqual(
            manager.verify_image_md5(self.image, self.digest_md5, chunk_size=3),
            self.digest_md5,
        )

    def test_md5_sidecar_is_accepted_for_upgrade_selection(self):
        self.md5_sidecar.write_text(self.digest_md5 + "\n", encoding="utf-8")
        self.assertEqual(manager.verify_image_checksum_sidecar(self.image), self.digest_md5)
        with patch.object(manager, "images_dir", return_value=self.root):
            image, _ = manager.Manager._check_image(self.image.name)
        self.assertEqual(Path(image), self.image.resolve())

    def test_md5_mismatch_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "MD5 mismatch"):
            manager.verify_image_md5(self.image, "0" * 32)

    def test_missing_sidecar_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "missing MD5 checksum sidecar"):
            manager.verify_image_checksum_sidecar(self.image)

    def test_legacy_sha256sum_sidecar_is_still_accepted(self):
        self.sha256_sidecar.write_text(
            f"{self.digest_sha256}  {self.image.name}\n", encoding="utf-8"
        )
        self.assertEqual(
            manager.verify_image_checksum_sidecar(self.image), self.digest_sha256
        )

    def test_sha256sum_format_is_verified_in_chunks(self):
        self.sha256_sidecar.write_text(
            f"{self.digest_sha256}  {self.image.name}\n", encoding="utf-8"
        )
        self.assertEqual(
            manager.verify_image_sha256_sidecar(self.image, chunk_size=3),
            self.digest_sha256,
        )

    def test_sha256_mismatch_fails_closed(self):
        self.sha256_sidecar.write_text("0" * 64 + "  image.bin\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_malformed_legacy_manifest_fails_closed(self):
        self.sha256_sidecar.write_text("not-a-digest\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "64-character hexadecimal digest"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_empty_legacy_manifest_fails_closed(self):
        self.sha256_sidecar.write_text(
            "# checksum intentionally missing\n\n", encoding="utf-8"
        )
        with self.assertRaisesRegex(ValueError, "64-character hexadecimal digest"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_image_selection_refuses_md5_sidecar_mismatch(self):
        self.md5_sidecar.write_text("0" * 32, encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=self.root):
            with self.assertRaisesRegex(manager.ApiError, "Image integrity check failed"):
                manager.Manager._check_image(self.image.name)

    def test_image_selection_accepts_matching_md5_sidecar(self):
        self.md5_sidecar.write_text(self.digest_md5, encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=self.root):
            image, size = manager.Manager._check_image(self.image.name)
        self.assertEqual(Path(image), self.image.resolve())
        self.assertEqual(size, len(self.payload))

    def test_image_selection_requires_a_checksum_sidecar(self):
        with patch.object(manager, "images_dir", return_value=self.root):
            with self.assertRaisesRegex(manager.ApiError, "missing MD5 checksum sidecar"):
                manager.Manager._check_image(self.image.name)

    def test_non_bin_file_is_rejected(self):
        other = self.root / "image.txt"
        other.write_text("not a firmware image", encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=self.root):
            with self.assertRaisesRegex(manager.ApiError, "must be a .bin file"):
                manager.Manager._check_image(str(other))


if __name__ == "__main__":
    unittest.main()
