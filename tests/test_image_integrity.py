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
        self.sidecar = Path(str(self.image) + ".sha256")
        self.digest = hashlib.sha256(self.payload).hexdigest()
        self.sha512_sidecar = Path(str(self.image) + ".sha512")
        self.digest512 = hashlib.sha512(self.payload).hexdigest()

    def test_sha512_digest_is_verified_in_chunks(self):
        self.assertEqual(manager.verify_image_sha512(self.image, self.digest512, chunk_size=3), self.digest512)

    def test_sha512_sidecar_is_accepted_for_upgrade_selection(self):
        self.sha512_sidecar.write_text(self.digest512 + "\n", encoding="utf-8")
        self.assertEqual(manager.verify_image_checksum_sidecar(self.image), self.digest512)
        with patch.object(manager, "images_dir", return_value=self.root):
            image, _ = manager.Manager._check_image(self.image.name)
        self.assertEqual(Path(image), self.image.resolve())

    def test_sha512_mismatch_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "SHA-512 mismatch"):
            manager.verify_image_sha512(self.image, "0" * 128)

    def test_missing_sidecar_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "missing SHA-256 sidecar"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_sha256sum_format_is_verified_in_chunks(self):
        self.sidecar.write_text(f"{self.digest}  {self.image.name}\n", encoding="utf-8")
        self.assertEqual(manager.verify_image_sha256_sidecar(self.image, chunk_size=3), self.digest)

    def test_mismatch_fails_closed(self):
        self.sidecar.write_text("0" * 64 + "  image.bin\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_malformed_manifest_fails_closed(self):
        self.sidecar.write_text("not-a-digest\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "64-character hexadecimal digest"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_empty_manifest_fails_closed(self):
        self.sidecar.write_text("# checksum intentionally missing\n\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "64-character hexadecimal digest"):
            manager.verify_image_sha256_sidecar(self.image)

    def test_image_selection_refuses_sidecar_mismatch(self):
        self.sidecar.write_text("0" * 64, encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=self.root):
            with self.assertRaisesRegex(manager.ApiError, "Image integrity check failed"):
                manager.Manager._check_image(self.image.name)

    def test_image_selection_accepts_matching_sidecar(self):
        self.sidecar.write_text(self.digest, encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=self.root):
            image, size = manager.Manager._check_image(self.image.name)
        self.assertEqual(Path(image), self.image.resolve())
        self.assertEqual(size, len(self.payload))

    def test_image_selection_requires_a_manifest(self):
        with patch.object(manager, "images_dir", return_value=self.root):
            with self.assertRaisesRegex(manager.ApiError, "missing SHA-512 or SHA-256 checksum sidecar"):
                manager.Manager._check_image(self.image.name)

    def test_non_bin_file_is_rejected(self):
        other = self.root / "image.txt"
        other.write_text("not a firmware image", encoding="utf-8")
        with self.assertRaisesRegex(manager.ApiError, "must be a .bin file"):
            manager.Manager._check_image(str(other))


if __name__ == "__main__":
    unittest.main()
