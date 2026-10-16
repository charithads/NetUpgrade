"""HTTP-level tests for checksum uploads without connecting to a switch."""
import hashlib
import json
import socket
import tempfile
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import patch

import ios_xe_upgrade_manager as manager


class ChecksumUploadEndpointTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.temp.name)
        cls.image_name = "cat9k_iosxe.bin"
        cls.image = cls.root / cls.image_name
        cls.payload = b"IOS-XE test image bytes"
        cls.image.write_bytes(cls.payload)
        cls.digest = hashlib.sha256(cls.payload).hexdigest()
        cls.digest512 = hashlib.sha512(cls.payload).hexdigest()
        cls.images_patch = patch.object(manager, "images_dir", return_value=cls.root)
        cls.images_patch.start()
        sock = socket.socket()
        sock.bind(("127.0.0.1", 0))
        port = sock.getsockname()[1]
        sock.close()
        cls.url, cls.shutdown = manager.gui_main(
            embedded=True, argv=["--web-port", str(port), "--mode", "tcp"])

    @classmethod
    def tearDownClass(cls):
        cls.shutdown()
        cls.images_patch.stop()
        cls.temp.cleanup()

    def post_image(self, filename, payload, digest):
        boundary = "codex-one-shot-boundary"
        body = (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            "Content-Type: application/octet-stream\r\n\r\n"
        ).encode("ascii") + payload + f"\r\n--{boundary}--\r\n".encode("ascii")
        headers = {"Content-Type": f"multipart/form-data; boundary={boundary}"}
        if digest is not None:
            headers["X-Image-SHA512"] = digest
        req = urllib.request.Request(
            self.url + "/api/images/upload", data=body,
            headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=10) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as exc:
            return exc.code, json.loads(exc.read())

    def test_image_and_pasted_sha512_are_verified_in_one_upload(self):
        filename = "one_shot.bin"
        status, response = self.post_image(filename, self.payload, self.digest512)
        self.assertEqual(status, 200)
        self.assertIn("SHA-512", response["message"])
        image = self.root / filename
        self.assertEqual(image.read_bytes(), self.payload)
        self.assertEqual(Path(str(image) + ".sha512").read_text(encoding="ascii").strip(),
                         self.digest512)
        self.assertEqual(manager.verify_image_checksum_sidecar(image), self.digest512)

    def test_one_shot_upload_rejects_mismatch_without_saving_image(self):
        filename = "bad_checksum.bin"
        status, response = self.post_image(filename, self.payload, "0" * 128)
        self.assertEqual(status, 400)
        self.assertIn("SHA-512 mismatch", response["error"])
        self.assertFalse((self.root / filename).exists())
        self.assertEqual(list(self.root.glob("*.upload")), [])

    def test_one_shot_upload_requires_pasted_checksum(self):
        filename = "missing_checksum.bin"
        status, response = self.post_image(filename, self.payload, None)
        self.assertEqual(status, 400)
        self.assertIn("Paste the 128-character SHA-512 checksum", response["error"])
        self.assertFalse((self.root / filename).exists())

    def test_bad_checksum_does_not_replace_existing_verified_image(self):
        filename = "existing.bin"
        image = self.root / filename
        image.write_bytes(b"previous verified image")
        sidecar = Path(str(image) + ".sha512")
        old_digest = hashlib.sha512(image.read_bytes()).hexdigest()
        sidecar.write_text(old_digest + "\n", encoding="ascii")

        status, response = self.post_image(filename, self.payload, "0" * 128)

        self.assertEqual(status, 400)
        self.assertIn("SHA-512 mismatch", response["error"])
        self.assertEqual(image.read_bytes(), b"previous verified image")
        self.assertEqual(manager.verify_image_checksum_sidecar(image), old_digest)

    def post_manifest(self, content):
        boundary = "codex-test-boundary"
        body = (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{self.image_name}.sha256"\r\n'
            "Content-Type: text/plain\r\n\r\n"
        ).encode("ascii") + content + f"\r\n--{boundary}--\r\n".encode("ascii")
        req = urllib.request.Request(
            self.url + "/api/images/checksum-upload", data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=10) as response:
                return response.status, json.loads(response.read())
        except urllib.error.HTTPError as exc:
            return exc.code, json.loads(exc.read())

    def test_valid_manifest_is_verified_and_saved(self):
        status, response = self.post_manifest(f"{self.digest}  {self.image_name}\n".encode("ascii"))
        self.assertEqual(status, 200)
        self.assertIn("Verified", response["message"])
        self.assertEqual(manager.verify_image_sha256_sidecar(self.image), self.digest)

    def test_invalid_manifest_is_rejected_without_replacing_valid_one(self):
        valid = f"{self.digest}\n".encode("ascii")
        status, _ = self.post_manifest(valid)
        self.assertEqual(status, 200)
        bad_status, response = self.post_manifest(("0" * 64).encode("ascii"))
        self.assertEqual(bad_status, 400)
        self.assertIn("SHA-256 mismatch", response["error"])
        self.assertEqual(manager.verify_image_sha256_sidecar(self.image), self.digest)
        self.assertEqual(list(self.root.glob("*.upload")), [])

    def test_manifest_upload_requires_matching_image_filename(self):
        boundary = "codex-test-boundary"
        filename = "other.bin.sha256"
        body = (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{filename}\"\r\n"
                "Content-Type: text/plain\r\n\r\n" + self.digest + f"\r\n--{boundary}--\r\n").encode("ascii")
        req = urllib.request.Request(
            self.url + "/api/images/checksum-upload", data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}, method="POST")
        with self.assertRaises(urllib.error.HTTPError) as error:
            urllib.request.urlopen(req, timeout=10)
        self.assertEqual(error.exception.code, 400)


if __name__ == "__main__":
    unittest.main()
