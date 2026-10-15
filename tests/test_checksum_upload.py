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
