"""Regression tests for the CodeQL fixes: SSH host keys (trust on first use), confined paths, safe headers."""
import socket
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import storage_sandbox  # noqa: F401  (must come before the app: keeps tests away from real data)
import ios_xe_upgrade_manager as manager
import paramiko
from netmiko.exceptions import NetmikoAuthenticationException, NetmikoTimeoutException


def _rsa():
    return paramiko.RSAKey.generate(2048)


class _FakeSSHServer(object):
    """A local SSH server that offers `host_key` and rejects every login (enough to test the host key check)."""

    def __init__(self, host_key):
        self.host_key = host_key
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(5)
        self.port = self.sock.getsockname()[1]
        self.transports = []
        threading.Thread(target=self._serve, daemon=True).start()

    def _serve(self):
        while True:
            try:
                conn, _ = self.sock.accept()
            except OSError:
                return
            t = paramiko.Transport(conn)
            t.add_server_key(self.host_key)
            self.transports.append(t)
            try:
                t.start_server(server=paramiko.ServerInterface())   # default: every auth method is refused
            except Exception:  # noqa: BLE001
                pass

    def close(self):
        self.sock.close()
        for t in self.transports:
            t.close()


class HostKeyPolicyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = str(Path(self.temp.name) / manager.KNOWN_HOSTS_FILENAME)
        self.policy = manager.TrustOnFirstUsePolicy(self.path)

    def test_first_key_is_saved_and_same_key_is_accepted_again(self):
        key = _rsa()
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", key)
        saved = manager.load_known_host_keys(self.path).lookup("10.0.0.1")
        self.assertEqual(saved["ssh-rsa"].asbytes(), key.asbytes())
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", key)   # no exception
        self.assertEqual(len(Path(self.path).read_text(encoding="utf-8").splitlines()), 1)

    def test_changed_key_is_refused(self):
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", _rsa())
        with self.assertRaisesRegex(manager.SSHHostKeyMismatch, "has changed"):
            self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", _rsa())

    def test_different_key_type_for_a_known_host_is_refused(self):
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", _rsa())
        with self.assertRaises(manager.SSHHostKeyMismatch):
            self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", paramiko.ECDSAKey.generate())

    def test_keys_are_kept_per_host_and_port(self):
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.1", _rsa())
        self.policy.missing_host_key(paramiko.SSHClient(), "10.0.0.2", _rsa())
        self.policy.missing_host_key(paramiko.SSHClient(), "[10.0.0.1]:2222", _rsa())
        self.assertEqual(len(manager.load_known_host_keys(self.path)), 3)

    def test_bad_host_key_inside_a_netmiko_error_is_reported_as_mismatch(self):
        try:
            try:
                raise paramiko.BadHostKeyException("10.0.0.1", _rsa(), _rsa())
            except paramiko.SSHException:
                raise NetmikoTimeoutException("A paramiko SSHException occurred")
        except NetmikoTimeoutException as exc:
            found = manager.host_key_error(exc)
        self.assertIsInstance(found, manager.SSHHostKeyMismatch)
        self.assertIn("10.0.0.1", str(found))
        self.assertIsNone(manager.host_key_error(NetmikoTimeoutException("TCP connection failed")))

    def test_known_hosts_file_follows_the_site_command_log_folder(self):
        with patch.object(manager, "COMMAND_LOG_DIR", self.temp.name):
            self.assertEqual(manager.known_hosts_path(), self.path)


class HostKeyConnectionTests(unittest.TestCase):
    """End to end through the real ConnectHandler wrapper (netmiko + paramiko) against a local SSH server."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        patcher = patch.object(manager, "COMMAND_LOG_DIR", self.temp.name)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _login(self, server):
        manager.ConnectHandler(device_type="cisco_xe", host="127.0.0.1", port=server.port,
                               username="u", password="p", fast_cli=False)

    def test_unknown_switch_is_trusted_then_a_changed_key_is_refused(self):
        first = _FakeSSHServer(_rsa())
        self.addCleanup(first.close)
        with self.assertRaises(NetmikoAuthenticationException):    # host key accepted, then the login refused
            self._login(first)
        saved = manager.load_known_host_keys(manager.known_hosts_path())
        self.assertIsNotNone(saved.lookup(f"[127.0.0.1]:{first.port}"))

        with self.assertRaises(NetmikoAuthenticationException):    # same key again: still accepted
            self._login(first)

        impostor = _FakeSSHServer(_rsa())
        self.addCleanup(impostor.close)
        Path(manager.known_hosts_path()).write_text(
            Path(manager.known_hosts_path()).read_text(encoding="utf-8").replace(
                f"[127.0.0.1]:{first.port}", f"[127.0.0.1]:{impostor.port}"), encoding="utf-8")
        with self.assertRaisesRegex(manager.SSHHostKeyMismatch, "has changed"):
            self._login(impostor)

    def test_web_console_refuses_a_changed_key(self):
        server = _FakeSSHServer(_rsa())
        self.addCleanup(server.close)
        manager.TrustOnFirstUsePolicy(manager.known_hosts_path()).missing_host_key(
            paramiko.SSHClient(), f"[127.0.0.1]:{server.port}", _rsa())         # a different key is saved
        session = manager.TerminalSession("127.0.0.1", "sw1", "u", "p", "", "test", "ssh", server.port)
        for _ in range(200):
            if session.closed:
                break
            threading.Event().wait(0.05)
        self.assertTrue(session.closed)
        self.assertIn("has changed", session.read(0)["text"])


class ConfinedPathTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "root"
        self.root.mkdir()

    def test_paths_inside_are_allowed(self):
        self.assertEqual(manager.confine_path(self.root, "a", "b.txt"), (self.root / "a" / "b.txt").resolve())
        self.assertEqual(manager.confine_path(self.root), self.root.resolve())

    def test_escapes_are_refused(self):
        outside = Path(self.temp.name) / "rootless"           # shares the prefix "root" but is outside it
        for parts in (("..", "x"), ("a", "..", "..", "x"), (str(outside),), ("../rootless",)):
            with self.assertRaises(ValueError, msg=parts):
                manager.confine_path(self.root, *parts)

    def test_storage_resolve_still_rejects_traversal(self):
        with self.assertRaises(manager.ApiError):
            manager.storage_resolve(self.root, f"{manager.SITE_BACKUPS_SUBDIR}/../../x")

    def test_image_selection_refuses_a_path_outside_the_images_folder(self):
        images = self.root / "images"
        images.mkdir()
        elsewhere = self.root / "elsewhere"
        elsewhere.mkdir()
        payload = b"image"
        for folder in (images, elsewhere):
            (folder / "cat9k.bin").write_bytes(payload)
            (folder / "cat9k.bin.md5").write_text(manager.hashlib.md5(payload).hexdigest(), encoding="utf-8")
        with patch.object(manager, "images_dir", return_value=images):
            with self.assertRaisesRegex(manager.ApiError, "must be in the IOS Images folder"):
                manager.Manager._check_image(str(elsewhere / "cat9k.bin"))
            with self.assertRaisesRegex(manager.ApiError, "must be in the IOS Images folder"):
                manager.Manager._check_image(r"\\server\share\cat9k.bin")
            image, _ = manager.Manager._check_image(str(images / "cat9k.bin"))   # a saved schedule's full path
            self.assertEqual(Path(image), (images / "cat9k.bin").resolve())
            self.assertIsNone(manager.image_store_path("cat9k.txt"))
            self.assertEqual(manager.image_store_path("../cat9k.bin"), (images / "cat9k.bin").resolve())


class HeaderTests(unittest.TestCase):
    def test_header_safe_removes_line_breaks(self):
        self.assertEqual(manager.header_safe("/x\r\nSet-Cookie: a=b"), "/xSet-Cookie: a=b")

    def test_content_disposition_cannot_inject_headers(self):
        value = manager.content_disposition('evil\r\nSet-Cookie: a=b".zip')
        self.assertNotIn("\r", value)
        self.assertNotIn("\n", value)
        self.assertTrue(value.startswith('attachment; filename="evilSet-Cookie__a_b_.zip"'))
        self.assertIn("filename*=UTF-8''evilSet-Cookie%3A%20a%3Db%22.zip", value)

    def test_content_disposition_keeps_unicode_names(self):
        value = manager.content_disposition("Büro Site.zip")
        self.assertIn('filename="B_ro_Site.zip"', value)
        self.assertIn("filename*=UTF-8''B%C3%BCro%20Site.zip", value)


if __name__ == "__main__":
    unittest.main()
