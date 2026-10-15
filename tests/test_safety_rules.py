"""Unit tests for authorization and pre-upgrade verdict rules."""
import unittest

import ios_xe_upgrade_manager as manager


class AuthorizationRuleTests(unittest.TestCase):
    def test_viewer_is_read_only(self):
        self.assertEqual(manager.role_allows(manager.ROLE_VIEWER, "GET", "/api/status"), (True, ""))
        allowed, message = manager.role_allows(manager.ROLE_VIEWER, "POST", "/api/upgrade")
        self.assertFalse(allowed)
        self.assertIn("read-only", message)

    def test_operator_cannot_stop_monitoring(self):
        allowed, message = manager.role_allows(
            manager.ROLE_OPERATOR, "POST", "/api//monitor/%73top?ignored=1")
        self.assertFalse(allowed)
        self.assertIn("Administrator", message)

    def test_unknown_role_fails_closed(self):
        allowed, message = manager.role_allows("unexpected", "GET", "/api/status")
        self.assertFalse(allowed)
        self.assertIn("no valid role", message)

    def test_admin_can_submit_upgrade(self):
        self.assertEqual(manager.role_allows(manager.ROLE_ADMIN, "POST", "/api/upgrade"), (True, ""))


class DashboardExposureTests(unittest.TestCase):
    def test_loopback_is_allowed_without_remote_auth(self):
        self.assertIsNone(manager.dashboard_exposure_error("127.0.0.1", False, None, None))

    def test_remote_binding_requires_explicit_flag(self):
        error = manager.dashboard_exposure_error("192.0.2.10", False, {"users": {"admin": {}}}, True)
        self.assertIn("--allow-remote", error)

    def test_remote_binding_requires_users_and_https(self):
        self.assertIn("configured dashboard users", manager.dashboard_exposure_error(
            "192.0.2.10", True, None, True))
        self.assertIn("requires HTTPS", manager.dashboard_exposure_error(
            "192.0.2.10", True, {"users": {"admin": {}}}, None))

    def test_remote_binding_with_auth_and_tls_is_allowed(self):
        self.assertIsNone(manager.dashboard_exposure_error(
            "192.0.2.10", True, {"users": {"admin": {}}}, ("cert.pem", "key.pem")))


class ReadinessVerdictTests(unittest.TestCase):
    @staticmethod
    def checks(**overrides):
        result = {
            "reachable": True, "login": True, "space_ok": True, "space_level": "ok",
            "space_detail": "Sufficient flash space", "boot": {"mode": "INSTALL", "boot_ok": True},
            "boot_error": "", "cpu_pct": 20, "mem_free_pct": 40, "stack": "Ready",
            "version": "17.12.4", "snapshot": None, "snapshot_text": "",
        }
        result.update(overrides)
        return result

    def test_unreachable_device_is_not_ready(self):
        state, _ = manager.compute_readiness_verdict(self.checks(reachable=False))
        self.assertEqual(state, "not_ready")

    def test_low_flash_is_a_hard_block(self):
        state, message = manager.compute_readiness_verdict(
            self.checks(space_ok=False, space_detail="Not enough flash"))
        self.assertEqual(state, "not_ready")
        self.assertIn("Not enough flash", message)

    def test_high_cpu_requires_operator_attention(self):
        state, message = manager.compute_readiness_verdict(self.checks(cpu_pct=95))
        self.assertEqual(state, "warning")
        self.assertIn("CPU 95%", message)

    def test_healthy_device_is_ready(self):
        state, _ = manager.compute_readiness_verdict(self.checks())
        self.assertEqual(state, "ready")


if __name__ == "__main__":
    unittest.main()
