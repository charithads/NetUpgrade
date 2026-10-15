"""Headless smoke test: fails fast if the tool can no longer start. No switches needed."""
import re, socket, sys, tempfile, urllib.request, urllib.error
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def const(path, name):
    m = re.search(rf'^{name}\s*=\s*["\']?([\w.]+)', (ROOT / path).read_text(encoding="utf-8"), re.M)
    assert m, f"{name} not found in {path}"
    return m.group(1)


def test_versions_in_sync():
    a = const("ios_xe_upgrade_manager.py", "APP_VERSION")
    b = const("launcher.py", "APP_VERSION")
    assert a == b, f"APP_VERSION mismatch: backend {a} vs launcher {b}"


def test_guard_version():
    import ios_xe_upgrade_manager as m
    need = int(const("launcher.py", "GUARD_VERSION_NEEDED"))
    have = getattr(m.apply_access_control, "version", 0)
    assert have >= need, f"launcher needs access-control v{need}, backend has v{have}"


def test_excel_loading():
    import ios_xe_upgrade_manager as m
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(["ip", "hostname"]); ws.append(["192.0.2.10", "sw1"]); ws.append(["192.0.2.11", "sw2"])
    with tempfile.TemporaryDirectory() as d:
        f = Path(d) / "switches.xlsx"; wb.save(f)
        assert len(m.load_hosts(str(f))) == 2


def test_dashboard_starts():
    import ios_xe_upgrade_manager as m
    s = socket.socket(); s.bind(("127.0.0.1", 0)); port = s.getsockname()[1]; s.close()
    url, shutdown = m.gui_main(embedded=True, argv=["--web-port", str(port), "--mode", "tcp"])
    try:
        try:
            code = urllib.request.urlopen(url, timeout=10).status
        except urllib.error.HTTPError as e:
            code = e.code
        assert code < 500, f"dashboard returned HTTP {code}"
    finally:
        shutdown()


if __name__ == "__main__":
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            try:
                fn(); print("PASS", name)
            except Exception as exc:
                failed += 1; print("FAIL", name, "-", exc)
    sys.exit(1 if failed else 0)
