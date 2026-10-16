"""
Launcher for the NetUpgrade.

One small window with two tabs:
  General  - network interface, port, HTTPS, tool status (running / stopped),
             public access on/off, users who are logged in right now
  Security - admin account, allowed client IPs, user accounts (applied instantly)
Everyone - including the person at this PC - signs in to the web dashboard.
The admin account is created the first time the program starts.
Minimizes to the system tray. No cmd window.

Build:  run build.bat   (or)
  pyinstaller --noconsole --onefile --name IOSXEUpgradeManager --icon app.ico ^
      --add-data "app.ico;." --hidden-import ios_xe_upgrade_manager ^
      --collect-all netmiko --collect-all cryptography launcher.py
Needs:  pip install pyinstaller pystray pillow psutil cryptography
"""
import os
import sys

import subprocess
import shutil


# --- No console flashes ----------------------------------------------------
# This exe has no console of its own (--noconsole). On Windows, every command-line
# program it starts (the "ping" that checks each device, every few seconds) then
# gets a brand-new console window, which flashes on screen. The fix: start child
# programs with CREATE_NO_WINDOW. A child that is deliberately given its own
# visible window (CREATE_NEW_CONSOLE) - such as the upgrade-job windows - is left
# alone, and output of the child is still captured normally.
CREATE_NEW_CONSOLE, DETACHED_PROCESS, CREATE_NO_WINDOW = 0x00000010, 0x00000008, 0x08000000


def hide_child_consoles(popen_cls):
    original = popen_cls.__init__

    def quiet_init(self, *args, **kwargs):
        flags = kwargs.get("creationflags", 0)
        if len(args) <= 13 and not flags & (CREATE_NEW_CONSOLE | DETACHED_PROCESS | CREATE_NO_WINDOW):
            kwargs["creationflags"] = flags | CREATE_NO_WINDOW
        original(self, *args, **kwargs)

    popen_cls.__init__ = quiet_init


if sys.platform == "win32":
    hide_child_consoles(subprocess.Popen)      # covers subprocess.run / check_output / Popen


# --- Worker / CLI dispatch: must run before any window is built --------------
# The GUI starts each switch upgrade as "<this exe> --job <file>". Without this, every
# job would open another launcher window. It comes after the console-hiding above so the
# pings made by a worker stay hidden too.
if len(sys.argv) == 3 and sys.argv[1] == "--job":
    import ios_xe_upgrade_manager as _m
    sys.exit(_m.job_main(sys.argv[2]))
if "--cli" in sys.argv:
    sys.argv.remove("--cli")
    import ios_xe_upgrade_manager as _m
    _m.interactive_main()
    sys.exit(0)

import ctypes
import hashlib
import hmac
import ipaddress
import json
import re
import socket
import threading
import time
import traceback
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox

import psutil
import pystray
from PIL import Image, ImageTk

APP_NAME = "NetUpgrade"
APP_VERSION = "5.6.1"          # keep equal to APP_VERSION in ios_xe_upgrade_manager.py
UI_FONT_SIZE = 8               # window size follows the font: 9 = the old, larger window; 8 = compact
APP_AUTHOR = "Charitha Dasanayaka"
APP_REPO_URL = "https://github.com/charithads/NetUpgrade"
DEFAULT_HTTPS_PORT, DEFAULT_HTTP_PORT = 443, 80
MIN_PW = 10                                   # minimum password length
USERNAME_RE = re.compile(r"^[A-Za-z0-9._-]{1,32}$")
GREEN, RED, ORANGE = "#1a7f37", "#b42318", "#b26a00"
GUARD_VERSION_NEEDED = 5


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


_DATA_ROOT = os.environ.get("LOCALAPPDATA", ".")
_LEGACY_DATA_DIR = os.path.join(_DATA_ROOT, "IOSXEUpgradeManager")
DATA_DIR = os.path.join(_DATA_ROOT, "NetUpgrade")
if os.path.isdir(_LEGACY_DATA_DIR) and not os.path.exists(DATA_DIR):
    try:
        shutil.copytree(_LEGACY_DATA_DIR, DATA_DIR)
    except OSError:
        DATA_DIR = _LEGACY_DATA_DIR  # Keep existing settings usable if migration is blocked.
os.makedirs(DATA_DIR, exist_ok=True)
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
AUDIT_FILE = os.path.join(DATA_DIR, "audit.log")
TLS_DIR = os.path.join(DATA_DIR, "tls")
CERT_FILE = os.path.join(TLS_DIR, "cert.pem")
KEY_FILE = os.path.join(TLS_DIR, "key.pem")
CERT_META = os.path.join(TLS_DIR, "cert.json")

# With --noconsole sys.stdout/stderr are None; send prints/tracebacks to a log.
_log = open(os.path.join(DATA_DIR, "app.log"), "a", buffering=1, encoding="utf-8")
sys.stdout = _log
sys.stderr = _log

ICON_PATH = resource_path("app.ico")

# Tell Windows this is its own app (not "python"), so the taskbar shows our icon.
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("NetUpgrade.Launcher")
except Exception:
    pass


# --- Settings --------------------------------------------------------------
def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = {}
    # drop the old single-password settings (replaced by user accounts)
    for k in ("require_password", "pw_salt", "pw_hash"):
        data.pop(k, None)
    # one-time move to HTTPS: turn it on and switch the old default port 80 to 443
    if not data.get("tls_default_applied"):
        data["tls_default_applied"] = True
        data.setdefault("https", True)
        if data.get("port") in (None, DEFAULT_HTTP_PORT):
            data["port"] = DEFAULT_HTTPS_PORT
    return data


def save_settings():
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f)
    except Exception:
        traceback.print_exc()


def hash_password(password, salt_hex):
    # Must match apply_access_control() in ios_xe_upgrade_manager.py
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                               bytes.fromhex(salt_hex), 200_000).hex()


def make_record(password):
    salt = os.urandom(16).hex()
    return {"salt": salt, "hash": hash_password(password, salt)}


settings = load_settings()

# Live access rules. The tool reads this dict on every request, so changing it
# takes effect immediately - no restart. Keys are always replaced wholesale.
# epoch goes up each time public access is blocked -> remote users are signed out.
live_access = {"epoch": 0}

# Roles for the dashboard accounts. The tool (ios_xe_upgrade_manager.py, ROLES) enforces them on the
# server, so keep these keys identical. An account without a role is read-only (fail closed).
USER_ROLES = [("administrator", "Administrator"), ("operator", "Operator"), ("viewer", "Viewer")]
ROLE_LABEL = dict(USER_ROLES)
DEFAULT_ROLE = "viewer"


def sync_live_access():
    live_access["public"] = bool(settings.get("public", True))
    live_access["allowed"] = list(settings.get("allowed", []))
    users = {}
    for name, rec in settings.get("users", {}).items():
        rec = dict(rec)
        rec.pop("admin", None)                      # only the launcher's admin account is an admin by identity
        rec["role"] = rec.get("role") if rec.get("role") in ROLE_LABEL else DEFAULT_ROLE
        users[name] = rec
    adm = settings.get("admin")
    if adm:
        users[adm["name"]] = {"salt": adm["salt"], "hash": adm["hash"], "admin": True, "role": "administrator"}
    live_access["users"] = users
    live_access["audit_path"] = AUDIT_FILE


sync_live_access()


def audit_event(text):
    """Write a line to the audit log (same file and format the web side uses)."""
    try:
        try:
            if os.path.getsize(AUDIT_FILE) > 5_000_000:
                os.replace(AUDIT_FILE, AUDIT_FILE + ".1")
        except OSError:
            pass
        clean = re.sub(r"[\x00-\x1f]+", " ", str(text))[:200]
        line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | launcher | this PC | {clean}\n"
        with open(AUDIT_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        traceback.print_exc()


# --- HTTPS certificate -----------------------------------------------------
def local_ipv4s():
    ips = {"127.0.0.1"}
    try:
        for addrs in psutil.net_if_addrs().values():
            for a in addrs:
                if a.family == socket.AF_INET and not a.address.startswith("169.254."):
                    ips.add(a.address)
    except Exception:
        traceback.print_exc()
    return sorted(ips)


def ensure_certificate():
    """Return (cert_path, key_path) of a self-signed certificate that covers this
    PC's current IP addresses. A new one is made only when an address is missing,
    the host name changed or it is about to expire (so browsers are not asked
    to re-accept it all the time). Raises ImportError if 'cryptography' is absent."""
    os.makedirs(TLS_DIR, exist_ok=True)
    ips = local_ipv4s()
    host = socket.gethostname()
    try:
        with open(CERT_META, encoding="utf-8") as f:
            meta = json.load(f)
    except Exception:
        meta = {}
    if (os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE) and set(ips) <= set(meta.get("ips", []))
            and meta.get("host") == host and meta.get("expires", 0) > time.time() + 30 * 86400):
        return CERT_FILE, KEY_FILE

    import datetime
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

    days = 3650
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, APP_NAME)])
    san = [x509.DNSName("localhost")]
    try:
        host.encode("ascii")
        san.append(x509.DNSName(host))
    except UnicodeError:
        pass
    san += [x509.IPAddress(ipaddress.ip_address(ip)) for ip in ips]
    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (x509.CertificateBuilder()
            .subject_name(name).issuer_name(name).public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(now - datetime.timedelta(days=1))
            .not_valid_after(now + datetime.timedelta(days=days))
            .add_extension(x509.SubjectAlternativeName(san), critical=False)
            .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
            .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
            .sign(key, hashes.SHA256()))
    with open(KEY_FILE, "wb") as f:
        f.write(key.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.TraditionalOpenSSL,
                                  serialization.NoEncryption()))
    with open(CERT_FILE, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))
    with open(CERT_META, "w", encoding="utf-8") as f:
        json.dump({"ips": ips, "host": host, "expires": time.time() + days * 86400}, f)
    audit_event(f"HTTPS certificate created for {', '.join(ips)}")
    return CERT_FILE, KEY_FILE


# --- Network interfaces ----------------------------------------------------
def build_options():
    """Choices: this PC only, each active IPv4 interface, all interfaces."""
    opts = [{"key": "__local__", "label": "This PC only (127.0.0.1)", "ip": "127.0.0.1"}]
    try:
        stats = psutil.net_if_stats()
        for name, addrs in psutil.net_if_addrs().items():
            st = stats.get(name)
            if st is not None and not st.isup:
                continue
            n = 0
            for a in addrs:
                if a.family != socket.AF_INET:
                    continue
                ip = a.address
                if ip.startswith("127.") or ip.startswith("169.254."):
                    continue
                n += 1
                key = name if n == 1 else f"{name}#{n}"
                opts.append({"key": key, "label": f"{name} - {ip}", "ip": ip})
    except Exception:
        traceback.print_exc()
    opts.append({"key": "__all__", "label": "All network interfaces (0.0.0.0)", "ip": "0.0.0.0"})
    return opts


def http_url(ip, port, https):
    default = DEFAULT_HTTPS_PORT if https else DEFAULT_HTTP_PORT
    return f"{'https' if https else 'http'}://{ip}" + ("" if port == default else f":{port}")


# --- Server control --------------------------------------------------------
_url = None
_shutdown = None
_running_cfg = None        # (ip, port, https) of the running server


def selected_option():
    label = combo.get()
    for o in options:
        if o["label"] == label:
            return o
    return options[0]


def read_port():
    try:
        port = int(port_var.get())
        if 1 <= port <= 65535:
            return port
    except ValueError:
        pass
    messagebox.showerror(APP_NAME, "Port must be a number between 1 and 65535.")
    return None


def stop_server():
    global _url, _shutdown, _running_cfg
    try:
        if _shutdown:
            _shutdown()
    except Exception:
        traceback.print_exc()
    _url = _shutdown = _running_cfg = None


def server_state():
    """'stopped' (not launched), 'running' (answers on its port) or 'down' (launched but not answering)."""
    if _url is None or _running_cfg is None:
        return "stopped"
    ip, port = _running_cfg[0], _running_cfg[1]
    host = "127.0.0.1" if ip in ("0.0.0.0", "") else ip
    try:
        with socket.create_connection((host, port), timeout=0.3):
            return "running"
    except OSError:
        return "down"


def launch_tool():
    global _url, _shutdown, _running_cfg
    if not ensure_admin():
        return
    opt = selected_option()
    port = read_port()
    if port is None:
        return
    use_https = bool(https_var.get())
    cfg = (opt["ip"], port, use_https)

    if _url is not None and cfg != _running_cfg:
        if not messagebox.askyesno(
            APP_NAME,
            "The tool is already running with different network settings.\n\n"
            "Restart it with the new settings? Open dashboard pages will "
            "disconnect. Upgrades already running keep going in their own windows.",
        ):
            return
        stop_server()

    if _url is None:
        shared = opt["ip"] not in ("127.0.0.1", "localhost", "::1")
        argv = [
            "--host", opt["ip"],
            "--web-port", str(port),
        ]
        if shared:
            argv.append("--allow-remote")
        tls = None
        if use_https:
            try:
                tls = ensure_certificate()
            except ImportError:
                messagebox.showerror(
                    APP_NAME,
                    "HTTPS needs the 'cryptography' package, which is missing from this program.\n\n"
                    "Install it (pip install cryptography) and rebuild with build.bat,\n"
                    "or untick HTTPS on the General tab (not recommended).")
                return
            except Exception as exc:
                traceback.print_exc()
                messagebox.showerror(APP_NAME, f"Could not create the HTTPS certificate:\n{exc}\n\nSee app.log in\n{DATA_DIR}")
                return
        try:
            import ios_xe_upgrade_manager as m
            guard = getattr(m, "apply_access_control", None)
            if getattr(guard, "version", 0) < GUARD_VERSION_NEEDED:
                messagebox.showerror(
                    APP_NAME,
                    "ios_xe_upgrade_manager.py does not contain the latest access-control code,\n"
                    "so login, HTTPS and the Block button would not work.\n\n"
                    "Replace apply_access_control and gui_main with the code from\n"
                    "gui_main_patch.py (delete any old copies), then rebuild the exe.")
                return
            import io
            import contextlib
            err_text = io.StringIO()
            try:
                with contextlib.redirect_stderr(err_text):
                    _url, _shutdown = m.gui_main(embedded=True, argv=argv, access=live_access, tls=tls)
            except SystemExit as exc:
                if isinstance(exc.code, str) and exc.code:
                    raise
                lines = [ln for ln in err_text.getvalue().splitlines() if ln.strip()]
                detail = lines[-1] if lines else "no further details"
                print("gui_main exited:", repr(exc.code), "args:", argv, "stderr:", err_text.getvalue(), flush=True)
                messagebox.showerror(APP_NAME, f"The tool could not start.\n\n{detail}\n\nArguments: {' '.join(argv)}")
                return
            _running_cfg = cfg
            audit_event(f"tool started on {opt['ip']}:{port} ({'HTTPS' if use_https else 'HTTP'})")
        except SystemExit as exc:       # the script's own sys.exit("message") errors
            messagebox.showerror(APP_NAME, str(exc.code) if exc.code else "Could not start the tool.")
            return
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror(APP_NAME, f"Could not start the tool:\n{exc}\n\nSee app.log in\n{DATA_DIR}")
            return

    settings.update({"interface": opt["key"], "port": port, "https": use_https})
    save_settings()
    update_share_text()
    update_status()
    webbrowser.open(_url)


def stop_tool():
    if _url is None:
        return
    n = len(client_rows())
    msg = "Stop the tool?\n\n"
    if n:
        msg += f"{n} user{'s are' if n != 1 else ' is'} logged in and will be disconnected.\n\n"
    msg += "Upgrades already running keep going in their own windows."
    if not messagebox.askyesno(APP_NAME, msg):
        return
    stop_server()
    audit_event("tool stopped")
    update_status()
    update_clients()


# --- Public access toggle (live) -------------------------------------------
def toggle_public():
    turning_on = not settings.get("public", True)
    settings["public"] = turning_on
    save_settings()
    sync_live_access()          # applied immediately, no restart
    if not turning_on:
        live_access["epoch"] += 1   # invalidates every remote login session (this PC stays signed in)
    audit_event("public access " + ("allowed" if turning_on else "BLOCKED"))
    update_protection_text()
    update_clients()


# --- Tray ------------------------------------------------------------------
tray = None


def make_tray():
    global tray
    menu = pystray.Menu(
        pystray.MenuItem("Open window", lambda: root.after(0, show_window), default=True),
        pystray.MenuItem("Open web interface", lambda: root.after(0, launch_tool)),
        pystray.MenuItem("Exit", lambda: root.after(0, quit_app)),
    )
    tray = pystray.Icon("iosxe_upgrade_manager", Image.open(ICON_PATH).convert("RGBA"), APP_NAME, menu)
    tray.run_detached()


def hide_window():
    root.withdraw()


def show_window():
    root.deiconify()
    root.state("normal")
    root.lift()
    root.focus_force()


def on_minimize(event):
    if event.widget is root and root.state() == "iconic":
        root.after(10, hide_window)


def quit_app():
    if not messagebox.askyesno(
        APP_NAME,
        "Exit the tool?\n\nUpgrades already running keep going in their own windows.",
    ):
        return
    stop_server()
    audit_event("launcher closed")
    if tray:
        tray.stop()
    root.destroy()
    os._exit(0)


# --- Window icon -----------------------------------------------------------
_icon_refs = []             # keep the images alive, or Tk drops the icon


def apply_icon(window):
    """Title-bar + taskbar icon. Built from app.ico with Pillow at several sizes,
    because Tk's own .ico loader often fails on modern (PNG-compressed) icons."""
    try:
        base = Image.open(ICON_PATH).convert("RGBA")      # largest frame in the .ico
        photos = [ImageTk.PhotoImage(base.resize((s, s), Image.LANCZOS)) for s in (16, 24, 32, 48, 64)]
        _icon_refs.extend(photos)
        window.iconphoto(True, *photos)
    except Exception:
        traceback.print_exc()
        try:
            window.iconbitmap(default=ICON_PATH)
        except Exception:
            traceback.print_exc()


# --- Admin account ---------------------------------------------------------
def set_admin(name, password=None):
    """Create the admin account (password given) or rename it (password None)."""
    if password is not None:
        rec = make_record(password)
        settings["admin"] = {"name": name, "salt": rec["salt"], "hash": rec["hash"]}
    else:
        settings["admin"]["name"] = name
    save_settings()
    sync_live_access()          # sessions of the old credentials stop working at once


def make_show_check(parent, entries):
    """'Show password' checkbox that reveals/hides the typed text of `entries`."""
    var = tk.BooleanVar(value=False)

    def toggle():
        for e in entries:
            e.configure(show="" if var.get() else "*")
    return ttk.Checkbutton(parent, text="Show password", variable=var, command=toggle, takefocus=False)


def admin_dialog(first_run):
    """Create the admin account (first start) or change its credentials. Returns True if saved."""
    adm = settings.get("admin")
    first_run = first_run or not adm
    dlg = tk.Toplevel(root)
    dlg.title("Create the admin account" if first_run else "Change admin credentials")
    dlg.transient(root)
    dlg.resizable(False, False)
    apply_icon(dlg)
    saved = {"ok": False}
    body = ttk.Frame(dlg, padding=14)
    body.pack(fill="both", expand=True)
    if first_run:
        ttk.Label(body, text="First start: create the administrator account.\n"
                             "Everyone signs in to the tool, also on this PC.",
                  justify="left").pack(anchor="w", pady=(0, 8))
    grid = ttk.Frame(body)
    grid.pack(fill="x")
    grid.columnconfigure(1, weight=1)
    entries = {}

    def add_row(r, label, key, show=None, initial=""):
        ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=2)
        e = ttk.Entry(grid, show=show, width=28)
        e.grid(row=r, column=1, sticky="ew", padx=(8, 0), pady=2)
        if initial:
            e.insert(0, initial)
        entries[key] = e

    r = 0
    if not first_run:
        add_row(r, "Current password", "cur", "*")
        r += 1
    add_row(r, "Admin username", "name", None, adm["name"] if adm else "admin")
    r += 1
    add_row(r, "Password" if first_run else "New password", "pw1", "*")
    r += 1
    add_row(r, "Confirm password", "pw2", "*")
    pw_keys = [k for k in ("cur", "pw1", "pw2") if k in entries]
    make_show_check(grid, [entries[k] for k in pw_keys]).grid(row=r + 1, column=1, sticky="w", padx=(8, 0))
    ttk.Label(body, foreground="gray",
              text=(f"At least {MIN_PW} characters." if first_run else
                    "Leave the new password empty to keep the current one.\n"
                    "Changing it signs you out of the web page.")).pack(anchor="w", pady=(6, 0))
    msg = tk.StringVar()
    tk.Label(body, textvariable=msg, fg=RED, justify="left").pack(anchor="w")

    def save(_e=None):
        name = entries["name"].get().strip().lower()
        pw1, pw2 = entries["pw1"].get(), entries["pw2"].get()
        if not USERNAME_RE.match(name):
            msg.set("Username: 1-32 characters, letters, digits, '.', '_' or '-'.")
            return
        if name in settings.get("users", {}):
            msg.set("That name is already used by another account.")
            return
        if not first_run:
            if not hmac.compare_digest(hash_password(entries["cur"].get(), adm["salt"]), adm["hash"]):
                msg.set("The current password is wrong.")
                audit_event("admin credential change refused: wrong current password")
                return
        if first_run or pw1 or pw2:
            if pw1 != pw2:
                msg.set("The passwords do not match.")
                return
            if len(pw1) < MIN_PW:
                msg.set(f"Use a password of at least {MIN_PW} characters.")
                return
        set_admin(name, pw1 or None)
        audit_event("admin account created" if first_run else "admin credentials changed")
        update_admin_label()
        update_protection_text()
        update_clients()
        saved["ok"] = True
        dlg.destroy()

    btns = ttk.Frame(body)
    btns.pack(fill="x", pady=(10, 0))
    ttk.Button(btns, text="Save", width=9, command=save).pack(side="right")
    ttk.Button(btns, text="Cancel", width=9, command=dlg.destroy).pack(side="right", padx=(0, 6))
    dlg.bind("<Return>", save)
    dlg.bind("<Escape>", lambda _e: dlg.destroy())
    first = entries["cur"] if not first_run else entries["name"]
    first.focus_set()
    dlg.grab_set()
    root.wait_window(dlg)
    return saved["ok"]


def ensure_admin():
    if settings.get("admin"):
        return True
    return admin_dialog(True)


def startup_admin():
    if not settings.get("admin") and not admin_dialog(True):
        messagebox.showwarning(APP_NAME, "The admin account is not set up yet.\n\n"
                                         "The tool cannot be launched until you create it\n"
                                         "(Launch Tool, or Security tab > Change...).")


# --- Window ----------------------------------------------------------------
root = tk.Tk()
root.title(f"{APP_NAME}  v{APP_VERSION}")
root.resizable(False, False)     # no fixed size: the window shrinks to fit its contents
for _font_name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont", "TkFixedFont"):
    try:
        tkfont.nametofont(_font_name).configure(size=UI_FONT_SIZE)
    except tk.TclError:
        pass
ttk.Style().configure("Treeview", rowheight=int(UI_FONT_SIZE * 2.1))
apply_icon(root)                 # window title bar + taskbar

options = build_options()

# Credit footer: packed first so it is the very bottom strip, below the buttons.
credit = ttk.Frame(root, padding=(6, 1, 6, 4))
credit.pack(side="bottom", fill="x")
ttk.Separator(root).pack(side="bottom", fill="x")
_credit_row = ttk.Frame(credit)
_credit_row.pack(anchor="center")
ttk.Label(_credit_row, text="Developed by ").pack(side="left")
ttk.Label(_credit_row, text=APP_AUTHOR, font=("Segoe UI", UI_FONT_SIZE, "bold")).pack(side="left")
ttk.Label(_credit_row, text="  |  ", foreground="gray").pack(side="left")
_gh_link = ttk.Label(_credit_row, text="GitHub", foreground="#0969da", cursor="hand2",
                     font=("Segoe UI", UI_FONT_SIZE, "underline"))
_gh_link.pack(side="left")
_gh_link.bind("<Button-1>", lambda _e: webbrowser.open(APP_REPO_URL))

# Buttons at the bottom are packed next so they never get squeezed out.
bottom = ttk.Frame(root, padding=(6, 0, 6, 5))
bottom.pack(side="bottom", fill="x")
bottom.columnconfigure(0, weight=2)
bottom.columnconfigure(1, weight=1)
bottom.columnconfigure(2, weight=1)
launch_btn = ttk.Button(bottom, text="Launch Tool", command=lambda: launch_tool())
launch_btn.grid(row=0, column=0, sticky="ew", ipady=2, padx=(0, 4))
stop_btn = ttk.Button(bottom, text="Stop", command=lambda: stop_tool())
stop_btn.grid(row=0, column=1, sticky="ew", ipady=2, padx=(0, 4))
ttk.Button(bottom, text="Exit", command=lambda: quit_app()).grid(row=0, column=2, sticky="ew", ipady=2)

nb = ttk.Notebook(root)
nb.pack(fill="both", expand=True, padx=6, pady=(6, 4))
gen = ttk.Frame(nb, padding=6)
sec = ttk.Frame(nb, padding=6)
nt = ttk.Frame(nb, padding=6)
nb.add(gen, text="  General  ")
nb.add(sec, text="  Security  ")
nb.add(nt, text="  Notifications  ")

# ============================ GENERAL TAB ==================================
gen.columnconfigure(1, weight=1)

ttk.Label(gen, text="Status").grid(row=0, column=0, sticky="w", pady=2)
status_lbl = tk.Label(gen, text="", font=("Segoe UI", UI_FONT_SIZE, "bold"), anchor="w")
status_lbl.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=2)

ttk.Label(gen, text="Interface").grid(row=1, column=0, sticky="w", pady=2)
row = ttk.Frame(gen)
row.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=2)
combo = ttk.Combobox(row, state="readonly", width=26, values=[o["label"] for o in options])
combo.pack(side="left", fill="x", expand=True)


def refresh_interfaces():
    global options
    current_key = selected_option()["key"] if combo.get() else settings.get("interface")
    options = build_options()
    combo["values"] = [o["label"] for o in options]
    match = next((o for o in options if o["key"] == current_key), options[0])
    combo.set(match["label"])
    update_share_text()


ttk.Button(row, text="Refresh", width=7, command=refresh_interfaces).pack(side="left", padx=(4, 0))

ttk.Label(gen, text="Port").grid(row=2, column=0, sticky="w", pady=2)
port_row = ttk.Frame(gen)
port_row.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=2)
port_var = tk.StringVar(value=str(settings.get("port", DEFAULT_HTTPS_PORT)))
ttk.Entry(port_row, textvariable=port_var, width=7).pack(side="left")
https_var = tk.BooleanVar(value=bool(settings.get("https", True)))


def on_https_toggle():
    if not https_var.get():
        if not messagebox.askyesno(
            APP_NAME,
            "Without HTTPS, user names and passwords travel unencrypted on the network.\n\n"
            "Turn HTTPS off anyway?", icon="warning"):
            https_var.set(True)
            return
    if port_var.get().strip() in (str(DEFAULT_HTTP_PORT), str(DEFAULT_HTTPS_PORT)):
        port_var.set(str(DEFAULT_HTTPS_PORT if https_var.get() else DEFAULT_HTTP_PORT))
    update_share_text()
    update_protection_text()


ttk.Checkbutton(port_row, text="Use HTTPS (encrypted)", variable=https_var,
                command=on_https_toggle).pack(side="left", padx=(12, 0))

ttk.Label(gen, text="Link").grid(row=3, column=0, sticky="w", pady=2)
share_var = tk.StringVar()
ttk.Entry(gen, textvariable=share_var, state="readonly").grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=2)

ttk.Label(gen, text="Public access").grid(row=4, column=0, sticky="w", pady=2)
pub_row = ttk.Frame(gen)
pub_row.grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=2)
pub_lbl = tk.Label(pub_row, text="", font=("Segoe UI", UI_FONT_SIZE, "bold"))
pub_lbl.pack(side="left")
pub_btn = ttk.Button(pub_row, text="", width=8, command=toggle_public)
pub_btn.pack(side="right")

ttk.Separator(gen).grid(row=5, column=0, columnspan=2, sticky="ew", pady=3)

clients_var = tk.StringVar(value="Logged-in users: 0")
ttk.Label(gen, textvariable=clients_var, font=("Segoe UI", UI_FONT_SIZE, "bold")).grid(row=6, column=0, columnspan=2, sticky="w")
tree_frame = ttk.Frame(gen)
tree_frame.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=(3, 0))
gen.rowconfigure(7, weight=1)
tree = ttk.Treeview(tree_frame, columns=("user", "ip", "since", "last"), show="headings", height=3)
for col, text, width in (("user", "Username", 95), ("ip", "IP address", 90),
                         ("since", "Logged in", 62), ("last", "Last active", 70)):
    tree.heading(col, text=text)
    tree.column(col, width=width, anchor="w", stretch=(col == "user"))
tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=tree_scroll.set)
tree.pack(side="left", fill="both", expand=True)
tree_scroll.pack(side="right", fill="y")

prot_var = tk.StringVar()
adm_var = tk.StringVar()


def update_share_text(*_):
    opt = selected_option()
    try:
        port = int(port_var.get())
    except ValueError:
        port = DEFAULT_HTTPS_PORT if https_var.get() else DEFAULT_HTTP_PORT
    https = bool(https_var.get())
    if opt["key"] == "__local__":
        share_var.set(http_url("127.0.0.1", port, https) + "  (this PC only)")
    elif opt["key"] == "__all__":
        ips = [o["ip"] for o in options if o["key"] not in ("__local__", "__all__")]
        share_var.set("  |  ".join(http_url(ip, port, https) for ip in ips) or "(no active interfaces found)")
    else:
        share_var.set(http_url(opt["ip"], port, https))


def update_admin_label():
    adm = settings.get("admin")
    adm_var.set(adm["name"] if adm else "(not set up yet)")


def update_protection_text():
    if live_access["public"]:
        pub_lbl.config(text="ALLOWED", fg=GREEN)
        pub_btn.config(text="Block")
    else:
        pub_lbl.config(text="BLOCKED", fg=RED)
        pub_btn.config(text="Allow")
    n_users = len(live_access["users"])
    n_ips = len(live_access["allowed"])
    parts = ["HTTPS" if https_var.get() else "HTTP (unencrypted)",
             f"{n_users} account{'s' if n_users != 1 else ''}",
             f"{n_ips} allowed address{'es' if n_ips != 1 else ''}" if n_ips else "any address"]
    prot_var.set("Protection: " + ", ".join(parts))


def update_status():
    state = server_state()
    if state == "running":
        ip, port = _running_cfg[0], _running_cfg[1]
        where = "all interfaces" if ip == "0.0.0.0" else ip
        text, color = f"● Running on {where}:{port}", GREEN
    elif state == "down":
        text, color = "● Not responding", ORANGE
    else:
        text, color = "● Stopped", RED
    status_lbl.config(text=text, fg=color)
    stop_btn.state(["!disabled"] if state != "stopped" else ["disabled"])
    launch_btn.config(text="Open in browser" if state == "running" else "Launch Tool")
    try:
        if tray:
            tray.title = f"{APP_NAME} - {'running' if state == 'running' else 'not running'}"
    except Exception:
        pass


def ago(t):
    d = max(0, int(time.time() - t))
    if d < 5:
        return "just now"
    if d < 60:
        return f"{d}s ago"
    if d < 3600:
        return f"{d // 60} min ago"
    return f"{d // 3600} h ago"


def client_rows():
    """Users logged in right now: [(username, ip, logged-in-at, last-active), ...]"""
    fn = live_access.get("active_clients")
    if _url is None or fn is None:
        return []
    return [(c["user"] + f" ({ROLE_LABEL.get(c.get('role'), 'Viewer')})",
             "this PC" if c.get("local") else c["ip"],
             time.strftime("%H:%M:%S", time.localtime(c["since"])), ago(c["last"]))
            for c in fn()]


_last_rows = None


def update_clients():
    global _last_rows
    rows = client_rows()
    if rows != _last_rows:                  # only redraw when something changed (no flicker)
        _last_rows = rows
        tree.delete(*tree.get_children())
        for r in rows:
            tree.insert("", "end", values=r)
    clients_var.set(f"Logged-in users: {len(rows)}")


def tick():
    try:
        update_status()
        update_clients()
    except Exception:
        traceback.print_exc()
    root.after(2000, tick)


combo.bind("<<ComboboxSelected>>", update_share_text)
port_var.trace_add("write", update_share_text)

# ============================ SECURITY TAB =================================
sec.columnconfigure(0, weight=1)


def open_audit():
    try:
        if not os.path.exists(AUDIT_FILE):
            open(AUDIT_FILE, "a", encoding="utf-8").close()
        os.startfile(AUDIT_FILE)                 # Windows: opens in Notepad
    except AttributeError:
        messagebox.showinfo(APP_NAME, f"Audit log file:\n{AUDIT_FILE}")
    except Exception as exc:
        messagebox.showerror(APP_NAME, f"Could not open the audit log:\n{exc}")


ttk.Label(sec, text="Admin account", font=("Segoe UI", UI_FONT_SIZE, "bold")).grid(row=0, column=0, sticky="w")
adm_row = ttk.Frame(sec)
adm_row.grid(row=1, column=0, sticky="ew", pady=(2, 0))
ttk.Label(adm_row, textvariable=adm_var).pack(side="left")
ttk.Button(adm_row, text="Audit log", width=9, command=open_audit).pack(side="right")
ttk.Button(adm_row, text="Change...", width=9, command=lambda: admin_dialog(False)).pack(side="right", padx=(0, 4))

ttk.Separator(sec).grid(row=2, column=0, sticky="ew", pady=3)

ttk.Label(sec, text="Allowed client IPs / subnets", font=("Segoe UI", UI_FONT_SIZE, "bold")).grid(row=3, column=0, sticky="w")
ttk.Label(sec, text="Other PCs only, one per line, e.g. 192.168.1.0/24. Empty = any address.",
          foreground="gray").grid(row=4, column=0, sticky="w")
allowed_txt = tk.Text(sec, width=40, height=2)
allowed_txt.grid(row=5, column=0, sticky="ew", pady=(2, 0))

ttk.Separator(sec).grid(row=6, column=0, sticky="ew", pady=3)

sec_users = {}                       # working copy of the user accounts, applied with "Save"
ttk.Label(sec, text="Other user accounts", font=("Segoe UI", UI_FONT_SIZE, "bold")).grid(row=7, column=0, sticky="w")
ttk.Label(sec, text="Administrator: everything. Operator: everything except stopping monitoring. "
                    "Viewer: read-only.",
          foreground="gray", wraplength=360, justify="left").grid(row=8, column=0, sticky="w")

lrow = ttk.Frame(sec)
lrow.grid(row=9, column=0, sticky="ew", pady=(2, 0))
user_list = tk.Listbox(lrow, height=2, exportselection=False)
user_list.pack(side="left", fill="x", expand=True)


def user_role_text(rec):
    role = rec.get("role")
    return ROLE_LABEL[role] if role in ROLE_LABEL else "Viewer (no role set - choose one)"


def refresh_user_list():
    user_list.delete(0, "end")
    for name in sorted(sec_users):
        user_list.insert("end", f"{name}  -  {user_role_text(sec_users[name])}")


def remove_user():
    sel = user_list.curselection()
    if not sel:
        return
    sec_users.pop(sorted(sec_users)[sel[0]], None)
    refresh_user_list()


ttk.Button(lrow, text="Remove", width=8, command=remove_user).pack(side="left", padx=(6, 0), anchor="n")

form = ttk.Frame(sec)
form.grid(row=10, column=0, sticky="ew", pady=(6, 0))
form.columnconfigure(1, weight=1)
ttk.Label(form, text="Username").grid(row=0, column=0, sticky="w")
ttk.Label(form, text="Password").grid(row=1, column=0, sticky="w")
ttk.Label(form, text="Confirm").grid(row=2, column=0, sticky="w")
ttk.Label(form, text="Role").grid(row=3, column=0, sticky="w")
u_ent = ttk.Entry(form)
p1_ent = ttk.Entry(form, show="*")
p2_ent = ttk.Entry(form, show="*")
u_ent.grid(row=0, column=1, sticky="ew", padx=(8, 6), pady=1)
p1_ent.grid(row=1, column=1, sticky="ew", padx=(8, 6), pady=1)
p2_ent.grid(row=2, column=1, sticky="ew", padx=(8, 6), pady=1)
role_var = tk.StringVar(value=ROLE_LABEL[DEFAULT_ROLE])
role_box = ttk.Combobox(form, textvariable=role_var, values=[label for _k, label in USER_ROLES],
                        state="readonly", width=16)
role_box.grid(row=3, column=1, sticky="w", padx=(8, 6), pady=1)
make_show_check(form, [p1_ent, p2_ent]).grid(row=4, column=1, sticky="w", padx=(8, 0))


def on_user_select(_e=None):
    sel = user_list.curselection()
    if sel:
        name = sorted(sec_users)[sel[0]]
        u_ent.delete(0, "end")
        u_ent.insert(0, name)
        role_var.set(ROLE_LABEL.get(sec_users[name].get("role"), ROLE_LABEL[DEFAULT_ROLE]))


user_list.bind("<<ListboxSelect>>", on_user_select)


def add_user():
    name = u_ent.get().strip().lower()
    adm = settings.get("admin")
    if not USERNAME_RE.match(name):
        messagebox.showerror("Security", "Username: 1-32 characters, letters, digits, '.', '_' or '-'.")
        return
    if adm and name == adm["name"]:
        messagebox.showerror("Security", "That name is the admin account. Change it with 'Change...' above.")
        return
    role = next((k for k, label in USER_ROLES if label == role_var.get()), DEFAULT_ROLE)
    if name in sec_users and not (p1_ent.get() or p2_ent.get()):
        sec_users[name]["role"] = role               # existing account, password left empty: only the role changes
    else:
        if p1_ent.get() != p2_ent.get():
            messagebox.showerror("Security", "The passwords do not match.")
            return
        if len(p1_ent.get()) < MIN_PW:
            messagebox.showerror("Security", f"Use a password of at least {MIN_PW} characters.")
            return
        sec_users[name] = {**make_record(p1_ent.get()), "role": role}
    for e in (u_ent, p1_ent, p2_ent):
        e.delete(0, "end")
    role_var.set(ROLE_LABEL[DEFAULT_ROLE])
    refresh_user_list()


ttk.Button(form, text="Add /\nchange", width=8, command=add_user).grid(row=0, column=2, rowspan=4, sticky="ns")

save_msg = tk.StringVar()


def load_security_fields():
    role_var.set(ROLE_LABEL[DEFAULT_ROLE])
    allowed_txt.delete("1.0", "end")
    allowed_txt.insert("1.0", "\n".join(settings.get("allowed", [])))
    sec_users.clear()
    sec_users.update(settings.get("users", {}))
    for e in (u_ent, p1_ent, p2_ent):
        e.delete(0, "end")
    refresh_user_list()


def save_security():
    entries = [ln.strip() for ln in allowed_txt.get("1.0", "end").replace(",", "\n").splitlines() if ln.strip()]
    for item in entries:
        try:
            ipaddress.ip_network(item, strict=False)
        except ValueError:
            messagebox.showerror("Security", f"'{item}' is not a valid IP address or subnet.")
            return
    if u_ent.get().strip() or p1_ent.get():
        if not messagebox.askyesno("Security",
                                   "You typed a username/password but didn't press\n"
                                   "'Add / change'. Save without adding it?"):
            return
    old_users = set(settings.get("users", {}))
    old_roles = {n: r.get("role") for n, r in settings.get("users", {}).items()}
    settings["allowed"] = entries
    settings["users"] = dict(sec_users)
    save_settings()
    sync_live_access()          # applied immediately, no restart
    added, removed = sorted(set(sec_users) - old_users), sorted(old_users - set(sec_users))
    new_roles = {n: r.get("role", DEFAULT_ROLE) for n, r in sec_users.items()}
    changed = [n for n in sorted(sec_users) if n in old_roles and old_roles[n] != new_roles[n]]
    note = f"security settings saved: {len(entries)} allowed address(es)"
    if added:
        note += ", users added: " + ", ".join(n + " (" + new_roles[n] + ")" for n in added)
    if changed:
        note += ", roles changed: " + ", ".join(n + " " + str(old_roles[n] or "none") + " -> " + new_roles[n] for n in changed)
    if removed:
        note += ", users removed: " + ", ".join(removed)
    audit_event(note)
    update_protection_text()
    update_clients()
    save_msg.set("Saved - applied now.")
    root.after(3000, lambda: save_msg.set(""))


sbtn = ttk.Frame(sec)
sbtn.grid(row=11, column=0, sticky="ew", pady=(8, 0))
ttk.Label(sbtn, textvariable=prot_var, foreground="gray").pack(side="left")
ttk.Button(sbtn, text="Save", width=7, command=save_security).pack(side="right")
ttk.Label(sec, textvariable=save_msg, foreground=GREEN).grid(row=12, column=0, sticky="e")

# ============================ NOTIFICATIONS TAB ============================
# E-mail copies of the dashboard notifications (set up for Gmail by default).
# Settings live in settings.json (the password is stored ENCRYPTED, never as text) and are
# pushed to the running tool through live_access["email"], so a change applies at once.
# EMAIL_KINDS must match the categories Manager._notify sends in ios_xe_upgrade_manager.py.
EMAIL_KINDS = [        # (key, label shown here, ticked by default)
    ("started", "Started", False),
    ("completed", "Completed", True),
    ("failed", "Failed", True),
    ("attention", "Needs attention", True),
    ("postcheck", "Post-check result", True),
    ("missed", "Schedule missed", True),
    ("login", "Sign-ins", False),
]
EMAIL_PROBLEM_KINDS = ("failed", "attention", "missed")
EMAIL_DEFAULTS = {"enabled": False, "host": "smtp.gmail.com", "port": 587, "security": "starttls",
                  "username": "", "password_enc": "", "to": "",
                  "kinds": [k for k, _label, on in EMAIL_KINDS if on]}
EMAIL_SECURITY = {"starttls": "STARTTLS (port 587)", "ssl": "SSL (port 465)"}
EMAIL_ADDR_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def email_settings():
    cfg = dict(EMAIL_DEFAULTS)
    cfg.update(settings.get("email") or {})
    return cfg


def sync_email_access():
    cfg = email_settings()
    cfg["from"] = cfg.get("username", "")
    live_access["email"] = cfg          # read by the running tool on every notification


def _backend():
    """The tool module (a slow import, so only loaded when a button really needs it)."""
    import ios_xe_upgrade_manager as backend
    return backend


nt.columnconfigure(0, weight=1)
em_enabled = tk.BooleanVar()
em_host, em_port, em_sec = tk.StringVar(), tk.StringVar(), tk.StringVar()
em_user, em_to, em_msg = tk.StringVar(), tk.StringVar(), tk.StringVar()
em_kinds = {key: tk.BooleanVar() for key, _label, _on in EMAIL_KINDS}
em_pw_state = tk.StringVar()
em_widgets = []                         # everything that is greyed out while e-mail is switched off


def update_email_state():
    flag = "!disabled" if em_enabled.get() else "disabled"
    for w in em_widgets:
        w.state([flag])


ttk.Checkbutton(nt, text="Send e-mail notifications", variable=em_enabled,
                command=update_email_state).grid(row=0, column=0, sticky="w")
ttk.Label(nt, foreground="gray", wraplength=360, justify="left",
          text="Gmail: smtp.gmail.com, port 587 (STARTTLS) and an App password "
               "(2-Step Verification > App passwords). Your normal password will not work."
          ).grid(row=1, column=0, sticky="w", pady=(0, 4))

em_form = ttk.Frame(nt)
em_form.grid(row=2, column=0, sticky="ew")
em_form.columnconfigure(1, weight=1)
for i, text in enumerate(("SMTP server", "Port / security", "Gmail address", "App password", "Send to")):
    ttk.Label(em_form, text=text).grid(row=i, column=0, sticky="w", pady=1)
e_host = ttk.Entry(em_form, textvariable=em_host)
e_host.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=1)
ps_row = ttk.Frame(em_form)
ps_row.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=1)
e_port = ttk.Entry(ps_row, textvariable=em_port, width=6)
e_port.pack(side="left")
c_sec = ttk.Combobox(ps_row, textvariable=em_sec, values=list(EMAIL_SECURITY.values()), state="readonly", width=20)
c_sec.pack(side="left", padx=(6, 0))
e_user = ttk.Entry(em_form, textvariable=em_user)
e_user.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=1)
pw_row = ttk.Frame(em_form)
pw_row.grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=1)
e_pw = ttk.Entry(pw_row, show="*")
e_pw.pack(side="left", fill="x", expand=True)
c_show = make_show_check(pw_row, [e_pw])
c_show.pack(side="left", padx=(6, 0))
e_to = ttk.Entry(em_form, textvariable=em_to)
e_to.grid(row=4, column=1, sticky="ew", padx=(8, 0), pady=1)
ttk.Label(em_form, textvariable=em_pw_state, foreground="gray").grid(row=5, column=1, sticky="w", padx=(8, 0))


def on_security_change(_e=None):
    """Keep the port in step with the security choice (only if it is still a standard port)."""
    if em_port.get().strip() in ("", "587", "465"):
        em_port.set("465" if em_sec.get() == EMAIL_SECURITY["ssl"] else "587")


c_sec.bind("<<ComboboxSelected>>", on_security_change)

ttk.Separator(nt).grid(row=3, column=0, sticky="ew", pady=3)
ttk.Label(nt, text="E-mail me about", font=("Segoe UI", UI_FONT_SIZE, "bold")).grid(row=4, column=0, sticky="w")
em_kind_frame = ttk.Frame(nt)
em_kind_frame.grid(row=5, column=0, sticky="ew", pady=(2, 0))
for i, (key, label, _on) in enumerate(EMAIL_KINDS):
    cb = ttk.Checkbutton(em_kind_frame, text=label, variable=em_kinds[key])
    cb.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 12), pady=0)
    em_widgets.append(cb)


def set_kinds(keys):
    for key, var in em_kinds.items():
        var.set(key in keys)


quick = ttk.Frame(nt)
quick.grid(row=6, column=0, sticky="w", pady=(3, 0))
b_all = ttk.Button(quick, text="Select all", command=lambda: set_kinds([k for k, _l, _o in EMAIL_KINDS]))
b_prob = ttk.Button(quick, text="Problems only", command=lambda: set_kinds(EMAIL_PROBLEM_KINDS))
b_none = ttk.Button(quick, text="Clear", command=lambda: set_kinds(()))
for b in (b_all, b_prob, b_none):
    b.pack(side="left", padx=(0, 4))


def collect_email(strict):
    """(settings dict, problems). strict=True also checks that the form is complete."""
    sec_key = next((k for k, v in EMAIL_SECURITY.items() if v == em_sec.get()), "starttls")
    kinds = [k for k, _label, _on in EMAIL_KINDS if em_kinds[k].get()]
    cfg = {"enabled": bool(em_enabled.get()), "host": em_host.get().strip(), "port": em_port.get().strip(),
           "security": sec_key, "username": em_user.get().strip(), "to": em_to.get().strip(), "kinds": kinds}
    problems = []
    if strict:
        if not cfg["host"]:
            problems.append("Enter the SMTP server (for Gmail: smtp.gmail.com).")
        try:
            port = int(cfg["port"])
            if not 1 <= port <= 65535:
                raise ValueError
            cfg["port"] = port
        except ValueError:
            problems.append("The port must be a number (587 for STARTTLS, 465 for SSL).")
        if not EMAIL_ADDR_RE.match(cfg["username"]):
            problems.append("Enter the e-mail address used to sign in to the mail server.")
        recipients = [r for r in re.split(r"[,;\s]+", cfg["to"]) if r]
        bad = [r for r in recipients if not EMAIL_ADDR_RE.match(r)]
        if not recipients:
            problems.append("Enter at least one address under 'Send to'.")
        elif bad:
            problems.append("Not a valid e-mail address: " + ", ".join(bad[:3]))
        if not e_pw.get() and not (settings.get("email") or {}).get("password_enc"):
            problems.append("Enter the App password.")
    return cfg, problems


def update_pw_state():
    saved = bool((settings.get("email") or {}).get("password_enc"))
    em_pw_state.set(("Password saved - leave empty to keep it.  " if saved else "") + "Several recipients: use commas.")


def save_email():
    cfg, problems = collect_email(strict=em_enabled.get())
    if problems:
        messagebox.showerror("Notifications", "\n".join(problems))
        return
    if cfg["enabled"] and not cfg["kinds"] and not messagebox.askyesno(
            "Notifications", "No notification type is ticked, so no e-mail would ever be sent.\n\nSave anyway?"):
        return
    cfg["password_enc"] = (settings.get("email") or {}).get("password_enc", "")
    if e_pw.get():
        try:
            cfg["password_enc"] = _backend().protect_secret(e_pw.get())
        except (Exception, SystemExit) as exc:
            messagebox.showerror("Notifications", f"Could not encrypt the password:\n{exc}")
            return
    was_on = bool((settings.get("email") or {}).get("enabled"))
    settings["email"] = cfg
    save_settings()
    sync_email_access()                                  # applied immediately, no restart
    e_pw.delete(0, "end")
    update_pw_state()
    audit_event("e-mail notification settings saved: " + (
        f"ON ({len(cfg['kinds'])} type(s), {cfg['host']})" if cfg["enabled"]
        else ("switched OFF" if was_on else "OFF")))
    em_msg_lbl.config(foreground=GREEN)
    em_msg.set("Saved - applied now." if cfg["enabled"] else "Saved - e-mail notifications are OFF.")
    root.after(4000, lambda: em_msg.set(""))


em_test_result = []


def test_email():
    cfg, problems = collect_email(strict=True)
    if problems:
        messagebox.showerror("Notifications", "\n".join(problems))
        return
    cfg["from"] = cfg["username"]
    if e_pw.get():
        cfg["password"] = e_pw.get()                     # only used for this test, not stored
    else:
        cfg["password_enc"] = (settings.get("email") or {}).get("password_enc", "")
    b_test.state(["disabled"])
    em_msg_lbl.config(foreground="gray")
    em_msg.set("Sending test e-mail ...")
    del em_test_result[:]

    def work():
        try:
            n = _backend().send_test_email(cfg)
            em_test_result.append((True, f"Test e-mail sent to {n} address(es). Check the inbox."))
        except (Exception, SystemExit) as exc:
            try:
                reason = _backend().describe_email_error(exc)
            except Exception:
                reason = type(exc).__name__
            em_test_result.append((False, reason))

    def poll():
        if not em_test_result:
            root.after(300, poll)
            return
        ok, text = em_test_result[0]
        em_msg_lbl.config(foreground=GREEN if ok else RED)
        em_msg.set(text)
        b_test.state(["!disabled"])
        update_email_state()

    threading.Thread(target=work, daemon=True).start()
    root.after(300, poll)


em_btns = ttk.Frame(nt)
em_btns.grid(row=8, column=0, sticky="ew", pady=(5, 0))
b_test = ttk.Button(em_btns, text="Send test e-mail", command=test_email)
b_test.pack(side="left")
ttk.Button(em_btns, text="Save", width=7, command=save_email).pack(side="right")
em_msg_lbl = ttk.Label(nt, textvariable=em_msg, foreground=GREEN, wraplength=360, justify="left")
em_msg_lbl.grid(row=9, column=0, sticky="w", pady=(2, 0))
em_last = tk.StringVar()
em_last_lbl = ttk.Label(nt, textvariable=em_last, foreground="gray", wraplength=360, justify="left")
em_last_lbl.grid(row=10, column=0, sticky="w")
em_widgets += [e_host, e_port, c_sec, e_user, e_pw, c_show, e_to, b_all, b_prob, b_none, b_test]


def load_email_fields():
    cfg = email_settings()
    em_enabled.set(bool(cfg.get("enabled")))
    em_host.set(cfg.get("host") or "")
    em_port.set(str(cfg.get("port") or ""))
    em_sec.set(EMAIL_SECURITY.get(cfg.get("security"), EMAIL_SECURITY["starttls"]))
    em_user.set(cfg.get("username") or "")
    em_to.set(cfg.get("to") or "")
    set_kinds(cfg.get("kinds") or [])
    e_pw.delete(0, "end")
    update_pw_state()
    update_email_state()


def poll_email_status():
    """Show the result of the latest real notification e-mail (sent / why it failed)."""
    try:
        fn = live_access.get("email_last")
        last = fn() if callable(fn) else None
        if last:
            em_last_lbl.config(foreground=GREEN if last.get("ok") else RED)
            em_last.set(("Last e-mail %s: " % last.get("time", "")) + str(last.get("detail", "")))
        else:
            em_last.set("")
    except Exception:
        pass
    root.after(3000, poll_email_status)


# ---- start-up -------------------------------------------------------------
# Restore the saved interface (falls back to "This PC only" if it's gone)
saved = next((o for o in options if o["key"] == settings.get("interface")), options[0])
combo.set(saved["label"])
load_security_fields()
sync_email_access()
load_email_fields()
update_share_text()
update_admin_label()
update_protection_text()
update_status()

root.bind("<Unmap>", on_minimize)
root.protocol("WM_DELETE_WINDOW", quit_app)

make_tray()
root.after(2000, tick)
root.after(3000, poll_email_status)
root.after(400, startup_admin)          # first start: ask for the admin account
root.mainloop()
