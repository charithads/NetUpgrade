#!/usr/bin/env python3
"""
NetUpgrade v5.10.1 - monitor, back up and upgrade Cisco Catalyst 9200 switches from one
web page, with the firmware-upgrade worker built into the same file/program.

What it does
  * Availability   - pings every switch from the Excel list and shows live UP/DOWN status.
  * Backup         - select switches in the GUI, click "Run backup": saves show running-config,
                     show version, show boot, show inventory and any extra commands typed in the
                     GUI - all outputs together in ONE file per switch - then shows COMPLETED and the
                     current software version in the table.
  * Install mode   - every upgrade uses IOS-XE INSTALL mode, following Cisco's "Upgrade Catalyst 9300 /
                     9200 Switches" guides: boot variable checked (and set to flash:packages.conf when it
                     is wrong), 1-1.5 GB flash expansion space checked on every stack member, and the
                     switch's platform + version + boot mode come from a backup or "Refresh version".
  * Firmware       - select switches, enter the IOS-XE image (.bin) path and either click
    upgrade          "Start now" or "Schedule" for a chosen date/time. A live web console
                     opens in the dashboard for each switch (immediately, or when its time arrives) and
                     runs this same program again in "--job" (worker) mode for that one switch.
                     Scheduled switches show their scheduled time in the table until they start.
                     The GUI follows the progress, the reboot (ping goes down, comes back) and
                     finally shows the NEW software version. Any number of switches can upgrade
                     at the same time.
  * CLI upgrade    - "--cli" runs the original interactive, whole-list upgrade flow with no GUI
                     (asks for credentials once, confirms, then upgrades switches one at a time by
                     default; --parallel N upgrades N at once in waves, and the rollout stops after a
                     wave with a failure unless --continue-on-failure is added). Independent of the GUI.

This file is also the single source used to build the one-click Windows .exe (see
build_exe.bat / BUILD.md) - "--job" is how the .exe re-launches itself as a per-switch worker,
so no second script needs to ship alongside it.

Install (running from source):
    pip install netmiko openpyxl

Usage:
    python ios_xe_upgrade_manager.py switches.xlsx
    python ios_xe_upgrade_manager.py switches.xlsx --mode tcp --port 22   # if ICMP is blocked
    python ios_xe_upgrade_manager.py switches.xlsx --duration 30          # stop monitoring after 30 min
    python ios_xe_upgrade_manager.py switches.xlsx --web-port 9000 --open
    python ios_xe_upgrade_manager.py --cli                 # no GUI; prompts for image/creds/xlsx

Packaged as IOS-XE-Upgrade-Manager.exe: double-click it with a switches.xlsx file in the same
folder (no command line needed) - see BUILD.md to build that .exe from this file.

Then open http://127.0.0.1:8080

Excel format (first sheet by default): a header row with "ip" (required), and optionally
"hostname" (label), "platform" (e.g. C9200L-48P-4G) and "type" (only 'cisco-xe' switches can be
backed up or upgraded).

Security notes
  * The dashboard can back up and upgrade switches, so it only listens on 127.0.0.1 by default.
    Remote binding is accepted only through the desktop launcher with configured users and HTTPS.
  * Credentials are typed into the GUI and sent only to this program. They are never written to
    logs. The upgrade worker receives them through a temporary file (owner-only permissions)
    that is deleted as soon as it has been read.
"""

# ---- Standard library ----
import argparse
import csv
import difflib
import functools
import getpass
import hashlib
import hmac
import io
import ipaddress
import json
import os
import platform
import queue
import re
import secrets
import shutil
import smtplib
import socket
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import warnings
import webbrowser
import zipfile
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from email.message import EmailMessage
from html import escape as html_escape
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from pprint import pprint
from urllib.parse import urlsplit, urlparse, parse_qs, quote

APP_NAME = "NetUpgrade"
APP_VERSION = "5.13.0"          # single place to change the release number (shown in the GUI, --version, reports)
APP_AUTHOR = "Charitha Dasanayaka"
APP_REPO_URL = "https://github.com/charithads/NetUpgrade"   # shown in the footer

# What each release added, newest first - shown under Settings > About. Add one entry here for every release
# (keep APP_VERSION above and APP_VERSION in launcher.py equal to the newest entry).
CHANGELOG = [
    {"version": "5.13.0", "date": "2026-10-25", "features": [
        "Choose separate folders during installation for IOS images and switch configuration backups." ]},
    {"version": "5.12.1", "date": "2026-10-25", "features": [
        "About page under Settings with the application details and the new features of every version.",
        "The Windows application has an About tab with a Check for new updates button; new releases are downloaded from the GitHub repository.",
        "Custom Checks removed from Settings (the built-in pre-checks and post-checks are unchanged).",
        "The 'Before you upgrade' section removed from the Overview tab of the device window."]},
    {"version": "5.12.0", "date": "2026-10-23", "features": [
        "Download template (.xlsx) in the device-list upload section.",
        "Notifications settings moved from the Windows application into Settings > Notifications of this dashboard."]},
    {"version": "5.11.0", "date": "2026-10-23", "features": [
        "One guided Upgrade button, with distribute and activate under Advanced.",
        "Timeout presets (Standard / Slow switch / Stack) with an automatic suggestion.",
        "CSV / Excel device import with a template and credential validation.",
        "A red readiness result blocks the upgrade.",
        "Maintenance window with stop-on-first-failure.",
        "Plain next-step guidance after a failed upgrade.",
        "Batch summary with an automatic report.",
        "Teams / Slack alerts."]},
    {"version": "5.10.1 and earlier", "date": "", "features": [
        "Availability monitoring, configuration backup and scheduled firmware upgrades for Catalyst 9200 / 9300 switches from one dashboard.",
        "Install-mode upgrades following Cisco's guides, with a boot-variable check and a flash-space check that includes the expansion room.",
        "Upgrade output in the Console tab of the device drawer; image stage view and repeat-distribute warning.",
        "Settings page with editable wait times and Command duration profiles; Storage page with MD5 image verification.",
        "Device-list search, filters and a column chooser; light and dark mode; read-only public access option."]},
]
APP_ICON_DATA_URI = (
    "data:image/png;base64,"
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAA"
    "DsMAAA7DAcdvqGQAAAxOSURBVGhD7ZcJUBVX2oZPd9+NxWCMSyaKgM6POol/4oYIAgqCKyKiBidBFgnKoqOJCqi4oMgmiyKi"
    "IIssymXfBBXuBRVQo1ERFxyDMm5o3FCG9fZ3vqnLmJS2UzOZ0X9S9RdP1Vu3v/O+p/vr7urTfQnppZdeevn/RTUdztW2L2TP"
    "tgQwF37aw9XdiRNfvR0ivXFjmUbjFSNyJUsinPLbk3xbxhbzbuzxrhOssq2LPY/IXkPkbqqQu92G4rvPUHrvPspuXkHNH8/d"
    "1Gk6E6zTdMxAuJvfBDa1y5GVqxrYo4isApE9iciepsjWvES2uhnY7+/w3OUmFDfeRUlTI2o21eMHj+qwzy1lq85fjm0lWCkS"
    "7vO/w+ZH2tx+SOMOI3JyRDYfkc1p/StT1JrHlrZ4iSqemZGqB6NI9b0RorM3J7EXGpaKLtcdll79/pnW3Quo3ViFfR/XYN+m"
    "ktM6l+T/5buxuXUgF8WfEyUicgcQucT2bia1LYLIO/51I9d/+ER67dR2jasVnX0aK7DvveOo05h7r1/9wVHC6P8NobQPt42/"
    "II5G5HYjstEdl0jMX8cKY2okF1SzZBeoiXBcjfhc2VjNq0eu9mkswz63C/CDG4dua11KHSjMvXfE67ryJDsQxSGIou1tCrL2"
    "cR9hRo24il8krUeU1UO77Ey3mdBXo12Z0V/jUu5FrRsF2OdOAWpfSj4izLxXxCvbnSXrEcUbEbn1bZfJd81awowacT7vJDuJ"
    "KKtGlJ1HlJ7j26RVHdbCnBqNyvghsh/kD7Xr5ah9Kwu1zscvEmbeD16oLVre9kD8LaJodXsXWfn0M2FEjTSF95AWI0pLEWVH"
    "VSCr6OZlNYiSar5LUvZirjCvRlKVZq9Rl4Na9RmocWbfdVK5+f2vTCLXFm+xF6JoBSLn8Txa6KuRRHWvlKYiStMRJZkqkOar"
    "QFasAumRLl5aASg53g3ivJZ/eIUlVSmVGhczUas+HTVrYmcL/XeGW/zTOZFbN3IuzzuJ07OhQl+0qdtXuhtREkNRsk8FkiQV"
    "SNNVIJV3gySni5dkt/PivC4U5Xcje/CJs3C+5FiKvez7TNS8lIHSE7GpQv/dmPdoOLugmeecOpFd3Hz8Dc8ZZcyqrhhRKKI4"
    "7O+SRHTzkj0qXrpfxUsSu3nxwU5enNLGc+kq5LLV7wtENqbZn2Qh98t+jqVqiY8mPZGdyUCpMvbPZPNm9o3jvAvs7NuO7IIX"
    "yC5uR3bhfb83TO+2T7iVnXLu2840kX93MrdFVS8K4UEc1sWLo7t4cUwXL47r4EWJncDue1HN7Hl6gNnz9CAT/SCBhL+5CHBl"
    "yVWS2iyUKvZ3yY7u1O0ZRGRez/xHMNMbtzDzXiBj/xTJvL/YCf3X4Xw7tnNBiKKtbSAKbufFYe28KKqNZxPU74zWr4V59ZXm"
    "Dh3IYovS67mClKeio+kgKU+ishP7ftSo3nVJ82zYx8Ip/zbMtJuxjO1zZGybkcy9ZS70X4dZ2RLFrkfk/F4Ct/ElcFtbeXbH"
    "S2BjENngFk9hXg2XlmjPluYiW3AIRUXpID6SDLIreSipio0hSN7DHZjaEMfMfIzM7IdIZjZZCv3XYb55vJtZBciseAbsmuc8"
    "69/CswEtPBtCkWx67C3M/wyTtH83W5yLbHYKz5XLUVSWdJ1kva9Pb4srQYzNI2RmPUUy49ZbK8gbOD2IYZZ1IuP+CBjvn3hm"
    "1ROe+fYJMJtUSHyfrBDGfyEiQoNJSbrJ5B5GtugwkpxkU2HkP8e0fglj9QCJzRMk1rf3C+03WHBnL7OkFRnn+zzj/oBnlj3k"
    "Gc+HwKztRLLywWph/HW4pHhLtqwEmfTkWKH3bpjUjSCm14GY3ENi2viIkP1iYeQX7Br3ky+fIvNlE8843eWJy32eWXqfZ1a0"
    "IVl+7zth/C0Sop1Jxt4PhcPvjKZ1w8lRvq23RmzpLphSSJ3n38Bp9teotf1laj3vIrWZc5XOtG2g/8vYNMQRu4fIzL3JMw63"
    "eGZxE898fYcn7i+RuN5ZN6uEmlsX4kJLuWqBWlY5KgerHOpgLacLrPNwoWM5HSM89jvzVTOdveBY63W7hNNd0yIr26ZEnkHz"
    "yB/QTK2d59E0/Dya7b6Kk0Mrc8nkhjgypwXJ7DtI7O4hs+ABMo6PkLh0I/lj02bTjaVnTLefx0mbT6HxplNoHHASJ26swokb"
    "KtFovRKNAmrQaNej2rG59N0/J2Y+puMtj3cXm0bW4qeOYThsyir4vdUa+j82vmBo40tHTPelhtP9euqRswLwD9NXFxGLZ07E"
    "+mkxsb5bQGbdLyRzmwvI/OZC4theRBY/dxhps/r7ETP80dBmHW9osw4MbdbA761Ww3DLP8HwqStgmIUXNZwXjJ9+dwo/i2ov"
    "+zSH/vsPs9EZavB5aXfC6OAqargwCAeP/Qp1jZeCnpnnK3nz+uZeYGDhDfpqmXv1SM/Cp9PA3CN1uLl7pIGFR7SBhXu0gZl7"
    "9HD1r7lblL65R46+uSfomS+neuaeoG/hBfoWnuq5vJ7ZMhhq5skPNVsGQya5gO4ERxw2ZzMO81Hg0OC2DL1Y+uv+sf2uUGWt"
    "e4R/oZvxHAf5leDHfqXwuwAFHRxQQYcEKGHwRgUM3qAA3Y1K0N2gBF1/JQz1V4KevwL0/JRosLEGh22oRQP/Whymll8tDldr"
    "XQ0OW1uN+qsUVH+FAvRXKEHP55W8FVTPWwF6XgoY6qUEXU9FjwZ/UwqD3QpxyJaHODCSb+8bSf/pV0APmrl8wQCPBNQa8yVw"
    "E1xAZuIOkoluwI13AbGRK2hMcgfReBfgxjmDzHgpyIzd/+5NcAFNk29AMmEpcONcQTLBDTSNvwHRWJceaRq5g3S8G3BjXEDy"
    "qhaNcQXuc2eQjXPrEff5EtA2XgYDp66Ej0x8oL+ZD3xk6gMfTVgC/edHo3ZwZ4Ww37dgMzoKpeYrcb6LLxQePQmxybk0RV5K"
    "c0oqQV6ogF3xcpAXKUBdJ2QU032phTSv9ARkFSlh14FsyCxUQnZxJaRkldGYpFzIOVIFOSVVao8ezD4K2SWVkJFfAVHxcppZ"
    "pOyp96cX0bi0QsgvOwFzXNeDVN+aHzTJDvp9YQsDje2g/5iZoGPqjbLA5+XCft+CSe/IJxOW4YG0Aqg6XQejpy2lc902QN21"
    "RtgQngyGpl/TkL2Z9OzFBmqxcDU1svWh6lzYvixqONmJ+ock0svXb8F8jy30D5aukFdWTeUlJ+jIKS50kWcgPX/5z9RnUyw1"
    "NHWioXsz4fSF62DmsAomzFpOL1z5ESLic4F8NBkmn1gKHzs5wPhDzqC/3gE0RnuhKPDFMWG/b8GkdOST8ctwb1IO2LlvoaSf"
    "CSUfToRVW+OgzyhbSj6ZSj8Zv4i6rdlJyQdGQD6cRG3dNtHBRot7PO2RtvTbbfGU9J1IyYDJdJL9n6jRXB9KBppR0t8Elm+I"
    "oYyeNSVDrOjH4xZSp9XhlOhMpERrPPUKiIWdCXmUDDCjg+bYg864OdDf2g76Gs8ATSMfZLf+mhNI7Mgn4zxwb3I+LSw/CwO/"
    "cOhpQH2lPPx30X6jbOna4ERQ1tbByKluVN/kj7S44ixdG5zU43n476anLzZQ0/mr6IAvHGhidgXEZZTRfqPt6VTHtfTE2at0"
    "kfeOnqxfWApU1NSB4RRXqm+8mFafvwah+/KA9Dejpgp3GPTVfPg89WsYsm4+lX3mg+yWX3ECJL6jRDTGFR1d1tCTJ89AxJ40"
    "mphWQI9X1EBWfjkNjUyiuUUKelxRA3sTs2lMfCYtr6yFnCIFDY1KplkF5bSsohoSMwpp+J40eqS8GkqOnYLQmIM05XAxPaao"
    "gYzcozQkKonmlSh76j2J2XR3gpxWnTxN7ZaspaKBxrS/tQ39YLQV/dB8GuiMNQetL9yQ2dhSJWz3bZIhUbamEtkZ25BM34ES"
    "u0jkbMOR2OxAMiMUJXMjkNiEILEKQtHMMBTP2onEcgeSacEoUW9bBSOZGoSsdSiKZ4QjmRKEZMoOlMyIQHZaKBLz7UgsQ3pq"
    "MiW4pxZZh6HIJhyJeRAyVmGoMXsXSi2jUDY9CqVWUSizCkWpaxkyWzszhe2+TSIdQLKpLzlEt5ODdBtJoYEk6ZUO0ECyjwaS"
    "uFeKpYEk5pV20UASIVA4DSShr7SDbuvRNhpIgl7VwTTwF6nH1Pp5++fxELqNhNLtJBL9SBAdJGy3l1566aWXXnr5Lfkba0f5"
    "o98C9T4AAAAASUVORK5CYII="
)

# ---- Third-party (pip install netmiko openpyxl) ----
try:
    import paramiko
    from openpyxl import load_workbook
    from netmiko import ConnectHandler, SCPConn
    from netmiko.exceptions import (NetmikoAuthenticationException, NetmikoBaseException,
                                    NetmikoTimeoutException, ReadTimeout)
except ImportError as exc:  # pragma: no cover
    sys.exit(f"Missing dependency ({exc.name}). Install with:  pip install netmiko openpyxl")


# --------------------------------------------------------------------------- #
# Per-device command log. Every command this program sends to a switch (backups, credential
# tests, readiness checks, version checks, and the upgrade worker's own session) is noted in
# one small file per device: <site>/Runs/commands_<ip>.jsonl. Only the command, who ran it
# (the "source"), when, whether it worked and how long the reply was are kept - NEVER the
# switch's output - so the file holds nothing sensitive. The device drawer in the dashboard
# reads it. It works in both the GUI process and the separate --job worker process.
# --------------------------------------------------------------------------- #
_netmiko_connect = ConnectHandler
COMMAND_LOG_DIR = None                 # set by the Manager (site's Runs folder); the worker uses LOG_FILE's folder
COMMAND_LOG_MAX_BYTES = 400_000        # trimmed to the newest COMMAND_LOG_KEEP_LINES lines beyond this
COMMAND_LOG_KEEP_LINES = 1200
_COMMAND_LOCK = threading.Lock()
_ACTIVITY = threading.local()
RECORDED_METHODS = ("send_command", "send_command_timing", "send_config_set")


class activity(object):
    """Name what the code inside is doing ("Backup", "Readiness check", ...). Use it as
    `with activity("Backup"):` or as a decorator. Commands recorded meanwhile carry that name."""

    def __init__(self, source):
        self.source = source
        self._prev = None

    def __enter__(self):
        self._prev = getattr(_ACTIVITY, "source", None)
        _ACTIVITY.source = self.source
        return self

    def __exit__(self, *exc):
        _ACTIVITY.source = self._prev
        return False

    def __call__(self, fn):
        source = self.source

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            with activity(source):
                return fn(*args, **kwargs)
        return wrapper


def record_command(host, command, ok, seconds, error="", lines=None):
    """Append one line to the device's command log. Never raises: logging must not affect a job."""
    try:
        base = COMMAND_LOG_DIR or (os.path.dirname(LOG_FILE) if LOG_FILE else None)
        if not base or not host:
            return
        entry = {"t": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "src": getattr(_ACTIVITY, "source", None) or "Other",
                 "cmd": str(command)[:200], "ok": bool(ok), "sec": round(float(seconds), 1)}
        if error:
            entry["err"] = str(error)[:160]
        if lines is not None:
            entry["lines"] = int(lines)
        path = os.path.join(str(base), "commands_" + re.sub(r"[^A-Za-z0-9_.-]", "_", str(host)) + ".jsonl")
        with _COMMAND_LOCK:
            os.makedirs(str(base), exist_ok=True)
            with open(path, "a", encoding="utf-8") as fh:
                fh.write(json.dumps(entry) + "\n")
            if os.path.getsize(path) > COMMAND_LOG_MAX_BYTES:
                with open(path, "r", encoding="utf-8", errors="replace") as fh:
                    keep = fh.readlines()[-COMMAND_LOG_KEEP_LINES:]
                with open(path, "w", encoding="utf-8") as fh:
                    fh.writelines(keep)
    except Exception:  # noqa: BLE001
        pass


class RecordedSession(object):
    """Wraps a netmiko session: behaves exactly like it, but notes every command sent."""

    def __init__(self, conn, host):
        self._conn = conn
        self._host = host

    def __getattr__(self, name):
        attr = getattr(self._conn, name)
        if name not in RECORDED_METHODS or not callable(attr):
            return attr

        def recorded(*args, **kwargs):
            command = args[0] if args else (kwargs.get("command_string") or kwargs.get("config_commands") or "")
            if isinstance(command, (list, tuple)):
                command = "; ".join(str(c) for c in command)
            started = time.time()
            try:
                output = attr(*args, **kwargs)
            except Exception as exc:  # noqa: BLE001 - recorded, then raised again unchanged
                record_command(self._host, command, False, time.time() - started, type(exc).__name__)
                raise
            text = output if isinstance(output, str) else ""
            rejected = "Invalid input" in text[:300]
            record_command(self._host, command, not rejected, time.time() - started,
                           "rejected by the switch (Invalid input)" if rejected else "",
                           text.count("\n") + 1 if text else 0)
            return output
        return recorded

    def __enter__(self):
        self._conn.__enter__()
        return self

    def __exit__(self, *exc):
        return self._conn.__exit__(*exc)


# Switches whose credential profile says Telnet: {ip: "telnet"}. Every other switch uses SSH. It is refreshed from the
# profile each time credentials are looked up, so one place (the ConnectHandler below) switches every connection - backup,
# credential test, readiness, topology, commands - between SSH and Telnet.
HOST_PROTOCOL = {}
HOST_PORT = {}      # {ip: port} - only when the credential profile uses a non-default port (SSH 22 / Telnet 23)


def default_port(protocol):
    return 23 if protocol == "telnet" else 22


def clean_port(value, protocol="ssh"):
    """Port typed in the profile form -> int. Empty means the protocol's default. Raises ValueError when invalid."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return default_port(protocol)
    if isinstance(value, bool):
        raise ValueError("Enter a port number between 1 and 65535.")
    try:
        port = int(str(value).strip())
    except (TypeError, ValueError):
        raise ValueError("Enter a port number between 1 and 65535.")
    if not 1 <= port <= 65535:
        raise ValueError("Enter a port number between 1 and 65535.")
    return port


def set_host_protocol(host, protocol, port=None):
    if protocol == "telnet":
        HOST_PROTOCOL[str(host)] = "telnet"
    else:
        HOST_PROTOCOL.pop(str(host), None)
    try:
        port = int(port) if port else 0
    except (TypeError, ValueError):
        port = 0
    if port and port != default_port(protocol):
        HOST_PORT[str(host)] = port
    else:
        HOST_PORT.pop(str(host), None)


def ConnectHandler(*args, **kwargs):  # noqa: N802 - same name on purpose: every call site records its commands
    host = kwargs.get("host") or kwargs.get("ip") or ""
    if HOST_PROTOCOL.get(str(host)) == "telnet":                       # cisco_xe -> cisco_xe_telnet, port 23
        dtype = str(kwargs.get("device_type") or "cisco_xe")
        dtype = dtype[:-4] if dtype.endswith("_ssh") else dtype
        if not dtype.endswith("_telnet"):
            kwargs["device_type"] = dtype + "_telnet"
        kwargs.setdefault("port", 23)
    if HOST_PORT.get(str(host)):                                        # profile asks for a non-default port
        kwargs["port"] = HOST_PORT[str(host)]
    kwargs["conn_timeout"] = kwargs["auth_timeout"] = kwargs["banner_timeout"] = tmo("ssh_connect")
    started = time.time()
    try:
        conn = _netmiko_connect(*args, **kwargs)
    except Exception as exc:  # noqa: BLE001
        record_command(host, "(connect and log in)", False, time.time() - started, type(exc).__name__)
        raise
    record_command(host, "(connect and log in)", True, time.time() - started)
    return RecordedSession(conn, host)


def this_program_path():
    """Path to re-invoke for a worker subprocess: the frozen .exe itself, or this .py file.
    Uses __file__ (this module's own location), not sys.argv[0], so it is correct even if
    this module were ever imported by another script rather than run directly."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(globals().get("__file__") or sys.argv[0]).resolve()


def self_invoke_argv(*extra_args):
    """argv to launch another copy of this same program (worker mode), whether running as a
    plain .py (via the current Python interpreter) or as a frozen one-file .exe."""
    if getattr(sys, "frozen", False):
        return [sys.executable, *extra_args]           # the exe IS the interpreter here
    return [sys.executable, str(this_program_path()), *extra_args]


def _frozen_excepthook(exc_type, exc, tb):
    """Keep a double-clicked .exe's console open on a startup crash, so the message is readable."""
    import traceback as _tb
    _tb.print_exception(exc_type, exc, tb)
    try:
        input("\nThe program hit an error (above) and is closing. Press Enter to exit ...")
    except (EOFError, OSError):
        pass


# =============================================================================
# Part 1 of 2 - firmware upgrade worker (runs standalone with --cli, or as a per-switch
# subprocess with --job; the web GUI below launches the latter automatically)
# =============================================================================

# Errors that are EXPECTED when the switch drops the SSH session because it reloads.
CONNECTION_LOST = (OSError, EOFError, paramiko.SSHException, NetmikoBaseException)

# Operator-tunable waits, in seconds. These are the DEFAULTS; the dashboard's Settings page changes them (saved
# in settings.json) and hands the current values to every upgrade worker in its job file. tmo(key) is always the
# value in force right now.
TIMEOUT_DEFAULTS = {
    'ssh_connect': 20,       # SSH connect / login / banner wait
    'show_cmd': 60,          # show version, show boot, show switch, dir, show install summary
    'show_run': 120,         # backup: show running-config
    'backup_cmd': 60,        # backup: every other command
    'snapshot_cmd': 45,      # pre/post-check snapshot commands
    'health_cmd': 20,        # CPU / memory / stack health commands
    'write_memory': 180,     # saving the configuration
    'install_commit': 300,   # install commit after the reload
    'md5_verify': 900,       # verify /md5 of the image on the switch
    'install_total': 3600,   # the whole install add / activate / reload
    'clean_total': 3600,     # package clean-up
    'reboot_down': 600,      # wait for the switch to stop answering ping
    'reboot_up': 2400,       # wait for it to answer ping again (stacks boot slowly)
    'reboot_ssh': 900,       # wait for SSH to accept a login again
}
TIMEOUTS = dict(TIMEOUT_DEFAULTS)


def tmo(key):
    return TIMEOUTS.get(key, TIMEOUT_DEFAULTS[key])


def apply_timeouts(values):
    for key, default in TIMEOUT_DEFAULTS.items():
        try:
            n = int((values or {}).get(key, default))
        except (TypeError, ValueError, AttributeError):
            n = default
        TIMEOUTS[key] = n if n > 0 else default


INSTALL_POLL = 2            # seconds between reads while the install runs
INSTALL_HEARTBEAT = 60      # print "still waiting" this often when the switch is quiet


# Waiting for the switch to reload after 'install ... activate commit' (--job mode)
REBOOT_POLL = 5             # seconds between pings
REBOOT_SSH_RETRY = 15       # seconds between login attempts

STATUS = None               # StatusFile, set in --job mode (progress shown in the web GUI)
LOG_FILE = None             # optional log file, set in --job mode

_print_lock = threading.Lock()


#------------------------------------------------------------------------------
def log(device, message):
    """Thread-safe, per-device prefixed output (never pass credentials here)."""
    line = '[{}] {}'.format(device.get('name') or device['ipaddr'], message)
    with _print_lock:
        print(line)
        if LOG_FILE:
            try:
                with open(LOG_FILE, 'a', encoding='utf-8') as fh:
                    fh.write('{} {}\n'.format(time.strftime('%Y-%m-%d %H:%M:%S'), line))
            except OSError:
                pass                                   # logging must never stop the upgrade
#------------------------------------------------------------------------------
def read_secret(prompt):
    """Read a hidden value. Abort rather than echo if the terminal can't hide input."""
    with warnings.catch_warnings():
        warnings.simplefilter('error', getpass.GetPassWarning)
        return getpass.getpass(prompt=prompt)
#------------------------------------------------------------------------------
class StatusFile(object):
    """Progress file that the web GUI reads. It never contains credentials.

    Fields: state (running | completed | failed | attention), phase, message, result,
    old_version, new_version, updated (epoch seconds; refreshed by a heartbeat thread so
    the GUI can tell when this window was closed).
    """
    def __init__(self, path):
        self.path = path
        self.data = {}
        self._lock = threading.Lock()

    def update(self, **fields):
        with self._lock:
            self.data.update(fields)
            self.data['updated'] = time.time()
            self._write()

    def _write(self):
        tmp = self.path + '.tmp'
        for _ in range(10):                            # the GUI may be reading the file right now
            try:
                with open(tmp, 'w', encoding='utf-8') as fh:
                    json.dump(self.data, fh)
                os.replace(tmp, self.path)
                return
            except OSError:
                time.sleep(0.1)

    def start_heartbeat(self, every=5):
        def beat():
            while True:
                time.sleep(every)
                self.update()
        threading.Thread(target=beat, daemon=True).start()
#------------------------------------------------------------------------------
PHASE_STATE = {'name': None, 'start': 0.0, 'done': []}     # how long each phase of THIS run took (per worker process)


def report(phase, message, **extra):
    """Publish progress to the web GUI (no-op when the script runs interactively). Also keeps the time spent in
    every phase; the list goes to the GUI as phase_times, where the Command time test turns it into timeouts."""
    if STATUS is not None:
        now = time.time()
        if phase != PHASE_STATE['name']:
            if PHASE_STATE['name']:
                PHASE_STATE['done'].append({'phase': PHASE_STATE['name'],
                                            'seconds': round(now - PHASE_STATE['start'], 1)})
            PHASE_STATE['name'], PHASE_STATE['start'] = phase, now
        STATUS.update(phase=phase, message=message, phase_times=list(PHASE_STATE['done']), **extra)
#------------------------------------------------------------------------------
def parse_version(show_version_output):
    """'Cisco IOS XE Software, Version 17.09.04a' -> '17.09.04a' (None if not found)."""
    match = re.search(r'Cisco IOS XE Software, Version\s+(\S+)', show_version_output)
    if not match:
        match = re.search(r'Version\s+([0-9][^\s,]*)', show_version_output)
    return match.group(1) if match else None
#------------------------------------------------------------------------------
def get_version(net_connect):
    try:
        return parse_version(net_connect.send_command('show version'))
    except Exception:                                  # a version lookup must never stop the upgrade
        return None
#------------------------------------------------------------------------------
def guess_version_from_filename(image_name):
    """Cisco IOS-XE image filenames normally embed the version, e.g.
    'cat9k_lite_iosxe.17.09.04a.SPA.bin' -> '17.09.04a'. Best-effort only: returns
    None if nothing that looks like a version number is found."""
    match = re.search(r'(\d+\.\d+\.\d+[a-z]?)', image_name)
    return match.group(1) if match else None
#------------------------------------------------------------------------------
# Image <-> switch model check. Catalyst 9200 / 9200L switches run "cat9k_lite" images;
# 9300 / 9400 / 9500 / 9600 run "cat9k" (cat9k_iosxe...) images. Installing the wrong family
# fails late (after the upload), so the dashboard checks the choice before anything starts.
IMAGE_FAMILY_LABEL = {"cat9k_lite": "Catalyst 9200 series image",
                      "cat9k": "Catalyst 9300/9400/9500/9600 series image"}
IMAGE_FAMILY_PREFIX = {"cat9k_lite": "cat9k_lite", "cat9k": "cat9k_iosxe"}


def parse_model(show_version_output):
    """'Model Number : C9200L-48P-4G' -> 'C9200L-48P-4G' (None if not found)."""
    out = show_version_output or ""
    match = re.search(r'Model Number\s*:\s*(\S+)', out, re.I)
    if not match:
        match = re.search(r'^\s*\*?\s*\d+\s+\d+\s+(C9\d{3}[\w-]*)\s', out, re.M)
    return match.group(1) if match else None


def version_key(version):
    """'17.09.04a' -> (17, 9, 4, 'a') so versions can be compared; None if it is not one."""
    match = re.search(r'(\d+)\.(\d+)\.(\d+)([a-z]*)', version or '')
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)), match.group(4))


def image_family(image_name):
    name = (image_name or '').lower()
    if name.startswith('cat9k_lite'):
        return 'cat9k_lite'
    if name.startswith('cat9k'):
        return 'cat9k'
    return None


def platform_family(model):
    m = (model or '').strip().upper()
    if re.match(r'C?9200', m):
        return 'cat9k_lite'
    if re.match(r'C?9[3-6]\d\d', m):
        return 'cat9k'
    return None


def dashboard_exposure_error(host, allow_remote, access, tls):
    """Refuse a remote dashboard unless authentication and HTTPS are active."""
    if host in ("127.0.0.1", "localhost", "::1"):
        return None
    if not allow_remote:
        return "Refusing to listen on a non-local address without --allow-remote."
    if access is None or not access.get("users"):
        return "Remote access requires configured dashboard users through the desktop launcher."
    if not tls:
        return "Remote access requires HTTPS. Enable HTTPS in the desktop launcher."
    return None


def check_image_for_device(image_name, model, current_version):
    """Reasons (plain English) why this image looks wrong for this switch; [] = looks fine.
    Only checks what it can be sure about: an unknown model or version is simply not judged."""
    problems = []
    img_fam, dev_fam = image_family(image_name), platform_family(model)
    if dev_fam:
        if img_fam and img_fam != dev_fam:
            problems.append("'{}' is a {}, but this switch is a {} ({})".format(
                image_name, IMAGE_FAMILY_LABEL[img_fam], model,
                IMAGE_FAMILY_LABEL[dev_fam].replace(' image', '')))
        elif not img_fam:
            problems.append("'{}' does not look like a Catalyst 9000 IOS-XE image (a {} image "
                            "name starts with '{}')".format(image_name, model, IMAGE_FAMILY_PREFIX[dev_fam]))
    new, cur = guess_version_from_filename(image_name), current_version
    nk, ck = version_key(new), version_key(cur)
    if nk and ck and nk < ck:
        problems.append("image version {} is older than the version the switch runs now ({}) - "
                        "this would be a downgrade".format(new, cur))
    return problems


#------------------------------------------------------------------------------
# On Windows, a program started without a console (the packaged .exe) would pop up a black
# window for every ping. This flag keeps child programs hidden.
HIDE_CHILD_WINDOW = ({'creationflags': 0x08000000} if os.name == 'nt' else {})   # CREATE_NO_WINDOW
#------------------------------------------------------------------------------
def is_reachable(host, timeout=2):
    """One ping (falls back to a TCP connect to port 22 when there is no ping command)."""
    if shutil.which('ping') is None:
        try:
            with socket.create_connection((host, 22), timeout=timeout):
                return True
        except OSError:
            return False
    system = platform.system().lower()
    if system == 'windows':
        cmd = ['ping', '-n', '1', '-w', str(int(timeout * 1000)), host]
    elif system == 'darwin':
        cmd = ['ping', '-c', '1', '-W', str(int(timeout * 1000)), host]
    else:
        cmd = ['ping', '-c', '1', '-W', str(max(1, int(timeout))), host]
    try:
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout + 3, **HIDE_CHILD_WINDOW)
    except (subprocess.TimeoutExpired, OSError):
        return False
    if proc.returncode != 0:
        return False
    if system == 'windows':                            # Windows exits 0 on "host unreachable" too
        return b'ttl=' in proc.stdout.lower()
    return True
#------------------------------------------------------------------------------
def reconnect(old_session, connect_args, secret):
    """Open a fresh CLI session (the old one is discarded)."""
    if old_session is not None:
        try:
            old_session.disconnect()
        except Exception:
            pass                                       # it is already closed
    new_session = ConnectHandler(**connect_args)
    if secret:
        new_session.enable()
    return new_session
#------------------------------------------------------------------------------
def ensure_session(session, connect_args, secret, device):
    """The image upload (SCP) can take many minutes and leaves this CLI session idle. The
    switch (line vty exec-timeout / session-timeout) or a firewall may close it meanwhile.
    Check that it still answers and reconnect if not."""
    try:
        session.find_prompt()
        return session
    except Exception:
        log(device, 'The CLI session was closed while the image was uploading (idle timeout) '
                    '- reconnecting ...')
        return reconnect(session, connect_args, secret)
#------------------------------------------------------------------------------
# Install mode, boot variable and flash space.
# Everything in this block follows Cisco's "Upgrade Catalyst 9300 Switches" (doc 222280) and
# "Upgrade Catalyst 9200 Switches" (doc 222282): the upgrade is always done in INSTALL mode.
GIB = 1024 ** 3
FLASH_EXPANSION_MIN = 1 * GIB                    # Cisco: "at least 1GB to 1.5GB of available space in the
FLASH_EXPANSION_RECOMMENDED = int(1.5 * GIB)     # flash memory for the expansion of the new image":
#                                                  below image + 1 GB = refuse, below image + 1.5 GB = warn
BOOT_TARGET = 'flash:packages.conf'
BOOT_FIX_COMMANDS = ['no boot system', 'boot system ' + BOOT_TARGET, 'no boot manual']    # Cisco steps 4 + 5
INSTALL_FAIL_RE = re.compile(r'FAILED|%\s*Error|Invalid input|Incomplete command|Ambiguous command')
ACTIVATE_FAIL_RE = re.compile(r'FAILED|%\s*Error|No such file|nothing to activate|Invalid input|Incomplete command|Ambiguous command', re.I)
PENDING_COMMIT = {'needed': False}              # set when the image was activated WITHOUT 'commit': 'install commit' must follow the reload


def model_series(model):
    """'C9300-48UXM' -> '9300', 'C9200L-48P-4G' -> '9200' (None if it is not a Catalyst 9x00)."""
    match = re.match(r'C?(9[2-6]\d\d)', (model or '').strip().upper())
    return match.group(1) if match else None


def install_flavor(model, version):
    """Which command set Cisco's guide prescribes for this switch.
      'install' - install remove inactive / install add file ... activate commit
                  (Catalyst 9200, and Catalyst 9300 on 16.6.2 or later)
      'legacy'  - request platform software package clean / install ... auto-copy
                  (Catalyst 9300 on 16.5.1a or 16.6.1 - the 9300 guide's old-release procedure)"""
    key = version_key(version)
    if model_series(model) == '9300' and key and key[:3] < (16, 6, 2):
        return 'legacy'
    return 'install'


def fmt_bytes(count):
    if count is None:
        return 'unknown'
    return '{:.2f} GB'.format(count / GIB) if count >= GIB else '{:.0f} MB'.format(count / 1048576)


def parse_run_mode(show_version_output):
    """'Installation mode is INSTALL' (or the Mode column of the version table) -> INSTALL | BUNDLE |
    MIXED (stack members differ) | None."""
    out = show_version_output or ''
    found = {m.upper() for m in re.findall(r'Installation mode is\s+(\w+)', out, re.I)}
    if not found:
        found = {m.upper() for m in re.findall(
            r'(?m)^\s*\*?\s*\d+\s+\d+\s+C9\S+\s+\S+\s+\S+\s+(INSTALL|BUNDLE)\b', out, re.I)}
    if not found:
        return None
    return found.pop() if len(found) == 1 else 'MIXED'


def _boot_entries(text):
    entries = []
    for part in (text or '').split(';'):
        part = re.sub(r',\s*\d+\s*$', '', part.strip()).strip()        # 'flash:packages.conf,12' -> path only
        if part:
            entries.append(part)
    return entries


def parse_boot(show_boot_output):
    """'show boot' -> one dict per stack member: {'switch', 'current': [...], 'next': [...], 'manual'}.
    Understands both the 'BOOT variable = ...;' and the older 'BOOT path-list : ...' layouts."""
    out = show_boot_output or ''
    parts = re.split(r'(?m)^\s*Switch\s*:?\s*(\d+)\s*$', out)
    if len(parts) < 3:
        blocks = [(1, out)]
    else:
        blocks = [(int(parts[i]), parts[i + 1]) for i in range(1, len(parts) - 1, 2)]
    boot_re = re.compile(r'(?im)^\s*BOOT\s*(?:variable\s*=|path-list\s*:|=|:)\s*(.*)$')
    members = []
    for number, text in blocks:
        sections = re.split(r'(?i)Boot Variables on next reload\s*:?', text, maxsplit=1)
        cur_txt, nxt_txt = (sections[0], sections[1]) if len(sections) == 2 else (text, text)
        cur, nxt = boot_re.search(cur_txt), boot_re.search(nxt_txt)
        if not nxt:
            continue
        manual = re.search(r'(?im)^\s*Manual Boot\s*(?:=|:)\s*(\w+)', text)
        members.append({'switch': number, 'current': _boot_entries(cur.group(1) if cur else ''),
                        'next': _boot_entries(nxt.group(1)), 'manual': manual.group(1).lower() if manual else ''})
    return members


def boot_state(show_version_output, show_boot_output):
    """What the 'Boot mode' column shows: the running mode (INSTALL / BUNDLE) and whether the boot
    variable of EVERY stack member is what Cisco prescribes for install mode (flash:packages.conf,
    manual boot off).  boot_ok is None when 'show boot' could not be read."""
    members = parse_boot(show_boot_output)
    problems = []
    for m in members:
        who = 'switch {}: '.format(m['switch']) if len(members) > 1 else ''
        if not m['next']:
            problems.append(who + 'no boot variable set')
        elif m['next'][0].lower() != BOOT_TARGET:
            problems.append('{}boots {}'.format(who, m['next'][0]))
        if m['manual'] == 'yes':
            problems.append(who + 'manual boot is ON')
    first = members[0] if members else None
    if not members:
        detail = "'show boot' could not be read"
    elif problems:
        detail = '; '.join(problems)
    else:
        detail = 'boots {} (auto boot)'.format(BOOT_TARGET)
    return {'mode': parse_run_mode(show_version_output),
            'boot_ok': (not problems) if members else None,
            'boot_var': first['next'][0] if first and first['next'] else '',
            'manual': first['manual'] if first else '',
            'detail': detail, 'time': now_str()}


def read_boot_state(net_connect, show_version_output=None):
    """-> (boot_state dict, raw 'show boot' output)."""
    if show_version_output is None:
        show_version_output = net_connect.send_command('show version', read_timeout=tmo('show_cmd'))
    raw = net_connect.send_command('show boot', read_timeout=tmo('show_cmd'))
    return boot_state(show_version_output, raw), raw


def set_boot_install_mode(net_connect, was=''):
    """Cisco steps 4-6: no boot system / boot system flash:packages.conf / no boot manual, write memory,
    then 'show boot' to confirm.  -> (boot_state dict, raw 'show boot' output)."""
    net_connect.send_config_set(BOOT_FIX_COMMANDS)
    net_connect.send_command('write memory', read_timeout=tmo('write_memory'))
    state, raw = read_boot_state(net_connect)
    if state['boot_ok']:
        state['detail'] = 'set to {}{}'.format(BOOT_TARGET, ' (was: {})'.format(was) if was else '')
    return state, raw


def stream_command(net_connect, command, device, timeout, label='clean-up', on_output=None, should_stop=None):
    """Send `command` and stream the output instead of waiting for one fixed prompt string.
    - every output line is logged as it arrives (a slow command no longer looks stuck)
    - any [y/n] question is answered with 'y'
    - returns as soon as the CLI prompt is back
    - raises ReadTimeout if nothing conclusive happens within `timeout` seconds (the output so far is on .output)
    device=None: nothing is logged (the Command time test); on_output(text) is called with all output so far;
    should_stop() returning True ends the wait (ReadTimeout with .stopped = True)"""
    net_connect.clear_buffer()
    net_connect.write_channel(command + '\n')
    output, line_buf, answered_upto = '', '', 0
    started = last_beat = time.time()
    while time.time() - started < timeout:
        time.sleep(INSTALL_POLL)
        if should_stop and should_stop():              # the Stop button of the Command time test
            err = ReadTimeout('{} stopped'.format(label))
            err.output, err.stopped = output, True
            raise err
        chunk = net_connect.read_channel()
        if chunk:
            output += chunk
            line_buf += chunk
            *lines, line_buf = line_buf.split('\n')
            for line in lines:
                if line.strip() and device:
                    log(device, '  {}> {}'.format(label, line.rstrip()))
            if on_output:
                on_output(output)
            last_beat = time.time()
        elif not net_connect.is_alive():
            raise EOFError('SSH session closed by the switch')
        elif time.time() - last_beat >= INSTALL_HEARTBEAT:
            if device:
                log(device, '  {}> still running, no new output ({}s elapsed) ...'.format(label, int(time.time() - started)))
            last_beat = time.time()
        pending = output[answered_upto:]
        if re.search(r'\[y/n\]\s*$', pending):
            net_connect.write_channel('y\n')
            answered_upto = len(output)
        elif answered_upto < len(output) and re.search(r'(?m)^\S+#\s*$', pending):
            return output
    err = ReadTimeout('{} did not finish within {}s'.format(label, timeout))
    err.output = output
    raise err


def remove_inactive(net_connect, device=None):
    """Cisco step 1 for the 'install' command set: install remove inactive (streamed; answers the
    [y/n] question; when there is nothing to remove the switch just returns to the prompt)."""
    return stream_command(net_connect, 'install remove inactive', device, tmo('clean_total'), 'remove-inactive')


def parse_free_bytes(text):
    match = re.search(r'(\d+)\s+bytes\s+(?:free|available)', text or '')
    return int(match.group(1)) if match else None


def stack_members(net_connect):
    """-> (active member number, [all member numbers]) from 'show switch'; (1, [1]) if unreadable."""
    try:
        out = net_connect.send_command('show switch', read_timeout=tmo('show_cmd'))
    except Exception:
        return 1, [1]
    rows = re.findall(r'(?m)^\s*\*?\s*(\d+)\s+(Active|Standby|Member|Master)\b', out, re.I)
    if not rows:
        return 1, [1]
    numbers = sorted({int(n) for n, _ in rows})
    active = next((int(n) for n, role in rows if role.lower() in ('active', 'master')), numbers[0])
    return active, numbers


def check_flash_space(net_connect, image_size=None, image_name=None):
    """Flash space the way Cisco's guides ask for it: room for the image (unless it is already on flash)
    PLUS 1-1.5 GB for its expansion, on every stack member (the image itself only lands on the active
    switch).  Returns {'level': 'ok' | 'low' | 'fail', 'detail', 'exists', 'has_packages_conf', ...}.
    'fail' = below image + 1 GB, 'low' = below image + 1.5 GB."""
    active, numbers = stack_members(net_connect)
    free, listing = {}, ''
    for n in numbers:
        try:
            out = net_connect.send_command('dir {}'.format('flash:' if n == active else 'flash-{}:'.format(n)),
                                           read_timeout=tmo('show_cmd'))
        except Exception:
            out = ''
        free[n] = parse_free_bytes(out)
        if n == active:
            listing = out
    if free.get(active) is None:
        raise RuntimeError('Could not parse free space from "dir flash:"')
    exists = bool(image_name) and re.search(r'(?m)^.*\s{}\s*$'.format(re.escape(image_name)), listing) is not None
    has_conf = re.search(r'(?m)\spackages\.conf\s*$', listing) is not None
    image_need = 0 if (exists or not image_size) else image_size
    rows = [(n, f, image_need if n == active else 0) for n, f in free.items() if f is not None]
    n, f, base = min(rows, key=lambda r: r[1] - r[2])                # the tightest switch decides
    level = 'fail' if f < base + FLASH_EXPANSION_MIN else 'low' if f < base + FLASH_EXPANSION_RECOMMENDED else 'ok'
    where = ' on switch {}'.format(n) if len(numbers) > 1 else ''
    need = '{} image + '.format(fmt_bytes(base)) if base else ''
    if level == 'fail':
        detail = ('NOT enough free flash{}: {} free, at least {}1 GB for the expansion = {} is needed '
                  '(Cisco: 1-1.5 GB free for the expansion of the new image)'.format(
                      where, fmt_bytes(f), need, fmt_bytes(base + FLASH_EXPANSION_MIN)))
    elif level == 'low':
        detail = ('Flash is tight{}: {} free; Cisco recommends {}1.5 GB for the expansion = {}'.format(
            where, fmt_bytes(f), need, fmt_bytes(base + FLASH_EXPANSION_RECOMMENDED)))
    else:
        detail = 'Enough free flash{}: {} free ({}1.5 GB expansion = {} needed)'.format(
            where, fmt_bytes(f), need, fmt_bytes(base + FLASH_EXPANSION_RECOMMENDED))
    unread = [str(k) for k, v in free.items() if v is None]
    if unread:
        detail += ' - could not read the flash of switch {}'.format(', '.join(unread))
    if exists:
        detail = 'Image already on flash. ' + detail
    return {'level': level, 'detail': detail, 'exists': exists, 'has_packages_conf': has_conf,
            'active': active, 'free': free}
#------------------------------------------------------------------------------
def check_md5(filename, chunk_size=1024 * 1024):
    md5 = hashlib.md5()
    with open(filename, 'rb') as image:
        for chunk in iter(lambda: image.read(chunk_size), b''):
            md5.update(chunk)
    return md5.hexdigest()
#------------------------------------------------------------------------------
def verify_image_sha256_sidecar(filename, chunk_size=1024 * 1024, sidecar_path=None):
    """Require and verify a <image>.sha256 manifest before an image can be queued.

    Obtain the expected digest independently from Cisco's trusted download page.
    """
    image_path = Path(filename)
    sidecar = Path(sidecar_path) if sidecar_path else Path(str(image_path) + '.sha256')
    if not sidecar.exists():
        raise ValueError('missing SHA-256 sidecar {}'.format(sidecar.name))
    try:
        line = next((line.strip() for line in sidecar.read_text(encoding='utf-8').splitlines()
                     if line.strip() and not line.lstrip().startswith('#')), '')
    except (OSError, UnicodeError) as exc:
        raise ValueError('could not read SHA-256 sidecar: {}'.format(exc)) from exc
    match = re.fullmatch(r'([0-9a-fA-F]{64})(?:\s+\*?.*)?', line)
    if not match:
        raise ValueError('SHA-256 sidecar must start with a 64-character hexadecimal digest')
    expected = match.group(1).lower()
    digest = hashlib.sha256()
    try:
        with image_path.open('rb') as image:
            for chunk in iter(lambda: image.read(chunk_size), b''):
                digest.update(chunk)
    except OSError as exc:
        raise ValueError('could not read image for SHA-256 verification: {}'.format(exc)) from exc
    actual = digest.hexdigest()
    if not hmac.compare_digest(actual, expected):
        raise ValueError('SHA-256 mismatch for {} (expected {}, got {})'.format(
            image_path.name, expected, actual))
    return actual

def _verify_image_digest(filename, algorithm, expected_digest, chunk_size=1024 * 1024):
    label = {"md5": "MD5", "sha512": "SHA-512"}.get(algorithm, "SHA-256")
    size = {"md5": 32, "sha512": 128}.get(algorithm, 64)
    expected = re.sub(r"\s+", "", str(expected_digest or "")).lower()
    if not re.fullmatch(r"[0-9a-f]{" + str(size) + r"}", expected):
        raise ValueError("{} checksum must be a {}-character hexadecimal digest".format(label, size))
    digest = hashlib.new(algorithm)
    image_path = Path(filename)
    try:
        with image_path.open("rb") as image:
            for chunk in iter(lambda: image.read(chunk_size), b""):
                digest.update(chunk)
    except OSError as exc:
        raise ValueError("could not read image for {} verification: {}".format(label, exc)) from exc
    actual = digest.hexdigest()
    if not hmac.compare_digest(actual, expected):
        raise ValueError("{} mismatch for {} (expected {}, got {})".format(
            label, image_path.name, expected, actual))
    return actual


def verify_image_md5(filename, expected_digest, chunk_size=1024 * 1024):
    """Verify image bytes against the MD5 digest shown on Cisco's download page."""
    return _verify_image_digest(filename, "md5", expected_digest, chunk_size)


def verify_image_sha512(filename, expected_digest, chunk_size=1024 * 1024):
    """Verify image bytes against a SHA-512 digest copied from Cisco's download page."""
    return _verify_image_digest(filename, "sha512", expected_digest, chunk_size)


def verify_image_checksum_sidecar(filename, chunk_size=1024 * 1024):
    """Verify a saved checksum sidecar: MD5 (what the Storage page writes), else a legacy SHA-512 / SHA-256 one."""
    image_path = Path(filename)
    for algorithm, label, size in (("md5", "MD5", 32), ("sha512", "SHA-512", 128), ("sha256", "SHA-256", 64)):
        sidecar = Path(str(image_path) + "." + algorithm)
        if not sidecar.exists():
            continue
        try:
            line = next((line.strip() for line in sidecar.read_text(encoding="utf-8").splitlines()
                         if line.strip() and not line.lstrip().startswith("#")), "")
        except (OSError, UnicodeError) as exc:
            raise ValueError("could not read {} sidecar: {}".format(label, exc)) from exc
        match = re.fullmatch(r"([0-9a-fA-F]{" + str(size) + r"})(?:\s+\*?.*)?", line)
        if not match:
            raise ValueError("{} sidecar must start with a {}-character hexadecimal digest".format(label, size))
        return _verify_image_digest(image_path, algorithm, match.group(1), chunk_size)
    raise ValueError("missing MD5 checksum sidecar - upload the image again in Storage > IOS-XE images")


#------------------------------------------------------------------------------
def software_install(net_connect, image_name, device, timeout=None, mode='full'):
    """IOS-XE install mode. `mode` selects which install command is actually run:
      'full'       - install add file ... activate commit   (upload+add+activate+reload, one step;
                                                               this is the original, unchanged behavior)
      'distribute' - install add file ...                   (copy to flash only - no activate, no reload;
                                                               lets you pre-stage the image ahead of time)
      'activate'   - install activate                       (activates an already-distributed image and
                                                               reloads; no filename needed or used. This
                                                               command has no 'commit' keyword - the commit
                                                               is a separate 'install commit' AFTER the reload)

      'legacy'     - request platform software package install switch all file ... auto-copy
                     (Catalyst 9300 on 16.5.1a / 16.6.1 only - the 9300 guide's old-release command)

    The command is streamed instead of waiting for one fixed prompt string:
    - progress lines are logged as they arrive (so a slow install does not look stuck)
    - any [y/n] prompt is answered with 'y'
    - it returns as soon as the switch prompt is back (install failed / nothing to do)
    - a dropped SSH session raises (= the switch is reloading)
    - ReadTimeout is raised if nothing conclusive happens within `timeout` seconds
    """
    timeout = tmo('install_total') if timeout is None else timeout
    if mode == 'activate':
        command = 'install activate\n'
    elif mode == 'legacy':
        net_connect.send_command('write memory', read_timeout=tmo('write_memory'))
        command = 'request platform software package install switch all file flash:{} auto-copy\n'.format(image_name)
    elif mode == 'distribute':
        net_connect.send_command('write memory', read_timeout=tmo('write_memory'))
        command = 'install add file flash:{}\n'.format(image_name)
    else:
        net_connect.send_command('write memory', read_timeout=tmo('write_memory'))
        command = 'install add file flash:{} activate commit\n'.format(image_name)
    net_connect.write_channel(command)

    output = ''            # everything received since the command was sent
    line_buf = ''          # incomplete trailing line, for logging
    answered_upto = 0      # prompts before this offset were already answered
    started = last_beat = time.time()

    while time.time() - started < timeout:
        time.sleep(INSTALL_POLL)
        chunk = net_connect.read_channel()      # raises ReadException if the channel was closed
        if chunk:
            output += chunk
            line_buf += chunk
            *lines, line_buf = line_buf.split('\n')
            for line in lines:
                if line.strip():
                    log(device, '  install> {}'.format(line.rstrip()))
            last_beat = time.time()
        elif not net_connect.is_alive():
            raise EOFError('SSH session closed by the switch')
        elif time.time() - last_beat >= INSTALL_HEARTBEAT:
            log(device, '  install> still running, no new output ({}s elapsed) ...'.format(
                int(time.time() - started)))
            last_beat = time.time()

        pending = output[answered_upto:]
        if re.search(r'\[y/n\]\s*$', pending):                 # e.g. "Do you want to proceed? [y/n]"
            net_connect.write_channel('y\n')
            answered_upto = len(output)
        elif re.search(r'(?m)^\S+#\s*$', pending) and answered_upto < len(output):
            return output                                       # CLI prompt is back: command finished

    raise ReadTimeout('install did not finish or drop the session within {}s'.format(timeout))
#------------------------------------------------------------------------------
def clean_packages(net_connect, device=None):
    """IOS-XE legacy: free flash with 'request platform software package clean switch all' (streamed,
    [y/n] answered with 'y')."""
    return stream_command(net_connect, 'request platform software package clean switch all',
                          device, tmo('clean_total'), 'clean')
#------------------------------------------------------------------------------
def verify_md5(net_connect, image_name, md5, member_fs='flash'):
    result = net_connect.send_command(
        'verify /md5 {}:{} {}'.format(member_fs, image_name, md5), read_timeout=tmo('md5_verify'))
    return re.search(r'Verified', result) is not None
#------------------------------------------------------------------------------
def probe_device_readiness(host, device_type, username, password, secret, image_local_path):
    """Connect, run a few read-only commands, and disconnect. No upload, no install, no reload -
    this is meant to answer "is it safe to upgrade this switch?" in a few seconds, not perform
    the upgrade itself. Returns a dict of individual findings; compute_readiness_verdict() turns
    that into an overall ready/warning/not_ready call.  It also reads the platform, the running
    mode and the boot variable ('show boot'), and checks the flash space the way Cisco asks
    (image + 1-1.5 GB for the expansion, on every stack member)."""
    device = {"device_type": device_type, "host": host, "username": username, "password": password,
              "secret": secret, "conn_timeout": 15, "auth_timeout": 15, "banner_timeout": 15}
    checks = {"reachable": False, "login": False, "version": None, "model": None, "boot": None,
              "space_ok": None, "space_level": None, "space_detail": "", "cpu_pct": None,
              "mem_free_pct": None, "stack": None}
    with ConnectHandler(**device) as conn:
        checks["reachable"] = True
        if secret:
            conn.enable()
        checks["login"] = True
        transcript = []                                # every command + output, saved as one file
        checks["_transcript"] = transcript
        _send = conn.send_command

        def _recording_send(*args, **kwargs):
            out = _send(*args, **kwargs)
            cmd = args[0] if args else kwargs.get("command_string")
            if isinstance(cmd, str) and isinstance(out, str):
                transcript.append((cmd, out))
            return out
        conn.send_command = _recording_send
        show_version = None
        try:
            show_version = conn.send_command("show version", read_timeout=tmo('show_cmd'))
        except Exception:  # noqa: BLE001 - a failed lookup is a finding, not a crash
            pass
        checks["version"] = parse_version(show_version) if show_version else None
        checks["model"] = parse_model(show_version) if show_version else None

        try:                                           # boot variable + install/bundle mode
            checks["boot"], _raw = read_boot_state(conn, show_version)
        except Exception as exc:  # noqa: BLE001
            checks["boot_error"] = f"Could not read the boot settings: {type(exc).__name__}"

        try:                                           # image + 1-1.5 GB expansion, every stack member
            size = os.path.getsize(image_local_path) if image_local_path else None
            space = check_flash_space(conn, size, Path(image_local_path).name if image_local_path else None)
            checks["space_ok"] = space["level"] != "fail"
            checks["space_level"] = space["level"]
            checks["space_detail"] = space["detail"]
        except Exception as exc:  # noqa: BLE001 - a failed space check is a finding, not a crash
            checks["space_detail"] = f"Could not check flash space: {exc}"

        health = read_health(conn)
        checks.update(health)

        try:                                           # interface / VLAN / trunk / port-channel snapshot
            outputs, snap = read_snapshot(conn)
            checks["snapshot"] = snap
            checks["snapshot_text"] = summarize_snapshot(snap)
            checks["_outputs"] = outputs               # saved to a file by the caller, then dropped
        except Exception:  # noqa: BLE001
            pass

    return checks


def compute_readiness_verdict(checks):
    """Turn probe_device_readiness()'s findings into (state, summary message)."""
    if not checks["reachable"] or not checks["login"]:
        return "not_ready", "Could not connect or log in"

    problems, warnings = [], []
    if checks["space_ok"] is False:
        problems.append(checks["space_detail"] or "not enough free flash space")
    elif checks.get("space_level") == "low":
        warnings.append(checks["space_detail"])
    boot = checks.get("boot")
    if boot:
        if boot.get("mode") == "BUNDLE":
            warnings.append(f"BUNDLE mode ({boot.get('detail')}) - the upgrade converts it to INSTALL mode "
                            f"and sets {BOOT_TARGET}")
        elif boot.get("boot_ok") is False:
            warnings.append(f"boot variable is not right for install mode ({boot.get('detail')}) - "
                            f"the upgrade sets {BOOT_TARGET}")
    elif checks.get("boot_error"):
        warnings.append(checks["boot_error"])
    if checks["cpu_pct"] is not None and checks["cpu_pct"] > 80:
        warnings.append(f"CPU {checks['cpu_pct']}% (1-min average)")
    if checks["mem_free_pct"] is not None and checks["mem_free_pct"] < 15:
        warnings.append(f"only {checks['mem_free_pct']}% memory free")
    if checks["stack"] and "not Ready" in checks["stack"]:
        warnings.append(checks["stack"])
    snap = checks.get("snapshot")
    if snap:
        if snap["err_disabled"]:
            warnings.append(f"{len(snap['err_disabled'])} err-disabled port(s): {_names(snap['err_disabled'])}")
        down_po = [n for n, v in sorted(snap["po"].items()) if "U" not in v["flags"]]
        if down_po:
            warnings.append(f"port-channel(s) down: {_names(down_po)}")

    if problems:
        return "not_ready", "; ".join(problems)
    if warnings:
        return "warning", "; ".join(warnings)

    parts = [f"version {checks['version']}" if checks["version"] else None,
             f"{boot.get('mode') or 'unknown'} mode, {boot.get('detail')}" if boot else None,
             checks["space_detail"] or None,
             f"CPU {checks['cpu_pct']}%" if checks["cpu_pct"] is not None else None,
             f"{checks['mem_free_pct']}% mem free" if checks["mem_free_pct"] is not None else None,
             checks["stack"], checks.get("snapshot_text") or None]
    return "ready", "; ".join(p for p in parts if p) or "Looks ready"


#------------------------------------------------------------------------------
# Network-state snapshot: taken before an upgrade (pre-check) and again after it (post-check),
# so the two can be compared. Read-only show commands only.
SNAPSHOT_COMMANDS = [
    ("ip_int_brief", "show ip interface brief"),
    ("int_status", "show interfaces status"),
    ("vlan_brief", "show vlan brief"),
    ("int_trunk", "show interfaces trunk"),
    ("etherchannel", "show etherchannel summary"),
]
POSTCHECK_SETTLE = 90         # seconds to let ports/port-channels come up after the reload
POSTCHECK_RETRIES = 3         # extra tries when something still differs from before the upgrade
POSTCHECK_RETRY_WAIT = 60


def parse_snapshot(outputs):
    """Raw `show` outputs -> a small comparable dict. 'ok' says which commands could be parsed."""
    ip_up, connected, err_disabled, vlans, trunks, po = [], [], [], [], [], {}
    for m in re.finditer(r'(?m)^(\S+)\s+(\S+)\s+(?:YES|NO)\s+\S+\s+(administratively down|up|down|deleted)\s+(up|down)\s*$',
                         outputs.get("ip_int_brief", "")):
        if m.group(3) == "up" and m.group(4) == "up":
            ip_up.append(m.group(1))
    ip_rows = bool(re.search(r'(?m)^\S+\s+\S+\s+(?:YES|NO)\s', outputs.get("ip_int_brief", "")))
    st_rows = 0
    for line in outputs.get("int_status", "").splitlines():
        if not line.strip() or line.startswith("Port"):
            continue
        m = re.match(r'^(\S+)\s+(?:.*?\s)?(connected|notconnect|disabled|err-disabled|inactive|suspended|monitoring|faulty|sfp-config-mismatch)\b', line)
        if m:
            st_rows += 1
            if m.group(2) == "connected":
                connected.append(m.group(1))
            elif m.group(2) == "err-disabled":
                err_disabled.append(m.group(1))
    for m in re.finditer(r'(?m)^(\d+)\s+(\S+)\s+(active|act/unsup|suspended|act/lshut|sus/lshut)\b',
                         outputs.get("vlan_brief", "")):
        vlans.append(m.group(1))
    tr_rows = 0
    for m in re.finditer(r'(?m)^(\S+)\s+(on|auto|desirable|nonegotiate)\s+\S+\s+(trunking|not-trunking)\s+\d+',
                         outputs.get("int_trunk", "")):
        tr_rows += 1
        if m.group(3) == "trunking":
            trunks.append(m.group(1))
    for m in re.finditer(r'(?m)^\s*\d+\s+(Po\d+)\((\w+)\)\s+\S+\s*(.*)$', outputs.get("etherchannel", "")):
        members = re.findall(r'\S+?\((\w+)\)', m.group(3))
        po[m.group(1)] = {"flags": m.group(2), "members": len(members),
                          "bundled": sum(1 for f in members if f == "P")}
    ec_text = outputs.get("etherchannel", "")
    return {"ip_up": sorted(ip_up), "connected": sorted(connected), "err_disabled": sorted(err_disabled),
            "vlans": sorted(vlans, key=int), "trunks": sorted(trunks), "po": po,
            "ok": {"ip_int_brief": ip_rows, "int_status": st_rows > 0,
                   "vlan_brief": bool(vlans),
                   "int_trunk": tr_rows > 0 or "Port" in outputs.get("int_trunk", ""),
                   "etherchannel": "channel-groups" in ec_text or bool(po)}}


def summarize_snapshot(snap):
    po_up = sum(1 for v in snap["po"].values() if "U" in v["flags"])
    parts = [f"{len(snap['connected'])} port(s) connected", f"{len(snap['ip_up'])} interface(s) up/up",
             f"{len(snap['vlans'])} VLAN(s)", f"{len(snap['trunks'])} trunk(s)"]
    if snap["po"]:
        parts.append(f"{po_up}/{len(snap['po'])} port-channel(s) up")
    if snap["err_disabled"]:
        parts.append(f"{len(snap['err_disabled'])} err-disabled")
    return ", ".join(parts)


def _names(items, limit=6):
    items = list(items)
    return ", ".join(items[:limit]) + (f" (+{len(items) - limit} more)" if len(items) > limit else "")


def compare_snapshots(pre, post):
    """What was there before the upgrade and is missing/worse now. [] = nothing to report."""
    issues = []
    both = lambda k: pre["ok"].get(k) and post["ok"].get(k)          # noqa: E731 - only compare parsed data
    if both("ip_int_brief"):
        lost = sorted(set(pre["ip_up"]) - set(post["ip_up"]))
        if lost:
            issues.append(f"{len(lost)} interface(s) up/up before but not now: {_names(lost)}")
    if both("int_status"):
        lost = sorted(set(pre["connected"]) - set(post["connected"]))
        if lost:
            issues.append(f"{len(lost)} port(s) connected before but not now: {_names(lost)}")
        new_bad = sorted(set(post["err_disabled"]) - set(pre["err_disabled"]))
        if new_bad:
            issues.append(f"{len(new_bad)} new err-disabled port(s): {_names(new_bad)}")
    if both("vlan_brief"):
        lost = sorted(set(pre["vlans"]) - set(post["vlans"]), key=int)
        if lost:
            issues.append(f"VLAN(s) missing: {_names(lost)}")
    if both("int_trunk"):
        lost = sorted(set(pre["trunks"]) - set(post["trunks"]))
        if lost:
            issues.append(f"trunk(s) not trunking any more: {_names(lost)}")
    if both("etherchannel"):
        for name, was in sorted(pre["po"].items()):
            now = post["po"].get(name)
            if now is None:
                issues.append(f"{name} is missing")
            elif "U" in was["flags"] and "U" not in now["flags"]:
                issues.append(f"{name} is down ({now['flags']})")
            elif now["bundled"] < was["bundled"]:
                issues.append(f"{name} has {now['bundled']} bundled member(s), had {was['bundled']}")
    return issues


def read_snapshot(conn):
    """Run the snapshot commands on an open session -> (raw outputs, parsed snapshot)."""
    outputs = {}
    for key, cmd in SNAPSHOT_COMMANDS:
        try:
            outputs[key] = conn.send_command(cmd, read_timeout=tmo('snapshot_cmd'))
        except Exception as exc:  # noqa: BLE001 - one failing show command must not stop the others
            outputs[key] = ""
            outputs[key + "_error"] = type(exc).__name__
    return outputs, parse_snapshot(outputs)


def snapshot_text(outputs, heading):
    parts = [heading, "=" * len(heading), ""]
    for key, cmd in SNAPSHOT_COMMANDS:
        parts += [f"#### {cmd}", outputs.get(key) or f"(no output{': ' + outputs[key + '_error'] if key + '_error' in outputs else ''})", ""]
    return "\n".join(parts)


COMMAND_HEADER_RE = re.compile(r"(?m)^!={5} COMMAND: (.*?) ={5}\s*$")


def command_sections_text(heading, sections):
    """ONE text file holding the output of many commands. `sections` = [(command, output)]; every
    block starts with a '!===== COMMAND: <command> =====' line so it can be split again later."""
    parts = [f"! {heading}", f"! Commands: {'; '.join(c for c, _ in sections)}", ""]
    for cmd, out in sections:
        parts += [f"!===== COMMAND: {cmd} =====", (out or "").rstrip("\n") or "(no output)", ""]
    return "\n".join(parts)


def split_command_sections(text):
    """Inverse of command_sections_text(): {command: output}."""
    marks = list(COMMAND_HEADER_RE.finditer(text))
    out = {}
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        out.setdefault(m.group(1), text[m.end():end].strip("\r\n"))
    return out


def read_health(conn):
    """CPU (1-min average), free memory and stack state - shared by the pre- and post-check."""
    out = {"cpu_pct": None, "mem_free_pct": None, "stack": None}
    try:
        cpu_out = conn.send_command("show processes cpu", read_timeout=tmo('health_cmd'))
        m = re.search(r"five seconds:\s*\d+%.*?one minute:\s*(\d+)%", cpu_out, re.S)
        if m:
            out["cpu_pct"] = int(m.group(1))
    except Exception:  # noqa: BLE001 - nice-to-have, never fatal
        pass
    try:
        mem_out = conn.send_command("show processes memory", read_timeout=tmo('health_cmd'))
        m = re.search(r"Processor Pool Total:\s*(\d+)\s*Used:\s*(\d+)\s*Free:\s*(\d+)", mem_out)
        if m:
            total, free = int(m.group(1)), int(m.group(3))
            if total:
                out["mem_free_pct"] = round(free / total * 100, 1)
    except Exception:  # noqa: BLE001
        pass
    try:
        sw_out = conn.send_command("show switch", read_timeout=tmo('health_cmd'))
        states = re.findall(r"^\*?\s*\d+\s+\S+\s+\S+\s+\S+\s+(\S+)\s*$", sw_out, re.M)
        if states:
            not_ready = [x for x in states if x.lower() != "ready"]
            out["stack"] = (f"{len(states)} member(s), all Ready" if not not_ready else
                            f"{len(states)} member(s), not Ready: {', '.join(not_ready)}")
    except Exception:  # noqa: BLE001 - a standalone switch may not support/need this at all
        pass
    return out


def evaluate_postcheck(conn, expected_version, pre):
    """One post-upgrade check on an open session. Returns (result dict, raw outputs, issues)."""
    version = get_version(conn)
    health = read_health(conn)
    outputs, snap = read_snapshot(conn)
    problems, warnings = [], []
    if expected_version and version != expected_version:
        problems.append(f"running {version or 'unknown'}, expected {expected_version}")
    if health["cpu_pct"] is not None and health["cpu_pct"] > 80:
        warnings.append(f"CPU {health['cpu_pct']}% (1-min average)")
    if health["mem_free_pct"] is not None and health["mem_free_pct"] < 15:
        warnings.append(f"only {health['mem_free_pct']}% memory free")
    stack_note = None
    if health["stack"] and "not Ready" in health["stack"]:
        if PRE_STACK is not None and health["stack"] == PRE_STACK:
            # e.g. a provisioned-but-removed stack member: it was already like this before the
            # upgrade, so the upgrade did not cause it - note it, but don't raise a warning
            stack_note = health["stack"] + " (unchanged from before the upgrade)"
        else:
            warnings.append(health["stack"] + (" (before the upgrade: {})".format(PRE_STACK)
                                               if PRE_STACK else ""))
    issues = compare_snapshots(pre, snap) if pre else []
    warnings += issues
    checks = {"version": version, "cpu_pct": health["cpu_pct"], "mem_free_pct": health["mem_free_pct"],
              "stack": health["stack"], "summary": summarize_snapshot(snap), "issues": issues,
              "compared": bool(pre)}
    if problems:
        state, detail = "failed", "; ".join(problems + warnings)
    elif warnings:
        state, detail = "warning", "; ".join(warnings)
    else:
        parts = [f"version {version}" if version else None, checks["summary"],
                 stack_note or health["stack"], f"CPU {health['cpu_pct']}%" if health["cpu_pct"] is not None else None,
                 "same as before the upgrade" if pre else "no pre-upgrade snapshot to compare"]
        state, detail = "passed", "; ".join(x for x in parts if x)
    return {"state": state, "detail": detail, "checks": checks}, outputs, snap


PRE_SNAPSHOT = None           # snapshot taken by this worker just before the upgrade
PRE_STACK = None              # stack state ("show switch") seen just before the upgrade


def save_check_file(suffix, heading, outputs, extra=None):
    """Write the raw `show` output of a pre-/post-check next to this run's log (web-GUI runs only)."""
    if not LOG_FILE:
        return None
    path = os.path.splitext(LOG_FILE)[0] + '.' + suffix + '.txt'
    try:
        text = snapshot_text(outputs, heading)
        if extra:
            text = '\n'.join(extra) + '\n\n' + text
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write(text)
        return path
    except OSError:
        return None


@activity('Pre-check')
def take_pre_snapshot(session, device):
    """Record interfaces / VLANs / trunks / port-channels before anything is changed."""
    global PRE_SNAPSHOT
    report('prechecking', 'Taking the pre-upgrade snapshot (interfaces, VLANs, trunks, port-channels) ...')
    log(device, 'Pre-check: reading the network state before the upgrade ...')
    try:
        outputs, snap = read_snapshot(session)
    except Exception as exc:  # noqa: BLE001 - a failed snapshot must not block the upgrade
        log(device, 'WARNING: pre-upgrade snapshot failed ({}); the post-check will not be able to '
                    'compare.'.format(type(exc).__name__))
        return
    PRE_SNAPSHOT = snap
    global PRE_STACK
    try:
        PRE_STACK = read_health(session).get("stack")
    except Exception:  # noqa: BLE001 - nice-to-have
        PRE_STACK = None
    save_check_file('precheck', 'Pre-upgrade snapshot - {}'.format(device['ipaddr']), outputs)
    if STATUS is not None:
        STATUS.update(pre_snapshot=snap)
    log(device, 'Pre-upgrade snapshot: {}'.format(summarize_snapshot(snap)))


@activity('Post-check')
def run_postcheck(device, creds, expected_version, pre):
    """After the reload: wait for ports to settle, then verify version, stack, CPU/memory and
    compare the network state with the pre-upgrade snapshot. Retries a few times so ports that
    are slow to come up don't raise a false alarm. Returns the result dict."""
    ip = device['ipaddr']
    secret = creds[2] if len(creds) > 2 else None
    log(device, 'Post-check: waiting {} s for ports and port-channels to settle ...'.format(POSTCHECK_SETTLE))
    report('postcheck', 'Post-upgrade check: letting ports settle ...')
    time.sleep(POSTCHECK_SETTLE)
    result, outputs = None, {}
    for attempt in range(POSTCHECK_RETRIES + 1):
        report('postcheck', 'Running post-upgrade checks ...')
        session = None
        try:
            connect_args = dict(device_type='cisco_xe', ip=ip, username=creds[0], password=creds[1])
            if secret:
                connect_args['secret'] = secret
            session = ConnectHandler(**connect_args)
            if secret:
                session.enable()
            result, outputs, _snap = evaluate_postcheck(session, expected_version, pre)
        except Exception as exc:  # noqa: BLE001
            result = {'state': 'failed', 'checks': {},
                      'detail': 'Post-check could not log in or read the switch ({})'.format(type(exc).__name__)}
        finally:
            if session is not None:
                try:
                    session.disconnect()
                except Exception:
                    pass
        log(device, 'Post-check result: {} - {}'.format(result['state'].upper(), result['detail']))
        if result['state'] == 'passed' or attempt == POSTCHECK_RETRIES:
            break
        log(device, 'Post-check found differences - checking again in {} s ({}/{}) ...'.format(
            POSTCHECK_RETRY_WAIT, attempt + 1, POSTCHECK_RETRIES))
        report('postcheck', 'Post-check found differences - checking again shortly ({}/{}) ...'.format(
            attempt + 1, POSTCHECK_RETRIES))
        time.sleep(POSTCHECK_RETRY_WAIT)
    result['time'] = now_str()
    save_check_file('postcheck', 'Post-upgrade check - {}'.format(ip), outputs,
                    ['Result : {}'.format(result['state'].upper()), 'Detail : {}'.format(result['detail'])] +
                    ['Issue  : {}'.format(i) for i in result['checks'].get('issues', [])])
    return result


SCP_STATE = {'added': False}      # True when THIS run switched 'ip scp server enable' on (the switch did not have it)
SCP_LINE_RE = re.compile(r'(?m)^\s*ip scp server enable\s*$')
SCP_SHOW = 'show running-config | include ip scp server enable'


def transfer_file(net_connect, local_path, image_name, state=None):
    state = SCP_STATE if state is None else state
    try:                                   # remember whether the switch already had it: only what WE add is removed later
        already = bool(SCP_LINE_RE.search(net_connect.send_command(SCP_SHOW, read_timeout=tmo('show_cmd')) or ''))
    except Exception:  # noqa: BLE001 - cannot tell: treat as "was there", so nothing is removed afterwards
        already = None
    net_connect.send_config_set(['ip scp server enable'])
    if already is False:
        state['added'] = True
    scp_conn = SCPConn(getattr(net_connect, '_conn', net_connect))     # the real session, not the recording wrapper
    try:
        scp_conn.scp_transfer_file(local_path, image_name)
    finally:
        scp_conn.close()
def remove_scp_server(device, creds):
    """The upgrade turned 'ip scp server enable' on only to upload the image. Once the upgrade has finished, turn it
    off again (and save, so it does not come back after the next reload). -> (ok, note)."""
    secret = creds[2] if len(creds) > 2 else None
    session = None
    try:
        connect_args = dict(device_type='cisco_xe', ip=device['ipaddr'], username=creds[0], password=creds[1],
                            conn_timeout=tmo('ssh_connect'))
        if secret:
            connect_args['secret'] = secret
        session = ConnectHandler(**connect_args)
        if secret:
            session.enable()
        session.send_config_set(['no ip scp server enable'])
        session.send_command('write memory', read_timeout=tmo('write_memory'))
        if SCP_LINE_RE.search(session.send_command(SCP_SHOW, read_timeout=tmo('show_cmd')) or ''):
            return False, "'ip scp server enable' is still on - remove it on the switch (no ip scp server enable)"
        return True, "'ip scp server enable' removed again (the switch did not have it before the upgrade)"
    except Exception as exc:  # noqa: BLE001
        return False, ("could not remove 'ip scp server enable' that the upgrade added ({}) - remove it on the "
                       "switch: no ip scp server enable".format(type(exc).__name__))
    finally:
        if session is not None:
            try:
                session.disconnect()
            except Exception:  # noqa: BLE001
                pass


#------------------------------------------------------------------------------
def read_devices(xlsx_filename, sheet_name='Switches'):
    """Read the switch list from an .xlsx file. Only cisco-xe devices are returned."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ValueError('openpyxl is required to read the switch list: pip install openpyxl')

    workbook = load_workbook(xlsx_filename, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook.worksheets[0]
        rows = sheet.iter_rows(values_only=True)

        header = next(rows, None)
        if not header:
            raise ValueError('{}: sheet {!r} is empty'.format(xlsx_filename, sheet.title))
        cols = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}

        def find_col(*names):
            for name in names:
                if name in cols:
                    return cols[name]
            return None

        ip_col = find_col('ip', 'ipaddr', 'ip address', 'ip_address')
        name_col = find_col('hostname', 'name')
        type_col = find_col('type', 'device_type', 'device type')
        if ip_col is None:
            raise ValueError("{}: sheet {!r} has no 'ip' column (headers found: {})".format(
                xlsx_filename, sheet.title, ', '.join(cols) or 'none'))

        def cell(row, col):
            if col is None or col >= len(row) or row[col] is None:
                return ''
            return str(row[col]).strip()

        devices = {}          # key = ipaddr
        skipped_type = []
        for row in rows:
            ipaddr = cell(row, ip_col)
            if not ipaddr:
                continue                                   # blank row
            dev_type = cell(row, type_col).lower().replace('_', '-')
            if dev_type and dev_type != 'cisco-xe':
                skipped_type.append('{} ({})'.format(ipaddr, dev_type))
                continue                                   # only cisco-xe is processed
            if ipaddr in devices:
                print('WARNING: duplicate ip {} in switch list - using the first entry'.format(ipaddr))
                continue
            devices[ipaddr] = {'ipaddr': ipaddr, 'type': 'cisco-xe',
                               'name': cell(row, name_col) or ipaddr}
    finally:
        workbook.close()

    print('\n----- devices (cisco-xe only) ----------')
    pprint(devices)
    if skipped_type:
        print('Skipped (not cisco-xe): {}'.format(', '.join(skipped_type)))

    return devices
#------------------------------------------------------------------------------
def report_boot(boot):
    """Publish the boot mode to the web GUI (the 'Boot mode' column follows it live)."""
    if STATUS is not None and boot:
        STATUS.update(boot=boot)
#------------------------------------------------------------------------------
@activity('Upgrade')
def command_worker(device, creds, local_path, image_name, md5, results, mode='full', ignore_space=False):
    """Returns nothing; outcome is stored in results[ipaddr].  IOS-XE INSTALL mode for the whole
    operation, following Cisco's "Upgrade Catalyst 9300 / 9200 Switches" guides:
      1. platform + version + boot mode are read ('show version', 'show boot')
      2. remove inactive packages  (install remove inactive; the 9300 on 16.5.1a/16.6.1 uses
         'request platform software package clean switch all')
      3. flash check: image + 1-1.5 GB expansion free on every stack member
      4. SCP the .bin to flash:, verify /md5
      5. boot variable: no boot system / boot system flash:packages.conf / no boot manual / write memory
         (only when it is not right already), then 'show boot' to confirm
      6. install add file flash:<image> activate commit  (the 9300 on 16.5.1a/16.6.1: request platform
         software package install switch all file ... auto-copy)
    mode: 'full' (default) = steps 1-6.
          'distribute'     = steps 1-4 + install add only - no activate, no reload.
          'activate'       = boot variable + install activate (commit follows the reload) only - assumes an image was already
                             distributed; no upload, no local_path/image_name/md5 needed (pass None).
    """
    username, password = creds[0], creds[1]
    secret = creds[2] if len(creds) > 2 else None       # optional enable secret
    device_type = BACKUP_DEVICE_TYPES.get(device['type'], 'cisco_ios')
    session = None
    status = 'FAILED'

    def ensure_boot(boot):
        """Cisco steps 4-6.  Leaves a correct boot variable alone; otherwise sets it and verifies it
        with 'show boot'.  Raises RuntimeError when it cannot be made right (nothing is installed)."""
        if boot and boot.get('boot_ok'):
            log(device, 'Boot variable is already correct ({}) - nothing to change'.format(boot['detail']))
            return boot
        before = boot['detail'] if boot else "'show boot' could not be read"
        report('boot', 'Setting the boot variable for install mode ...')
        log(device, 'Boot check: {} - setting it for install mode ({} / write memory) ...'.format(
            before, ' / '.join(BOOT_FIX_COMMANDS)))
        new, raw = set_boot_install_mode(session, before)
        log(device, 'show boot after the change:\n{}'.format(raw.strip()))
        report_boot(new)
        if not new.get('boot_ok'):
            raise RuntimeError('the boot variable could not be set to {} ({})'.format(BOOT_TARGET, new['detail']))
        log(device, 'Boot variable {} and saved'.format(new['detail']))
        return new

    def add_image():
        """install add file flash:<image>   (no activation, no reload).  True = added."""
        try:
            output = software_install(session, image_name, device, mode='distribute')
        except (ReadTimeout, CONNECTION_LOST) as exc:
            log(device, 'Problem while adding the image: {}: {}'.format(type(exc).__name__, exc))
            return False
        if INSTALL_FAIL_RE.search(output):
            if re.search(r'already\s+(been\s+)?added', output, re.I):
                log(device, 'The image is already added (staged) on the switch - nothing more to do')
                return True
            log(device, 'Add reported a failure (see output above)')
            return False
        return True

    def run_install(install_mode, image, fail_re):
        """Run the activating install command; returns the status string for this switch."""
        if install_mode == 'activate':
            PENDING_COMMIT['needed'] = True            # 'install activate' does not commit: done after the reload
            if STATUS is None:
                log(device, 'NOTE: after the reload run "install commit" on the switch, or it rolls the activation back')
        try:
            output = software_install(session, image, device, mode=install_mode)
        except ReadTimeout:
            # NOT a reload: the switch never produced the expected prompt/output
            log(device, 'TIMED OUT waiting for the install command - the switch has probably '
                        'NOT reloaded. Check "show install summary" and "show install log".')
            return 'CHECK MANUALLY (timeout)'
        except CONNECTION_LOST:
            log(device, 'Session dropped - switch is reloading. '
                        'Please check if device comes up again as desired ...')
            report('reloading', 'Session dropped - the switch is reloading ...')
            return 'OK (reloading)'
        if fail_re.search(output):
            log(device, 'The install command reported a failure (see output above)')
            if install_mode == 'activate' and re.search(r'install\s+add', output, re.I):
                log(device, 'Nothing is added (staged) on this switch, so there is nothing to activate. '
                            'Run "Distribute only" first (or use "Full upgrade").')
                return 'FAILED (nothing distributed on the switch - run Distribute only first)'
            if 'Super package already added' in output:
                log(device, 'Hint: a previous install left packages added. Run "install remove inactive" '
                            'on the switch, then re-run.')
            return 'FAILED (install)'
        if STATUS is not None:                         # web GUI mode
            # The prompt can come back a few seconds BEFORE the switch reloads.
            # job_main() waits for the reload and verifies the new version.
            log(device, 'Install finished; the session has not dropped yet. The switch normally '
                        'reloads a few seconds later - waiting and verifying the new version ...')
            report('reloading', 'Install finished - waiting for the switch to reload ...')
        else:
            log(device, 'Install finished but the session did not drop - '
                        'no reload happened, check the device state manually!')
        return 'CHECK MANUALLY'

    try:
        if device_type != 'cisco_xe':
            log(device, 'This tool upgrades IOS-XE switches in INSTALL mode only - device type {!r} is not supported'
                .format(device['type']))
            status = 'FAILED (only IOS-XE install mode is supported)'
            return

        # ---- Connect to the device (credentials are never printed)
        report('connecting', 'Connecting to {} ...'.format(device['ipaddr']))
        log(device, 'Connecting to device {}, username={}'.format(device['ipaddr'], username))
        connect_args = dict(device_type=device_type, ip=device['ipaddr'],
                            username=username, password=password, keepalive=30)
        if secret:
            connect_args['secret'] = secret
        session = ConnectHandler(**connect_args)
        if secret:
            session.enable()

        # ---- Platform, version and boot mode (read fresh; the dashboard values are only a hint)
        show_version = session.send_command('show version', read_timeout=tmo('show_cmd'))
        old_version, model = parse_version(show_version), parse_model(show_version)
        series, flavor = model_series(model), install_flavor(model, parse_version(show_version))
        log(device, 'Running software version: {}, platform: {}'.format(old_version or 'unknown', model or 'unknown'))
        report('connected', 'Connected - {} running {}'.format(model or 'switch', old_version or 'unknown'),
               old_version=old_version, model=model)
        if series in ('9200', '9300'):
            log(device, 'Catalyst {0} - following the Cisco "Upgrade Catalyst {0} Switches" guide, INSTALL mode '
                        '({1})'.format(series, 'request platform software package ... commands for this old release'
                                       if flavor == 'legacy' else 'install commands'))
        else:
            log(device, 'Platform {} - using the generic IOS-XE INSTALL mode commands'.format(model or 'unknown'))
        if flavor == 'legacy' and mode != 'full':
            log(device, 'Release {} only supports the one-step "request platform software package install" '
                        'upgrade - use the normal Upgrade, not Distribute/Activate'.format(old_version))
            status = 'FAILED (this release does not support distribute/activate)'
            return
        try:
            boot, raw_boot = read_boot_state(session, show_version)
        except Exception as exc:  # noqa: BLE001 - handled below: the boot variable is then simply set
            boot = None
            log(device, 'WARNING: could not read "show boot": {}: {}'.format(type(exc).__name__, exc))
        if boot:
            report_boot(boot)
            log(device, 'Boot check: {} mode, {}'.format(boot['mode'] or 'unknown', boot['detail']))
            if boot['mode'] == 'BUNDLE':
                log(device, 'The switch runs in BUNDLE mode - this upgrade converts it to INSTALL mode')

        if mode == 'activate':
            take_pre_snapshot(session, device)
            boot = ensure_boot(boot)
            report('installing', 'Activating the previously distributed image ...')
            log(device, 'Activating the already-distributed image. Reload will follow!')
            status = run_install('activate', None, ACTIVATE_FAIL_RE)
            return

        guessed_version = guess_version_from_filename(image_name)
        if old_version and guessed_version and old_version == guessed_version:
            log(device, 'The image file name suggests version {} - the switch is already running '
                        'that version. Skipping this switch (no upload, no install, no reload).'
                        .format(guessed_version))
            report('skipped', 'Already running version {} - no upgrade needed'.format(guessed_version))
            status = 'SKIPPED (already running this version)'
            return

        if mode == 'full':
            take_pre_snapshot(session, device)

        # ---- Cisco step 1: remove inactive packages (frees flash before anything else is done)
        report('cleaning', 'Removing inactive packages (can take several minutes) ...')
        log(device, 'Removing inactive packages ({}) - this can take several minutes ...'.format(
            'request platform software package clean switch all' if flavor == 'legacy' else 'install remove inactive'))
        try:
            clean_output = (clean_packages(session, device) if flavor == 'legacy'
                            else remove_inactive(session, device))
        except ReadTimeout:
            log(device, 'TIMED OUT during package clean-up - stopping for this switch')
            status = 'FAILED (clean-up timeout)'
            return
        log(device, 'Clean-up output:\n{}'.format(clean_output))

        # ---- flash: image + 1-1.5 GB for the expansion, on every stack member
        report('space_check', 'Checking free space on flash (image + 1-1.5 GB expansion) ...')
        log(device, 'Verifying sufficient space on the file system (image + 1-1.5 GB for the expansion) ...')
        space = check_flash_space(session, os.path.getsize(local_path), image_name)
        log(device, space['detail'])
        skip_upload = False
        if space['exists']:
            log(device, 'File already uploaded on device ...')
            if mode == 'activate':
                status = 'SKIPPED (image already on flash)'
                return
            # A .bin on flash is NOT a distributed image: only 'install add' stages it (it then shows in
            # 'show install summary'). The file may be there from an interrupted run or a manual copy, so
            # skip just the upload, verify the file, and still run the install command for this mode.
            skip_upload = True
            log(device, 'The file is on flash but is not necessarily added (staged) yet - skipping the upload, '
                        'checking the file and running the install ...')
        if space['level'] == 'fail':
            if ignore_space:
                log(device, 'WARNING: not enough free flash by Cisco\'s rule - the operator chose to IGNORE the flash '
                            'space check, so the upgrade continues. The install may still fail if flash really runs out.')
            else:
                log(device, 'Not enough free space on device ...')
                status = 'FAILED (not enough space)'
                return
        if space['level'] == 'low':
            log(device, 'WARNING: free flash is below the 1.5 GB Cisco recommends - continuing, the install '
                        'may still succeed')

        # ---- transferring file to device
        if not skip_upload:
            report('uploading', 'Uploading {} to flash (SCP) ...'.format(image_name))
            log(device, 'Success! - proceed with image upload')
            log(device, 'Uploading file : {} ...'.format(image_name))
            transfer_file(session, local_path, image_name)
            log(device, 'Success! - upload file: {} was successfull ...'.format(image_name))

        # ---- verifying md5
        report('md5', 'Verifying MD5 checksum on the switch ...')
        log(device, 'Verifying md5 checksum on device ...')
        session = ensure_session(session, connect_args, secret, device)
        try:
            md5_ok = verify_md5(session, image_name, md5)
        except CONNECTION_LOST:
            log(device, 'Session lost during the MD5 check - reconnecting and retrying once ...')
            session = reconnect(session, connect_args, secret)
            md5_ok = verify_md5(session, image_name, md5)
        if not md5_ok:
            log(device, 'Error verifying md5 checksum on device, quitting !!!')
            status = 'FAILED (md5 mismatch)'
            return

        if mode == 'distribute':
            report('installing', 'Adding the image to flash (no activation, no reload) ...')
            log(device, 'MD5 check OK - adding the image to flash (distribute only, no reload) ...')
            if add_image():
                if not (boot and boot.get('boot_ok')):
                    log(device, 'NOTE: the boot variable is not set for install mode - it is set when the '
                                'image is activated')
                log(device, 'Image added to flash. Not activated - no reload has happened.')
                status = 'DISTRIBUTED'
            else:
                status = 'FAILED (distribute)'
            return

        # ---- boot variable + install (mode == 'full')
        log(device, 'MD5 Check... Success! - Starting installing, activating and commiting '
                    'new Image. Reload will follow!')
        if flavor == 'legacy':
            boot = ensure_boot(boot)
            report('installing', 'Installing the new image (request platform software package install) ...')
            status = run_install('legacy', image_name, INSTALL_FAIL_RE)
        elif not (boot and boot.get('boot_ok')) and not space['has_packages_conf']:
            # BUNDLE mode: flash has no packages.conf yet, so pointing the boot variable at it NOW
            # could leave a switch that cannot boot if anything went wrong. Add the image first (that
            # creates packages.conf), then set the boot variable, then activate.
            report('installing', 'Adding the image (install add) ...')
            log(device, 'No {} on flash yet (BUNDLE mode): adding the image first, then setting the boot '
                        'variable, then activating'.format(BOOT_TARGET))
            if not add_image():
                status = 'FAILED (install add)'
                return
            listing = session.send_command('dir {}'.format(BOOT_TARGET), read_timeout=tmo('show_cmd'))
            if not re.search(r'packages\.conf', listing) or re.search(r'%\s*Error|No such file', listing):
                log(device, 'install add finished but {} is still missing on flash - NOT activating'.format(BOOT_TARGET))
                status = 'FAILED (packages.conf missing after install add)'
                return
            boot = ensure_boot(boot)
            report('installing', 'Activating and committing the new image ...')
            status = run_install('activate', None, ACTIVATE_FAIL_RE)
        else:
            boot = ensure_boot(boot)
            report('installing', 'Installing, activating and committing the new image ...')
            status = run_install('full', image_name, INSTALL_FAIL_RE)

    except Exception as exc:                           # one bad device must not kill the others
        log(device, 'ERROR: {}: {}'.format(type(exc).__name__, exc))
        log(device, 'Details:\n{}'.format(traceback.format_exc()))
        status = 'FAILED ({})'.format(type(exc).__name__)
    finally:
        if session is not None:
            try:
                session.disconnect()
            except Exception:
                pass                                   # session is already gone after a reload
        results[device['ipaddr']] = status

#==============================================================================
# ---- Web GUI mode: one switch, non-interactive, waits for the reload
#==============================================================================
def commit_install(device, creds):
    """After 'install activate' and the reload: 'install commit', then check 'show install summary'.
    Returns (ok, note). Without the commit the switch rolls the activation back when its timer ends."""
    ip = device['ipaddr']
    username, password = creds[0], creds[1]
    secret = creds[2] if len(creds) > 2 else None
    report('committing', 'Committing the new image (install commit) ...')
    log(device, 'Committing the activated image: install commit')
    session = None
    try:
        args = dict(device_type='cisco_xe', ip=ip, username=username, password=password)
        if secret:
            args['secret'] = secret
        session = ConnectHandler(**args)
        if secret:
            session.enable()
        out = session.send_command_timing('install commit', read_timeout=tmo('install_commit'))
        for line in out.splitlines():
            if line.strip():
                log(device, '  install> {}'.format(line.rstrip()))
        if INSTALL_FAIL_RE.search(out):
            return False, 'install commit reported a failure'
        summary = session.send_command('show install summary', read_timeout=tmo('show_cmd'))
        if re.search(r'(?m)^\s*IMG\s+U\b', summary):
            return False, 'the image is still shown as activated but not committed'
        log(device, 'Commit OK - the new image is committed')
        return True, 'committed'
    except Exception as exc:  # noqa: BLE001
        return False, 'install commit could not run ({})'.format(type(exc).__name__)
    finally:
        if session is not None:
            try:
                session.disconnect()
            except Exception:
                pass


def wait_for_reboot(device, creds, old_version):
    """After 'install ... activate commit' the switch reloads. Wait until it stops answering
    ping, answers again and accepts a login, then read the NEW software version.

    Returns (state, message, new_version); state is completed | attention.
    """
    ip = device['ipaddr']
    username, password = creds[0], creds[1]
    secret = creds[2] if len(creds) > 2 else None

    # ---- 1. wait for the switch to go down
    report('waiting_down', 'Waiting for the switch to stop answering ping (reload) ...')
    log(device, 'Waiting for the switch to go down (ping stops) ...')
    started = time.time()
    misses = 0
    went_down = False
    while time.time() - started < tmo('reboot_down'):
        if is_reachable(ip):
            misses = 0
        else:
            misses += 1
            if misses >= 2:
                went_down = True
                break
        time.sleep(REBOOT_POLL)

    # ---- 2. wait for the switch to come back
    if went_down:
        down_at = time.time()
        log(device, 'Switch is down (reloading). Waiting for it to answer ping again ...')
        report('waiting_up', 'Switch is down - waiting for it to come back (usually 5-15 min) ...')
        hits = 0
        last_note = time.time()
        while True:
            if time.time() - down_at > tmo('reboot_up'):
                msg = 'Switch did not answer ping again within {} minutes'.format(tmo('reboot_up') // 60)
                log(device, msg)
                return 'attention', msg, None
            if is_reachable(ip):
                hits += 1
                if hits >= 2:
                    break
            else:
                hits = 0
            if time.time() - last_note >= 30:
                elapsed = int(time.time() - down_at)
                note = 'Switch still down ({}m {:02d}s) ...'.format(elapsed // 60, elapsed % 60)
                log(device, note)
                report('waiting_up', note)
                last_note = time.time()
            time.sleep(REBOOT_POLL)
        elapsed = int(time.time() - down_at)
        log(device, 'Switch answers ping again after {}m {:02d}s'.format(elapsed // 60, elapsed % 60))
    else:
        log(device, 'WARNING: the switch never stopped answering ping within {} s - '
                    'checking the software version anyway'.format(tmo('reboot_down')))

    # ---- 3. wait until SSH accepts a login, then read the version
    report('waiting_ssh', 'Switch answers ping - waiting for SSH login ...')
    log(device, 'Waiting for SSH login ...')
    started = time.time()
    connected = False
    new_version = None
    new_boot = None
    last_error = ''
    while time.time() - started < tmo('reboot_ssh'):
        session = None
        try:
            connect_args = dict(device_type='cisco_xe', ip=ip, username=username, password=password)
            if secret:
                connect_args['secret'] = secret
            session = ConnectHandler(**connect_args)
            if secret:
                session.enable()
            connected = True
            show_version = None
            try:
                show_version = session.send_command('show version', read_timeout=tmo('show_cmd'))
            except Exception:                          # a version lookup must never stop the check
                pass
            new_version = parse_version(show_version) if show_version else None
            if show_version:
                try:                                   # the "Boot mode" column follows the new state
                    new_boot, _raw = read_boot_state(session, show_version)
                    report_boot(new_boot)
                except Exception as exc:  # noqa: BLE001
                    log(device, 'WARNING: could not read "show boot" after the reload: {}'.format(type(exc).__name__))
            break
        except Exception as exc:                       # SSH/AAA are often not ready right after boot
            last_error = type(exc).__name__
            time.sleep(REBOOT_SSH_RETRY)
        finally:
            if session is not None:
                try:
                    session.disconnect()
                except Exception:
                    pass

    if not connected:
        msg = 'Switch is back but SSH login failed within {} minutes ({})'.format(
            tmo('reboot_ssh') // 60, last_error or 'no answer')
        log(device, msg)
        return 'attention', msg, None
    if not new_version:
        msg = 'Switch is back but the software version could not be read'
        log(device, msg)
        return 'attention', msg, None

    log(device, 'Software version after reload: {}'.format(new_version))
    if old_version and new_version == old_version:
        msg = 'Switch is back but the software version is unchanged ({})'.format(new_version)
        log(device, 'WARNING: ' + msg)
        return 'attention', msg, new_version
    mode_txt = new_boot['mode'] if new_boot and new_boot.get('mode') else None
    if new_boot:
        log(device, 'Boot mode after reload: {} ({})'.format(mode_txt or 'unknown', new_boot['detail']))
    if new_boot and (mode_txt != 'INSTALL' or new_boot.get('boot_ok') is False):
        msg = ('Upgraded {} -> {}, but the switch is NOT fully in INSTALL mode: {} mode, {}'
               .format(old_version or 'unknown', new_version, mode_txt or 'unknown', new_boot['detail']))
        log(device, 'WARNING: ' + msg)
        return 'attention', msg, new_version
    msg = 'Upgrade completed: {} -> {}{}'.format(old_version or 'unknown', new_version,
                                                  ' ({} mode)'.format(mode_txt) if mode_txt else '')
    log(device, msg)
    return 'completed', msg, new_version


def watch_reload_timed(ip, username, password, secret, limits, emit=None, should_stop=None, extra=None):
    """The Command time test's version of wait_for_reboot: the same waits (ping stops, ping answers again, SSH login),
    but every one is TIMED and reported through emit(text). limits: seconds for 'down', 'up' and 'ssh'.
    -> {down, up, ssh (seconds or None), version, boot, stopped, error}"""
    emit = emit or (lambda text: None)
    stop = should_stop or (lambda: False)
    res = {"down": None, "up": None, "ssh": None, "version": None, "boot": None, "stopped": False, "error": "",
           "extra": {}}
    started, misses, went_down = time.time(), 0, False
    emit("Waiting for the switch to stop answering ping ...")
    while time.time() - started < limits["down"]:
        if stop():
            res["stopped"] = True
            return res
        misses = 0 if is_reachable(ip) else misses + 1
        if misses >= 2:
            went_down = True
            break
        time.sleep(REBOOT_POLL)
    if went_down:
        res["down"] = round(time.time() - started, 1)
        down_at, hits, last_note = time.time(), 0, time.time()
        emit("Ping stopped after {} s - the switch is reloading. Waiting for it to answer again ...".format(res["down"]))
        while True:
            if stop():
                res["stopped"] = True
                return res
            if time.time() - down_at > limits["up"]:
                res["error"] = "The switch did not answer ping again within {} s.".format(limits["up"])
                emit(res["error"])
                return res
            hits = hits + 1 if is_reachable(ip) else 0
            if hits >= 2:
                break
            if time.time() - last_note >= 30:
                emit("Still down ({} s) ...".format(int(time.time() - down_at)))
                last_note = time.time()
            time.sleep(REBOOT_POLL)
        res["up"] = round(time.time() - down_at, 1)
        emit("Ping answers again after {} s.".format(res["up"]))
    else:
        emit("The switch never stopped answering ping within {} s - checking the version anyway.".format(limits["down"]))
    emit("Waiting for SSH login ...")
    ssh_started, last_error = time.time(), ""
    while time.time() - ssh_started < limits["ssh"]:
        if stop():
            res["stopped"] = True
            return res
        session = None
        try:
            args = dict(device_type='cisco_xe', ip=ip, username=username, password=password)
            if secret:
                args['secret'] = secret
            session = ConnectHandler(**args)
            if secret:
                session.enable()
            res["ssh"] = round(time.time() - ssh_started, 1)
            show_version = None
            try:
                show_version = session.send_command('show version', read_timeout=tmo('show_cmd'))
            except Exception:  # noqa: BLE001
                pass
            res["version"] = parse_version(show_version) if show_version else None
            if show_version:
                try:
                    res["boot"], _raw = read_boot_state(session, show_version)
                except Exception:  # noqa: BLE001
                    pass
            emit("SSH login works after {} s. Software version: {}".format(res["ssh"], res["version"] or "unknown"))
            for command in (extra or []):                   # for example show install summary
                try:
                    res["extra"][command] = session.send_command(command, read_timeout=tmo('show_cmd')) or ""
                    emit("{}# {}\n{}".format(ip, command, res["extra"][command].rstrip()))
                except Exception:  # noqa: BLE001
                    pass
            return res
        except Exception as exc:  # noqa: BLE001 - SSH/AAA are often not ready right after boot
            last_error = type(exc).__name__
            time.sleep(REBOOT_SSH_RETRY)
        finally:
            if session is not None:
                try:
                    session.disconnect()
                except Exception:  # noqa: BLE001
                    pass
    res["error"] = "SSH login did not work within {} s ({}).".format(limits["ssh"], last_error or "no answer")
    emit(res["error"])
    return res
#------------------------------------------------------------------------------
def set_console_title(title):
    if os.environ.get('IOSXE_BACKGROUND_WORKER') or not sys.stdout.isatty():   # background worker: no window to title
        return
    title = re.sub(r'[\x00-\x1f\x7f]', '', title)
    try:
        if os.name == 'nt':
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW(title)
        else:
            sys.stdout.write('\033]0;{}\007'.format(title))
            sys.stdout.flush()
    except Exception:
        pass
#------------------------------------------------------------------------------
def pause_before_exit():
    """Keep a real console window open so the result can be read. Nothing to wait for when
    the worker runs in the background (output goes to the dashboard's web console)."""
    if os.environ.get('IOSXE_BACKGROUND_WORKER') or not (sys.stdin and sys.stdin.isatty()):
        return                                         # (Windows reports NUL as a terminal, hence the flag)
    print('\nThis window can be closed. Press Enter to close it ...')
    try:
        input()
    except (EOFError, KeyboardInterrupt, OSError):
        pass
#------------------------------------------------------------------------------
def job_main(job_path):
    """Upgrade ONE switch described by a job file written by switch_manager.py."""
    global STATUS, LOG_FILE
    for stream in (sys.stdout, sys.stderr):           # switch output must never crash the console
        try:                                           # line-buffered so the web console is live
            stream.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
        except Exception:
            pass

    job = None
    try:
        with open(job_path, 'r', encoding='utf-8') as fh:
            job = json.load(fh)
    except Exception as exc:
        print('Cannot read the job file: {}'.format(exc))
    finally:
        try:
            os.remove(job_path)                        # it contains the password: delete at once
        except OSError:
            pass
    if job is None:
        pause_before_exit()
        return 2

    LOG_FILE = job.get('log_path')
    apply_timeouts(job.get('timeouts'))             # the waits chosen on the dashboard's Settings page
    device = {'ipaddr': job['ip'], 'type': 'cisco-xe', 'name': job.get('name') or job['ip']}
    STATUS = StatusFile(job['status_path'])
    STATUS.update(state='running', phase='starting', message='Starting ...')
    STATUS.start_heartbeat()
    set_console_title('Upgrade {} ({})'.format(device['name'], device['ipaddr']))
    creds = (job['username'], job['password'], job.get('secret') or None)
    del job['password']
    set_host_protocol(job['ip'], 'ssh', job.get('port'))

    action = job.get('action') or 'full'          # 'full' (default) | 'distribute' | 'activate'
    ignore_space = bool(job.get('ignore_space'))  # operator chose to bypass the flash space check
    log(device, 'Worker started - action: {}'.format(action))
    state, message, new_version = 'failed', 'Unknown error', None
    reloaded = False                              # True only when a real reload + new version was verified
    try:
        if action == 'activate':
            # No image involved here at all - this activates whatever was already added to
            # flash by an earlier 'distribute' run.
            results = {}
            command_worker(device, creds, None, None, None, results, mode='activate')
            status = results.get(device['ipaddr'], 'UNKNOWN')
        else:
            image_path = job['image_path']
            if not os.path.isfile(image_path):
                raise RuntimeError('Image file not found: {}'.format(image_path))
            try:
                verify_image_checksum_sidecar(image_path)
            except ValueError as exc:
                raise RuntimeError('Image integrity check failed before worker execution: {}'.format(exc)) from exc
            image_name = os.path.basename(image_path)      # name on the switch's flash

            report('md5', 'Calculating the local MD5 of {} ...'.format(image_name))
            log(device, 'Calculating MD5 of {} ...'.format(image_name))
            md5 = check_md5(image_path)
            log(device, 'Local MD5: {}'.format(md5))

            results = {}
            command_worker(device, creds, image_path, image_name, md5, results, mode=action, ignore_space=ignore_space)
            status = results.get(device['ipaddr'], 'UNKNOWN')

        if status in ('OK (reloading)', 'CHECK MANUALLY'):
            # 'CHECK MANUALLY' = the install/activate command returned before the session
            # dropped. wait_for_reboot() sorts it out: it reports the new version, or 'version
            # unchanged' if the switch never reloaded. Applies to both 'full' and 'activate'.
            state, message, new_version = wait_for_reboot(
                device, creds, STATUS.data.get('old_version'))
            reloaded = (state == 'completed' and bool(new_version))
            if PENDING_COMMIT['needed']:
                if reloaded:
                    ok, note = commit_install(device, creds)
                    if ok:
                        message += ' - committed'
                    else:
                        state = 'attention'
                        message += ' - BUT {}: run "install commit" on the switch or it will roll back'.format(note)
                        log(device, 'WARNING: ' + message)
                elif new_version and new_version != STATUS.data.get('old_version'):
                    log(device, 'WARNING: the new image is active but was not committed - run "install commit" '
                                'on the switch or it will roll back')
        elif status == 'DISTRIBUTED' or status.startswith('DISTRIBUTED'):
            # Image added to flash only - no activation, no reload. This IS the successful
            # outcome for a 'distribute' run.
            state, message = 'completed', status
        elif status.startswith('SKIPPED (already running'):
            # Nothing was uploaded or installed - the switch already has the requested version.
            state, message = 'completed', 'Already running the requested version - no upgrade needed'
            new_version = STATUS.data.get('old_version')
        elif status.startswith('FAILED'):
            state, message = 'failed', status
        else:                                          # SKIPPED (image already on flash) / UNKNOWN
            state, message = 'attention', status
    except Exception as exc:
        state, message = 'failed', 'FAILED ({}: {})'.format(type(exc).__name__, exc)
        log(device, 'ERROR: {}'.format(message))

    if reloaded and action in ('full', 'activate'):
        try:                                       # post-upgrade check, shown in the device list
            STATUS.update(postcheck=run_postcheck(device, creds, new_version, PRE_SNAPSHOT))
        except Exception as exc:  # noqa: BLE001 - never turn a finished upgrade into a failure
            log(device, 'WARNING: the post-check itself failed: {}: {}'.format(type(exc).__name__, exc))
            STATUS.update(postcheck={'state': 'failed', 'checks': {}, 'time': now_str(),
                                     'detail': 'Post-check could not run ({})'.format(type(exc).__name__)})

    if SCP_STATE['added']:                         # undo the one configuration change the upgrade made
        ok, note = remove_scp_server(device, creds)
        log(device, ('SCP clean-up: ' if ok else 'WARNING: SCP clean-up: ') + note)
        if ok:
            message += ' - ip scp server setting removed again'
        else:
            message += ' - WARNING: ' + note
            if state == 'completed':
                state = 'attention'
    report('done', message, state=state, result=message, new_version=new_version)
    print('\n---- Result -----------------------------')
    print('{:<16} {:<20} {}'.format(device['ipaddr'], device['name'], message))
    pause_before_exit()
    return 0 if state == 'completed' else 1


#==============================================================================
# ---- Main: Run Commands
#==============================================================================
def interactive_main():
    image_path = input('Enter Image Filename: ').strip()
    if not os.path.isfile(image_path):
        sys.exit('Image file not found: {}'.format(image_path))
    image_name = os.path.basename(image_path)          # name on the switch's flash
    try:
        verify_image_checksum_sidecar(image_path)
    except ValueError as exc:
        sys.exit('Image integrity check failed: {}'.format(exc))

    username = input('Enter Username: ').strip()
    try:
        password = read_secret('Enter Password: ')
    except getpass.GetPassWarning:
        sys.exit('This terminal cannot hide password input - run the script in a real terminal.')
    if not username or not password:
        sys.exit('Username and password must not be empty.')

    devices_file = input('Enter switch list (.xlsx) location [switches.xlsx]: ').strip() or 'switches.xlsx'
    try:
        devices = read_devices(devices_file)
    except Exception as exc:                            # missing file, not an xlsx, bad headers ...
        sys.exit('Cannot read switch list: {}'.format(exc))
    if not devices:
        sys.exit('No cisco-xe switches found in {}'.format(devices_file))

    creds = (username, password)
    print('Calculating MD5 of {} ...'.format(image_name))
    md5 = check_md5(image_path)                        # once, not once per device
    print('Local MD5: {}'.format(md5))

    parallel, stop_on_failure = 1, True
    argv = sys.argv[1:]
    if '--parallel' in argv:
        try:
            parallel = max(1, min(MAX_PARALLEL_LIMIT, int(argv[argv.index('--parallel') + 1])))
        except (IndexError, ValueError):
            sys.exit('--parallel needs a number from 1 to {}'.format(MAX_PARALLEL_LIMIT))
    if '--continue-on-failure' in argv:
        stop_on_failure = False

    results = {}
    print('\n---- Upgrading {} switch(es) at a time; {} ----\n'.format(
        parallel, 'stopping after a wave with a failed upgrade' if stop_on_failure
        else 'continuing even if an upgrade fails'))
    items = list(devices.items())
    for start in range(0, len(items), parallel):
        wave = items[start:start + parallel]
        threads = []
        for ipaddr, device in wave:
            print('Starting upgrade for: ', device)
            threads.append(threading.Thread(
                target=command_worker, name=device['name'],
                args=(device, creds, image_path, image_name, md5, results)))
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        failed = [device['name'] for ipaddr, device in wave
                  if results.get(ipaddr, 'UNKNOWN').startswith('FAILED')]
        if failed and stop_on_failure and start + parallel < len(items):
            print('Stopping the rollout: {} failed. The remaining switches were not started.'.format(
                ', '.join(failed)))
            break

    print('\n---- Summary --------------------------')
    for ipaddr, device in devices.items():
        print('{:<16} {:<20} {}'.format(ipaddr, device['name'], results.get(ipaddr, 'UNKNOWN')))




# =============================================================================
# Part 2 of 2 - web GUI: availability monitor, backup runner, upgrade scheduler/launcher
# =============================================================================

MIN_DURATION_MIN = 10         # minimum monitoring time when --duration is used
BACKUP_WORKERS = 10           # parallel backup connections
MAX_CONCURRENT_RELOADS = 1    # DEFAULT number of switches upgraded (reloaded) at the same time; an Administrator
                              # can change it in the GUI (saved per site); the rest of a batch waits in a queue
MAX_PARALLEL_LIMIT = 10       # the most an Administrator can set
STALE_SECONDS = 90            # no heartbeat from an upgrade worker for this long = lost
SCHEDULE_POLL = 2             # seconds between checks of pending scheduled upgrades
MIN_SCHEDULE_LEAD = 30        # a scheduled time must be at least this far in the future
ACTIVE = ("starting", "running")
TERMINAL = ("completed", "failed", "attention")

ADDRESS_HEADERS = ["ip", "ip address", "ipaddress", "address", "host", "hostname", "device"]
LABEL_HEADERS = ["hostname", "name", "label", "description"]
TYPE_HEADERS = ["type", "device_type", "device type"]
PLATFORM_HEADERS = ["platform", "model", "device model", "switch model", "hardware"]
MAX_PLATFORM_LEN = 40
BACKUP_DEVICE_TYPES = {"cisco-ios": "cisco_ios", "cisco-xe": "cisco_xe"}
EXTRA_COMMANDS = [("show version", "show_version"), ("show boot", "show_boot"),
                  ("show inventory", "show_inventory")]

IMAGES_DIRNAME = "IOS Images"       # folder (next to the program) where uploaded images are kept
MAX_IMAGE_UPLOAD_BYTES = 2 * 1024 * 1024 * 1024    # 2 GB - generous for any IOS-XE image
MAX_CHECKSUM_UPLOAD_BYTES = 8192                   # includes multipart headers for a few lines of text
MAX_DEVICE_LIST_UPLOAD_BYTES = 10 * 1024 * 1024    # 10 MB - an .xlsx switch list is tiny


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def validate_ip_address(text):
    """The IP address in its normal written form, or an ApiError saying what is wrong with it."""
    text = str(text or "").strip()
    try:
        addr = ipaddress.ip_address(text)
    except ValueError:
        raise ApiError(f"'{text[:60]}' is not a valid IP address (for example 10.10.0.11).")
    if addr.is_unspecified or addr.is_multicast or addr.is_loopback:
        raise ApiError(f"{addr} cannot be the management address of a switch.")
    return str(addr)


class ApiError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def now_str():
    return datetime.now().strftime("%H:%M:%S")


def safe_name(text):
    return re.sub(r"[^A-Za-z0-9_.-]", "_", text)


def read_json_file(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


_VERSION_STORE_LOCK = threading.Lock()

# --------------------------------------------------------------------------- #
# Site profiles - everything under one site (device list, backups, reports,
# version history) lives in its own folder tree, so multiple sites never mix:
#
#   Switch Configuration Backups/
#     <Site A>/
#       devices.json              <- this site's device list (persists across restarts)
#       device_versions.json      <- this site's known-version store
#       Switch Backups/
#         <Hostname>_<ip-with-underscores>/     <- one PERMANENT folder per device
#           config_backup_<timestamp>.txt       <- every backup run adds ONE new timestamped file
#                                                  (all command outputs inside) instead of overwriting
#       Reports/
#         upgrade_report_<timestamp>.csv
#         backup_summary_<timestamp>.csv
#       Runs/                      <- firmware upgrade worker logs for this site
#     <Site B>/
#       ...
#
# IOS Images are NOT site-scoped - the "IOS Images" folder (see images_dir()) is shared
# across every site, since the same image is often used on more than one site's switches.
# --------------------------------------------------------------------------- #
SITES_ROOT_DIRNAME = "Switch Configuration Backups"
SITE_BACKUPS_SUBDIR = "Switch Backups"
SITE_REPORTS_SUBDIR = "Reports"
SITE_RUNS_SUBDIR = "Runs"

STORAGE_REGISTRY_KEY = r"Software\Innovate Nexus\NetUpgrade"


def _configured_storage_path(value_name, legacy_dirname):
    """Resolve an installer-selected data folder and retain data from older installs."""
    env_names = {
        "ImagesPath": "NETUPGRADE_IMAGES_DIR",
        "BackupsPath": "NETUPGRADE_BACKUPS_DIR",
    }
    configured = os.environ.get(env_names[value_name])
    if not configured and sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STORAGE_REGISTRY_KEY) as key:
                configured, _ = winreg.QueryValueEx(key, value_name)
        except (ImportError, OSError):
            configured = None

    legacy_path = this_program_path().parent / legacy_dirname
    if not configured:
        return legacy_path

    selected_path = Path(os.path.expandvars(os.path.expanduser(str(configured)))).resolve()
    if selected_path != legacy_path.resolve() and legacy_path.is_dir():
        if not selected_path.exists():
            # Copy once so an installer-selected location change does not hide existing user data.
            shutil.copytree(legacy_path, selected_path)
        elif selected_path.is_dir() and not any(selected_path.iterdir()):
            # The installer pre-creates its selected folders, so migrate into an empty one.
            shutil.copytree(legacy_path, selected_path, dirs_exist_ok=True)
    return selected_path


# --------------------------------------------------------------------------- #
# Global settings (the gear icon page). One settings.json next to the program, shared by every site.
# --------------------------------------------------------------------------- #
SETTINGS_FILENAME = "settings.json"
# Every launch starts with a clean device list: the status columns (software version, boot mode, backup,
# readiness, image stage, upgrade, post-check) begin empty. The devices, the saved backups and the credential
# profiles of the site are always kept. This used to be a switch on the Settings page; it is now fixed.
FRESH_START = True
SETTINGS_GROUPS = [
    {"id": "rollout", "title": "Upgrade Rollout",
     "intro": "How a batch of switches is upgraded. These apply to the whole dashboard, not to one site."},
    {"id": "timeouts", "title": "Command Outputs",
     "intro": "How long the tool waits for a switch before it gives up. Raise a value when a production switch is slow "
              "to return a command's output or to save its configuration. The default is shown next to each value "
              "(click it to put the default back). A change applies to everything that starts after you save; an "
              "upgrade that is already running keeps the values it started with. Command duration profiles "
              "keep named sets of these values, and the Command time test can measure a switch to build one."},
    {"id": "notifications", "title": "Notifications",
     "intro": "E-mail copies of the notifications shown under the bell icon. Off until you switch it on. Gmail: "
              "smtp.gmail.com, port 587 (STARTTLS) and an App password (2-Step Verification > App passwords) - "
              "your normal password will not work. The password is stored encrypted and is never shown again."},
    {"id": "about", "title": "About",
     "intro": "Details of this installation and what each version added. The application is updated only from the "
              "Windows application (About tab > Check for new updates), never from this page."},
]


def _t_item(key, section, label, help_text, lo, hi):
    return {"key": key, "group": "timeouts", "section": section, "label": label, "help": help_text,
            "type": "int", "unit": "s", "default": TIMEOUT_DEFAULTS[key], "min": lo, "max": hi}


SETTINGS_ITEMS = [
    {"key": "max_parallel", "group": "rollout", "section": "", "label": "Switches upgraded at the same time",
     "help": "How many switches reload together (1 to 10). A bigger batch starts this many and queues the rest; each "
             "queued switch starts when a running one finishes.",
     "type": "int", "unit": "switches", "default": MAX_CONCURRENT_RELOADS, "min": 1, "max": MAX_PARALLEL_LIMIT},
    {"key": "stop_on_failure", "group": "rollout", "section": "",
     "label": "Stop starting more switches if one fails",
     "help": "If a switch fails or needs attention, switches still waiting in the queue are not started. Switches "
             "already running finish normally.",
     "type": "bool", "unit": "", "default": True},
    _t_item("ssh_connect", "Connection", "SSH connect and login",
            "Waiting for the TCP connection, the login prompt and the banner when the tool connects to a switch.", 5, 300),
    _t_item("show_cmd", "Command Outputs", "Show commands",
            "show version, show boot, show switch, dir flash: and show install summary.", 10, 1800),
    _t_item("show_run", "Command Outputs", "Backup: show running-config",
            "Large configurations on a busy switch can take minutes to print.", 30, 1800),
    _t_item("backup_cmd", "Command Outputs", "Backup: other commands",
            "Every other command saved by a backup, including the additional commands you list.", 10, 1800),
    _t_item("snapshot_cmd", "Command Outputs", "Pre-check and post-check snapshot commands",
            "The commands that capture the switch state before and after an upgrade.", 10, 1800),
    _t_item("health_cmd", "Command Outputs", "Health checks (CPU, memory, stack)",
            "show processes cpu, show processes memory and show switch during the readiness check.", 5, 600),
    _t_item("write_memory", "Saving And Installing", "Saving the configuration (write memory)",
            "Used before and after the install steps. A switch with a large configuration can need several minutes.", 30, 1800),
    _t_item("install_commit", "Saving And Installing", "Install commit after the reload",
            "The install commit command that finishes the upgrade once the switch is back.", 30, 3600),
    _t_item("md5_verify", "Saving And Installing", "Image MD5 check on the switch (verify /md5)",
            "Calculating the MD5 of the copied image on the switch's flash. Big images on slow flash take a while.", 60, 7200),
    _t_item("install_total", "Saving And Installing", "Whole install (add, activate, reload)",
            "How long the install may run before the tool reports a failure.", 600, 14400),
    _t_item("clean_total", "Saving And Installing", "Package clean-up (install remove inactive)",
            "How long removing inactive packages may take.", 300, 14400),
    _t_item("reboot_down", "Reload And Recovery", "Switch to stop answering ping",
            "After the install command, how long to wait for the switch to go down.", 60, 3600),
    _t_item("reboot_up", "Reload And Recovery", "Switch to answer ping again",
            "How long to wait for the switch to come back. Stacks boot slowly.", 300, 7200),
    _t_item("reboot_ssh", "Reload And Recovery", "SSH to accept a login again",
            "After it answers ping, how long to keep trying to log in.", 60, 3600),
]


def settings_defaults():
    return {i["key"]: i["default"] for i in SETTINGS_ITEMS}


def settings_path():
    return this_program_path().parent / SETTINGS_FILENAME


def clean_settings(data, strict=False):
    """Every known setting, valid. strict=True (a save from the page) raises ApiError for a bad value; otherwise
    (a damaged or old file) a bad or missing value silently becomes its default."""
    data = data if isinstance(data, dict) else {}
    out = {}
    for item in SETTINGS_ITEMS:
        key, default = item["key"], item["default"]
        if key not in data:
            out[key] = default
            continue
        value = data[key]
        if item["type"] == "bool":
            if not isinstance(value, bool):
                if strict:
                    raise ApiError(f"{item['label']} must be on or off.")
                value = default
            out[key] = value
            continue
        n = None
        if not isinstance(value, bool) and isinstance(value, (int, str)):
            try:
                n = int(value)
            except ValueError:
                n = None
        if n is None or not item["min"] <= n <= item["max"]:
            if strict:
                raise ApiError(f"{item['label']} must be a whole number from {item['min']} to {item['max']}.")
            n = default
        out[key] = n
    return out


def load_global_settings():
    return clean_settings(read_json_file(settings_path()))


def save_global_settings(values):
    path = settings_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(values, indent=2), encoding="utf-8")
    tmp.replace(path)

# --------------------------------------------------------------------------- #
# Command duration profiles (Settings > Command Outputs). A profile is a named set of the command timeouts above.
# "Default" is built in and always there; the others are kept in command_profiles.json next to the program.
# Applying a profile copies its values into the settings (and saves them), so it works like typing them in.
# The Command time test on the same page runs the preset commands below on one switch, times each of them and
# can turn the measured times into a new profile.
# --------------------------------------------------------------------------- #
TIMEOUT_PROFILES_FILENAME = "command_profiles.json"
DEFAULT_TIMEOUT_PROFILE = "Default"
TIMEOUT_PROFILE_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._-]{0,39}$")
_TIMEOUT_PROFILES_LOCK = threading.Lock()
TIMEOUT_ITEMS = {i["key"]: i for i in SETTINGS_ITEMS if i["group"] == "timeouts"}

# The steps of the Command time test, in the order a manual upgrade of ONE switch runs them (the browser sends only
# the step id, never a command line). "keys" are the timeout settings a step's duration tells something about.
# kind: "command" (one command), "stream" (streamed and answered like the real upgrade does, can be stopped),
# "bundle" (several commands in one login, timed one by one and in total) or a special step (upload, md5, boot,
# activate). "warn" marks the steps that change the switch; the GUI asks for confirmation first.
_PRECHECK_STEPS = ([("show version", "show_cmd"), ("show boot", "show_cmd")]
                   + [(cmd, "snapshot_cmd") for _key, cmd in SNAPSHOT_COMMANDS]
                   + [("show processes cpu", "health_cmd"), ("show processes memory", "health_cmd"),
                      ("show switch", "health_cmd")])
_POSTCHECK_STEPS = ([("show version", "show_cmd"), ("show processes cpu", "health_cmd"),
                     ("show processes memory", "health_cmd"), ("show switch", "health_cmd")]
                    + [(cmd, "snapshot_cmd") for _key, cmd in SNAPSHOT_COMMANDS])
G_PRE, G_PREP, G_INST, G_INST1, G_POST = ("1. Pre-Check", "2. Prepare The Switch", "3. Install - Activate/Distribute",
                                          "3. Install - Full Upgrade", "4. Post-Check")
COMMAND_TEST_PRESETS = [
    {"id": "precheck_bundle", "title": "Pre-Check Bundle", "command": "show version, show boot, snapshot and health commands",
     "group": G_PRE, "kind": "bundle", "steps": _PRECHECK_STEPS,
     "desc": "Everything the upgrade reads before it changes anything. Each command and the total are timed."},
    {"id": "dir_flash", "title": "Check free space", "command": "dir flash:", "group": G_PREP, "kind": "command",
     "keys": ("show_cmd",), "desc": "The flash needs room for the image plus about 1.5 GB for the expansion."},
    {"id": "install_remove_inactive", "title": "Remove inactive packages", "command": "install remove inactive",
     "group": G_PREP, "kind": "stream", "keys": ("clean_total",),
     "warn": "removes the inactive software packages from the flash of {host} (the upgrade does this as its first step).",
     "desc": "Frees the flash. This is the step that can stay silent for a long time - the live output shows what the "
             "switch prints; use Stop if nothing happens."},
    {"id": "upload_image", "title": "Copy the image to the flash", "command": "SCP upload of the image", "group": G_PREP,
     "kind": "upload", "keys": (), "needs_image": True,
     "warn": "copies the selected image to the flash of {host}. 'ip scp server enable' is switched on for the copy and "
             "removed again afterwards if the switch did not have it.",
     "desc": "Needs the image chosen above. Not a timeout setting - the time is for your information."},
    {"id": "verify_md5", "title": "Check the image MD5", "command": "verify /md5 flash:<image> <md5>", "group": G_PREP,
     "kind": "md5", "keys": ("md5_verify",), "needs_image": True,
     "desc": "Compares the file on the flash with the checksum of the image on this PC."},
    {"id": "boot_variable", "title": "Set the boot variable", "command": " / ".join(BOOT_FIX_COMMANDS),
     "group": G_PREP, "kind": "boot", "keys": (),
     "warn": "changes the boot variable of {host} to " + BOOT_TARGET + " (install mode); it is not saved until write memory.",
     "desc": "Configuration commands, then show boot to confirm."},
    {"id": "write_memory", "title": "Save the configuration", "command": "write memory", "group": G_PREP, "kind": "command",
     "keys": ("write_memory",), "warn": "saves the running configuration of {host} to its startup configuration.",
     "desc": "Slow on some switches - this is the one that broke an upgrade before."},
    {"id": "install_add", "title": "Add the image", "command": "install add file flash:<image>", "group": G_INST,
     "kind": "stream", "keys": ("install_total",), "needs_image": True, "fail_re": "INSTALL_FAIL_RE",
     "warn": "adds (expands) the selected image on {host}. Nothing is activated and the switch does not reload.",
     "desc": "Expands the packages. Together with the activation it makes up 'Whole install'."},
    {"id": "install_activate", "title": "Activate (the switch reloads)", "command": "install activate", "group": G_INST,
     "kind": "activate", "keys": ("install_total", "reboot_down", "reboot_up", "reboot_ssh"),
     "warn": "ACTIVATES the added image and RELOADS {host}. The switch is offline for several minutes and everything "
             "connected to it loses connectivity.",
     "desc": "The switch reloads. The test then times how long ping stays down, how long until ping and SSH work again "
             "and checks the software version it came back with."},
    {"id": "install_commit", "title": "Commit", "command": "install commit", "group": G_INST, "kind": "stream",
     "keys": ("install_commit",), "fail_re": "INSTALL_FAIL_RE",
     "warn": "commits the activated image on {host}. Without it the switch rolls the activation back.",
     "desc": "Run it once the switch is back after the activation."},
    {"id": "install_full", "title": "Add, activate and commit (one command)",
     "command": "install add file flash:<image> activate commit", "group": G_INST1, "kind": "activate",
     "keys": ("install_total", "reboot_down", "reboot_up", "reboot_ssh"), "needs_image": True,
     "warn": "ADDS, ACTIVATES and COMMITS the selected image on {host} with ONE command and RELOADS it. The switch is "
             "offline for several minutes and everything connected to it loses connectivity.",
     "desc": "The one-command path the upgrade normally uses. Use it INSTEAD of the three-step path above (Add, Activate, "
             "Commit) - running both would upgrade the switch twice. After the reload the test checks the version and "
             "whether the image is committed."},
    {"id": "postcheck_bundle", "title": "Post-Check Bundle", "command": "version, health and snapshot commands",
     "group": G_POST, "kind": "bundle", "steps": _POSTCHECK_STEPS,
     "desc": "What the upgrade checks after the reload. Each command and the total are timed."},
]
for _p in COMMAND_TEST_PRESETS:
    _p.setdefault("keys", ())
    if _p["kind"] == "bundle":
        _p["keys"] = tuple(dict.fromkeys(k for _c, k in _p["steps"]))
# upgrade phase (as the worker reports it) -> the timeout setting it tells something about
PHASE_TIMEOUT_KEYS = {"cleaning": "clean_total", "md5": "md5_verify", "installing": "install_total",
                      "reloading": "install_total", "committing": "install_commit", "waiting_down": "reboot_down",
                      "waiting_up": "reboot_up", "waiting_ssh": "reboot_ssh"}
COMMAND_TEST_OUTPUT_LIMIT = 20000      # characters of a command's output sent back to the browser
COMMAND_TEST_BAD_OUTPUT = re.compile(r"(?m)^\s*% ?(Invalid input|Ambiguous command|Incomplete command|Unknown command)")


def clean_timeout_values(values, strict=False):
    """Every command timeout, valid. strict=True raises ApiError for a bad value; otherwise a bad or missing
    value becomes its default (a damaged or old file)."""
    values = values if isinstance(values, dict) else {}
    out = {}
    for key, item in TIMEOUT_ITEMS.items():
        raw, n = values.get(key, item["default"]), None
        if isinstance(raw, float) and raw.is_integer():
            raw = int(raw)
        if not isinstance(raw, bool) and isinstance(raw, (int, str)):
            try:
                n = int(raw)
            except ValueError:
                n = None
        if n is None or not item["min"] <= n <= item["max"]:
            if strict:
                raise ApiError(f"{item['label']} must be a whole number from {item['min']} to {item['max']}.")
            n = item["default"]
        out[key] = n
    return out


def timeout_profiles_path():
    return this_program_path().parent / TIMEOUT_PROFILES_FILENAME


def load_timeout_profiles():
    """{"profiles": {name: record}, "last_applied": name}. The built-in Default is not stored here."""
    data = read_json_file(timeout_profiles_path())
    data = data if isinstance(data, dict) else {}
    raw = data.get("profiles") if isinstance(data.get("profiles"), dict) else {}
    profiles = {}
    for name, rec in raw.items():
        if (not isinstance(name, str) or not TIMEOUT_PROFILE_NAME_RE.match(name)
                or name.lower() == DEFAULT_TIMEOUT_PROFILE.lower() or not isinstance(rec, dict)):
            continue
        profiles[name] = {"values": clean_timeout_values(rec.get("values")),
                          "note": str(rec.get("note") or "")[:300], "saved_at": str(rec.get("saved_at") or "")[:30],
                          "saved_by": str(rec.get("saved_by") or "")[:60], "site": str(rec.get("site") or "")[:80],
                          "source": "command test" if rec.get("source") == "command test" else "manual",
                          "switch": str(rec.get("switch") or "")[:80]}
    last = data.get("last_applied")
    return {"profiles": profiles, "last_applied": last if isinstance(last, str) else ""}


def save_timeout_profiles(store):
    path = timeout_profiles_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps({"profiles": store["profiles"], "last_applied": store.get("last_applied", "")}, indent=2),
                   encoding="utf-8")
    tmp.replace(path)


def image_checksum_kind(image_path):
    for ext, label in ((".md5", "MD5"), (".sha512", "SHA-512"), (".sha256", "SHA-256")):
        if Path(str(image_path) + ext).exists():
            return label
    return ""


def list_image_details():
    """[{name, size, checksum}] for the Storage page's image list, newest first."""
    out = []
    for name in list_images():
        p = images_dir() / name
        try:
            size = p.stat().st_size
        except OSError:
            continue
        out.append({"name": name, "size": size, "checksum": image_checksum_kind(p)})
    return out


class SiteContext:
    """The currently active site profile. Mutated only by Manager.switch_site(); read by
    version_store_path() so every part of the program - including code that doesn't have a
    Manager reference - saves to the right site's store."""
    def __init__(self):
        self.name = None


SITE = SiteContext()


def sites_root():
    d = _configured_storage_path("BackupsPath", SITES_ROOT_DIRNAME)
    d.mkdir(parents=True, exist_ok=True)
    return d


def safe_site_name(name):
    """A site name typed by the user, made safe to use as a folder name. Raises ValueError
    (caught by the caller and turned into an ApiError) if nothing usable is left."""
    name = (name or "").strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)   # strip characters a folder name can't contain
    name = name.strip(" .")                     # Windows disallows a trailing dot/space too
    if not name:
        raise ValueError("Enter a site profile name.")
    return name[:100]


def list_site_profiles():
    """Every site profile that has a folder on disk, newest-first by last activity."""
    root = sites_root()
    sites = [p for p in root.iterdir() if p.is_dir()]
    sites.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return [p.name for p in sites]


def site_dir(name):
    return sites_root() / safe_site_name(name)


def site_devices_path(name):
    return site_dir(name) / "devices.json"


def site_backups_dir(name):
    return site_dir(name) / SITE_BACKUPS_SUBDIR


def site_reports_dir(name):
    return site_dir(name) / SITE_REPORTS_SUBDIR


def site_runs_dir(name):
    return site_dir(name) / SITE_RUNS_SUBDIR


def site_history_path(name):
    return site_dir(name) / "upgrade_history.json"


HISTORY_MAX = 2000                 # most recent finished runs kept per site
_HISTORY_LOCK = threading.Lock()


def load_history(name):
    """Finished upgrade runs of a site, oldest first (survives restarts; used by the reports)."""
    data = read_json_file(site_history_path(name))
    return data if isinstance(data, list) else []


def append_history(name, record):
    with _HISTORY_LOCK:
        items = load_history(name)
        items.append(record)
        path = site_history_path(name)
        tmp = str(path) + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(items[-HISTORY_MAX:], fh)
        os.replace(tmp, path)


def compute_image_stage(history):
    """Per switch, from a site's permanent run history (oldest first):
         staged   - the image this tool distributed to the switch and has not activated yet
         upgraded - the last upgrade/activation that completed
       Only what this tool did is known; a change made on the switch by hand is not seen."""
    out = {}
    for r in history:
        ip = r.get("ip_address")
        if not ip or r.get("status") != "completed":
            continue
        entry = out.setdefault(ip, {"staged": None, "upgraded": None})
        action, message = r.get("action"), str(r.get("message") or "")
        if action == "Image distribution":
            if message.startswith("DISTRIBUTED"):          # not the 'already running this version' skip
                entry["staged"] = {"image": r.get("image_file") or "", "when": r.get("finished") or ""}
        elif action in ("Firmware upgrade", "Image activation"):
            if message.startswith("Already running"):      # skipped: nothing was installed
                continue
            prev = entry["staged"]
            entry["upgraded"] = {"image": r.get("image_file") or (prev and prev["image"]) or "",
                                 "version": r.get("current_image") or "", "when": r.get("finished") or ""}
            entry["staged"] = None                         # activating uses the staged image up
    return out


def device_folder_name(label, host):
    """'SwitchA', '10.10.10.10' -> 'SwitchA_10_10_10_10' - the permanent per-device backup
    folder name, dots replaced with underscores to match the requested layout exactly."""
    name = re.sub(r"[^A-Za-z0-9_-]", "_", (label or host).strip() or host)
    ip = re.sub(r"[^A-Za-z0-9_-]", "_", host)
    return f"{name}_{ip}"


def load_site_devices(name):
    """This site's saved device list: [{'host', 'label', 'type'}, ...]."""
    data = read_json_file(site_devices_path(name))
    return data if isinstance(data, list) else []


def save_site_devices(name, hosts):
    """Write-through so the device list survives a restart and shows up again the next time
    this site profile is selected. Best-effort: never blocks the action that triggered it."""
    try:
        site_dir(name).mkdir(parents=True, exist_ok=True)
        tmp = site_devices_path(name).with_suffix(".json.tmp")
        tmp.write_text(json.dumps(hosts, indent=2), encoding="utf-8")
        tmp.replace(site_devices_path(name))
    except OSError as exc:
        print(f"Could not save the device list for site '{name}': {exc}", flush=True)


def version_store_path():
    """Site-scoped once a site is active (the normal case); falls back to a path next to the
    program only if somehow called before any site has been selected."""
    if SITE.name:
        return site_dir(SITE.name) / "device_versions.json"
    return this_program_path().parent / "device_versions.json"


def load_known_versions():
    """{ip: {'hostname':..., 'version':..., 'updated_at': 'YYYY-MM-DD HH:MM:SS'}}, persisted to
    disk so a device's last-known software version survives restarting the program - the whole
    point being to have something to compare a new upgrade's result against."""
    data = read_json_file(version_store_path())
    return data if isinstance(data, dict) else {}


def save_known_version(ip, hostname, version, model=None, boot=None):
    """Record a device's current software version (plus its platform and boot mode, when known) to
    disk (overwrites any previous entry for that IP; a model/boot that is not passed is kept).
    Best-effort: a failure here never interrupts the backup/upgrade/check that found this version."""
    if not version:
        return
    with _VERSION_STORE_LOCK:
        try:
            data = load_known_versions()
            old = data.get(ip) or {}
            entry = {"hostname": hostname or ip, "version": version,
                     "updated_at": now_str() + " " + datetime.now().strftime("%Y-%m-%d")}
            model = model or old.get("model")
            boot = boot or old.get("boot")
            if model:
                entry["model"] = model
            if isinstance(boot, dict) and boot:
                entry["boot"] = boot
            data[ip] = entry
            tmp = version_store_path().with_suffix(".json.tmp")
            tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
            tmp.replace(version_store_path())
        except OSError as exc:
            print(f"Could not save the known-version store: {exc}", flush=True)


def images_dir():
    """The installer-selected folder for uploaded IOS images."""
    d = _configured_storage_path("ImagesPath", IMAGES_DIRNAME)
    d.mkdir(parents=True, exist_ok=True)
    return d


def list_images():
    """Image filenames available for upload/upgrade, newest first."""
    d = images_dir()
    files = [p for p in d.iterdir() if p.is_file() and p.suffix.lower() == ".bin"]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return [p.name for p in files]


def safe_upload_filename(name):
    """A plain filename for something saved into IOS_IMAGES - no path components, no traversal."""
    name = os.path.basename((name or "").replace("\\", "/")).strip()
    name = re.sub(r"[^A-Za-z0-9_.() -]", "_", name)
    return name


def read_multipart_file(rfile, content_length, content_type, max_bytes, dest_path_fn=None):
    """Parse a multipart/form-data body containing exactly one file field, streamed rather
    than buffered whole, so a large (hundreds of MB) upload doesn't have to fit in memory.

    If dest_path_fn is given, it is called with the uploaded filename once that's been parsed
    out of the body, and must return a filesystem path; the file's bytes are then streamed
    straight to that path (used for IOS images). If dest_path_fn is None, the bytes are
    collected in memory instead and returned (used for the small .xlsx device-list upload).
    Returns (filename, data_or_None, size_written, dest_path_or_None).
    Raises ValueError for anything malformed, missing, or over max_bytes.
    """
    m = re.search(r'boundary=(?:"([^"]+)"|([^;]+))', content_type or "")
    if not m:
        raise ValueError("Malformed upload (no multipart boundary).")
    delim = b"--" + (m.group(1) or m.group(2)).strip().encode()
    if content_length <= 0:
        raise ValueError("No file was uploaded.")
    if content_length > max_bytes:
        raise ValueError(f"File is too large (max {max_bytes // (1024 * 1024)} MB).")

    remaining = content_length
    CHUNK = 1024 * 1024

    def read_chunk():
        nonlocal remaining
        if remaining <= 0:
            return b""
        data = rfile.read(min(CHUNK, remaining))
        if not data:
            remaining = 0
            return b""
        remaining -= len(data)
        return data

    buf = b""
    while b"\r\n\r\n" not in buf:
        if len(buf) > 65536:
            raise ValueError("Malformed upload (headers too large).")
        chunk = read_chunk()
        if not chunk:
            raise ValueError("Upload ended unexpectedly while reading its headers.")
        buf += chunk
    head, _, data = buf.partition(b"\r\n\r\n")
    fm = re.search(rb'filename="([^"]*)"', head)
    if not fm or not fm.group(1):
        raise ValueError("No file was selected.")
    filename = fm.group(1).decode("utf-8", "replace")

    dest_path = dest_path_fn(filename) if dest_path_fn else None
    out = open(dest_path, "wb") if dest_path else io.BytesIO()
    written = 0
    tail_keep = len(delim) + 8  # room for "\r\n" + delimiter + a possible trailing "--"
    try:
        while True:
            idx = data.find(b"\r\n" + delim)
            if idx != -1:
                out.write(data[:idx])
                written += idx
                break
            chunk = read_chunk()
            if not chunk:
                raise ValueError("Upload ended unexpectedly (missing closing boundary).")
            data += chunk
            if written + len(data) > max_bytes + tail_keep:
                raise ValueError(f"File is too large (max {max_bytes // (1024 * 1024)} MB).")
            if len(data) > tail_keep * 2:
                safe = len(data) - tail_keep
                out.write(data[:safe])
                written += safe
                data = data[safe:]
    except BaseException:
        if dest_path:
            out.close()
            try:
                os.remove(dest_path)
            except OSError:
                pass
        raise
    if dest_path:
        out.close()
        return filename, None, written, dest_path
    return filename, out.getvalue(), written, None


# --------------------------------------------------------------------------- #
# Storage browser - read-only access to each site's "Switch Backups", "Reports" and "Runs"
# folder trees (the site's devices.json and other settings files are never exposed).
#
#   * Only <site>/Switch Backups/**, <site>/Reports/** and <site>/Runs/** can be reached; every path from the browser is validated
#     and must resolve to somewhere inside that folder (no "..", no absolute paths, no
#     symlinks that point elsewhere).
#   * Any site can be browsed without making it the active site, so looking at another
#     site's backups never stops monitoring or reloads the device list.
# --------------------------------------------------------------------------- #
STORAGE_PREVIEW_BYTES = 8 * 1024 * 1024       # how much of a file the "View" pop-up loads (a big stack's show run is >500 KB)
STORAGE_SEARCH_LIMIT = 500                    # max search hits returned
STORAGE_MAX_ZIP_BYTES = 1024 * 1024 * 1024    # refuse to zip more than 1 GB of backups
STORAGE_CHUNK = 1024 * 1024
STORAGE_FOLDERS = (SITE_BACKUPS_SUBDIR, SITE_REPORTS_SUBDIR, SITE_RUNS_SUBDIR)   # top level of the browser


def storage_root(site):
    """(site name, the site's folder) for a site that really exists. The browser only ever
    shows STORAGE_FOLDERS inside it - see storage_resolve()."""
    try:
        name = safe_site_name(site)
    except ValueError:
        raise ApiError("Choose a site.")
    if name not in list_site_profiles():
        raise ApiError("That site does not exist.", 404)
    return name, site_dir(name)


def storage_resolve(root, rel):
    """Turn a browser-supplied relative path into a real path that is guaranteed to sit
    inside `root`. Returns (real_path, clean_relative_path)."""
    rel = (rel or "").replace("\\", "/")
    parts = [p for p in rel.split("/") if p not in ("", ".")]
    for p in parts:
        if p == ".." or "\x00" in p or (os.name == "nt" and ":" in p):
            raise ApiError("Invalid path.")
    if parts and parts[0] not in STORAGE_FOLDERS:      # devices.json etc. are not browsable
        raise ApiError("Invalid path.")
    try:
        real_root = root.resolve()
        real = root.joinpath(*parts).resolve()
        real.relative_to(real_root)
    except (ValueError, OSError):
        raise ApiError("Invalid path.")
    return real, "/".join(parts)


def _storage_entry(entry_name, rel, st, is_dir, item_count=None):
    item = {"name": entry_name, "rel": rel, "type": "dir" if is_dir else "file",
            "mtime": st.st_mtime, "size": None if is_dir else st.st_size}
    if is_dir:
        item["items"] = item_count
    return item


def _storage_order(entries):
    """Folders first (A-Z), then files newest-first - newest backup on top."""
    dirs = sorted((e for e in entries if e["type"] == "dir"), key=lambda e: e["name"].lower())
    files = sorted((e for e in entries if e["type"] != "dir"), key=lambda e: e["mtime"], reverse=True)
    return dirs + files


def storage_list(site, rel):
    name, root = storage_root(site)
    path, rel = storage_resolve(root, rel)
    result = {"site": name, "path": rel, "entries": [], "exists": root.is_dir()}
    if not root.is_dir():
        if rel:
            raise ApiError("Folder not found.", 404)
        return result                       # site has no files yet
    if not path.is_dir():
        raise ApiError("Folder not found.", 404)
    entries = []
    try:
        with os.scandir(path) as it:
            for e in it:
                if not rel and e.name not in STORAGE_FOLDERS:
                    continue                # site root: only Switch Backups / Reports / Runs
                try:
                    if e.is_symlink():
                        continue
                    is_dir = e.is_dir()
                    st = e.stat()
                    count = None
                    if is_dir:
                        try:
                            count = len(os.listdir(e.path))
                        except OSError:
                            count = None
                except OSError:
                    continue
                entries.append(_storage_entry(e.name, (rel + "/" + e.name).strip("/"), st, is_dir, count))
    except OSError as exc:
        raise ApiError(f"Could not read that folder ({type(exc).__name__}).", 500)
    result["entries"] = _storage_order(entries)
    return result


def storage_search(site, query):
    """Case-insensitive search of every folder/file name under the site's backups. All the
    words typed must appear somewhere in the item's path, so 'core-sw show_run' finds the
    show_run files inside the core-sw folders."""
    name, root = storage_root(site)
    tokens = (query or "").lower().split()
    if not tokens:
        raise ApiError("Type something to search for.")
    hits, truncated = [], False
    if root.is_dir():
        base = str(root)
        for dirpath, dirnames, filenames in os.walk(base):
            dirnames[:] = [d for d in dirnames if not os.path.islink(os.path.join(dirpath, d))
                           and (dirpath != base or d in STORAGE_FOLDERS)]
            if dirpath == base:
                filenames = []              # files at the site root (devices.json ...) stay hidden
            rel_dir = os.path.relpath(dirpath, base).replace(os.sep, "/")
            rel_dir = "" if rel_dir == "." else rel_dir
            for is_dir, names in ((True, dirnames), (False, filenames)):
                for n in names:
                    full = os.path.join(dirpath, n)
                    rel = (rel_dir + "/" + n).strip("/")
                    hay = rel.lower()
                    if not all(t in hay for t in tokens):
                        continue
                    try:
                        if os.path.islink(full):
                            continue
                        st = os.stat(full)
                    except OSError:
                        continue
                    item = _storage_entry(n, rel, st, is_dir)
                    item["parent"] = rel_dir
                    hits.append(item)
                    if len(hits) >= STORAGE_SEARCH_LIMIT:
                        truncated = True
                        break
                if truncated:
                    break
            if truncated:
                break
    return {"site": name, "query": query, "entries": _storage_order(hits), "truncated": truncated}


def storage_read(site, rel, tail=False):
    """First STORAGE_PREVIEW_BYTES of a text file (or the LAST ones when tail is set), for the in-page viewer."""
    name, root = storage_root(site)
    path, rel = storage_resolve(root, rel)
    if not path.is_file():
        raise ApiError("File not found.", 404)
    try:
        with open(path, "rb") as fh:
            size = os.fstat(fh.fileno()).st_size
            skipped = bool(tail) and size > STORAGE_PREVIEW_BYTES
            if skipped:
                fh.seek(size - STORAGE_PREVIEW_BYTES)
            data = fh.read(STORAGE_PREVIEW_BYTES + 1)
    except OSError as exc:
        raise ApiError(f"Could not read that file ({type(exc).__name__}).", 500)
    if b"\x00" in data[:8192]:
        raise ApiError("This looks like a binary file and cannot be previewed - use Download.", 415)
    truncated = skipped or len(data) > STORAGE_PREVIEW_BYTES
    text = data[:STORAGE_PREVIEW_BYTES].decode("utf-8", errors="replace")
    if skipped and "\n" in text:
        text = text.split("\n", 1)[1]              # the cut usually lands mid-line: start at a whole line
    return {"name": path.name, "rel": rel, "size": size, "truncated": truncated, "tail": skipped,
            "limit": STORAGE_PREVIEW_BYTES, "text": text}


def _storage_backup_folder(site, rel):
    """The folder of ONE switch inside 'Switch Backups' (the only place backups can be compared)."""
    name, root = storage_root(site)
    path, rel = storage_resolve(root, rel)
    parts = rel.split("/") if rel else []
    if len(parts) != 2 or parts[0] != SITE_BACKUPS_SUBDIR or not path.is_dir():
        raise ApiError("Open a switch's folder under Switch Backups to compare its backups.")
    return path


def storage_backup_runs(site, rel):
    """The backup runs (newest first) saved in one switch's backup folder."""
    path = _storage_backup_folder(site, rel)
    return {"runs": [{"stamp": r["stamp"], "time": r["time"], "files": len(r["files"])}
                     for r in list_device_backups(path, limit=200)]}


def storage_compare(site, rel, a, b, ignore_volatile=True, cmd=None):
    return compare_backup_dir(_storage_backup_folder(site, rel), a, b, ignore_volatile, cmd)


def storage_zip_folder(site, rel):
    """ZIP a backup folder into a temp file. Returns (open file at 0, size, download name)."""
    name, root = storage_root(site)
    path, rel = storage_resolve(root, rel)
    if not path.is_dir():
        raise ApiError("Folder not found.", 404)
    base = path.name if rel else f"{name}_Storage"
    files, total = [], 0
    for dirpath, dirnames, filenames in os.walk(str(path)):
        dirnames[:] = [d for d in dirnames if not os.path.islink(os.path.join(dirpath, d))
                       and (dirpath != str(path) or rel or d in STORAGE_FOLDERS)]
        if dirpath == str(path) and not rel:
            filenames = []                  # site root: only the three folders
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            try:
                if os.path.islink(full):
                    continue
                total += os.path.getsize(full)
            except OSError:
                continue
            files.append(full)
    if not files:
        raise ApiError("That folder has no files to download.", 404)
    if total > STORAGE_MAX_ZIP_BYTES:
        raise ApiError("That folder is too large to download as one ZIP. Open it and download files individually.", 413)
    tmp = tempfile.TemporaryFile()
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for full in files:
                arc = base + "/" + os.path.relpath(full, str(path)).replace(os.sep, "/")
                try:
                    zf.write(full, arc)
                except ValueError:          # timestamp outside what ZIP can store
                    with open(full, "rb") as fh:
                        zf.writestr(arc, fh.read())
                except OSError:
                    continue                # file vanished or unreadable - skip it
        size = tmp.tell()
        tmp.seek(0)
    except Exception:
        tmp.close()
        raise
    return tmp, size, base + ".zip"


class KeepAwake(object):
    """Stops Windows from going to sleep while `wanted()` is true (an upgrade is running or a
    scheduled upgrade is waiting). Windows ties the request to the thread that makes it, so a
    dedicated thread holds it. Does nothing on other systems."""
    ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001

    def __init__(self, wanted, on_change=None, poll=10):
        self.wanted, self.on_change, self.poll = wanted, on_change, poll
        self.active = False
        self._stop = threading.Event()
        threading.Thread(target=self._run, daemon=True).start()

    def _apply(self, on):
        if os.name != "nt":
            return True
        try:
            import ctypes
            fn = ctypes.windll.kernel32.SetThreadExecutionState
            fn.argtypes = [ctypes.c_uint32]
            fn.restype = ctypes.c_uint32
            return bool(fn(self.ES_CONTINUOUS | (self.ES_SYSTEM_REQUIRED if on else 0)))
        except Exception:  # noqa: BLE001
            return False

    def _run(self):
        while not self._stop.is_set():
            try:
                want = bool(self.wanted())
            except Exception:  # noqa: BLE001
                want = self.active
            if want != self.active:
                ok = self._apply(want)
                self.active = want
                if self.on_change:
                    try:
                        self.on_change(want, ok)
                    except Exception:  # noqa: BLE001
                        pass
            elif want:
                self._apply(True)                    # refresh the request
            self._stop.wait(self.poll)
        self._apply(False)

    def stop(self):
        self._stop.set()


NOTES_MAX = 200                                   # notifications kept per site


class EventLog:
    """Thread-safe, newest-first list of things that happened (shown in the GUI)."""

    def __init__(self, maxlen=300):
        self._d = deque(maxlen=maxlen)
        self._lock = threading.Lock()

    def add(self, host, label, text, level="info"):
        ev = {"time": now_str(), "host": host, "label": label, "text": text, "level": level}
        with self._lock:
            self._d.appendleft(ev)
        print(f"[{ev['time']}] {host}: {text}", flush=True)

    def items(self, limit=120):
        with self._lock:
            return list(self._d)[:limit]


# --------------------------------------------------------------------------- #
# Excel loading
# --------------------------------------------------------------------------- #
# --------------------------------------------------------------------------- #
# E-mail notifications. Optional and OFF until switched on in the dashboard's Settings page
# (gear icon > Notifications). They are saved in email_settings.json next to the program, the
# password encrypted. Every dashboard
# notification goes through Manager._notify, which hands a copy to EMAILER; the e-mail is
# sent from a background thread so a slow mail server can never slow down an upgrade.
# --------------------------------------------------------------------------- #
EMAIL_BATCH_SECONDS = 20          # events close together are combined into ONE e-mail
EMAIL_MAX_RECIPIENTS = 20
EMAIL_SUBJECT_PREFIX = "[NetUpgrade]"
_EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


EMAIL_SETTINGS_FILENAME = "email_settings.json"
EMAIL_KINDS = [        # (key, label shown in Settings, ticked by default) - must match the categories Manager._notify sends
    ("started", "Started", False),
    ("completed", "Completed", True),
    ("failed", "Failed", True),
    ("attention", "Needs attention", True),
    ("postcheck", "Post-check result", True),
    ("missed", "Schedule missed", True),
    ("login", "Sign-ins", False),
]
EMAIL_PROBLEM_KINDS = ("failed", "attention", "missed")
EMAIL_SECURITY = {"starttls": "STARTTLS (port 587)", "ssl": "SSL (port 465)"}
EMAIL_DEFAULTS = {"enabled": False, "host": "smtp.gmail.com", "port": 587, "security": "starttls",
                  "username": "", "password_enc": "", "to": "",
                  "kinds": [k for k, _label, on in EMAIL_KINDS if on]}


def email_settings_path():
    return this_program_path().parent / EMAIL_SETTINGS_FILENAME


def load_email_settings():
    """The saved e-mail settings, always complete and valid in shape (a damaged file gives the defaults)."""
    cfg = dict(EMAIL_DEFAULTS)
    data = read_json_file(email_settings_path())
    if isinstance(data, dict):
        for key in EMAIL_DEFAULTS:
            if key in data:
                cfg[key] = data[key]
    cfg["enabled"] = bool(cfg["enabled"])
    cfg["kinds"] = [k for k in (cfg["kinds"] if isinstance(cfg["kinds"], list) else [])
                    if k in [x[0] for x in EMAIL_KINDS]]
    if cfg["security"] not in EMAIL_SECURITY:
        cfg["security"] = "starttls"
    return cfg


def save_email_settings(cfg):
    path = email_settings_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    tmp.replace(path)


def parse_recipients(text):
    """'a@x.com, b@y.com; c@z.com' -> ['a@x.com', 'b@y.com', 'c@z.com'] (valid, no duplicates)."""
    out = []
    for part in re.split(r"[,;\s]+", text or ""):
        if part and _EMAIL_RE.match(part) and part.lower() not in [o.lower() for o in out]:
            out.append(part)
    return out[:EMAIL_MAX_RECIPIENTS]


def describe_email_error(exc):
    """A short, plain-English reason for a failed send (never contains the password)."""
    if isinstance(exc, smtplib.SMTPAuthenticationError):
        return ("The mail server refused the login. For Gmail use an App Password "
                "(2-Step Verification must be on), not your normal password.")
    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        return "The mail server refused the recipient address(es)."
    if isinstance(exc, ssl.SSLError):
        return ("The secure connection failed - use STARTTLS for port 587 or SSL for port 465.")
    if isinstance(exc, socket.gaierror):
        return "Could not find the mail server - check the SMTP server name."
    if isinstance(exc, (socket.timeout, TimeoutError)):
        return "Timed out talking to the mail server."
    if isinstance(exc, ConnectionError):
        return "Could not connect to the mail server ({}).".format(type(exc).__name__)
    if isinstance(exc, smtplib.SMTPException):
        return "Mail server error ({}).".format(type(exc).__name__)
    if isinstance(exc, OSError):
        return "Network error ({}).".format(type(exc).__name__)
    return "{}: {}".format(type(exc).__name__, exc)


def send_email_message(cfg, subject, body, timeout=20):
    """Send one plain-text e-mail using the launcher's settings dict. Raises on failure.
    cfg: host, port, security ('starttls' | 'ssl'), username, password_enc, from, to."""
    recipients = parse_recipients(cfg.get("to"))
    if not recipients:
        raise ValueError("No valid recipient address.")
    host = (cfg.get("host") or "").strip()
    if not host:
        raise ValueError("No SMTP server.")
    username = (cfg.get("username") or "").strip()
    sender = (cfg.get("from") or username).strip()
    if not sender:
        raise ValueError("No sender address.")
    password = cfg.get("password") or ""
    if not password and cfg.get("password_enc"):
        password = unprotect_secret(cfg["password_enc"])
    if host.lower().endswith(("gmail.com", "googlemail.com")):
        password = password.replace(" ", "")              # Google shows app passwords in groups of four
    msg = EmailMessage()
    msg["Subject"] = " ".join(str(subject).split())
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)
    port = int(cfg.get("port") or 587)
    security = (cfg.get("security") or "starttls").lower()
    context = ssl.create_default_context()
    if security == "ssl":
        server = smtplib.SMTP_SSL(host, port, timeout=timeout, context=context)
    else:
        server = smtplib.SMTP(host, port, timeout=timeout)
    with server:
        server.ehlo()
        if security == "starttls":
            server.starttls(context=context)
            server.ehlo()
        if username:
            server.login(username, password)
        server.send_message(msg, from_addr=sender, to_addrs=recipients)


def send_test_email(cfg):
    """Used by the launcher's 'Send test e-mail' button. Returns the number of recipients."""
    send_email_message(
        cfg, EMAIL_SUBJECT_PREFIX + " Test e-mail",
        "This is a test message from NetUpgrade.\n\nIf you can read this, e-mail notifications are "
        "set up correctly.\n\n--\nNetUpgrade\n")
    return len(parse_recipients(cfg.get("to")))


class EmailNotifier(object):
    """Queues copies of dashboard notifications and sends them as e-mail (see above)."""

    def __init__(self):
        self.cfg_provider = lambda: None       # set by gui_main: returns the launcher's live settings
        self.last = {}                         # {"time", "ok", "detail"} of the latest attempt
        self._q = queue.Queue()
        self._lock = threading.Lock()
        self._thread = None

    def submit(self, site, host, label, kind, category, title, text):
        """Called for every notification; returns at once. Never raises."""
        try:
            cfg = self.cfg_provider() or {}
            if not cfg.get("enabled") or not cfg.get("host") or not cfg.get("to"):
                return
            cat = category or kind
            cat = {"logout": "login"}.get(cat, cat)
            if cat not in (cfg.get("kinds") or []):
                return
            self._q.put({"time": now_str(), "site": site or "", "host": host or "", "label": label or "",
                         "title": title, "text": text})
            with self._lock:
                if self._thread is None:
                    self._thread = threading.Thread(target=self._run, daemon=True)
                    self._thread.start()
        except Exception:  # noqa: BLE001 - notifications must never break an upgrade
            pass

    def _run(self):
        while True:
            batch = [self._q.get()]
            deadline = time.time() + EMAIL_BATCH_SECONDS
            while time.time() < deadline:
                try:
                    batch.append(self._q.get(timeout=max(0.1, deadline - time.time())))
                except queue.Empty:
                    break
            self._send_batch(batch)

    @staticmethod
    def _device(ev):
        if ev["host"] in ("", "-"):
            return ""
        return "{} ({})".format(ev["label"], ev["host"]) if ev["label"] else ev["host"]

    def _send_batch(self, batch):
        cfg = self.cfg_provider() or {}
        if not cfg.get("enabled"):                         # switched off while it was waiting
            return
        if len(batch) == 1:
            ev = batch[0]
            dev = self._device(ev)
            subject = "{} {}{}".format(EMAIL_SUBJECT_PREFIX, ev["title"], " - " + dev if dev else "")
        else:
            subject = "{} {} notifications".format(EMAIL_SUBJECT_PREFIX, len(batch))
        lines = []
        for ev in batch:
            lines.append("[{}] {}".format(ev["time"], ev["title"]))
            if ev["site"]:
                lines.append("  Site:   {}".format(ev["site"]))
            if self._device(ev):
                lines.append("  Device: {}".format(self._device(ev)))
            if ev["text"]:
                lines.append("  {}".format(ev["text"]))
            lines.append("")
        lines += ["--", "NetUpgrade", ""]
        error = None
        for attempt in range(2):                           # one retry for a passing network hiccup
            try:
                send_email_message(cfg, subject, "\n".join(lines))
                error = None
                break
            except Exception as exc:  # noqa: BLE001
                error = describe_email_error(exc)
                if attempt == 0:
                    time.sleep(5)
        self.last = {"time": now_str(), "ok": error is None,
                     "detail": ("Sent: {}".format(subject) if error is None else error)}
        if error:
            print("E-mail notification failed: {}".format(error), flush=True)


EMAILER = EmailNotifier()


# =============================================================================
# Roles (who may do what). The access guard (apply_access_control) looks the role up on
# EVERY request, so a role change made in the launcher applies at once, and the rules below
# are enforced on the server - hiding a button in the page is only cosmetic.
#   administrator  everything
#   operator       everything except stopping monitoring
#   viewer         read-only: the Home page, live status, web consoles and reports
# A missing or unknown role is treated as 'viewer' (fail closed).
# =============================================================================
ROLE_ADMIN, ROLE_OPERATOR, ROLE_VIEWER = "administrator", "operator", "viewer"
ROLES = (ROLE_ADMIN, ROLE_OPERATOR, ROLE_VIEWER)
ROLE_LABELS = {ROLE_ADMIN: "Administrator", ROLE_OPERATOR: "Operator", ROLE_VIEWER: "Viewer"}
OPERATOR_DENIED_POSTS = ("/api/monitor/stop", "/api/settings", "/api/notifications/email",
                         "/api/notifications/email/test", "/api/timeout-profiles/save",
                         "/api/timeout-profiles/delete", "/api/timeout-profiles/apply")
VIEWER_PAGES = ("/", "/index.html", "/favicon.ico")
VIEWER_GETS = ("/api/status", "/api/console", "/api/me", "/api/device", "/api/upgrade-report.csv",
               "/api/upgrade-report.xlsx", "/api/upgrade-report/print")


def role_of(record):
    """The role of a user record from access["users"] (the launcher's admin account has admin=True)."""
    if record and record.get("admin"):
        return ROLE_ADMIN
    role = (record or {}).get("role")
    return role if role in ROLES else ROLE_VIEWER


def _canonical_path(path):
    """Collapse the ways one URL can be written, so a deny rule cannot be slipped past."""
    import posixpath
    from urllib.parse import unquote
    text = unquote(str(path).split("?", 1)[0].split(";", 1)[0]).lower()
    return posixpath.normpath(re.sub(r"/+", "/", "/" + text))          # also resolves '.' and '..' segments


def role_allows(role, method, path):
    """(allowed, message) - the single place that decides what each role may request."""
    method = (method or "").upper()
    raw = str(path).split("?", 1)[0]
    if role == ROLE_ADMIN:
        return True, ""
    if role == ROLE_OPERATOR:
        if method == "POST" and _canonical_path(raw) in OPERATOR_DENIED_POSTS:
            return False, "Operators cannot stop monitoring or change the settings. Ask an Administrator."
        return True, ""
    if role == ROLE_VIEWER:
        if method in ("GET", "HEAD") and (raw in VIEWER_GETS or raw in VIEWER_PAGES):
            return True, ""
        return False, "Your account has Viewer (read-only) access, so this action is not allowed."
    return False, "Your account has no valid role."


def parse_hosts_from_workbook(wb, sheet=None, column=None):
    """Return a de-duplicated list of {'host', 'label', 'type'} from an open Workbook.
    Raises ValueError on anything wrong with the sheet - the caller decides how to report it
    (sys.exit for the CLI, an ApiError for the GUI's device-list upload)."""
    try:
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        raise ValueError("The Excel sheet is empty.")

    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    col_idx, start_row = None, 0
    if column:
        wanted = column.strip().lower()
        if wanted not in header:
            raise ValueError(f"Column '{column}' not found. Headers seen: {header}")
        col_idx, start_row = header.index(wanted), 1
    else:
        for cand in ADDRESS_HEADERS:
            if cand in header:
                col_idx, start_row = header.index(cand), 1
                break
        if col_idx is None:  # no recognizable header: first column, row 1 is data
            col_idx, start_row = 0, 0

    def find(cands):
        if start_row != 1:
            return None
        for cand in cands:
            if cand in header and header.index(cand) != col_idx:
                return header.index(cand)
        return None

    label_idx, type_idx = find(LABEL_HEADERS), find(TYPE_HEADERS)
    platform_idx = find(PLATFORM_HEADERS)

    def cell(row, idx):
        if idx is None or idx >= len(row) or row[idx] is None:
            return ""
        return str(row[idx]).strip()

    hosts, seen = [], set()
    for row in rows[start_row:]:
        host = cell(row, col_idx)
        if not host or host in seen:
            continue
        seen.add(host)
        dev_type = cell(row, type_idx).lower().replace("_", "-") or "cisco-xe"
        hosts.append({"host": host, "label": cell(row, label_idx), "type": dev_type,
                      "platform": cell(row, platform_idx)[:MAX_PLATFORM_LEN]})
    return hosts


def load_hosts(xlsx_path, sheet=None, column=None):
    """CLI-facing: load the switch list from a file path, exiting the program on a bad file."""
    try:
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        return parse_hosts_from_workbook(wb, sheet, column)
    except ValueError as exc:
        sys.exit(str(exc))


def load_hosts_from_bytes(data, sheet=None, column=None):
    """GUI-facing: load the switch list from uploaded bytes. Raises ValueError (caught by the
    HTTP handler and shown to the user) instead of exiting the whole program on a bad file."""
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl raises various errors for a bad/corrupt file
        raise ValueError(f"Could not read that as an Excel (.xlsx) file: {exc}")
    return parse_hosts_from_workbook(wb, sheet, column)


def device_list_template_xlsx():
    """The blank device-list workbook offered next to 'Upload device list'. Sheet 1 ('Switches') has only the
    header row - nothing is imported until the user types devices in. Sheet 2 explains the columns."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Switches"
    headers = ["ip", "hostname", "platform"]
    ws.append(headers)
    notes = {"ip": "Required. The management IP address (or DNS name) of the switch.",
             "hostname": "Optional. A friendly name shown in the dashboard.",
             "platform": "Optional. The model, for example C9200L-48P-4G. It is read from the switch later if left empty."}
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")
        cell.comment = Comment(notes[cell.value], APP_NAME)
    for col, width in zip("ABC", (22, 28, 26)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    help_ws = wb.create_sheet("Instructions")
    for line in [
            (APP_NAME + " - device list template",),
            ("",),
            ("1. Type one switch per row on the 'Switches' sheet, below the header row. Keep the header row as it is.",),
            ("2. Save the file as .xlsx.",),
            ("3. In the dashboard open Device List > Upload device list, choose the file and click Upload list.",),
            ("",),
            ("Column", "Needed?", "Example"),
            ("ip", "Required", "10.10.1.11"),
            ("hostname", "Optional", "SW-FLOOR1-01"),
            ("platform", "Optional", "C9200L-48P-4G"),
            ("",),
            ("Devices that are already in the list are skipped. Blank rows are ignored.",),
            ("This sheet is ignored by the upload - only the first sheet is read.",)]:
        help_ws.append(list(line))
    help_ws["A1"].font = Font(bold=True, size=14)
    for cell in help_ws[7]:
        cell.font = Font(bold=True)
    help_ws.column_dimensions["A"].width = 26
    help_ws.column_dimensions["B"].width = 14
    help_ws.column_dimensions["C"].width = 24
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------- #
# Availability monitor
# --------------------------------------------------------------------------- #
class Monitor:
    def __init__(self, hosts, events, mode, port, interval, timeout, duration_min,
                 fail_threshold, history_len=60):
        self.hosts = hosts
        self.events = events
        self.mode = mode
        self.port = port
        self.interval = interval
        self.timeout = timeout
        self.duration_sec = duration_min * 60
        self.fail_threshold = fail_threshold
        self.history_len = history_len

        self.lock = threading.Lock()
        self.ctl = threading.Lock()
        self.stop_event = threading.Event()
        self.thread = None
        self.running = False
        self._reset()

    def _reset(self):
        with self.lock:
            self.cycles = 0
            self.started_at = None
            self._t_start = None
            self._t_end = None
            self.stats = {}
            for h in self.hosts:
                self.stats[h["host"]] = {
                    "host": h["host"], "label": h["label"], "status": "PENDING",
                    "checks": 0, "successes": 0, "consec_fail": 0, "last_rtt": None,
                    "rtt_sum": 0.0, "rtt_min": None, "rtt_max": None,
                    "last_checked": None, "last_change": None,
                    "history": deque(maxlen=self.history_len),
                }

    def replace_hosts(self, new_hosts):
        """Swap the entire device list at once (used when switching site profiles). Stops
        monitoring first if it was running - the operator restarts it explicitly for the
        new site, same as at first launch."""
        self.stop()
        with self.lock:
            self.hosts[:] = new_hosts
        self._reset()

    def add_host(self, host_entry):
        """Add one device while the monitor is idle or already running. Returns False if
        that IP/host is already tracked (the caller should treat this as a duplicate)."""
        with self.lock:
            if host_entry["host"] in self.stats:
                return False
            self.hosts.append(host_entry)
            self.stats[host_entry["host"]] = {
                "host": host_entry["host"], "label": host_entry["label"], "status": "PENDING",
                "checks": 0, "successes": 0, "consec_fail": 0, "last_rtt": None,
                "rtt_sum": 0.0, "rtt_min": None, "rtt_max": None,
                "last_checked": None, "last_change": None,
                "history": deque(maxlen=self.history_len),
            }
            return True

    def remove_host(self, host):
        """Drop one device. Safe to call while monitoring is running."""
        with self.lock:
            self.stats.pop(host, None)
            self.hosts[:] = [h for h in self.hosts if h["host"] != host]

    # ---- control ---------------------------------------------------------- #
    def start(self):
        with self.ctl:
            if self.running:
                return False
            self._reset()
            self.stop_event.clear()
            self.running = True
            self.thread = threading.Thread(target=self._run, daemon=True)
            self.thread.start()
            return True

    def stop(self):
        self.stop_event.set()

    # ---- probes ----------------------------------------------------------- #
    def probe(self, host):
        """Return (ok, rtt_ms) for one check of one host."""
        if self.mode == "tcp":
            t0 = time.perf_counter()
            try:
                with socket.create_connection((host, self.port), timeout=self.timeout):
                    pass
                return True, (time.perf_counter() - t0) * 1000.0
            except OSError:
                return False, None
        return self._ping(host)

    def _ping(self, host):
        system = platform.system().lower()
        if system == "windows":
            cmd = ["ping", "-n", "1", "-w", str(int(self.timeout * 1000)), host]
        elif system == "darwin":
            cmd = ["ping", "-c", "1", "-W", str(int(self.timeout * 1000)), host]
        else:
            cmd = ["ping", "-c", "1", "-W", str(max(1, int(self.timeout))), host]
        try:
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                  timeout=self.timeout + 3, **HIDE_CHILD_WINDOW)
        except (subprocess.TimeoutExpired, OSError):
            return False, None
        out = proc.stdout.decode("ascii", "ignore").lower()  # bytes: localized Windows output
        ok = proc.returncode == 0
        if system == "windows":  # Windows returns 0 even for "host unreachable"
            ok = ok and "ttl=" in out
        if not ok:
            return False, None
        m = re.search(r"time[=<]\s*([\d.]+)\s*ms", out)
        return True, float(m.group(1)) if m else None

    # ---- state ------------------------------------------------------------ #
    def record(self, host, ok, rtt):
        now = datetime.now()
        event = None
        with self.lock:
            s = self.stats.get(host)
            if s is None:
                return
            s["checks"] += 1
            s["last_checked"] = now.strftime("%H:%M:%S")
            s["history"].append(1 if ok else 0)
            if ok:
                s["successes"] += 1
                s["consec_fail"] = 0
                s["last_rtt"] = rtt
                if rtt is not None:
                    s["rtt_sum"] += rtt
                    s["rtt_min"] = rtt if s["rtt_min"] is None else min(s["rtt_min"], rtt)
                    s["rtt_max"] = rtt if s["rtt_max"] is None else max(s["rtt_max"], rtt)
                new = "UP"
            else:
                s["consec_fail"] += 1
                s["last_rtt"] = None
                new = "DOWN" if s["consec_fail"] >= self.fail_threshold else s["status"]

            prev = s["status"]
            if new != prev:
                s["status"] = new
                s["last_change"] = now.strftime("%H:%M:%S")
                if not (prev == "PENDING" and new == "UP"):  # don't log the initial "came up"
                    event = (s["label"], f"{prev} -> {new}", "ok" if new == "UP" else "bad")
        if event:
            self.events.add(host, event[0], event[1], event[2])

    def _run(self):
        with self.lock:
            self.started_at = datetime.now()
            self._t_start = time.monotonic()
        deadline = self._t_start + self.duration_sec if self.duration_sec > 0 else None
        workers = min(64, max(1, len(self.hosts)))
        try:
            with ThreadPoolExecutor(max_workers=workers) as pool:
                while not self.stop_event.is_set():
                    cycle_start = time.monotonic()
                    futures = {pool.submit(self.probe, h["host"]): h["host"] for h in self.hosts}
                    for fut in as_completed(futures):
                        ok, rtt = fut.result()
                        self.record(futures[fut], ok, rtt)
                    with self.lock:
                        self.cycles += 1
                    if deadline and time.monotonic() >= deadline:
                        break
                    wait = self.interval - (time.monotonic() - cycle_start)
                    self.stop_event.wait(max(0.0, wait))
        finally:
            with self.lock:
                self._t_end = time.monotonic()
            self.running = False

    # ---- views ------------------------------------------------------------ #
    def snapshot(self):
        with self.lock:
            hosts, total_checks, total_ok = [], 0, 0
            for s in self.stats.values():
                total_checks += s["checks"]
                total_ok += s["successes"]
                hosts.append({
                    "host": s["host"], "status": s["status"], "last_rtt": s["last_rtt"],
                    "availability": (s["successes"] / s["checks"] * 100.0) if s["checks"] else None,
                    "checks": s["checks"], "successes": s["successes"],
                    "last_change": s["last_change"], "history": list(s["history"]),
                })
            statuses = [h["status"] for h in hosts]
            cycles, started_at, t_start, t_end = self.cycles, self.started_at, self._t_start, self._t_end

        if t_start is None:
            elapsed = 0.0
        else:
            elapsed = (t_end if t_end is not None else time.monotonic()) - t_start
        if self.running:
            state = "running"
        elif t_start is None:
            state = "idle"
        elif self.stop_event.is_set():
            state = "stopped"
        else:
            state = "completed"
        return {
            "monitor": {
                "state": state,
                "mode": self.mode if self.mode == "icmp" else f"tcp/{self.port}",
                "interval": self.interval, "duration": self.duration_sec,
                "elapsed": elapsed, "cycles": cycles,
                "started_at": started_at.strftime("%Y-%m-%d %H:%M:%S") if started_at else None,
            },
            "summary": {
                "total": len(hosts), "up": statuses.count("UP"),
                "down": statuses.count("DOWN"), "pending": statuses.count("PENDING"),
                "availability": (total_ok / total_checks * 100.0) if total_checks else None,
            },
            "hosts": hosts,
        }


# --------------------------------------------------------------------------- #
# Backup (runs inside this process, in worker threads)
# --------------------------------------------------------------------------- #
BACKUP_FILE_PREFIX = "config_backup"


def do_backup(host, device_type, username, password, secret, device_dir, run_stamp, custom_commands=None):
    """Run show running-config / show version / show boot / show inventory, plus any extra commands
    typed in the GUI, and save ALL the outputs in ONE file inside device_dir - a PERMANENT folder for
    this one device. The file name carries run_stamp, so repeated backups accumulate a history in that
    same folder instead of overwriting each other. Returns a result dict including a per-command log
    for the GUI to display."""
    device = {
        "device_type": device_type, "host": host, "username": username,
        "password": password, "secret": secret or "",
        "conn_timeout": 20, "auth_timeout": 20, "banner_timeout": 20,
    }
    warnings, commands_run, sections = [], [], []
    device_dir = Path(device_dir)

    def run_command(command, timeout=None):
        try:
            output = conn.send_command(command, read_timeout=timeout or tmo("backup_cmd"))
            if not output or "Invalid input" in output[:200]:
                raise ValueError("empty or invalid output")
            sections.append((command, output))
            commands_run.append({"command": command, "status": "OK"})
            return output
        except Exception as exc:  # noqa: BLE001 - one bad command shouldn't stop the others
            warnings.append(f"'{command}' failed ({type(exc).__name__})")
            commands_run.append({"command": command, "status": "FAILED", "detail": str(exc)})
            sections.append((command, f"!! FAILED: {type(exc).__name__}: {exc}"))
            return None

    with ConnectHandler(**device) as conn:
        if secret:
            conn.enable()
        hostname = conn.find_prompt().rstrip("#>").strip()
        device_dir.mkdir(parents=True, exist_ok=True)

        config = run_command("show running-config", timeout=tmo("show_run"))
        if config is None:
            raise ValueError("Unexpected/empty output from 'show running-config'")

        version, model, sv_out, sb_out = None, None, None, None
        for command, suffix in EXTRA_COMMANDS:
            output = run_command(command)
            if suffix == "show_version" and output:
                version, model, sv_out = parse_version(output), parse_model(output), output
            elif suffix == "show_boot" and output:
                sb_out = output

        for command in (custom_commands or []):
            run_command(command)

        filename = f"{BACKUP_FILE_PREFIX}_{run_stamp}.txt"
        heading = f"Configuration backup - {hostname or host} ({host}) - {now_str()}"
        (device_dir / filename).write_text(command_sections_text(heading, sections) + "\n", encoding="utf-8")

    return {"hostname": hostname, "files": [filename], "version": version, "model": model,
            "boot": boot_state(sv_out, sb_out) if sb_out else None, "warnings": warnings,
            "commands_run": commands_run}


# --------------------------------------------------------------------------- #
# Running each upgrade worker in the background (its output feeds the web console)
# --------------------------------------------------------------------------- #
def launch_background_worker(argv, console_log):
    """Run argv (an upgrade worker) with NO console window. Everything it prints goes to the
    console_log file, which the dashboard's per-switch web console reads and streams to the
    browser. Because the output is a plain file (not a pipe to this process), a worker keeps
    running - and logging - even if the dashboard is closed in the middle of a reload.
    Returns the Popen object so the dashboard can tell when the worker has ended."""
    env = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8", IOSXE_BACKGROUND_WORKER="1")
    kwargs = {}
    if platform.system() == "Windows":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    out = open(console_log, "ab")
    try:
        return subprocess.Popen(argv, stdin=subprocess.DEVNULL, stdout=out,
                                stderr=subprocess.STDOUT, env=env, **kwargs)
    finally:
        out.close()


# --------------------------------------------------------------------------- #
# Credential profiles: named username/password/enable-secret sets, saved per site and
# assigned to devices. Passwords are stored ENCRYPTED, never in plain text:
#   * Windows: DPAPI - only the Windows user who saved them can read them back.
#   * Fallback (DPAPI unavailable / other OS): Fernet with a key file in the user's profile.
# --------------------------------------------------------------------------- #
PROFILE_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._-]{0,39}$")


def _profile_key_path():
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home())
    legacy = root / "IOSXEUpgradeManager"
    base = root / "NetUpgrade"
    if legacy.is_dir() and not base.exists():
        try:
            import shutil
            shutil.copytree(legacy, base)
        except OSError:
            base = legacy  # Preserve access to existing credentials if migration is blocked.
    base.mkdir(parents=True, exist_ok=True)
    return base / "profile.key"


def _fernet():
    from cryptography.fernet import Fernet
    path = _profile_key_path()
    if not path.is_file():
        fd = os.open(str(path), os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "wb") as fh:
            fh.write(Fernet.generate_key())
    return Fernet(path.read_bytes().strip())


def _dpapi(data, protect):
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    fn.argtypes = [ctypes.POINTER(BLOB), ctypes.c_void_p, ctypes.POINTER(BLOB), ctypes.c_void_p,
                   ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(BLOB)]
    fn.restype = wintypes.BOOL
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = BLOB()
    if not fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)):
        raise OSError("DPAPI call failed")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def protect_secret(text):
    """Encrypt a string for storage. '' stays ''."""
    if not text:
        return ""
    import base64
    raw = text.encode("utf-8")
    if os.name == "nt":
        try:
            blob = _dpapi(raw, True)
            if _dpapi(blob, False) == raw:                 # prove it round-trips before trusting it
                return "dpapi:" + base64.b64encode(blob).decode("ascii")
        except Exception:                                  # noqa: BLE001 - fall back below
            pass
    return "fernet:" + _fernet().encrypt(raw).decode("ascii")


def unprotect_secret(blob):
    if not blob:
        return ""
    import base64
    kind, _, body = blob.partition(":")
    if kind == "dpapi":
        return _dpapi(base64.b64decode(body), False).decode("utf-8")
    if kind == "fernet":
        return _fernet().decrypt(body.encode("ascii")).decode("utf-8")
    raise ValueError("unknown format")


def site_schedule_path(name):
    return site_dir(name) / "schedules.json"


def site_upgrade_settings_path(name):
    return site_dir(name) / "upgrade_settings.json"


def site_notes_path(name):
    return site_dir(name) / "notifications.json"


def site_profiles_path(name):
    return site_dir(name) / "credential_profiles.json"


def load_site_profiles(name):
    """(profiles, assignments) saved for this site: {name: {username, password, secret}} and
    {host: profile name}. Passwords are still encrypted here."""
    data = read_json_file(site_profiles_path(name))
    if not isinstance(data, dict):
        return {}, {}
    profiles = {}
    for pname, rec in (data.get("profiles") or {}).items():
        if isinstance(rec, dict) and isinstance(rec.get("username"), str) and isinstance(rec.get("password"), str):
            profiles[str(pname)] = {"username": rec["username"], "password": rec["password"],
                                    "secret": rec.get("secret") if isinstance(rec.get("secret"), str) else "",
                                    "protocol": "telnet" if rec.get("protocol") == "telnet" else "ssh"}
            try:
                profiles[str(pname)]["port"] = clean_port(rec.get("port"), profiles[str(pname)]["protocol"])
            except ValueError:
                profiles[str(pname)]["port"] = default_port(profiles[str(pname)]["protocol"])
    assign = {h: p for h, p in (data.get("assignments") or {}).items()
              if isinstance(h, str) and p in profiles}
    return profiles, assign


def save_site_profiles(name, profiles, assignments):
    path = site_profiles_path(name)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    fd = os.open(str(tmp), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump({"profiles": profiles, "assignments": assignments}, fh, indent=2)
    tmp.replace(path)


# --------------------------------------------------------------------------- #
# Manager: ties monitor, backups and upgrades together
# --------------------------------------------------------------------------- #
# --------------------------------------------------------------------------- #
# Device drawer - everything this program knows about ONE switch, gathered for the dashboard's
# right-hand panel: what was read from it (backups, readiness / pre- / post-checks), the commands
# that were run on it, its events and past upgrades, and which steps are still missing before an
# upgrade. Read-only: it never connects to the switch.
# --------------------------------------------------------------------------- #
STAMP_RE = re.compile(r"_(\d{8}_\d{6})\.[A-Za-z0-9]+$")
DEVICE_FILE_MAX_BYTES = 8_000_000        # the most of one file the drawer will show
_FACTS_CACHE = {}                        # (path, mtime) -> parsed facts, so a poll does not re-read files


def _stamp_to_time(stamp):
    try:
        return datetime.strptime(stamp, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return stamp


def backup_file_command(name, device_dir=None):
    """config_backup_<stamp>.txt (all commands in one file) -> the list of commands it holds."""
    m = STAMP_RE.search(name)
    base = name[:m.start()] if m else name
    if base == BACKUP_FILE_PREFIX:
        try:
            with open(os.path.join(str(device_dir), name), "r", encoding="utf-8", errors="replace") as fh:
                head = fh.read(4000)
            mm = re.search(r"(?m)^! Commands: (.+)$", head)
            if mm:
                return "All commands: " + mm.group(1)
        except (OSError, TypeError):
            pass
        return "All commands (one file)"
    return base


def list_device_backups(device_dir, limit=12):
    """The device's permanent backup folder grouped by backup run, newest first:
    [{'stamp', 'time', 'files': [{'name', 'size', 'command'}]}]"""
    runs = {}
    try:
        with os.scandir(device_dir) as it:
            for e in it:
                m = STAMP_RE.search(e.name)
                if not m or not e.is_file():
                    continue
                try:
                    size = e.stat().st_size
                except OSError:
                    continue
                runs.setdefault(m.group(1), []).append(
                    {"name": e.name, "size": size, "command": backup_file_command(e.name, device_dir)})
    except OSError:
        return []
    out = []
    for stamp in sorted(runs, reverse=True)[:limit]:
        out.append({"stamp": stamp, "time": _stamp_to_time(stamp), "files": sorted(runs[stamp], key=lambda f: f["name"])})
    return out


BACKUP_STAMP_ONLY_RE = re.compile(r"^\d{8}_\d{6}$")
BACKUP_COMPARE_MAX_FILE = 3_000_000       # most of one backup file read for a comparison
BACKUP_COMPARE_MAX_LINES = 3000           # diff lines kept per command
# Lines that change on every backup without meaning anything: skipped when "ignore changing lines" is on.
VOLATILE_LINE_RE = re.compile(
    r"^\s*(Building configuration|Current configuration\s*:|! ?Last configuration change|! ?NVRAM config last updated|"
    r"! ?No configuration change since last restart|!Time:|ntp clock-period|.*\buptime is\b|"
    r"System restarted at|System returned to ROM by|Last reload|Configuration register is|"
    r"\w{3} \w{3}\s+\d+ \d\d:\d\d:\d\d)", re.I)


def backup_run_sections(device_dir, stamp):
    """All the saved output of ONE backup run of a device as {command: text}. A run is one file holding every
    command (the '!===== COMMAND: x =====' format) or, in older backups, one file per command."""
    sections = {}
    try:
        names = sorted(e.name for e in os.scandir(device_dir)
                       if e.is_file() and (STAMP_RE.search(e.name) or [None, ""])[1] == stamp)
    except OSError:
        return sections
    for name in names:
        try:
            with open(Path(device_dir) / name, "rb") as fh:
                text = fh.read(BACKUP_COMPARE_MAX_FILE).decode("utf-8", "replace")
        except OSError:
            continue
        parts = split_command_sections(text)
        if parts:
            for cmd, out in parts.items():
                sections.setdefault(cmd, out)
        else:
            sections.setdefault(backup_file_command(name, device_dir) or name, text)
    return sections


def diff_backup_sections(a, b, ignore_volatile=True):
    """Compare two {command: text} dicts -> (summary, per-command results). Each result:
    {command, status: same|changed|added|removed, added, removed, diff: [unified-diff lines], truncated}."""
    def lines(text):
        rows = (text or "").replace("\r", "").split("\n")
        rows = [r.rstrip() for r in rows]
        while rows and not rows[-1]:
            rows.pop()
        return [r for r in rows if not (ignore_volatile and VOLATILE_LINE_RE.match(r))]
    results = []
    summary = {"same": 0, "changed": 0, "added": 0, "removed": 0, "lines_added": 0, "lines_removed": 0}
    for cmd in list(a) + [c for c in b if c not in a]:
        la, lb = lines(a.get(cmd)), lines(b.get(cmd))
        if cmd not in b:
            status = "removed"
        elif cmd not in a:
            status = "added"
        else:
            status = "same" if la == lb else "changed"
        diff, add, rem = [], 0, 0
        if status != "same":
            for ln in difflib.unified_diff(la, lb, "", "", lineterm="", n=3):
                if ln.startswith(("---", "+++")):
                    continue
                if ln.startswith("+"):
                    add += 1
                elif ln.startswith("-"):
                    rem += 1
                diff.append(ln)
        truncated = len(diff) > BACKUP_COMPARE_MAX_LINES
        summary[status] += 1
        summary["lines_added"] += add
        summary["lines_removed"] += rem
        results.append({"command": cmd, "status": status, "added": add, "removed": rem,
                        "diff": diff[:BACKUP_COMPARE_MAX_LINES], "truncated": truncated})
    order = {"changed": 0, "added": 1, "removed": 2, "same": 3}
    results.sort(key=lambda r: order[r["status"]])
    return summary, results


BACKUP_SIDE_MAX_LINES = 30000             # lines per side kept for the side-by-side view of ONE command


def side_by_side_rows(text_a, text_b, ignore_volatile=True):
    """Align two outputs line by line for a left/right view -> (rows, added, removed, truncated).
    Row: [line no. left (0 = none), left text, line no. right, right text, kind]; kind is '=' same,
    '~' changed, '-' only on the left (older), '+' only on the right (newer)."""
    def raw(text):
        rows = [r.rstrip() for r in (text or "").replace("\r", "").split("\n")]
        while rows and not rows[-1]:
            rows.pop()
        return rows
    ra, rb = raw(text_a), raw(text_b)
    cut = len(ra) > BACKUP_SIDE_MAX_LINES or len(rb) > BACKUP_SIDE_MAX_LINES
    ra, rb = ra[:BACKUP_SIDE_MAX_LINES], rb[:BACKUP_SIDE_MAX_LINES]
    key = lambda r: "\x00" if (ignore_volatile and VOLATILE_LINE_RE.match(r)) else r      # noqa: E731
    ka, kb = [key(r) for r in ra], [key(r) for r in rb]
    pre = 0                                    # identical start / end are skipped before the (slow) matcher runs
    while pre < len(ka) and pre < len(kb) and ka[pre] == kb[pre]:
        pre += 1
    suf = 0
    while suf < len(ka) - pre and suf < len(kb) - pre and ka[-1 - suf] == kb[-1 - suf]:
        suf += 1
    rows, add, rem = [], 0, 0
    for k in range(pre):
        rows.append([k + 1, ra[k], k + 1, rb[k], "="])
    ea, eb = len(ka) - suf, len(kb) - suf
    sm = difflib.SequenceMatcher(None, ka[pre:ea], kb[pre:eb], autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        i1, i2, j1, j2 = i1 + pre, i2 + pre, j1 + pre, j2 + pre
        if tag == "equal":
            for k in range(i2 - i1):
                rows.append([i1 + k + 1, ra[i1 + k], j1 + k + 1, rb[j1 + k], "="])
            continue
        for k in range(max(i2 - i1, j2 - j1)):
            has_a, has_b = k < i2 - i1, k < j2 - j1
            kind = "~" if has_a and has_b else ("-" if has_a else "+")
            add += 1 if has_b else 0
            rem += 1 if has_a else 0
            rows.append([i1 + k + 1 if has_a else 0, ra[i1 + k] if has_a else "",
                         j1 + k + 1 if has_b else 0, rb[j1 + k] if has_b else "", kind])
    for k in range(suf):
        rows.append([ea + k + 1, ra[ea + k], eb + k + 1, rb[eb + k], "="])
    return rows, add, rem, cut


def compare_backup_dir(folder, a, b, ignore_volatile=True, cmd=None):
    """Compare two backup runs (by timestamp) found in ONE device's backup folder, command by command."""
    a, b = str(a or ""), str(b or "")
    if not (BACKUP_STAMP_ONLY_RE.match(a) and BACKUP_STAMP_ONLY_RE.match(b)):
        raise ApiError("Pick two backups to compare.")
    if a == b:
        raise ApiError("Pick two different backups.")
    sec_a, sec_b = backup_run_sections(folder, a), backup_run_sections(folder, b)
    if not sec_a or not sec_b:
        raise ApiError("One of the backups was not found.", 404)
    if cmd:                                    # left/right view of ONE command's full output
        if cmd not in sec_a and cmd not in sec_b:
            raise ApiError("That command is not in either backup.", 404)
        rows, add, rem, cut = side_by_side_rows(sec_a.get(cmd), sec_b.get(cmd), bool(ignore_volatile))
        return {"a": {"stamp": a, "time": _stamp_to_time(a)}, "b": {"stamp": b, "time": _stamp_to_time(b)},
                "command": cmd, "in_a": cmd in sec_a, "in_b": cmd in sec_b, "rows": rows,
                "added": add, "removed": rem, "truncated": cut}
    summary, results = diff_backup_sections(sec_a, sec_b, bool(ignore_volatile))
    return {"a": {"stamp": a, "time": _stamp_to_time(a)}, "b": {"stamp": b, "time": _stamp_to_time(b)},
            "ignore_volatile": bool(ignore_volatile), "summary": summary, "commands": results}


def parse_show_version_facts(text):
    facts = {}
    m = re.search(r"(?m)^(\S+)\s+uptime is\s+(.+?)\s*$", text)
    if m:
        facts["hostname"], facts["uptime"] = m.group(1), m.group(2)
    m = (re.search(r"System Serial Number\s*:\s*(\S+)", text, re.I)
         or re.search(r"Processor board ID\s+(\S+)", text, re.I))
    if m:
        facts["serial"] = m.group(1)
    m = re.search(r'System image file is "([^"]+)"', text)
    if m:
        facts["image_file"] = m.group(1)
    m = re.search(r"Last reload reason:\s*(.+?)\s*$", text, re.M)
    if m:
        facts["last_reload"] = m.group(1)
    members = []
    for m in re.finditer(r"(?m)^\s*(\*?)\s*(\d+)\s+(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(INSTALL|BUNDLE)\b", text):
        members.append({"switch": int(m.group(2)), "active": bool(m.group(1)), "ports": int(m.group(3)),
                        "model": m.group(4), "version": m.group(5), "mode": m.group(7)})
    if members:
        facts["members"] = members
    return facts


def parse_inventory_facts(text):
    """[{'name', 'pid', 'vid', 'serial'}] from 'show inventory'."""
    out = []
    for m in re.finditer(r'NAME:\s*"([^"]*)",\s*DESCR:\s*"[^"]*"\s*\r?\n\s*PID:\s*(\S*)\s*,\s*VID:\s*(\S*)\s*,\s*SN:\s*(\S*)', text):
        out.append({"name": m.group(1), "pid": m.group(2), "vid": m.group(3), "serial": m.group(4)})
    return out


def read_device_facts(device_dir):
    """Serial number, hardware version, uptime, stack members ... read from the newest saved
    'show version' / 'show inventory' output of this device - from the newest one-file backup
    (config_backup_<stamp>.txt) ({} if never backed up)."""
    facts = {}
    try:
        all_names = [n for n in os.listdir(device_dir) if STAMP_RE.search(n)]
    except OSError:
        return facts
    for prefix, command in (("show_version", "show version"), ("show_inventory", "show inventory")):
        cands = [(STAMP_RE.search(n).group(1), n) for n in all_names
                 if n.startswith(BACKUP_FILE_PREFIX + "_")]
        if not cands:
            continue
        stamp, name = max(cands)
        path = os.path.join(str(device_dir), name)
        try:
            key = (path, os.path.getmtime(path), prefix)
        except OSError:
            continue
        if key not in _FACTS_CACHE:
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as fh:
                    text = fh.read(8_000_000)
            except OSError:
                continue
            text = split_command_sections(text).get(command, "")
            parsed = (parse_show_version_facts(text) if prefix == "show_version"
                      else {"inventory": parse_inventory_facts(text)})
            parsed["_time"] = _stamp_to_time(stamp)
            if len(_FACTS_CACHE) > 400:
                _FACTS_CACHE.clear()
            _FACTS_CACHE[key] = parsed
        for k, v in _FACTS_CACHE[key].items():
            if k == "_time":
                facts["read_from_backup"] = max(facts.get("read_from_backup", ""), v)
            else:
                facts[k] = v
    inv = facts.get("inventory") or []
    if inv:
        pick = next((i for i in inv if re.fullmatch(r"(?i)switch\s*1|chassis|1", i["name"].strip())), inv[0])
        facts.setdefault("serial", pick["serial"])
        facts["hardware_version"] = pick["vid"] or None
        facts["inventory"] = inv[:12]
    return facts


def read_command_log(runs_dir, host, limit=300):
    """The recorded commands of one device, newest first (see record_command)."""
    path = Path(runs_dir) / ("commands_" + safe_name(host) + ".jsonl")
    try:
        with open(path, "rb") as fh:
            fh.seek(0, os.SEEK_END)
            fh.seek(max(0, fh.tell() - 262144))
            raw = fh.read().decode("utf-8", "replace")
    except OSError:
        return []
    out = []
    for line in raw.splitlines()[-limit:]:
        try:
            item = json.loads(line)
        except ValueError:
            continue
        if isinstance(item, dict) and item.get("cmd"):
            out.append(item)
    out.reverse()
    return out


def run_file_pattern(host):
    key = re.escape(safe_name(host))
    return re.compile(r"^(?:precheck_%s_(\d{8}_\d{6})\.txt|upgrade_%s_(\d{8}_\d{6})\.(precheck\.txt|postcheck\.txt|log|console\.txt))$"
                      % (key, key))


RUN_FILE_LABELS = {"precheck.txt": "Pre-upgrade snapshot", "postcheck.txt": "Post-upgrade check",
                   "log": "Upgrade log", "console.txt": "Upgrade console output"}


def list_run_files(runs_dir, host, limit=24):
    """Check results and logs this program saved for the device (readiness snapshot, pre-/post-check, upgrade log)."""
    pat, out = run_file_pattern(host), []
    try:
        with os.scandir(runs_dir) as it:
            for e in it:
                m = pat.match(e.name)
                if not m or not e.is_file():
                    continue
                try:
                    size = e.stat().st_size
                except OSError:
                    continue
                stamp = m.group(1) or m.group(2)
                label = "Readiness check snapshot" if m.group(1) else RUN_FILE_LABELS.get(m.group(3), m.group(3))
                out.append({"name": e.name, "size": size, "label": label, "stamp": stamp, "time": _stamp_to_time(stamp)})
    except OSError:
        return []
    out.sort(key=lambda f: (f["stamp"], f["name"]), reverse=True)
    return out[:limit]


def latest_upgrade_status(runs_dir, host):
    """The status file of this device's newest upgrade run (holds the pre-upgrade snapshot), or None."""
    pat = re.compile(r"^upgrade_%s_(\d{8}_\d{6})\.status\.json$" % re.escape(safe_name(host)))
    best = None
    try:
        with os.scandir(runs_dir) as it:
            for e in it:
                m = pat.match(e.name)
                if m and (best is None or m.group(1) > best[0]):
                    best = (m.group(1), e.path)
    except OSError:
        return None
    if not best:
        return None
    data = read_json_file(best[1])
    return dict(data, _stamp=best[0]) if isinstance(data, dict) else None


# --------------------------------------------------------------------------- #
# Web terminal - an interactive SSH session to one switch, shown in the device drawer's Console
# tab. The browser talks to this program over plain HTTP (open / send / read-with-offset, the same
# way the upgrade consoles are fed); THIS program holds the SSH connection. The switch's login comes
# from the credential profile attached to it and never reaches the browser. Every line typed is also
# written to the device's command log (source "Web console"). Sessions end after TERM_IDLE_SECONDS
# without use, and cannot be opened while an upgrade is running on that switch.
# --------------------------------------------------------------------------- #
TERM_MAX_SESSIONS = 10
TERM_IDLE_SECONDS = 900
TERM_BUFFER_MAX = 1_000_000
TERM_MAX_LINE = 500
_ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]|\x1b[()][A-Za-z0-9]|\x1b[=>78]|\x1b\][^\x07]*\x07")


def term_clean(text):
    """Make switch output readable in a plain text box: no colour / cursor codes, no stray CRs, backspaces applied."""
    text = _ANSI_RE.sub("", text).replace("\r\n", "\n").replace("\r", "")
    out = []
    for ch in text:
        if ch == "\x08":
            if out and out[-1] != "\n":
                out.pop()
        elif ch == "\n" or ch == "\t" or ord(ch) >= 32:
            out.append(ch)
    return "".join(out)


class _TelnetChannel(object):
    """Just enough Telnet for the web console, with the same send / recv / close calls as an SSH channel.
    Telnet option negotiation is answered (only 'suppress go-ahead' and the server's echo are accepted)."""
    IAC, DONT, DO, WONT, WILL, SB, SE = 255, 254, 253, 252, 251, 250, 240

    def __init__(self, host, port=23, timeout=15):
        self.sock = socket.create_connection((host, port), timeout=timeout)
        self.sock.settimeout(0.5)
        self._pend = b""

    def settimeout(self, seconds):
        self.sock.settimeout(seconds)

    def send(self, data):
        if isinstance(data, str):
            data = data.encode("utf-8", "replace")
        data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n").replace(b"\xff", b"\xff\xff")
        self.sock.sendall(data)

    def recv(self, n):
        raw = self.sock.recv(n)
        if not raw:
            return b""                                   # the switch closed the connection
        out = self._strip(raw)
        if not out:
            raise socket.timeout()                       # only negotiation bytes arrived: nothing to show yet
        return out

    def _strip(self, raw):
        data, self._pend, out, i = self._pend + raw, b"", bytearray(), 0
        n = len(data)
        while i < n:
            b = data[i]
            if b != self.IAC:
                out.append(b); i += 1; continue
            if i + 1 >= n:
                self._pend = data[i:]; break
            c = data[i + 1]
            if c == self.IAC:
                out.append(self.IAC); i += 2
            elif c in (self.WILL, self.WONT, self.DO, self.DONT):
                if i + 2 >= n:
                    self._pend = data[i:]; break
                opt = data[i + 2]
                if c == self.WILL:
                    self.sock.sendall(bytes([self.IAC, self.DO if opt in (1, 3) else self.DONT, opt]))
                elif c == self.DO:
                    self.sock.sendall(bytes([self.IAC, self.WILL if opt == 3 else self.WONT, opt]))
                i += 3
            elif c == self.SB:
                end = data.find(bytes([self.IAC, self.SE]), i + 2)
                if end < 0:
                    self._pend = data[i:]; break
                i = end + 2
            else:
                i += 2
        return bytes(out)

    def close(self):
        try:
            self.sock.close()
        except OSError:
            pass


class TerminalSession(object):
    def __init__(self, host, label, username, password, secret, opened_by, protocol="ssh", port=None):
        self.protocol = "telnet" if protocol == "telnet" else "ssh"
        self.port = port or default_port(self.protocol)
        self.id = os.urandom(8).hex()
        self.host, self.label, self.opened_by = host, label, opened_by
        self.lock = threading.Lock()
        self.buf, self.base = "", 0               # output so far; absolute offset of buf[0]
        self.closed = False
        self.ready = False
        self.last_used = time.time()
        self.client = self.chan = None
        threading.Thread(target=self._run, args=(username, password, secret), daemon=True).start()

    # ---- output buffer ---- #
    def _add(self, text):
        if not text:
            return
        with self.lock:
            self.buf += text
            if len(self.buf) > TERM_BUFFER_MAX:
                cut = len(self.buf) - TERM_BUFFER_MAX // 2
                self.buf, self.base = self.buf[cut:], self.base + cut

    def _tail(self, n=200):
        with self.lock:
            return self.buf[-n:]

    def read(self, offset):
        with self.lock:
            offset = max(self.base, min(offset, self.base + len(self.buf)))
            return {"text": self.buf[offset - self.base:], "offset": self.base + len(self.buf),
                    "closed": self.closed, "ready": self.ready}

    # ---- the connection ---- #
    def _run(self, username, password, secret):
        import codecs
        try:
            self._add(f"Connecting to {self.label or self.host} ({self.host}) over {'Telnet' if self.protocol == 'telnet' else 'SSH'} ...\n")
            if self.protocol == "telnet":
                chan = self.chan = _TelnetChannel(self.host, self.port)
                if not self._telnet_login(chan, username, password):
                    return
            else:
                client = paramiko.SSHClient()
                client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                client.connect(self.host, port=self.port, username=username, password=password, look_for_keys=False,
                               allow_agent=False, timeout=15, auth_timeout=20, banner_timeout=20)
                self.client = client
                chan = self.chan = client.invoke_shell(term="vt100", width=200, height=50)
                chan.settimeout(0.5)
        except paramiko.AuthenticationException:
            self._finish("Login failed: wrong username or password in the credential profile.")
            return
        except Exception as exc:  # noqa: BLE001
            self._finish(f"Could not connect: {type(exc).__name__}: {exc}"[:300])
            return
        finally:
            password = None
        threading.Thread(target=self._after_login, args=(secret,), daemon=True).start()
        decoder = codecs.getincrementaldecoder("utf-8")("replace")
        try:
            while not self.closed:
                try:
                    data = chan.recv(4096)
                except socket.timeout:
                    continue
                if not data:
                    break
                self._add(term_clean(decoder.decode(data)))
        except Exception:  # noqa: BLE001 - the connection dropped
            pass
        self._finish("Connection closed by the switch." if not self.closed else "")

    def _telnet_login(self, chan, username, password):
        """Answer the switch's Username: / Password: prompts from the profile. True once the prompt is reached."""
        import codecs
        decoder = codecs.getincrementaldecoder("utf-8")("replace")
        seen, sent_user, sent_pw, end = "", False, 0, time.time() + 25
        while time.time() < end and not self.closed:
            try:
                data = chan.recv(4096)
            except socket.timeout:
                continue
            except OSError:
                break
            if not data:
                self._finish("Connection closed by the switch.")
                return False
            text = term_clean(decoder.decode(data))
            self._add(text)
            seen = (seen + text)[-300:]
            tail = seen.rstrip().lower()
            if "authentication failed" in tail or "login invalid" in tail or "access denied" in tail:
                self._finish("Login failed: wrong username or password in the credential profile.")
                return False
            if tail.endswith(("username:", "login:")) and not sent_user:
                chan.send(username + "\n"); sent_user = True; seen = ""
            elif tail.endswith("password:"):
                if sent_pw >= 2:                                   # asked again: the password was refused
                    self._finish("Login failed: wrong username or password in the credential profile.")
                    return False
                chan.send(password + "\n"); sent_pw += 1; seen = ""
            elif tail.endswith((">", "#")):
                return True
        return True

    def _wait_for(self, predicate, seconds):
        end = time.time() + seconds
        while time.time() < end and not self.closed:
            if predicate(self._tail()):
                return True
            time.sleep(0.15)
        return False

    def _after_login(self, secret):
        """Once the prompt appears: enter enable mode if the profile has an enable secret, and turn paging off."""
        try:
            self._wait_for(lambda t: t.rstrip().endswith((">", "#")), 8)
            if secret and self._tail().rstrip().endswith(">"):
                self.chan.send("enable\n")
                if self._wait_for(lambda t: t.rstrip().lower().endswith("password:"), 6):
                    self.chan.send(secret + "\n")
                    self._wait_for(lambda t: t.rstrip().endswith("#"), 6)
            self.chan.send("terminal length 0\n")
        except Exception:  # noqa: BLE001
            pass
        self.ready = True

    def send(self, data):
        if self.closed or self.chan is None:
            raise ApiError("The console is not connected.", 409)
        self.last_used = time.time()
        try:
            self.chan.send(data)
        except Exception:  # noqa: BLE001
            self._finish("Connection lost.")
            raise ApiError("The console connection was lost.", 409)

    def _finish(self, message):
        with self.lock:
            already = self.closed
            self.closed = True
        if not already:
            self._add(("\n" + message if message else "") + "\n[Session closed]\n")
        for obj in (self.chan, self.client):
            try:
                if obj is not None:
                    obj.close()
            except Exception:  # noqa: BLE001
                pass

    def close(self):
        self._finish("")


# --------------------------------------------------------------------------- #
# Diagram tab: CDP topology (what is connected to what)
# --------------------------------------------------------------------------- #
def site_topology_path(name):
    return site_dir(name) / "topology.json"


_CDP_SPLIT_RE = re.compile(r"^-{10,}\s*$", re.M)
_CORE_HINT_RE = re.compile(r"9[456]\d\d|3850|3750|65\d\d|60\d\d|N[3579]K", re.I)


def parse_cdp_neighbors_detail(text):
    """'show cdp neighbors detail' -> [{device_id, ips, platform, caps, kind, local_if, remote_if}].
    kind is 'net' for a switch or router and 'other' for phones, access points and hosts."""
    out = []
    for block in _CDP_SPLIT_RE.split(text or ""):
        m = re.search(r"^Device ID:\s*(\S.*?)\s*$", block, re.M)
        if not m:
            continue
        ips = []
        for ip in re.findall(r"IP(?:v4)? address:\s*(\d{1,3}(?:\.\d{1,3}){3})", block):
            if ip not in ips:
                ips.append(ip)
        pm = re.search(r"^Platform:\s*(.*?),\s*Capabilities:\s*(.*?)\s*$", block, re.M)
        platform = re.sub(r"^cisco\s+", "", pm.group(1).strip(), flags=re.I) if pm else ""
        caps = pm.group(2).strip() if pm else ""
        im = re.search(r"^Interface:\s*(.*?),\s*Port ID \(outgoing port\):\s*(.*?)\s*$", block, re.M)
        is_net = bool(re.search(r"\b(Switch|Router)\b", caps)) and not re.search(r"\b(Phone|Host)\b", caps)
        out.append({"device_id": m.group(1)[:80], "ips": ips[:4], "platform": platform[:60], "caps": caps[:60],
                    "kind": "net" if is_net else "other",
                    "local_if": (im.group(1) if im else "")[:48], "remote_if": (im.group(2) if im else "")[:48]})
        if len(out) >= 128:
            break
    return out


def cdp_name_key(name):
    """'SW-DIST1.corp.local(FOC1234)' and 'sw-dist1' are the same device."""
    text = re.sub(r"\(.*?\)\s*$", "", str(name or "").strip()).strip().lower()
    if re.fullmatch(r"\d{1,3}(?:\.\d{1,3}){3}", text):
        return text
    return text.split(".", 1)[0]


def build_topology_graph(info, order, devices, roots):
    """Turn what each scanned switch reported into {nodes, edges}. Neighbours are matched to onboarded
    devices by IP address or hostname; anything else shows as a 'not onboarded' node. Each node gets a
    level (0 = top of its tree) so the page can draw distribution above access."""
    managed = {h: info[h] for h in order if h in info}
    by_name = {}
    for h, rec in managed.items():
        for nm in (rec.get("label"), (devices.get(h) or {}).get("hostname")):
            key = cdp_name_key(nm)
            if key:
                by_name.setdefault(key, h)
    nodes, links = {}, {}

    def managed_node(h):
        if h not in nodes:
            rec = managed[h]
            dev = devices.get(h) or {}
            nodes[h] = {"id": h, "host": h, "name": rec.get("label") or dev.get("hostname") or h, "managed": True,
                        "platform": rec.get("platform") or "", "kind": "net", "scanned": False, "scan_ok": None,
                        "scan_error": "", "scan_time": "", "is_root": h in roots, "level": 0, "comp": 0}
        return nodes[h]

    for h in order:
        dev = devices.get(h)
        if h not in managed or not isinstance(dev, dict):
            continue
        me = managed_node(h)
        me.update(scanned=True, scan_ok=bool(dev.get("ok")), scan_error=str(dev.get("error") or ""),
                  scan_time=str(dev.get("time") or ""))
        for nb in dev.get("neighbors") or []:
            target = next((ip for ip in nb.get("ips", []) if ip in managed), None) \
                or by_name.get(cdp_name_key(nb.get("device_id")))
            if target is None:
                key = cdp_name_key(nb.get("device_id")) or str(nb.get("device_id"))
                target = "x:" + key
                if target not in nodes:
                    nodes[target] = {"id": target, "host": (nb.get("ips") or [None])[0], "managed": False,
                                     "name": re.sub(r"\(.*?\)\s*$", "", str(nb.get("device_id"))).strip(),
                                     "platform": nb.get("platform") or "", "kind": nb.get("kind") or "other",
                                     "scanned": False, "scan_ok": None, "scan_error": "", "scan_time": "",
                                     "is_root": False, "level": 0, "comp": 0}
            else:
                # Only show switches that were part of the scan: a managed switch that
                # was deselected (no scan data) must not reappear as a neighbour.
                if not isinstance(devices.get(target), dict):
                    continue
                managed_node(target)
            if target == h:
                continue
            ends = sorted([(h, nb.get("local_if") or ""), (target, nb.get("remote_if") or "")])
            links.setdefault((ends[0][0], ends[1][0]), set()).add((ends[0][1], ends[1][1]))

    net = {i for i, n in nodes.items() if n["kind"] == "net"}
    adj = {i: set() for i in nodes}
    for a, b in links:
        adj[a].add(b)
        adj[b].add(a)

    def score(i):
        n = nodes[i]
        return (len(adj[i] & net) * 10 + (6 if n["managed"] else 0)
                + (4 if _CORE_HINT_RE.search(n["platform"] or "") else 0))

    seen, comp = set(), 0
    for start in sorted(net, key=lambda i: (-score(i), nodes[i]["name"].lower())):
        if start in seen:
            continue
        members, stack = [], [start]
        seen.add(start)
        while stack:                                         # one connected group of switches
            cur = stack.pop()
            members.append(cur)
            for nxt in adj[cur] & net:
                if nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        def reach(src):                                      # distances from src within this group
            seen_d, q = {src: 0}, [src]
            while q:
                cur = q.pop(0)
                for nxt in adj[cur] & net:
                    if nxt not in seen_d:
                        seen_d[nxt] = seen_d[cur] + 1
                        q.append(nxt)
            return seen_d
        chosen = [i for i in members if nodes[i]["is_root"]]
        # Automatic top = the centre of the group (core above distribution above access); the most links breaks a tie.
        root = chosen[0] if chosen else min(members, key=lambda i: (max(reach(i).values()), -score(i), nodes[i]["name"].lower()))
        depth, queue = {root: 0}, [root]
        while queue:
            cur = queue.pop(0)
            for nxt in sorted(adj[cur] & net):
                if nxt not in depth:
                    depth[nxt] = depth[cur] + 1
                    queue.append(nxt)
        for i in members:
            nodes[i]["level"], nodes[i]["comp"] = depth.get(i, 0), comp
        comp += 1
    for i, n in nodes.items():                               # phones, access points ...: just below their switch
        if n["kind"] != "net":
            parents = [nodes[j] for j in adj[i] if j in net]
            if parents:
                top = min(parents, key=lambda p: p["level"])
                n["level"], n["comp"] = top["level"] + 1, top["comp"]
            else:
                n["comp"], comp = comp, comp + 1
    edges = [{"a": a, "b": b, "links": [{"a_if": x, "b_if": y} for x, y in sorted(ports)]}
             for (a, b), ports in sorted(links.items())]
    stamps = [n["scan_time"] for n in nodes.values() if n["scan_time"]]
    return {"nodes": sorted(nodes.values(), key=lambda n: (n["comp"], n["level"], n["name"].lower())),
            "edges": edges, "scanned": sum(1 for n in nodes.values() if n["scanned"]),
            "updated": max(stamps) if stamps else ""}


class Manager:
    def __init__(self, monitor, events):
        self.monitor = monitor
        self.events = events
        self.site_name = None               # no site profile active until switch_site() is called
        self.backup_dir = None
        self.runs_dir = None
        self.reports_dir = None
        self.order = []
        self.info = {}

        self.lock = threading.RLock()
        self.terms = {}                     # session id -> TerminalSession (the drawer's web console)
        self.term_lock = threading.Lock()
        self._cmdtest_busy = set()          # switches a Command time test is running on right now
        self._cmdtest_jobs = {}             # Command time test runs: job id -> state, live output, result
        self._img_md5 = {}                  # (path, mtime, size) -> MD5 of an image on this PC
        self.ops = {}
        self.jobs = {}                      # host -> private upgrade job data (paths, process)
        self.schedule = {}                  # host -> pending scheduled / queued upgrade (saved with the site)
        self.max_parallel = MAX_CONCURRENT_RELOADS   # switches upgraded at the same time (global setting)
        self.stop_on_failure = True         # stop starting queued switches once one in the rollout fails
        self.settings = load_global_settings()          # the Settings page (gear icon); one file for every site
        self.email_cfg = load_email_settings()          # Settings > Notifications (e-mail copies of the notifications)
        self.session_started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")   # "fresh start" = nothing older than this
        self._apply_settings()
        self.rollouts = {}                  # rollout id -> {"hosts": set, "failed": bool, "by": str}
        self.rollout_of = {}                # host -> rollout id
        self.profiles = {}                  # credential profiles of the active site (passwords encrypted)
        self.assign = {}                    # host -> credential profile name
        self.notes = []                     # upgrade notifications of the active site (oldest first)
        self.notes_lock = threading.Lock()
        self.keep_awake = KeepAwake(self._need_awake, self._awake_changed)
        self.backup_run = None
        self.target_image = ""              # image last picked in an Operator/Administrator drawer (shown to Viewers)
        self.topo_lock = threading.RLock()  # Diagram tab: guards topology.json and the scan progress below
        self.topo_run = None                # the latest CDP topology scan of the active site
        self.pool = ThreadPoolExecutor(max_workers=BACKUP_WORKERS)
        self.job_tmp = tempfile.mkdtemp(prefix="swmgr_")  # private dir (0700) for credential job files
        self._stop = threading.Event()
        threading.Thread(target=self._scheduler_loop, daemon=True).start()

    def switch_site(self, name, mode="any"):
        """Activate a site profile: load its saved device list (and version history) into this
        same running Manager/Monitor, pointing backups/reports/run-logs at that site's folder.
        Can be called for the very first site too - not just to switch between two.

        mode: "select" = the site must already exist (dashboard's 'Select site'),
              "create" = the site must NOT exist yet (dashboard's 'Create site'),
              "any"    = either (used by the --site command-line option)."""
        try:
            name = safe_site_name(name)
        except ValueError as exc:
            raise ApiError(str(exc))
        exists = site_dir(name).is_dir()
        if mode == "select" and not exists:
            raise ApiError(f"Site '{name}' does not exist - pick one from the list or create it.", 404)
        if mode == "create" and exists:
            raise ApiError(f"Site '{name}' already exists - select it from the list instead.", 409)
        with self.lock:
            self._poll_all()
            if any(self._upgrade_active(h) for h in self.order):
                raise ApiError("An upgrade is still running - wait for it to finish before "
                               "switching sites.", 409)
            if any(self.ops[h]["backup"]["state"] in ("queued", "running") for h in self.order):
                raise ApiError("A backup is still running - wait for it to finish before "
                               "switching sites.", 409)

            if self.site_name:                              # save the outgoing site's list
                save_site_devices(self.site_name, [self.info[h] for h in self.order])
            cancelled_schedules = len(self.schedule)
            self.schedule.clear()
            self.rollouts.clear()
            self.rollout_of.clear()
            self._save_schedule()                            # the outgoing site's file, still the active one here

            site_dir(name).mkdir(parents=True, exist_ok=True)  # so it's listed even with 0 devices
            SITE.name = name                                 # so version_store_path() follows along
            self.site_name = name
            self.backup_dir = site_backups_dir(name)
            self.runs_dir = site_runs_dir(name)
            global COMMAND_LOG_DIR
            COMMAND_LOG_DIR = self.runs_dir
            self.term_close_all()
            self.reports_dir = site_reports_dir(name)

            new_hosts = load_site_devices(name)
            self.order = [h["host"] for h in new_hosts]
            self.info = {h["host"]: h for h in new_hosts}
            self.ops = {h: self._new_ops() for h in self.order}
            self.jobs = {}
            self.backup_run = None
            self.topo_run = None
            try:
                self.profiles, self.assign = load_site_profiles(name)
            except Exception as exc:  # noqa: BLE001 - a damaged profile file must not block the site
                print(f"Could not read credential profiles for '{name}': {type(exc).__name__}", flush=True)
                self.profiles, self.assign = {}, {}
            self.assign = {h: p for h, p in self.assign.items() if h in self.info}
            data = read_json_file(site_notes_path(name))
            with self.notes_lock:
                carried = [n for n in self.notes if n.get("kind") in ("login", "logout")]   # not tied to a site
                loaded = [n for n in (data if isinstance(data, list) else [])
                          if isinstance(n, dict) and isinstance(n.get("id"), int)]
                self.notes = sorted(loaded + carried, key=lambda n: n["id"])[-NOTES_MAX:]
            sched = read_json_file(site_schedule_path(name))     # saved schedules (they hold no passwords)
            self.schedule = {}
            for h, e in (sched.items() if isinstance(sched, dict) else []):
                if h in self.info and isinstance(e, dict) and isinstance(e.get("run_at"), (int, float)):
                    self.schedule[h] = {"run_at": e["run_at"], "image_path": e.get("image_path"),
                                        "created": e.get("created") or "", "backup_first": bool(e.get("backup_first")),
                                        "action": e.get("action") or "full",
                                        "scheduled_by": e.get("scheduled_by") or "local",
                                        "queued": bool(e.get("queued")),
                                        "ignore_space": bool(e.get("ignore_space")),
                                        "rollout": e.get("rollout") if isinstance(e.get("rollout"), str) else None}
            self._rebuild_rollouts()
            known = load_known_versions()
            for h in self.order:
                self._apply_known_version(h, known)

            self.monitor.replace_hosts(new_hosts)
        msg = f"Switched to site '{name}' - {len(new_hosts)} device(s) loaded."
        if cancelled_schedules:
            msg += (f" {cancelled_schedules} pending scheduled run(s) from the previous site "
                   "were cancelled.")
        self.events.add("-", "", f"Active site changed to '{name}'", "info")
        return {"message": msg, "site": name, "device_count": len(new_hosts)}

    def _require_site(self):
        if not self.site_name:
            raise ApiError("Assign this dashboard to a site first (select or create one at the top).")

    def start_monitoring(self):
        """Monitoring is only allowed once the dashboard is assigned to a site."""
        self._require_site()
        with self.lock:
            self._report_missed()
        if not self.monitor.start():
            return {"message": "Monitoring is already running."}
        return {"message": "Monitoring started."}

    def stop_monitoring(self):
        with self.lock:
            self._poll_all()
            busy = [h for h in self.order if self._upgrade_active(h)]
            if busy:
                raise ApiError(f"Monitoring cannot be stopped while {len(busy)} upgrade(s) are running - "
                               "wait until they finish.", 409)
        self.monitor.stop()
        return {"message": "Monitoring stopped."}

    def _report_missed(self):
        """Scheduled upgrades whose time passed while monitoring was stopped never ran - say so.
        Called with self.lock held, right before monitoring starts."""
        now = time.time()
        missed = [(h, e) for h, e in self.schedule.items() if e["run_at"] <= now]
        for h, e in missed:
            when = datetime.fromtimestamp(e["run_at"]).strftime("%Y-%m-%d %H:%M")
            verb = self.ACTION_VERB.get(e.get("action") or "full", "Upgrade")
            self._notify(h, "missed", f"Scheduled {verb.lower()} missed",
                         f"It was planned for {when}, but monitoring was stopped, so it did not run. "
                         "Schedule it again if it is still needed.")
            self.events.add(h, self.label(h), f"Scheduled {verb.lower()} for {when} did not run: monitoring was stopped", "warn")
            del self.schedule[h]
        if missed:
            self._save_schedule()

    def _apply_known_version(self, host, known_versions=None):
        """Pre-fill a device's version from the on-disk store (survives a restart), so there's
        something to compare a future upgrade's result against even before the first backup or
        version check in this run."""
        if FRESH_START:
            return                          # every launch starts with empty status columns
        known = known_versions if known_versions is not None else load_known_versions()
        entry = known.get(host)
        if entry and entry.get("version"):
            o = self.ops[host]
            o["version"] = entry["version"]
            o["version_note"] = f"last known, {entry.get('updated_at', '')}".strip()
            if entry.get("model"):
                o["model"] = entry["model"]
            if isinstance(entry.get("boot"), dict):
                o["boot"] = dict(entry["boot"])

    @staticmethod
    def _new_ops():
        return {
            "version": None, "prev_version": None, "version_note": "", "model": None, "boot": {},
            "backup": {"state": "idle", "detail": "", "time": None},
            "creds": {"state": "idle", "detail": "", "time": None},
            "readiness": {"state": "idle", "detail": "", "checks": {}, "time": None},
            "postcheck": {"state": "idle", "detail": "", "checks": {}, "time": None},
            "upgrade": {"state": "idle", "phase": "", "message": "", "old_version": None,
                        "new_version": None, "started": None, "finished": None,
                        "started_ts": None, "finished_ts": None, "image": None, "action": None,
                        "started_by": None, "phase_times": []},
        }

    def label(self, host):
        return self.info[host]["label"]

    def add_host(self, host_entry):
        """Register one device, added live via the GUI (device-list upload or manual add).
        Returns False if that IP is already tracked."""
        self._require_site()
        with self.lock:
            if host_entry["host"] in self.info:
                return False
            if not self.monitor.add_host(host_entry):
                return False
            self.order.append(host_entry["host"])
            self.info[host_entry["host"]] = host_entry
            self.ops[host_entry["host"]] = self._new_ops()
            self._apply_known_version(host_entry["host"])
            save_site_devices(self.site_name, [self.info[h] for h in self.order])
            return True

    def add_device_manual(self, hostname, ip, platform_name=""):
        hostname = (hostname or "").strip()
        ip = (ip or "").strip()
        platform_name = (platform_name or "").strip()[:MAX_PLATFORM_LEN]
        if not hostname or not ip:
            raise ApiError("Both hostname and IP address are required.")
        if re.search(r"\s", ip) or re.search(r"\s", hostname):
            raise ApiError("Hostname and IP address cannot contain spaces.")
        ip = validate_ip_address(ip)
        if not self.add_host({"host": ip, "label": hostname, "type": "cisco-xe", "platform": platform_name}):
            raise ApiError(f"{ip} is already in the device list.", 409)
        self.events.add(ip, hostname, "Added to the device list", "info")
        return {"message": f"Added {hostname} ({ip}) to the device list."}

    def add_devices_bulk(self, hosts, source):
        added, skipped = 0, 0
        for h in hosts:
            if self.add_host(h):
                added += 1
            else:
                skipped += 1
        self.events.add("-", "", f"Device list upload ({source}): {added} added"
                        + (f", {skipped} already in the list" if skipped else ""), "info")
        return added, skipped

    def delete_devices(self, hosts):
        """Remove devices from the live list. Never touches anything already written to disk
        (backup folders, upgrade logs) - those are independent files, not linked to the list."""
        hosts = self._validate_hosts(hosts)
        with self.lock:
            self._poll_all()
            for h in hosts:
                if self.ops[h]["backup"]["state"] in ("queued", "running"):
                    raise ApiError(f"A backup is still running for {h} - wait for it to finish.", 409)
                if self._upgrade_active(h):
                    raise ApiError(f"{h} is being upgraded right now - wait for it to finish.", 409)
                if self.ops[h]["creds"]["state"] == "testing":
                    raise ApiError(f"A credentials test is still running for {h} - wait for it to finish.", 409)
            for h in hosts:
                label = self.info[h]["label"]
                self.monitor.remove_host(h)
                self.order.remove(h)
                del self.info[h]
                del self.ops[h]
                self.schedule.pop(h, None)
                self._rollout_forget(h)
                self.jobs.pop(h, None)
                self.assign.pop(h, None)
                self.term_close_host(h)
                self.events.add(h, label, "Removed from the device list", "info")
            save_site_devices(self.site_name, [self.info[h] for h in self.order])
            self._save_profiles()
            self._save_schedule()
        return {"message": f"Removed {len(hosts)} device(s) from the list."}

    # ---- validation -------------------------------------------------------- #
    def _validate_hosts(self, hosts):
        if not isinstance(hosts, list) or not hosts:
            raise ApiError("Select at least one switch.")
        out = []
        for h in hosts:
            if not isinstance(h, str) or h not in self.info:
                raise ApiError(f"Unknown switch: {str(h)[:60]!r}")
            if h not in out:
                out.append(h)
        return out

    # ---- credential profiles -------------------------------------------------- #
    def _save_profiles(self):
        if not self.site_name:
            return
        try:
            save_site_profiles(self.site_name, self.profiles, self.assign)
        except OSError as exc:
            raise ApiError(f"Could not save the credential profiles ({type(exc).__name__}).", 500)

    def profile_list(self):
        with self.lock:
            counts = {}
            for h, pname in self.assign.items():
                counts[pname] = counts.get(pname, 0) + 1
            return [{"name": n, "username": r["username"], "has_secret": bool(r.get("secret")),
                     "protocol": "telnet" if r.get("protocol") == "telnet" else "ssh",
                     "port": r.get("port") or default_port(r.get("protocol")),
                     "devices": counts.get(n, 0)} for n, r in sorted(self.profiles.items(), key=lambda kv: kv[0].lower())]

    def save_profile(self, name, username, password, secret, clear_secret=False, original=None, protocol="ssh", port=None):
        """Create a profile, or edit one (original = its current name). When editing, an empty
        password / secret means 'keep the saved one'."""
        self._require_site()
        name = name.strip() if isinstance(name, str) else ""
        username = username.strip() if isinstance(username, str) else ""
        if not PROFILE_NAME_RE.match(name):
            raise ApiError("Profile name: 1-40 characters - letters, digits, spaces, '.', '_' or '-'.")
        if not username or len(username) > 128:
            raise ApiError("Enter the switch username (up to 128 characters).")
        protocol = "telnet" if str(protocol or "ssh").strip().lower() == "telnet" else "ssh"
        try:
            port = clean_port(port, protocol)
        except ValueError as exc:
            raise ApiError(str(exc))
        for label, val in (("password", password), ("enable secret", secret)):
            if val is not None and (not isinstance(val, str) or len(val) > 256):
                raise ApiError(f"Invalid {label}.")
        with self.lock:
            editing = isinstance(original, str) and original in self.profiles
            if editing and name != original and name in self.profiles:
                raise ApiError(f"A profile called '{name}' already exists.", 409)
            if not editing and name in self.profiles:
                raise ApiError(f"A profile called '{name}' already exists - pick it from the list to edit it.", 409)
            old = self.profiles.get(original) if editing else None
            if not password and not old:
                raise ApiError("Enter the password for this profile.")
            try:
                new_pw = protect_secret(password) if password else old["password"]
                if secret:
                    new_secret = protect_secret(secret)
                else:
                    new_secret = "" if (clear_secret or not old) else old.get("secret", "")
            except Exception as exc:  # noqa: BLE001
                raise ApiError(f"Could not encrypt the password ({type(exc).__name__}). "
                               "Make sure the 'cryptography' package is installed.", 500)
            if editing and name != original:
                del self.profiles[original]
                for h, p in list(self.assign.items()):
                    if p == original:
                        self.assign[h] = name
            self.profiles[name] = {"username": username, "password": new_pw, "secret": new_secret, "protocol": protocol,
                                   "port": port}
            self._save_profiles()
        self.events.add("-", "", f"Credential profile '{name}' {'updated' if editing else 'created'}", "info")
        return {"message": f"Profile '{name}' saved ({'Telnet' if protocol == 'telnet' else 'SSH'}, port {port}).", "name": name}

    def delete_profile(self, name):
        with self.lock:
            if name not in self.profiles:
                raise ApiError("That profile does not exist.", 404)
            users = [h for h, p in self.assign.items() if p == name]
            if users:
                raise ApiError(f"'{name}' is still assigned to {len(users)} device(s) - assign them "
                               "another profile first.", 409)
            del self.profiles[name]
            self._save_profiles()
        self.events.add("-", "", f"Credential profile '{name}' deleted", "info")
        return {"message": f"Profile '{name}' deleted."}

    def assign_profile(self, hosts, profile):
        """Give every selected device this profile."""
        self._require_site()
        hosts = self._validate_hosts(hosts)
        with self.lock:
            if profile not in self.profiles:
                raise ApiError("Choose a profile to assign.")
            self._poll_all()
            for h in hosts:
                if self._upgrade_active(h) or self.ops[h]["backup"]["state"] in ("queued", "running"):
                    raise ApiError(f"{h} is busy - wait for its backup/upgrade to finish.", 409)
                self.assign[h] = profile
            self._save_profiles()
        self.events.add("-", "", f"Profile '{profile}' assigned to {len(hosts)} device(s)", "info")
        return {"message": f"Profile '{profile}' assigned to {len(hosts)} device(s)."}

    def _set_boot(self, host, boot):
        """Store a boot-mode reading (shown in the 'Boot mode' column, kept across restarts) and log an
        event when the mode or the boot variable changed. Called with self.lock held."""
        if not isinstance(boot, dict) or not (boot.get("mode") or boot.get("boot_ok") is not None):
            return
        o = self.ops.get(host)
        if o is None:
            return
        old = o.get("boot") or {}
        keys = ("mode", "boot_ok", "boot_var", "manual", "detail", "time")
        new = {k: boot.get(k) for k in keys}
        if old and all(old.get(k) == new[k] for k in keys):
            return                                      # the same reading again (status polls repeat it)
        new["prev"] = old.get("prev")
        if old.get("mode") and (old.get("mode") != new["mode"] or old.get("boot_var") != new["boot_var"]):
            new["prev"] = f"{old.get('mode') or '?'}, boot {old.get('boot_var') or '?'}"
            good = new["mode"] == "INSTALL" and new["boot_ok"]
            self.events.add(host, self.label(host),
                            f"Boot mode: {new['prev'].replace(', boot ', ' (boot ')}) -> "
                            f"{new['mode'] or '?'} (boot {new['boot_var'] or '?'})", "ok" if good else "warn")
        o["boot"] = new
        if o.get("version"):
            save_known_version(host, self.label(host), o["version"], o.get("model"), new)

    def _require_platform_known(self, hosts):
        """The platform, software version and boot mode come from a backup run or from 'Refresh
        version' - without them the switch's command set cannot be chosen. Called with self.lock held."""
        unknown = [h for h in hosts if not (self.ops[h].get("model") and self.ops[h].get("version"))]
        if unknown:
            shown = ", ".join(unknown[:5]) + (f" (+{len(unknown) - 5} more)" if len(unknown) > 5 else "")
            raise ApiError(f"The platform / software version of {shown} is not known yet. Run a backup, or "
                           "select the switch(es) and click 'Refresh version' just before upgrading.", 409)

    def _creds_map(self, hosts):
        """{host: (username, password, secret)} from each device's profile. Every device must
        have one - otherwise nothing is started."""
        with self.lock:
            missing = [h for h in hosts if self.assign.get(h) not in self.profiles]
            if missing:
                shown = ", ".join(missing[:5]) + (f" (+{len(missing) - 5} more)" if len(missing) > 5 else "")
                raise ApiError(f"No credentials profile is attached to: {shown}. Create a credentials "
                               "profile and attach it to the switch first (Credential profiles: pick a "
                               "profile, select the switch, press 'Assign to selected'). The platform, "
                               "software version and boot mode are read from the switch with it.")
            out = {}
            for h in hosts:
                rec = self.profiles[self.assign[h]]
                set_host_protocol(h, rec.get("protocol"), rec.get("port"))
                try:
                    out[h] = (rec["username"], unprotect_secret(rec["password"]),
                              unprotect_secret(rec.get("secret", "")))
                except Exception:  # noqa: BLE001
                    raise ApiError(f"The saved password of profile '{self.assign[h]}' cannot be read "
                                   "(was it saved by another Windows user?). Open the profile and "
                                   "enter the password again.")
            return out

    @staticmethod
    def _check_image(path):
        if not isinstance(path, str) or not path.strip():
            raise ApiError("Choose an IOS-XE image.")
        raw = path.strip().strip('"').strip("'")
        # A bare filename (normal case: picked from the dropdown) resolves against the
        # managed IOS Images folder; a path containing a separator is used as-is, so the
        # API still accepts a full path from another tool if ever called directly.
        p = Path(raw)
        p = (images_dir() / raw) if p.name == raw else Path(os.path.expanduser(raw))
        if p.suffix.lower() != ".bin":
            raise ApiError("The OS image must be a .bin file.")
        if not p.is_file():
            raise ApiError(f"Image not found: {p.name} - upload it first.")
        size = p.stat().st_size
        if size == 0:
            raise ApiError("The image file is empty.")
        try:
            verify_image_checksum_sidecar(p)
        except ValueError as exc:
            raise ApiError(f"Image integrity check failed: {exc}") from exc
        return str(p.resolve()), size

    def _upgrade_active(self, host):
        return self.ops[host]["upgrade"]["state"] in ACTIVE

    # ---- global settings (gear icon page), rollout queue and stop-on-failure ---------------- #
    def upgrade_settings_view(self):
        return {"max_parallel": self.max_parallel, "stop_on_failure": self.stop_on_failure,
                "limit": MAX_PARALLEL_LIMIT}

    def _apply_settings(self):
        st = self.settings
        self.max_parallel = st["max_parallel"]
        self.stop_on_failure = st["stop_on_failure"]
        apply_timeouts(st)                  # in-process commands (backup, readiness, ...) use these at once

    def settings_view(self):
        return {"values": dict(self.settings), "defaults": settings_defaults(), "groups": SETTINGS_GROUPS,
                "items": SETTINGS_ITEMS, "limit": MAX_PARALLEL_LIMIT,
                "about": {"name": APP_NAME, "version": APP_VERSION, "author": APP_AUTHOR, "repo": APP_REPO_URL,
                          "changelog": CHANGELOG}}

    def update_settings(self, values, who=None):
        if not isinstance(values, dict):
            raise ApiError("Send the settings as a JSON object.")
        with self.lock:
            merged = clean_settings({**self.settings, **values}, strict=True)
            changed = [i["label"] for i in SETTINGS_ITEMS if merged[i["key"]] != self.settings.get(i["key"])]
            if changed:
                try:
                    save_global_settings(merged)
                except OSError as exc:
                    raise ApiError(f"Could not save the settings file ({type(exc).__name__}).", 500)
                self.settings = merged
                self._apply_settings()
        if not changed:
            return {"message": "No changes to save.", **self.settings_view()}
        self.events.add("-", "", f"Settings changed by {who or 'local'}: " + "; ".join(changed[:4])
                        + (f" and {len(changed) - 4} more" if len(changed) > 4 else ""), "info")
        return {"message": "Settings saved. They apply to everything that starts from now on; upgrades already "
                           "running keep the values they started with.", **self.settings_view()}

    # ---- e-mail notifications (Settings > Notifications) ----------------------------------------- #
    def email_live_config(self):
        """What EMAILER reads on every notification (the saved settings, applied at once after a Save)."""
        cfg = dict(self.email_cfg)
        cfg["from"] = cfg.get("username", "")
        return cfg

    def email_settings_view(self):
        """For the Settings page. The password is never sent back - only whether one is saved."""
        cfg = self.email_cfg
        return {"values": {"enabled": cfg["enabled"], "host": cfg["host"], "port": cfg["port"],
                           "security": cfg["security"], "username": cfg["username"], "to": cfg["to"],
                           "kinds": list(cfg["kinds"])},
                "password_saved": bool(cfg.get("password_enc")),
                "kinds": [{"key": k, "label": label} for k, label, _on in EMAIL_KINDS],
                "problem_kinds": list(EMAIL_PROBLEM_KINDS),
                "security_options": [{"key": k, "label": v} for k, v in EMAIL_SECURITY.items()],
                "last": dict(EMAILER.last)}

    @staticmethod
    def _clean_email_values(values, strict):
        """(cfg without password, problems). strict=True also demands a complete form."""
        if not isinstance(values, dict):
            raise ApiError("Send the e-mail settings as a JSON object.")
        known = [k for k, _l, _o in EMAIL_KINDS]
        kinds = values.get("kinds")
        kinds = [k for k in (kinds if isinstance(kinds, list) else []) if k in known]
        cfg = {"enabled": bool(values.get("enabled")), "host": str(values.get("host") or "").strip()[:200],
               "port": values.get("port"), "security": values.get("security") if values.get("security") in EMAIL_SECURITY else "starttls",
               "username": str(values.get("username") or "").strip()[:200], "to": str(values.get("to") or "").strip()[:2000],
               "kinds": kinds}
        problems = []
        if strict:
            if not cfg["host"]:
                problems.append("Enter the SMTP server (for Gmail: smtp.gmail.com).")
            try:
                port = int(cfg["port"])
                if isinstance(cfg["port"], bool) or not 1 <= port <= 65535:
                    raise ValueError
                cfg["port"] = port
            except (TypeError, ValueError):
                problems.append("The port must be a number (587 for STARTTLS, 465 for SSL).")
            if not _EMAIL_RE.match(cfg["username"]):
                problems.append("Enter the e-mail address used to sign in to the mail server.")
            recipients = [r for r in re.split(r"[,;\s]+", cfg["to"]) if r]
            bad = [r for r in recipients if not _EMAIL_RE.match(r)]
            if not recipients:
                problems.append("Enter at least one address under 'Send to'.")
            elif bad:
                problems.append("Not a valid e-mail address: " + ", ".join(bad[:3]))
        else:
            try:
                cfg["port"] = int(cfg["port"])
            except (TypeError, ValueError):
                cfg["port"] = EMAIL_DEFAULTS["port"]
        return cfg, problems

    def update_email_settings(self, values, who=None):
        if not isinstance(values, dict):
            raise ApiError("Send the e-mail settings as a JSON object.")
        cfg, problems = self._clean_email_values(values, strict=bool((values or {}).get("enabled")))
        password = str((values or {}).get("password") or "")
        if cfg["enabled"] and not password and not self.email_cfg.get("password_enc"):
            problems.append("Enter the App password.")
        if problems:
            raise ApiError(" ".join(problems))
        cfg["password_enc"] = self.email_cfg.get("password_enc", "")
        if password:
            try:
                cfg["password_enc"] = protect_secret(password)
            except (Exception, SystemExit) as exc:  # noqa: BLE001
                raise ApiError(f"Could not encrypt the password ({type(exc).__name__}).", 500)
        was_on = self.email_cfg.get("enabled")
        try:
            save_email_settings(cfg)
        except OSError as exc:
            raise ApiError(f"Could not save the e-mail settings ({type(exc).__name__}).", 500)
        self.email_cfg = cfg                              # applied at once, no restart
        self.events.add("-", "", f"E-mail notification settings saved by {who or 'local'}: " + (
            f"ON ({len(cfg['kinds'])} type(s), {cfg['host']})" if cfg["enabled"]
            else ("switched OFF" if was_on else "OFF")), "info")
        out = self.email_settings_view()
        out["message"] = ("Saved - applied now." if cfg["enabled"] else "Saved - e-mail notifications are OFF.")
        if cfg["enabled"] and not cfg["kinds"]:
            out["message"] += " No notification type is ticked, so no e-mail will be sent."
        return out

    def test_email_settings(self, values, who=None):
        cfg, problems = self._clean_email_values(values, strict=True)
        password = str((values or {}).get("password") or "")
        if not password and not self.email_cfg.get("password_enc"):
            problems.append("Enter the App password.")
        if problems:
            raise ApiError(" ".join(problems))
        cfg["from"] = cfg["username"]
        if password:
            cfg["password"] = password                    # used for this test only, not stored
        else:
            cfg["password_enc"] = self.email_cfg.get("password_enc", "")
        try:
            n = send_test_email(cfg)
        except (Exception, SystemExit) as exc:  # noqa: BLE001
            raise ApiError("Test e-mail failed: " + describe_email_error(exc))
        self.events.add("-", "", f"Test e-mail sent by {who or 'local'} to {n} address(es)", "info")
        return {"message": f"Test e-mail sent to {n} address(es). Check the inbox."}

    # ---- command duration profiles + command time test (Settings > Command Outputs) ------------ #
    def _current_timeouts(self):
        return {k: self.settings[k] for k in TIMEOUT_ITEMS}

    def timeout_profiles_view(self):
        with _TIMEOUT_PROFILES_LOCK:
            store = load_timeout_profiles()
        profiles = [{"name": DEFAULT_TIMEOUT_PROFILE, "builtin": True,
                     "values": {k: i["default"] for k, i in TIMEOUT_ITEMS.items()},
                     "note": "The timeouts the tool ships with.", "saved_at": "", "saved_by": "", "site": "",
                     "source": "", "switch": ""}]
        for name in sorted(store["profiles"], key=str.lower):
            profiles.append({"name": name, "builtin": False, **store["profiles"][name]})
        by_name = {p["name"]: p for p in profiles}
        current = self._current_timeouts()
        order = ([store["last_applied"]] if store["last_applied"] in by_name else []) + list(by_name)
        active = next((n for n in order if by_name[n]["values"] == current), None)
        presets = [{"id": p["id"], "title": p["title"], "command": p["command"], "group": p["group"],
                    "kind": p["kind"], "keys": list(p["keys"]), "warn": p.get("warn", ""), "desc": p.get("desc", ""),
                    "needs_image": bool(p.get("needs_image")),
                    "steps": [{"command": c, "key": k} for c, k in p.get("steps", ())]}
                   for p in COMMAND_TEST_PRESETS]
        return {"profiles": profiles, "active": active, "presets": presets, "phase_map": PHASE_TIMEOUT_KEYS,
                "images": list_images()}

    @staticmethod
    def _find_timeout_profile(store, name):
        name = str(name or "").strip()
        if name.lower() == DEFAULT_TIMEOUT_PROFILE.lower():
            return DEFAULT_TIMEOUT_PROFILE
        return next((n for n in store["profiles"] if n.lower() == name.lower()), None)

    def timeout_profile_save(self, name, values, note="", source="manual", site="", switch="", overwrite=False,
                             who=None):
        name = str(name or "").strip()
        if not TIMEOUT_PROFILE_NAME_RE.match(name):
            raise ApiError("Give the profile a name of 1 to 40 characters: letters, digits, spaces, dot, dash or "
                           "underscore, starting with a letter or digit.")
        if name.lower() == DEFAULT_TIMEOUT_PROFILE.lower():
            raise ApiError("'Default' is the built-in profile and cannot be changed. Choose another name.")
        clean = clean_timeout_values(values, strict=True)
        with _TIMEOUT_PROFILES_LOCK:
            store = load_timeout_profiles()
            existing = self._find_timeout_profile(store, name)
            if existing and not overwrite:
                raise ApiError(f"A profile named '{existing}' already exists.", 409)
            name = existing or name
            store["profiles"][name] = {
                "values": clean, "note": str(note or "")[:300], "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "saved_by": str(who or "local")[:60], "site": str(site or "")[:80],
                "source": "command test" if source == "command test" else "manual", "switch": str(switch or "")[:80]}
            try:
                save_timeout_profiles(store)
            except OSError as exc:
                raise ApiError(f"Could not save the profile file ({type(exc).__name__}).", 500)
        self.events.add("-", "", f"Command duration profile '{name}' {'updated' if existing else 'created'} by "
                                 f"{who or 'local'}", "info")
        return {"message": f"Profile '{name}' {'updated' if existing else 'saved'}. Pick it in the list and click "
                           "Apply to use it.", "saved": name, **self.timeout_profiles_view()}

    def timeout_profile_delete(self, name, who=None):
        if str(name or "").strip().lower() == DEFAULT_TIMEOUT_PROFILE.lower():
            raise ApiError("The built-in Default profile cannot be deleted.")
        with _TIMEOUT_PROFILES_LOCK:
            store = load_timeout_profiles()
            match = self._find_timeout_profile(store, name)
            if not match:
                raise ApiError("That profile does not exist.", 404)
            del store["profiles"][match]
            if store["last_applied"] == match:
                store["last_applied"] = ""
            try:
                save_timeout_profiles(store)
            except OSError as exc:
                raise ApiError(f"Could not save the profile file ({type(exc).__name__}).", 500)
        self.events.add("-", "", f"Command duration profile '{match}' deleted by {who or 'local'}", "info")
        return {"message": f"Profile '{match}' deleted. The timeouts in use did not change.",
                **self.timeout_profiles_view()}

    def timeout_profile_apply(self, name, who=None):
        with _TIMEOUT_PROFILES_LOCK:
            store = load_timeout_profiles()
            match = self._find_timeout_profile(store, name)
            if not match:
                raise ApiError("That profile does not exist.", 404)
            values = ({k: i["default"] for k, i in TIMEOUT_ITEMS.items()} if match == DEFAULT_TIMEOUT_PROFILE
                      else dict(store["profiles"][match]["values"]))
        self.update_settings(values, who)               # validates, saves settings.json, applies, logs the change
        with _TIMEOUT_PROFILES_LOCK:
            store = load_timeout_profiles()
            store["last_applied"] = match
            try:
                save_timeout_profiles(store)
            except OSError:
                pass                                    # the values are applied; only the 'last used' hint is lost
        self.events.add("-", "", f"Command duration profile '{match}' applied by {who or 'local'}", "info")
        return {"message": f"Profile '{match}' applied and saved. Everything that starts from now on uses it; "
                           "upgrades already running keep the values they started with.",
                "settings": self.settings_view(), **self.timeout_profiles_view()}

    def command_test_start(self, host, cmd_id, who=None, image=""):
        """Command time test: start ONE step on ONE switch in the background and return a job id. The browser polls
        command_test_poll for the elapsed time, the live output and finally the result (a step such as install add
        or the reload can take many minutes, which must not hold one HTTP request open)."""
        preset = next((p for p in COMMAND_TEST_PRESETS if p["id"] == cmd_id), None)
        if preset is None:
            raise ApiError("Unknown test step.", 404)
        host = self._validate_hosts([host])[0]
        image = str(image or "").strip()
        image_path = ""
        if image:
            if image not in list_images():
                raise ApiError(f"Image not found: {image} - upload it first (Storage > IOS-XE images).")
            image_path = str(images_dir() / image)
        elif preset.get("needs_image"):
            raise ApiError("Choose the IOS-XE image for this step first.")
        username, password, secret = self._creds_map([host])[host]
        with self.lock:
            self._poll_all()
            if self._upgrade_active(host):
                raise ApiError(f"{host} is being upgraded right now - wait for it to finish.", 409)
            if host in self._cmdtest_busy:
                raise ApiError(f"A test step is already running on {host}.", 409)
            self._cmdtest_busy.add(host)
            dtype = BACKUP_DEVICE_TYPES.get(self.info[host]["type"], "cisco_xe")
            now = time.time()                              # forget old finished jobs
            for jid in [j for j, v in self._cmdtest_jobs.items() if v["state"] == "done" and now - v["ended"] > 900]:
                del self._cmdtest_jobs[jid]
            job = {"job": secrets.token_hex(6), "host": host, "id": cmd_id, "state": "running", "started": now,
                   "ended": None, "output": "", "log": "", "result": None, "stop": False}
            self._cmdtest_jobs[job["job"]] = job
        device = {"device_type": dtype, "host": host, "username": username, "password": password,
                  "secret": secret or "", "conn_timeout": 60, "auth_timeout": 60, "banner_timeout": 60,
                  "keepalive": 30}
        threading.Thread(target=self._command_test_worker, args=(job, preset, device, image_path, image),
                         daemon=True, name=f"cmdtest-{host}").start()
        return {"job": job["job"], "host": host, "id": cmd_id}

    def command_test_poll(self, job_id, offset=0):
        """state, the elapsed time and the console text written since `offset` (the Console panel shows it live)."""
        job = self._cmdtest_jobs.get(str(job_id or ""))
        if job is None:
            raise ApiError("That test run is no longer known (the program was restarted?).", 404)
        try:
            offset = max(0, int(offset))
        except (TypeError, ValueError):
            offset = 0
        log = job["log"]
        base = {"text": log[offset:], "offset": len(log)}
        if job["state"] == "running":
            return {"state": "running", "elapsed": round(time.time() - job["started"], 1),
                    "output": job["output"][-8000:], "stopping": bool(job["stop"]), **base}
        return {"state": "done", "result": job["result"], **base}

    @staticmethod
    def _cmdtest_emit(job, text):
        """Append text to the job's console log (capped; the Command time test never needs more than this)."""
        if len(job["log"]) < 2_000_000:
            job["log"] += text.replace("\r\n", "\n").replace("\r", "")
            if len(job["log"]) >= 2_000_000:
                job["log"] += "\n[console output truncated]\n"

    def command_test_stop(self, job_id):
        job = self._cmdtest_jobs.get(str(job_id or ""))
        if job is None:
            raise ApiError("That test run is no longer known.", 404)
        if job["state"] == "running":
            job["stop"] = True
        return {"stopping": job["state"] == "running"}

    def _image_md5(self, path):
        stat = os.stat(path)
        key = (path, stat.st_mtime_ns, stat.st_size)
        if key not in self._img_md5:
            self._img_md5[key] = check_md5(path)
        return self._img_md5[key]

    def _command_test_worker(self, job, preset, device, image_path="", image_name=""):
        host = device["host"]
        self._cmdtest_emit(job, f"\n=== {now_str()}  {preset['title']} ===\n")
        try:
            result = self._command_test_exec(job, preset, device, image_path, image_name)
        except Exception as exc:  # noqa: BLE001 - never leave a job running forever
            result = {"id": preset["id"], "command": preset["title"], "host": host, "kind": preset["kind"],
                      "status": "error", "detail": f"{type(exc).__name__}: {exc}", "connect_seconds": None,
                      "seconds": None, "lines": 0, "chars": 0, "output": job["output"], "truncated": False,
                      "parts": [], "time": now_str()}
        took = f"{result['seconds']:.2f} s" if result.get("seconds") is not None else "no time"
        self._cmdtest_emit(job, f"--- {preset['title']}: {result['status']}, {took}"
                                + (f" - {result['detail']}" if result.get("detail") else "") + " ---\n")
        with self.lock:
            self._cmdtest_busy.discard(host)
            job["result"], job["ended"], job["state"] = result, time.time(), "done"
        if host in self.info:
            name = preset["title"]
            if result["status"] in ("ok", "warning", "timeout") and result["seconds"] is not None:
                self.events.add(host, self.label(host), f"Command time test: '{name}' took {result['seconds']:.2f} s"
                                + (f" - {result['detail']}" if result["detail"] else ""),
                                "info" if result["status"] == "ok" else "warn")
            else:
                self.events.add(host, self.label(host), f"Command time test: '{name}' - {result['detail'] or result['status']}",
                                "warn" if result["status"] in ("stopped", "skipped") else "bad")

    def _command_test_exec(self, job, preset, device, image_path="", image_name=""):
        kind, host = preset["kind"], device["host"]
        result = {"id": preset["id"], "command": preset["title"], "host": host, "kind": kind, "status": "ok",
                  "detail": "", "connect_seconds": None, "seconds": None, "lines": 0, "chars": 0, "output": "",
                  "truncated": False, "parts": [], "time": now_str()}

        def stop():
            return bool(job.get("stop"))

        lab = self.label(host)
        sent = [0]

        def emit(text):
            self._cmdtest_emit(job, text)

        def keep(out):                                       # cumulative output of a streamed command -> live console
            emit(out[sent[0]:])
            sent[0] = len(out)
            job["output"] = out[-20000:]
        keep.reset = lambda: sent.__setitem__(0, 0)

        def note(line):
            emit(line + "\n")
            job["output"] = (job["output"] + line + "\n")[-20000:]

        def set_output(text):
            result["output"], result["lines"], result["chars"] = text[:COMMAND_TEST_OUTPUT_LIMIT], len(text.splitlines()), len(text)
            result["truncated"] = len(text) > COMMAND_TEST_OUTPUT_LIMIT

        def run_one(conn, command, limit):
            """One command -> (status, seconds, output, detail). The ceiling is the largest value the timeout can be
            set to, so a slow command is measured, not cut short."""
            began = time.perf_counter()
            emit(f"{lab}# {command}\n")
            try:
                out = conn.send_command(command, read_timeout=limit) or ""
                emit(out.rstrip("\n")[:40000] + ("\n" if out.strip() else ""))
                status, detail = "ok", ""
                if COMMAND_TEST_BAD_OUTPUT.search(out[:400]):
                    status, detail = "unsupported", "The switch did not accept this command (not supported here)."
                return status, round(time.perf_counter() - began, 2), out, detail
            except ReadTimeout:
                return "timeout", round(time.perf_counter() - began, 2), "", f"No complete output within {limit} s."
            except CONNECTION_LOST as exc:
                return "error", None, "", f"The session ended while the command ran ({type(exc).__name__})."

        def run_stream(conn, command, limit, fail_re=None):
            """A streamed command (answers [y/n], returns when the prompt is back, can be stopped)."""
            began = time.perf_counter()
            keep.reset()
            emit(f"{lab}# {command}\n")
            try:
                out = stream_command(conn, command, None, limit, "test", on_output=keep, should_stop=stop)
            except ReadTimeout as exc:
                out = getattr(exc, "output", "") or ""
                set_output(out)
                result["seconds"] = round(time.perf_counter() - began, 2)
                if getattr(exc, "stopped", False):
                    result["status"], result["detail"] = "stopped", "Stopped by you after " + f"{result['seconds']:.0f} s. The output received so far is shown."
                else:
                    result["status"], result["detail"] = "timeout", f"No prompt back within {limit} s. The output received so far is shown."
                return None
            except EOFError as exc:
                result["status"], result["detail"] = "error", str(exc)
                return None
            result["seconds"] = round(time.perf_counter() - began, 2)
            set_output(out)
            if fail_re is not None and fail_re.search(out):
                if re.search(r"already\s+(been\s+)?added|nothing to commit|already committed", out, re.I):
                    result["status"], result["detail"] = "skipped", "Nothing to do - the switch says it is already done."
                else:
                    result["status"], result["detail"] = "error", "The switch reported a failure - see the output."
            return out

        try:
            started = time.perf_counter()
            try:
                with ConnectHandler(**device) as conn:
                    if device["secret"]:
                        conn.enable()
                    conn.find_prompt()
                    result["connect_seconds"] = round(time.perf_counter() - started, 2)
                    began = time.perf_counter()
                    if kind == "bundle":
                        lines = []
                        for command, key in preset["steps"]:
                            st, sec, out, det = run_one(conn, command, TIMEOUT_ITEMS[key]["max"])
                            result["parts"].append({"command": command, "key": key, "status": st, "seconds": sec,
                                                    "detail": det})
                            lines.append(f"{command:<28} {('%.2f s' % sec) if sec is not None else '-':>10}   {st}")
                            job["output"] = "\n".join(lines)
                        result["seconds"] = round(time.perf_counter() - began, 2)
                        result["output"] = "\n".join(lines)
                        bad = [p for p in result["parts"] if p["status"] in ("error", "timeout")]
                        if bad:
                            result["status"] = bad[0]["status"]
                            result["detail"] = f"'{bad[0]['command']}': {bad[0]['detail']}"
                    elif kind == "stream":
                        cmd = preset["command"].replace("<image>", image_name)
                        fail = {"INSTALL_FAIL_RE": INSTALL_FAIL_RE}.get(preset.get("fail_re"))
                        if preset["id"] == "install_add":
                            cmd = f"install add file flash:{image_name}"
                        run_stream(conn, cmd, TIMEOUT_ITEMS[preset["keys"][0]]["max"], fail)
                    elif kind == "upload":
                        try:
                            verify_image_checksum_sidecar(Path(image_path))
                        except ValueError as exc:
                            raise ApiError(f"Image integrity check failed: {exc}")
                        listing = conn.send_command(f"dir flash:{image_name}", read_timeout=tmo("show_cmd")) or ""
                        existed = image_name in listing and not re.search(r"No such file|%\s*Error", listing)
                        note(f"Copying {image_name} ({fmt_bytes(os.path.getsize(image_path))}) ...")
                        state = {"added": False}
                        began = time.perf_counter()
                        transfer_file(conn, image_path, image_name, state=state)
                        result["seconds"] = round(time.perf_counter() - began, 2)
                        lines = [f"Copied {image_name} in {result['seconds']:.1f} s"]
                        if state["added"]:
                            conn.send_config_set(["no ip scp server enable"])
                            lines.append("'ip scp server enable' was switched on for the copy and is removed again.")
                        if existed:
                            result["status"] = "warning"
                            result["detail"] = "The image was already on the flash, so the copy was skipped or very short - not a real upload time."
                        lines.append(conn.send_command(f"dir flash:{image_name}", read_timeout=tmo("show_cmd")) or "")
                        emit(f"{lab}# dir flash:{image_name}\n{lines[-1].rstrip()}\n")
                        set_output("\n".join(lines))
                    elif kind == "md5":
                        note("Calculating the MD5 of the image on this PC ...")
                        expected = self._image_md5(image_path)
                        note(f"Local MD5: {expected}\nAsking the switch (verify /md5) ...")
                        out = run_stream(conn, f"verify /md5 flash:{image_name} {expected}", TIMEOUT_ITEMS["md5_verify"]["max"])
                        if out is not None and not re.search(r"Verified", out):
                            result["status"] = "error"
                            result["detail"] = "The switch did not report 'Verified' - MD5 mismatch, or the image is not on the flash."
                    elif kind == "boot":
                        emit(f"{lab}# configure terminal\n")
                        began = time.perf_counter()
                        out = conn.send_config_set(BOOT_FIX_COMMANDS)
                        result["seconds"] = round(time.perf_counter() - began, 2)
                        emit((out or "").rstrip("\n") + "\n")
                        state, raw = read_boot_state(conn)
                        emit(f"{lab}# show boot\n{(raw or '').rstrip()}\n")
                        set_output((out or "") + "\n" + (raw or ""))
                        if not state.get("boot_ok"):
                            result["status"], result["detail"] = "warning", "show boot does not confirm the change: " + state["detail"]
                        else:
                            result["detail"] = state["detail"]
                    elif kind == "activate":
                        self._test_activate(job, preset, device, conn, result, image_name, note, keep, stop)
                    else:
                        ceiling = max(TIMEOUT_ITEMS[k]["max"] for k in preset["keys"]) if preset["keys"] else 120
                        st, sec, out, det = run_one(conn, preset["command"], ceiling)
                        result["status"], result["seconds"], result["detail"] = st, sec, det
                        set_output(out)
            except NetmikoAuthenticationException:
                result["status"], result["detail"] = "connect_failed", "Authentication failed (bad username/password or enable secret)."
            except NetmikoTimeoutException:
                result["status"], result["detail"] = "connect_failed", "Connection timed out."
            except ApiError as exc:
                result["status"], result["detail"] = "error", str(exc)
            except Exception as exc:  # noqa: BLE001
                result["status"], result["detail"] = ("connect_failed" if result["connect_seconds"] is None else "error"), f"{type(exc).__name__}: {exc}"
        finally:
            pass
        return result

    def _test_activate(self, job, preset, device, conn, result, image_name, note, keep, stop):
        """install activate, then watch the reload: times the activation itself and then how long ping stays down, how
        long until ping and SSH answer again, and what version the switch came back with."""
        host = device["host"]
        full = preset["id"] == "install_full"
        command = f"install add file flash:{image_name} activate commit" if full else "install activate"
        fail_re = INSTALL_FAIL_RE if full else ACTIVATE_FAIL_RE
        show_version = conn.send_command("show version", read_timeout=tmo("show_cmd")) or ""
        old_version, model = parse_version(show_version), parse_model(show_version)
        if install_flavor(model, old_version) == "legacy":
            result["status"] = "error"
            result["detail"] = ("This release (Catalyst 9300 before 16.6.2) has no 'install activate'; it uses the one-step "
                                "'request platform software package install'. Use the normal Upgrade for it.")
            return
        expected = guess_version_from_filename(image_name) if image_name else None
        note(f"Running {old_version or 'unknown'}" + (f", image {image_name} (version {expected})" if expected else ""))
        keep.reset()
        note(f"{device['host']}# {command}")
        began = time.perf_counter()
        activate_s, dropped = None, False
        try:
            out = stream_command(conn, command, None, TIMEOUT_ITEMS["install_total"]["max"], "test",
                                 on_output=keep, should_stop=stop)
            activate_s = round(time.perf_counter() - began, 2)
            if fail_re.search(out):
                result["seconds"], result["status"] = activate_s, "error"
                result["detail"] = ("The switch reported a failure - see the output." if full else
                                    "The switch reported a failure - nothing to activate? Run 'Add the image' first. See the output.")
                result["output"] = out[:COMMAND_TEST_OUTPUT_LIMIT]
                return
            note("The prompt came back; the switch normally reloads a few seconds later.")
        except CONNECTION_LOST:
            activate_s, dropped = round(time.perf_counter() - began, 2), True
            note(f"The session dropped after {activate_s} s - the switch is reloading.")
        except ReadTimeout as exc:
            result["seconds"] = round(time.perf_counter() - began, 2)
            result["output"] = (getattr(exc, "output", "") or job["output"])[:COMMAND_TEST_OUTPUT_LIMIT]
            if getattr(exc, "stopped", False):
                result["status"], result["detail"] = "stopped", "Stopped by you before the switch reloaded."
            else:
                result["status"], result["detail"] = "timeout", "The activation neither finished nor dropped the session."
            return
        username, password, secret = device["username"], device["password"], device["secret"] or None
        limits = {"down": tmo("reboot_down"), "up": TIMEOUT_ITEMS["reboot_up"]["max"], "ssh": TIMEOUT_ITEMS["reboot_ssh"]["max"]}
        w = watch_reload_timed(host, username, password, secret, limits, emit=note, should_stop=stop,
                               extra=["show install summary"])
        result["seconds"] = round(time.perf_counter() - began, 2)
        parts = [{"command": ("add + activate + commit" if full else "install activate") + " (until the session dropped)",
                  "key": "install_total", "status": "ok",
                  "seconds": activate_s, "detail": "" if dropped else "the prompt came back before the drop"}]
        parts.append({"command": "Ping stops answering", "key": "reboot_down",
                      "status": "ok" if w["down"] is not None else "timeout", "seconds": w["down"],
                      "detail": "" if w["down"] is not None else f"never stopped answering within {limits['down']} s"})
        if w["down"] is not None:
            parts.append({"command": "Ping answers again", "key": "reboot_up", "status": "ok" if w["up"] is not None else "timeout",
                          "seconds": w["up"], "detail": "" if w["up"] is not None else w["error"]})
        if w["up"] is not None or w["down"] is None:
            parts.append({"command": "SSH login works", "key": "reboot_ssh", "status": "ok" if w["ssh"] is not None else "timeout",
                          "seconds": w["ssh"], "detail": "" if w["ssh"] is not None else w["error"]})
        result["parts"] = parts
        result["output"] = job["output"][-COMMAND_TEST_OUTPUT_LIMIT:]
        if w["stopped"]:
            result["status"], result["detail"] = "stopped", "Stopped by you while waiting for the switch. The times so far are shown."
            return
        problems = []
        if w["error"]:
            result["status"] = "timeout"
            problems.append(w["error"])
        elif w["version"]:
            if old_version and w["version"] == old_version:
                problems.append(f"the switch came back on the SAME version ({w['version']}) - the new image is not running")
            elif expected and w["version"] != expected:
                problems.append(f"the switch runs {w['version']}, the image file name suggests {expected}")
            boot = w.get("boot") or {}
            if boot and (boot.get("mode") != "INSTALL" or boot.get("boot_ok") is False):
                problems.append(f"boot mode {boot.get('mode') or 'unknown'}: {boot.get('detail', '')}")
        elif w["ssh"] is not None:
            problems.append("the software version could not be read after the reload")
        summary = w["extra"].get("show install summary", "")
        uncommitted = bool(re.search(r"(?m)^\s*IMG\s+U\b", summary))
        if full and uncommitted:
            problems.append("the image is activated but NOT committed - run 'install commit' (or look at show install summary)")
        if problems:
            if result["status"] == "ok":
                result["status"] = "warning"
            result["detail"] = "; ".join(problems)
        elif full:
            result["detail"] = f"back on {w['version']}" + (", committed" if re.search(r"(?m)^\s*IMG\s+C\b", summary) else "")
        else:
            result["detail"] = f"back on {w['version']} - now run Commit"

    def _rebuild_rollouts(self):
        """Re-create the rollout records of schedule entries loaded from disk."""
        self.rollouts.clear()
        self.rollout_of.clear()
        for h, e in self.schedule.items():
            rid = e.get("rollout")
            if rid:
                self.rollouts.setdefault(rid, {"hosts": set(), "failed": False, "by": e.get("scheduled_by")})
                self.rollouts[rid]["hosts"].add(h)
                self.rollout_of[h] = rid

    def _new_rollout(self, hosts, by):
        """A group of switches started together; if one fails, the ones still waiting are not started."""
        rid = secrets.token_hex(4)
        for h in hosts:
            self._rollout_forget(h)
            self.rollout_of[h] = rid
        self.rollouts[rid] = {"hosts": set(hosts), "failed": False, "by": by or "local"}
        return rid

    def _rollout_forget(self, host):
        rid = self.rollout_of.pop(host, None)
        if rid in self.rollouts:
            self.rollouts[rid]["hosts"].discard(host)

    def _check_rollouts(self):
        """Called with self.lock held. If a switch of a rollout failed (or needs attention) and
        'stop on failure' is on, the switches of that rollout that have not started yet are
        removed from the queue. Switches already running are left to finish."""
        for rid, ro in list(self.rollouts.items()):
            hosts = [h for h in ro["hosts"] if h in self.ops]
            waiting = [h for h in hosts if (self.schedule.get(h) or {}).get("rollout") == rid]
            started = [h for h in hosts if h not in waiting]
            if self.stop_on_failure and not ro["failed"] and waiting:
                bad = [h for h in started if self.ops[h]["upgrade"]["state"] in ("failed", "attention")]
                if bad:
                    ro["failed"] = True
                    for h in waiting:
                        self.schedule.pop(h, None)
                        self._rollout_forget(h)
                        self.events.add(h, self.label(h), "Upgrade not started: "
                                        f"{self.label(bad[0])} failed earlier in this rollout", "warn")
                    self._save_schedule()
                    self._notify(bad[0], "failed", "Rollout stopped",
                                 f"{self.label(bad[0])} did not finish its upgrade, so {len(waiting)} "
                                 "switch(es) that were still waiting were not started. Running ones continue.")
                    waiting = []
            if not waiting and not any(self._upgrade_active(h) for h in started):
                for h in hosts:
                    self.rollout_of.pop(h, None)
                del self.rollouts[rid]

    def _active_reload_count(self):
        return sum(1 for host in self.order
                   if self._upgrade_active(host)
                   and self.ops[host]["upgrade"].get("action") in ("full", "activate"))

    # ---- backup ------------------------------------------------------------ #
    MAX_CUSTOM_COMMANDS = 20
    MAX_COMMAND_LENGTH = 200

    @classmethod
    def _parse_custom_commands(cls, text):
        """Free-text box, one command per line -> a clean list. Not required."""
        if not text:
            return []
        if not isinstance(text, str):
            raise ApiError("Additional commands must be text.")
        lines = [ln.strip() for ln in text.splitlines()]
        commands = [ln for ln in lines if ln]
        if len(commands) > cls.MAX_CUSTOM_COMMANDS:
            raise ApiError(f"Enter at most {cls.MAX_CUSTOM_COMMANDS} additional commands.")
        for c in commands:
            if len(c) > cls.MAX_COMMAND_LENGTH:
                raise ApiError(f"Command too long (max {cls.MAX_COMMAND_LENGTH} characters): {c[:40]}...")
        return commands

    def start_backup(self, hosts, custom_commands_text=None):
        self._require_site()
        hosts = self._validate_hosts(hosts)
        cm = self._creds_map(hosts)
        custom_commands = self._parse_custom_commands(custom_commands_text)
        with self.lock:
            self._poll_all()
            for h in hosts:
                if self._upgrade_active(h):
                    raise ApiError(f"{h} is being upgraded right now - deselect it and try again.", 409)
                if self.ops[h]["backup"]["state"] in ("queued", "running"):
                    raise ApiError(f"A backup is already running for {h}.", 409)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            run = {"id": stamp, "total": len(hosts), "done": 0, "ok": 0,
                   "failed": 0, "started": now_str(), "finished": None, "results": [],
                   "commands": [], "custom_commands": custom_commands}
            self.backup_run = run
            for h in hosts:
                self.ops[h]["backup"] = {"state": "queued", "detail": "Waiting ...", "time": None}
        for h in hosts:
            self.pool.submit(self._backup_one, run, h, *cm[h], custom_commands)
        msg = f"Backup started for {len(hosts)} switch(es)."
        if custom_commands:
            msg += f" Running {len(custom_commands)} additional command(s) on each."
        return {"message": msg, "run": run["id"]}

    @activity("Backup")
    def _backup_one(self, run, host, username, password, secret, custom_commands):
        with self.lock:
            self.ops[host]["backup"].update(state="running", detail="Connecting ...")
        dtype = BACKUP_DEVICE_TYPES.get(self.info[host]["type"], "cisco_xe")
        label = self.info[host]["label"] or host
        # One PERMANENT folder per device (not per run) - every backup adds timestamped files
        # here, so a device's history accumulates in one place instead of scattering across
        # a new top-level folder each time.
        device_dir = self.backup_dir / device_folder_name(label, host)
        ok, detail, hostname, commands_run = False, "", "", []
        try:
            res = do_backup(host, dtype, username, password, secret, device_dir, run["id"], custom_commands)
            ok, hostname, commands_run = True, res["hostname"], res["commands_run"]
            detail = f"Saved 1 file with {len(res['commands_run'])} command output(s)"
            if res["warnings"]:
                detail += " | WARNING: " + "; ".join(res["warnings"])
            with self.lock:
                o = self.ops[host]
                if res.get("model"):
                    o["model"] = res["model"]
                if res["version"]:
                    o["version"], o["version_note"] = res["version"], f"from backup {now_str()}"
                    o["prev_version"] = None
                    save_known_version(host, label, res["version"], o.get("model"))
                self._set_boot(host, res.get("boot"))
        except NetmikoAuthenticationException:
            detail = "Authentication failed"
        except NetmikoTimeoutException:
            detail = "Connection timed out"
        except Exception as exc:  # noqa: BLE001
            detail = f"{type(exc).__name__}: {exc}"

        with self.lock:
            self.ops[host]["backup"] = {"state": "completed" if ok else "failed",
                                        "detail": detail, "time": now_str()}
            run["done"] += 1
            run["ok" if ok else "failed"] += 1
            run["results"].append({"host": host, "hostname": hostname,
                                   "status": "OK" if ok else "FAILED", "detail": detail})
            for c in commands_run:
                run["commands"].append({"host": host, "label": label, **c})
            if not commands_run and not ok:
                run["commands"].append({"host": host, "label": label, "command": "(connect)",
                                        "status": "FAILED", "detail": detail})
            finished = run["done"] == run["total"]
            if finished:
                run["finished"] = now_str()
        self.events.add(host, self.label(host), "Backup completed" if ok else f"Backup FAILED: {detail}",
                        "ok" if ok else "bad")
        if finished:
            self._write_backup_summary(run)
            self.events.add("-", "", f"Backup run finished: {run['ok']} of {run['total']} completed"
                            + (f", {run['failed']} failed" if run["failed"] else ""),
                            "ok" if not run["failed"] else "warn")

    def _write_backup_summary(self, run):
        try:
            self.reports_dir.mkdir(parents=True, exist_ok=True)
            with open(self.reports_dir / f"backup_summary_{run['id']}.csv", "w", newline="",
                      encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["host", "hostname", "status", "detail"])
                w.writeheader()
                w.writerows(sorted(run["results"], key=lambda r: r["host"]))
        except OSError:
            pass

    # ---- version check ------------------------------------------------------ #
    def start_version_check(self, hosts):
        hosts = self._validate_hosts(hosts)
        cm = self._creds_map(hosts)
        with self.lock:
            self._poll_all()
            for h in hosts:
                if self._upgrade_active(h):
                    raise ApiError(f"{h} is being upgraded right now - its version is checked "
                                   "automatically when the upgrade ends.", 409)
        for h in hosts:
            self.pool.submit(self._version_one, h, *cm[h])
        return {"message": f"Reading the software version from {len(hosts)} switch(es) ..."}

    @activity("Version check")
    def _version_one(self, host, username, password, secret):
        dtype = BACKUP_DEVICE_TYPES.get(self.info[host]["type"], "cisco_xe")
        device = {"device_type": dtype, "host": host, "username": username, "password": password,
                  "secret": secret, "conn_timeout": 20, "auth_timeout": 20, "banner_timeout": 20}
        try:
            with ConnectHandler(**device) as conn:
                if secret:
                    conn.enable()
                sv = conn.send_command("show version", read_timeout=tmo('show_cmd'))
                version, model = parse_version(sv), parse_model(sv)
                try:
                    sb = conn.send_command("show boot", read_timeout=tmo('show_cmd'))
                    boot = boot_state(sv, sb) if sb else None
                except Exception:  # noqa: BLE001 - the version is still worth having
                    boot = None
            if not version:
                raise ValueError("could not find the version in 'show version'")
        except NetmikoAuthenticationException:
            self.events.add(host, self.label(host), "Version check failed: authentication failed", "bad")
            return
        except NetmikoTimeoutException:
            self.events.add(host, self.label(host), "Version check failed: connection timed out", "bad")
            return
        except Exception as exc:  # noqa: BLE001
            self.events.add(host, self.label(host), f"Version check failed: {type(exc).__name__}: {exc}", "bad")
            return

        confirmed = None
        with self.lock:
            o, up = self.ops[host], self.ops[host]["upgrade"]
            old = up.get("old_version")
            if o["version"] and o["version"] != version:
                o["prev_version"] = o["version"]
            o["version"], o["version_note"] = version, f"checked {now_str()}"
            if model:
                o["model"] = model
            self._set_boot(host, boot)
            if up["state"] == "attention" and old and version != old:
                # An upgrade that ended in CHECK is now proven: the switch runs a new version.
                o["prev_version"] = old
                confirmed = f"Upgrade completed: {old} -> {version} (confirmed by version check)"
                up.update(state="completed", new_version=version, message=confirmed, finished=now_str())
        save_known_version(host, self.label(host), version, model)
        info = (f" ({model})" if model else "") + (f", {boot['mode']} mode" if boot and boot.get("mode") else "")
        self.events.add(host, self.label(host), confirmed or f"Software version: {version}{info}",
                        "ok" if confirmed else "info")

    def test_credentials(self, hosts, profile=None, adhoc=None):
        """Log in to the switches. With `profile` (a saved profile name) that profile is tried on every
        switch WITHOUT assigning it; without it each switch's own assigned profile is used.
        `adhoc` = the details typed in the profile form (NOT saved): {username, password, secret, protocol, port,
        original}. An empty password / secret while editing falls back to the saved profile `original`."""
        hosts = self._validate_hosts(hosts)
        if adhoc is not None:
            cm, protocol, port = self._adhoc_creds(hosts, adhoc)
            for h in hosts:
                set_host_protocol(h, protocol, port)
            profile = None
        elif profile:
            with self.lock:
                rec = self.profiles.get(profile)
            if rec is None:
                raise ApiError(f"The credentials profile '{profile}' does not exist.")
            try:
                one = (rec["username"], unprotect_secret(rec["password"]), unprotect_secret(rec.get("secret", "")))
            except Exception:  # noqa: BLE001
                raise ApiError(f"The saved password of profile '{profile}' cannot be read "
                               "(was it saved by another Windows user?). Open the profile and "
                               "enter the password again.")
            cm = {h: one for h in hosts}
            for h in hosts:
                set_host_protocol(h, rec.get("protocol"), rec.get("port"))   # the chosen profile decides SSH or Telnet + port
        else:
            cm = self._creds_map(hosts)
        with self.lock:
            self._poll_all()
            for h in hosts:
                if self._upgrade_active(h):
                    raise ApiError(f"{h} is being upgraded right now - wait for it to finish.", 409)
                self.ops[h]["creds"] = {"state": "testing", "detail": "Connecting ...", "time": None}
        for h in hosts:
            self.pool.submit(self._test_credentials_one, h, *cm[h], profile or None,
                             "the details in the form (not saved)" if adhoc is not None else None)
        if adhoc is not None:
            return {"message": f"Testing the entered credentials ({'Telnet' if protocol == 'telnet' else 'SSH'}, port {port}) "
                               f"on {len(hosts)} switch(es) - nothing is saved ..."}
        return {"message": f"Testing profile '{profile}' on {len(hosts)} switch(es) (not assigned) ..." if profile
                else f"Testing credentials on {len(hosts)} switch(es) ..."}

    def _adhoc_creds(self, hosts, adhoc):
        """Credentials typed in the profile form -> ({host: (user, password, secret)}, protocol, port)."""
        if not isinstance(adhoc, dict):
            raise ApiError("Nothing to test.")
        username = adhoc.get("username").strip() if isinstance(adhoc.get("username"), str) else ""
        if not username or len(username) > 128:
            raise ApiError("Enter the switch username first (up to 128 characters).")
        protocol = "telnet" if str(adhoc.get("protocol") or "ssh").strip().lower() == "telnet" else "ssh"
        try:
            port = clean_port(adhoc.get("port"), protocol)
        except ValueError as exc:
            raise ApiError(str(exc))
        password, secret = adhoc.get("password"), adhoc.get("secret")
        for label, val in (("password", password), ("enable secret", secret)):
            if val is not None and (not isinstance(val, str) or len(val) > 256):
                raise ApiError(f"Invalid {label}.")
        password, secret = password or "", secret or ""
        original = adhoc.get("original")
        if not password or (not secret and not adhoc.get("clear_secret")):
            with self.lock:
                old = self.profiles.get(original) if isinstance(original, str) else None
            try:
                if not password and old:
                    password = unprotect_secret(old["password"])
                if not secret and old and not adhoc.get("clear_secret"):
                    secret = unprotect_secret(old.get("secret", ""))
            except Exception:  # noqa: BLE001
                raise ApiError("The saved password of this profile cannot be read - type the password again.")
        if not password:
            raise ApiError("Enter the password first.")
        one = (username, password, secret)
        return {h: one for h in hosts}, protocol, port

    @activity("Credentials test")
    def _test_credentials_one(self, host, username, password, secret, profile=None, source=None):
        dtype = BACKUP_DEVICE_TYPES.get(self.info[host]["type"], "cisco_xe")
        device = {"device_type": dtype, "host": host, "username": username, "password": password,
                  "secret": secret, "conn_timeout": 15, "auth_timeout": 15, "banner_timeout": 15}
        ok, detail = False, ""
        try:
            with ConnectHandler(**device) as conn:
                prompt = conn.find_prompt()
                if secret:
                    conn.enable()
                    prompt = conn.find_prompt()
                ok = True
                detail = "Login OK" + (" + enable OK" if secret else "") + f" ({prompt.strip()})"
                if profile:
                    detail += f" - tested with profile '{profile}'"
                elif source:
                    detail += f" - tested with {source}"
        except NetmikoAuthenticationException:
            detail = "Authentication failed (bad username/password or enable secret)" + (
                f" - tested with profile '{profile}'" if profile else (f" - tested with {source}" if source else ""))
        except NetmikoTimeoutException:
            detail = "Connection timed out"
        except Exception as exc:  # noqa: BLE001
            detail = f"{type(exc).__name__}: {exc}"

        with self.lock:
            if host not in self.ops:
                return  # the device was deleted while this test was still running
            self.ops[host]["creds"] = {"state": "ok" if ok else "failed", "detail": detail,
                                       "time": now_str()}
        self.events.add(host, self.label(host), f"Credentials test: {detail}", "ok" if ok else "bad")

    # ---- readiness / health check -------------------------------------------- #
    def start_readiness_check(self, hosts, image_path=None):
        """Read-only pre-upgrade check: login, current version, free flash space against the
        chosen image (if any), CPU, memory, and stack health. No upload, no reload."""
        hosts = self._validate_hosts(hosts)
        cm = self._creds_map(hosts)
        image_local = None
        if image_path:
            image_local, _ = self._check_image(image_path)
        with self.lock:
            self._poll_all()
            for h in hosts:
                if self._upgrade_active(h):
                    raise ApiError(f"{h} is being upgraded right now - wait for it to finish.", 409)
                self.ops[h]["readiness"] = {"state": "checking", "detail": "Connecting ...",
                                            "checks": {}, "time": None}
        for h in hosts:
            self.pool.submit(self._readiness_one, h, *cm[h], image_local)
        msg = f"Running readiness checks on {len(hosts)} switch(es)"
        msg += f" against {Path(image_local).name}" if image_local else \
               " (no image selected - skipping the flash-space check)"
        return {"message": msg + " ..."}

    @activity("Readiness check")
    def _readiness_one(self, host, username, password, secret, image_local):
        dtype = BACKUP_DEVICE_TYPES.get(self.info[host]["type"], "cisco_xe")
        checks = {}
        try:
            checks = probe_device_readiness(host, dtype, username, password, secret, image_local)
            state, detail = compute_readiness_verdict(checks)
        except NetmikoAuthenticationException:
            state, detail = "not_ready", "Authentication failed (bad username/password or enable secret)"
        except NetmikoTimeoutException:
            state, detail = "not_ready", "Connection timed out"
        except Exception as exc:  # noqa: BLE001
            state, detail = "not_ready", f"{type(exc).__name__}: {exc}"

        outputs = checks.pop("_outputs", None) if isinstance(checks, dict) else None
        transcript = checks.pop("_transcript", None) if isinstance(checks, dict) else None
        if (transcript or outputs) and self.runs_dir:
            try:                  # every readiness command in ONE file, next to the pre-/post-check files
                self.runs_dir.mkdir(parents=True, exist_ok=True)
                heading = f"Readiness / pre-check - {host} - {now_str()}"
                text = (command_sections_text(heading, transcript) if transcript
                        else snapshot_text(outputs, heading))
                (self.runs_dir / f"precheck_{safe_name(host)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt").write_text(
                    text + "\n", encoding="utf-8")
            except OSError:
                pass
        with self.lock:
            if host not in self.ops:
                return  # deleted while the check was running
            self.ops[host]["readiness"] = {"state": state, "detail": detail, "checks": checks,
                                           "time": now_str()}
            version = checks.get("version")
            if version and not self.ops[host]["version"]:
                self.ops[host]["version"] = version
                self.ops[host]["version_note"] = f"from readiness check {now_str()}"
            if checks.get("model"):
                self.ops[host]["model"] = checks["model"]
            self._set_boot(host, checks.get("boot"))
        if version:
            save_known_version(host, self.label(host), version, checks.get("model"))
        level = {"ready": "ok", "warning": "warn", "not_ready": "bad"}[state]
        self.events.add(host, self.label(host), f"Readiness check: {detail}", level)

    # ---- upgrade ----------------------------------------------------------- #
    def _port_of(self, host):
        with self.lock:
            rec = self.profiles.get(self.assign.get(host))
        return (rec or {}).get("port") or default_port((rec or {}).get("protocol"))

    def _protocol_of(self, host):
        with self.lock:
            rec = self.profiles.get(self.assign.get(host))
        return "telnet" if rec and rec.get("protocol") == "telnet" else "ssh"

    def _check_upgradable(self, hosts):
        """Common checks for starting an upgrade now or scheduling one. Called with self.lock held."""
        for h in hosts:
            rec = self.profiles.get(self.assign.get(h)) or {}
            if rec.get("protocol") == "telnet":
                raise ApiError(f"{h} uses the Telnet profile '{self.assign.get(h)}'. Upgrades copy the image with SCP, "
                               "which needs SSH - switch the device to an SSH profile first. (Backups, credential tests, "
                               "readiness checks, the console and topology discovery work over Telnet.)")
            if self.info[h]["type"] != "cisco-xe":
                raise ApiError(f"{h} is type '{self.info[h]['type']}' - only cisco-xe switches "
                               "can be upgraded.")
            if self._upgrade_active(h):
                raise ApiError(f"{h} is already upgrading right now.", 409)

    def start_upgrade(self, hosts, image_path, backup_first=False, action="full",
                      started_by=None, skip_image_check=False, ignore_space=False):
        """action: 'full' (upload+activate+reload, default) | 'distribute' (upload+add to flash
        only, no reload - stage it ahead of time) | 'activate' (activate an image already
        distributed to flash and reload - no image_path needed)."""
        self._require_site()
        hosts = self._validate_hosts(hosts)
        if not self.monitor.running:
            raise ApiError("Monitoring is stopped. Upgrades only run while monitoring is on - "
                           "click 'Start monitoring' first.", 409)
        cm = self._creds_map(hosts)
        if action not in ("full", "distribute", "activate"):
            raise ApiError("Invalid action.")
        if action == "activate":
            image, size = None, 0
        else:
            image, size = self._check_image(image_path)
        verb = self.ACTION_VERB.get(action, "Upgrade")
        with self.lock:
            self._poll_all()
            self._check_upgradable(hosts)
            run_hosts, queued_hosts = list(hosts), []
            if action in ("full", "activate"):           # these reload the switch: respect the parallel limit
                free = max(0, self.max_parallel - self._active_reload_count())
                run_hosts, queued_hosts = list(hosts[:free]), list(hosts[free:])
            if not backup_first:                 # a backup-first run reads the platform itself
                self._require_platform_known(hosts)
            if image:
                self._enforce_image_check(hosts, image, skip_image_check, started_by)
            if not backup_first:
                for h in hosts:
                    if self.ops[h]["backup"]["state"] in ("queued", "running"):
                        raise ApiError(f"A backup is still running for {h} - wait for it to finish.", 409)
            for h in hosts:
                self.schedule.pop(h, None)   # a manual start replaces any pending schedule
                self._rollout_forget(h)
            rid = self._new_rollout(hosts, started_by) if (action in ("full", "activate") and len(hosts) > 1) else None
            for h in queued_hosts:               # waits for a free slot; started by the scheduler loop
                self.schedule[h] = {"run_at": time.time(), "image_path": image, "created": now_str(),
                                    "backup_first": bool(backup_first), "action": action,
                                    "scheduled_by": started_by or "local", "queued": True, "rollout": rid,
                                    "ignore_space": bool(ignore_space)}
            self._save_schedule()
            queue_note = ""
            if queued_hosts:
                queue_note = (f" {len(queued_hosts)} more switch(es) are queued and start as running ones finish "
                              f"({self.max_parallel} at a time"
                              + ("; the queue stops if one fails)." if self.stop_on_failure else ")."))
            if not run_hosts:
                return {"message": f"All {self.max_parallel} upgrade slot(s) are busy - "
                                   f"{len(queued_hosts)} switch(es) queued; each starts when a slot frees up."
                                   + (" The queue stops if one fails." if self.stop_on_failure else "")}

            if backup_first:
                # start_backup() itself refuses a host it sees as "already upgrading", so the
                # backup must be kicked off BEFORE this placeholder state is set, not after.
                backup_result = self.start_backup(run_hosts)
                for h in run_hosts:
                    self.ops[h]["upgrade"] = {**self._new_ops()["upgrade"], "state": "starting",
                                              "phase": "waiting_backup", "started": now_str(),
                                              "started_ts": time.time(), "action": action,
                                              "started_by": started_by or "local",
                                              "message": "Backing up before the upgrade starts ..."}
                threading.Thread(target=self._launch_after_backup,
                                 args=(list(run_hosts), cm, image, action, started_by or "local", ignore_space),
                                 daemon=True).start()
                return {"message": f"Backing up {len(run_hosts)} switch(es) first ({backup_result['message']}) "
                                   f"- each switch's {verb.lower()} will start automatically once its "
                                   "backup finishes, and will be skipped if the backup fails." + queue_note,
                       }

            self.runs_dir.mkdir(parents=True, exist_ok=True)
            for h in run_hosts:
                self._launch_job(h, *cm[h], image, action, started_by, ignore_space)
        msg = f"{verb} started on {len(run_hosts)} switch(es)"
        msg += f" with {Path(image).name} ({size / 1_048_576:.0f} MB)." if image else "."
        msg += queue_note + " "
        msg += "A live web console opens at the bottom of this page for each switch."
        return {"message": msg}

    def _launch_after_backup(self, hosts, cm, image, action="full", started_by=None, ignore_space=False):
        """Background watcher for backup_first upgrades: launches each device's action the
        moment its own backup finishes, or marks it failed (skipped) if the backup failed."""
        pending = set(hosts)
        while pending and not self._stop.is_set():
            time.sleep(2)
            with self.lock:
                done_now = [h for h in pending
                           if h in self.ops and self.ops[h]["backup"]["state"] in ("completed", "failed")]
                for h in done_now:
                    pending.discard(h)
                    if self.ops[h]["backup"]["state"] == "completed":
                        self.runs_dir.mkdir(parents=True, exist_ok=True)
                        self._launch_job(h, *cm[h], image, action, started_by, ignore_space)
                    else:
                        self.ops[h]["upgrade"] = {**self._new_ops()["upgrade"], "started_by": started_by}
                        self._finish(h, "failed", "Upgrade skipped: the pre-upgrade backup failed")

    # ---- scheduled upgrade --------------------------------------------------- #
    def schedule_upgrade(self, hosts, image_path, run_at, backup_first=False, action="full",
                         started_by=None, skip_image_check=False, ignore_space=False):
        self._require_site()
        hosts = self._validate_hosts(hosts)
        self._creds_map(hosts)          # every device needs a profile now; read again when the run starts
        if action not in ("full", "distribute", "activate"):
            raise ApiError("Invalid action.")
        image = None
        if action != "activate":
            image, size = self._check_image(image_path)      # must exist now; re-checked at run time
        if not isinstance(run_at, (int, float)):
            raise ApiError("Invalid scheduled time.")
        if run_at < time.time() + MIN_SCHEDULE_LEAD:
            raise ApiError("The scheduled time must be at least "
                           f"{MIN_SCHEDULE_LEAD} seconds in the future.")
        with self.lock:
            self._poll_all()
            self._check_upgradable(hosts)
            if not backup_first:
                self._require_platform_known(hosts)
            if image:
                self._enforce_image_check(hosts, image, skip_image_check, started_by)
            rid = None
            if action in ("full", "activate") and len(hosts) > 1:
                rid = self._new_rollout(hosts, started_by)
            else:
                for h in hosts:
                    self._rollout_forget(h)
            for h in hosts:
                self.schedule[h] = {"run_at": run_at, "image_path": image, "created": now_str(),
                                    "backup_first": bool(backup_first), "action": action,
                                    "scheduled_by": started_by or "local", "queued": False, "rollout": rid,
                                    "ignore_space": bool(ignore_space)}
            self._save_schedule()
        when = datetime.fromtimestamp(run_at).strftime("%Y-%m-%d %H:%M")
        verb = self.ACTION_VERB.get(action, "Upgrade")
        for h in hosts:
            self.events.add(h, self.label(h), f"{verb} scheduled for {when}", "info")
        msg = f"{verb} scheduled for {len(hosts)} switch(es) at {when}."
        if not self.monitor.running:
            msg += " NOTE: monitoring is stopped - scheduled upgrades only run while monitoring is on."
        return {"message": msg}

    def cancel_schedule(self, hosts):
        hosts = self._validate_hosts(hosts)
        with self.lock:
            cancelled = [h for h in hosts if self.schedule.pop(h, None) is not None]
            for h in cancelled:
                self._rollout_forget(h)
            self._save_schedule()
        for h in cancelled:
            self.events.add(h, self.label(h), "Scheduled upgrade cancelled", "info")
        if not cancelled:
            raise ApiError("None of the selected switches had a pending scheduled upgrade.")
        return {"message": f"Cancelled the scheduled upgrade for {len(cancelled)} switch(es)."}

    def _scheduler_loop(self):
        while not self._stop.is_set():
            due = []
            with self.lock:
                self._poll_all()
                self._check_rollouts()
                now = time.time()
                mon_on = bool(self.monitor.running)       # nothing is started while monitoring is stopped
                for h, entry in list(self.schedule.items()):
                    if entry["run_at"] <= now and mon_on:
                        due.append((h, entry))
                for h, entry in due:
                    if self._upgrade_active(h) or self.ops[h]["backup"]["state"] in ("queued", "running"):
                        # busy right this second (e.g. a manual backup just started) - try again shortly
                        entry["run_at"] = now + SCHEDULE_POLL
                        continue
                    action = entry.get("action") or "full"
                    if action in ("full", "activate") and self._active_reload_count() >= self.max_parallel:
                        entry["run_at"] = now + SCHEDULE_POLL
                        continue
                    del self.schedule[h]
                    self._save_schedule()
                    by = f"{entry.get('scheduled_by') or 'local'} (scheduled)"
                    if entry.get("image_path") and not Path(entry["image_path"]).is_file():
                        self.ops[h]["upgrade"] = self._new_ops()["upgrade"]
                        self._finish(h, "failed", "Scheduled run failed: the image file is no "
                                     f"longer at {entry['image_path']}")
                        entry.clear()
                        continue
                    try:
                        cm = self._creds_map([h])     # the profile as it is NOW (it may have been edited)
                    except ApiError as exc:
                        self.ops[h]["upgrade"] = self._new_ops()["upgrade"]
                        self._finish(h, "failed", f"Scheduled run failed: {exc}")
                        entry.clear()
                        continue
                    if entry.get("backup_first"):
                        try:
                            self.start_backup([h])
                        except ApiError as exc:
                            self._finish(h, "failed", f"Scheduled run: could not start the "
                                         f"pre-upgrade backup ({exc})")
                            entry.clear()
                            continue
                        self.ops[h]["upgrade"] = {**self._new_ops()["upgrade"], "state": "starting",
                                                  "phase": "waiting_backup", "started": now_str(),
                                                  "started_ts": time.time(), "action": action,
                                                  "started_by": by,
                                                  "message": "Backing up before the upgrade starts ..."}
                        threading.Thread(target=self._launch_after_backup,
                                         args=([h], cm, entry["image_path"], action, by, bool(entry.get("ignore_space"))),
                                         daemon=True).start()
                    else:
                        self.runs_dir.mkdir(parents=True, exist_ok=True)
                        self._launch_job(h, *cm[h], entry["image_path"], action, by, bool(entry.get("ignore_space")))
                    entry.clear()
            self._stop.wait(SCHEDULE_POLL)

    ACTION_VERB = {"full": "Firmware upgrade", "distribute": "Image distribution",
                  "activate": "Image activation"}

    def _launch_job(self, host, username, password, secret, image, action="full", started_by=None, ignore_space=False):
        """Start the upgrade worker for one switch (output goes to its web console).
        Called with self.lock held.
        action: 'full' (upload+activate+reload, default) | 'distribute' (upload+add only,
        no reload) | 'activate' (activate an already-distributed image, no image needed)."""
        self.term_close_host(host)                  # no console may stay open on a switch that is about to reload
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = self.runs_dir / f"upgrade_{safe_name(host)}_{stamp}"
        job = {"id": stamp, "status_path": f"{base}.status.json", "log_path": f"{base}.log",
               "job_path": os.path.join(self.job_tmp, f"{safe_name(host)}_{stamp}.job.json"),
               "console_log": f"{base}.console.txt", "beat": time.time(), "proc": None, "direct": False}
        verb = self.ACTION_VERB.get(action, "Firmware upgrade")
        self.ops[host]["postcheck"] = {"state": "idle", "detail": "", "checks": {}, "time": None}
        up = self.ops[host]["upgrade"] = {
            "state": "starting", "phase": "starting", "message": f"Starting {verb.lower()} ...",
            "old_version": None, "new_version": None, "started": now_str(), "finished": None,
            "started_ts": time.time(), "finished_ts": None,
            "image": Path(image).name if image else None, "action": action,
            "started_by": started_by or "local", "phase_times": []}
        self._notify(host, "started", f"{verb} started",
                     (f"with {Path(image).name}" if image else "") or "Worker launched")
        payload = {"ip": host, "name": self.label(host) or host, "action": action,
                   "username": username, "password": password, "secret": secret,
                   "status_path": job["status_path"], "log_path": job["log_path"],
                   "timeouts": dict(TIMEOUTS)}
        _rec = self.profiles.get(self.assign.get(host)) or {}
        if _rec.get("port") and _rec.get("port") != default_port("ssh"):
            payload["port"] = int(_rec["port"])         # custom SSH port: the worker (and its SCP copy) must use it too
        if image:
            payload["image_path"] = image
        if ignore_space:
            payload["ignore_space"] = True
            self.events.add(host, self.label(host), "Flash space check will be ignored for this run "
                            f"(chosen by {started_by or 'local'})", "warn")
        try:
            fd = os.open(job["job_path"], os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                json.dump(payload, fh)
            argv = self_invoke_argv("--job", job["job_path"])
            proc = launch_background_worker(argv, job["console_log"])
            job.update(proc=proc, direct=True)
        except Exception as exc:  # noqa: BLE001
            self.jobs[host] = job
            self._finish(host, "failed", f"Could not start the {verb.lower()} worker: {type(exc).__name__}: {exc}")
            return
        finally:
            payload.clear()
        self.jobs[host] = job
        detail = f" ({Path(image).name})" if image else ""
        self.events.add(host, self.label(host), f"{verb} started{detail}", "info")
        up["message"] = f"{verb} started - see the live console ..."

    CONSOLE_CHUNK = 65536          # most bytes returned to the browser per poll

    def console_read(self, host, job_id, offset):
        """Web console feed: the text a switch's upgrade worker has printed since byte `offset`.
        The browser polls this with the offset it was last given. If `job_id` is not the
        switch's current run (a new upgrade started), reading restarts from the beginning and
        `reset` tells the browser to clear its console first."""
        with self.lock:
            if host not in self.info:
                raise ApiError("Unknown switch.", 404)
            job = self.jobs.get(host)
            if job is None:
                return {"job_id": None, "offset": 0, "text": "", "reset": False, "ended": True}
            path, current, proc = job["console_log"], job["id"], job.get("proc")
        try:
            offset = max(0, int(offset))
        except (TypeError, ValueError):
            offset = 0
        reset = job_id != current
        if reset:
            offset = 0
        ended = proc is not None and proc.poll() is not None
        try:
            with open(path, "rb") as fh:
                fh.seek(offset)
                chunk = fh.read(self.CONSOLE_CHUNK)
        except OSError:
            chunk = b""
        # Only hand out whole lines (so a multi-byte character is never split between two
        # polls); a trailing partial line waits for its newline unless the worker has ended.
        if ended:
            pass                                       # worker is done: hand out everything left
        elif b"\n" in chunk:
            chunk = chunk[:chunk.rindex(b"\n") + 1]
        elif len(chunk) < self.CONSOLE_CHUNK:
            chunk = b""                                # a partial line: wait for its newline
        text = chunk.decode("utf-8", "replace").replace("\r\n", "\n").replace("\r", "\n")
        return {"job_id": current, "offset": offset + len(chunk), "text": text, "reset": reset,
                "ended": ended}

    def _poll_all(self):
        for h in self.order:
            self._poll_job(h)

    def _poll_job(self, host):
        """Refresh one upgrade from its status file. Called with self.lock held."""
        job, up = self.jobs.get(host), self.ops[host]["upgrade"]
        if job is None or up["state"] in TERMINAL:
            return
        data = read_json_file(job["status_path"])
        if data:
            self._apply_status(host, job, data)
        if up["state"] in TERMINAL:
            return
        proc = job.get("proc")
        if job["direct"] and proc is not None and proc.poll() is not None:
            data = read_json_file(job["status_path"])  # final state may have been written just now
            if data:
                self._apply_status(host, job, data)
            if up["state"] not in TERMINAL:
                self._finish(host, "attention", "The upgrade process ended "
                             f"unexpectedly (exit code {proc.returncode}). Check the switch manually.")
        elif time.time() - job["beat"] > STALE_SECONDS:
            self._finish(host, "attention", "Lost contact with the upgrade process. "
                         "Check the switch manually.")

    def _apply_status(self, host, job, data):
        up, o = self.ops[host]["upgrade"], self.ops[host]
        job["beat"] = data.get("updated") or job["beat"]
        up["phase"] = data.get("phase") or up["phase"]
        up["message"] = data.get("message") or up["message"]
        if isinstance(data.get("phase_times"), list):      # how long each phase took (Command time test uses it)
            up["phase_times"] = [{"phase": str(p.get("phase") or "")[:30], "seconds": float(p.get("seconds") or 0)}
                                 for p in data["phase_times"][:60]
                                 if isinstance(p, dict) and isinstance(p.get("seconds"), (int, float))]
        if data.get("old_version"):
            up["old_version"] = data["old_version"]
            if not o["version"]:
                o["version"], o["version_note"] = data["old_version"], "read before upgrade"
        if data.get("new_version"):
            up["new_version"] = data["new_version"]
        if data.get("model"):
            o["model"] = data["model"]
        if isinstance(data.get("boot"), dict):
            self._set_boot(host, data["boot"])           # before / after the boot variable fix / after reload
        pc = data.get("postcheck")
        if isinstance(pc, dict) and pc.get("state") in ("passed", "warning", "failed"):
            if o["postcheck"].get("time") != pc.get("time") or o["postcheck"].get("state") != pc["state"]:
                o["postcheck"] = {"state": pc["state"], "detail": str(pc.get("detail") or ""),
                                  "checks": pc.get("checks") if isinstance(pc.get("checks"), dict) else {},
                                  "time": pc.get("time")}
                title = {"passed": "Post-check passed", "warning": "Post-check found differences",
                         "failed": "Post-check failed"}[pc["state"]]
                self._notify(host, {"passed": "completed", "warning": "attention", "failed": "failed"}[pc["state"]],
                             title, o["postcheck"]["detail"], category="postcheck")
                self.events.add(host, self.label(host), f"{title}: {o['postcheck']['detail']}",
                                {"passed": "ok", "warning": "warn", "failed": "bad"}[pc["state"]])
        elif data.get("phase") == "postcheck" and up["state"] in ACTIVE and o["postcheck"]["state"] in ("idle", "checking"):
            o["postcheck"] = {"state": "checking", "detail": data.get("message") or "", "checks": {}, "time": None}
        state = data.get("state")
        if state == "running" and up["state"] in ACTIVE:
            up["state"] = "running"
        elif state in TERMINAL:
            self._finish(host, state, up["message"])

    def _notes_view(self):
        with self.notes_lock:
            return list(reversed(self.notes[-100:]))          # newest first

    def _notify(self, host, kind, title, text, category=None):
        """Store a notification (shown under the bell icon) and save it with the site, so it is
        still there after a restart. Never raises - notifications must not break an upgrade."""
        try:
            with self.notes_lock:
                nid = max(int(time.time() * 1000), (self.notes[-1]["id"] + 1) if self.notes else 0)   # always rising
                self.notes.append({"id": nid, "time": now_str(), "host": host,
                                   "label": (self.info.get(host) or {}).get("label") or "",
                                   "kind": kind, "title": title, "text": text})
                del self.notes[:-NOTES_MAX]
                if self.site_name:
                    try:
                        site_notes_path(self.site_name).write_text(json.dumps(self.notes), encoding="utf-8")
                    except OSError:
                        pass
        except Exception:  # noqa: BLE001
            pass
        EMAILER.submit(self.site_name, host, (self.info.get(host) or {}).get("label") or "",
                       kind, category, title, text)      # optional e-mail copy; returns at once

    def clear_notifications(self):
        with self.notes_lock:
            self.notes = []
            if self.site_name:
                try:
                    site_notes_path(self.site_name).write_text("[]", encoding="utf-8")
                except OSError:
                    pass
        return {"message": "Notifications cleared."}

    def _finish(self, host, state, message):
        """Mark an upgrade finished and release its slot. Called with self.lock held."""
        up, o = self.ops[host]["upgrade"], self.ops[host]
        if up["state"] in TERMINAL:
            return
        up.update(state=state, message=message, finished=now_str(), finished_ts=time.time())
        new, old = up.get("new_version"), up.get("old_version")
        if new:
            o["prev_version"] = old if old and old != new else None
            o["version"], o["version_note"] = new, f"after upgrade {now_str()}"
            save_known_version(host, self.label(host), new, o.get("model"))
        job = self.jobs.get(host)
        if job:
            try:
                os.remove(job["job_path"])  # leftover credential file, if the worker never read it
            except OSError:
                pass
        level = {"completed": "ok", "failed": "bad", "attention": "warn"}[state]
        self.events.add(host, self.label(host), message, level)
        action = up.get("action")
        title = ({"distribute": "Image distributed", "activate": "Image activated - upgrade completed"}.get(action, "Upgrade completed")
                 if state == "completed" else "Upgrade failed" if state == "failed" else "Upgrade needs attention")
        self._notify(host, state, title, message)
        self._record_history(host)

    # ---- views ------------------------------------------------------------- #
    def _stage_map(self):
        """compute_image_stage() for the current site, re-read only when the history file changes."""
        if not self.site_name:
            return {}
        try:
            stamp = os.path.getmtime(site_history_path(self.site_name))
        except OSError:
            stamp = None
        fresh = FRESH_START
        key = (self.site_name, stamp, fresh)
        if getattr(self, "_stage_key", None) != key:
            history = load_history(self.site_name)
            if fresh:                       # only what this tool did since it was launched
                history = [r for r in history if str(r.get("finished") or "") >= self.session_started]
            self._stage_cache = compute_image_stage(history)
            self._stage_key = key
        return self._stage_cache

    def snapshot(self):
        mon = self.monitor.snapshot()
        if not self.monitor.running:
            # Up/down is only monitored while monitoring is enabled: otherwise every device reads as DOWN.
            mon = dict(mon)
            mon["hosts"] = [{**h, "status": "DOWN"} for h in mon["hosts"]]
            mon["summary"] = {**mon["summary"], "up": 0, "down": len(mon["hosts"]), "pending": 0}
        by_host = {h["host"]: h for h in mon["hosts"]}
        with self.lock:
            self._poll_all()
            stage_map = self._stage_map()
            hosts = []
            for host in self.order:
                m = dict(by_host[host])
                o = self.ops[host]
                sched = self.schedule.get(host)
                m.update(profile=self.assign.get(host), label=self.info[host]["label"], type=self.info[host]["type"],
                         platform=o.get("model") or self.info[host].get("platform", ""),
                         boot=dict(o.get("boot") or {}),
                         version=o["version"], prev_version=o["prev_version"],
                         version_note=o["version_note"], backup=dict(o["backup"]),
                         creds=dict(o["creds"]), readiness=dict(o["readiness"]), postcheck=dict(o["postcheck"]),
                         stage=stage_map.get(host) or {"staged": None, "upgraded": None},
                         upgrade={**o["upgrade"], "job_id": (self.jobs.get(host) or {}).get("id")},
                         scheduled=None if sched is None else {
                             "run_at": sched["run_at"],
                             "run_at_str": datetime.fromtimestamp(sched["run_at"]).strftime("%Y-%m-%d %H:%M"),
                             "queued": bool(sched.get("queued"))})
                hosts.append(m)
            active = sum(1 for h in self.order if self._upgrade_active(h))
            scheduled_count = len(self.schedule)
            completed_count = sum(1 for h in self.order
                                  if self.ops[h]["upgrade"]["state"] == "completed"
                                  and self.ops[h]["upgrade"]["action"] != "distribute")
            run = self.backup_run
            run_info = None if run is None else {
                "total": run["total"], "done": run["done"], "ok": run["ok"], "failed": run["failed"],
                "running": run["done"] < run["total"], "started": run["started"], "finished": run["finished"],
                "commands": run["commands"][-200:]}
        return {"app_version": APP_VERSION, "site": self.site_name, "monitor": mon["monitor"], "summary": mon["summary"], "hosts": hosts,
                "profiles": self.profile_list(),
                "keep_awake": self.keep_awake.active, "monitoring_on": bool(self.monitor.running),
                "notifications": self._notes_view(),
                "events": self.events.items(), "upgrade": {"active": active, "scheduled": scheduled_count,
                            "completed": completed_count, "settings": self.upgrade_settings_view()},
                "backup_run": run_info}

    # ---- Diagram tab: CDP topology ------------------------------------------- #
    def _topology_load(self):
        data = read_json_file(site_topology_path(self.site_name)) if self.site_name else None
        data = data if isinstance(data, dict) else {}
        devices = data.get("devices") if isinstance(data.get("devices"), dict) else {}
        roots = [r for r in (data.get("roots") or []) if isinstance(r, str)]
        return {"devices": devices, "roots": roots}

    def _topology_save(self, store):
        try:
            site_dir(self.site_name).mkdir(parents=True, exist_ok=True)
            tmp = site_topology_path(self.site_name).with_suffix(".json.tmp")
            tmp.write_text(json.dumps(store, indent=1), encoding="utf-8")
            tmp.replace(site_topology_path(self.site_name))
        except OSError as exc:
            print(f"Could not save the topology for site '{self.site_name}': {exc}", flush=True)

    def topology_view(self):
        """Everything the Diagram tab draws: the saved CDP graph plus one row per device for the picker."""
        with self.lock:
            site = self.site_name
            order = list(self.order)
            info = {h: dict(self.info[h]) for h in order}
            cred_ok = {h: self.assign.get(h) in self.profiles for h in order}
            profile = {h: self.assign.get(h) for h in order}
        with self.topo_lock:
            store = self._topology_load()
            run = self.topo_run if (self.topo_run and self.topo_run.get("site") == site) else None
            run_copy = None if run is None else {
                "started": run["started"], "finished": run["finished"], "total": run["total"], "done": run["done"],
                "hosts": {h: dict(v) for h, v in run["hosts"].items()}}
        graph = build_topology_graph(info, order, store["devices"], set(store["roots"]))
        devices = []
        for h in order:
            scan = store["devices"].get(h) if isinstance(store["devices"].get(h), dict) else {}
            devices.append({"host": h, "label": info[h].get("label") or h, "platform": info[h].get("platform") or "",
                            "has_creds": cred_ok[h], "profile": profile[h] if cred_ok[h] else None,
                            "scan_ok": scan.get("ok"), "scan_time": scan.get("time") or "",
                            "scan_error": scan.get("error") or "", "neighbors": len(scan.get("neighbors") or [])})
        return {"site": site, "devices": devices, "graph": graph, "run": run_copy}

    def start_topology_discovery(self, hosts):
        """Read 'show cdp neighbors detail' from the selected switches (each must have a credentials profile)."""
        self._require_site()
        hosts = self._validate_hosts(hosts)
        with self.lock:
            missing = [h for h in hosts if self.assign.get(h) not in self.profiles]
        if missing:
            shown = ", ".join(self.label(h) for h in missing[:5]) + (f" (+{len(missing) - 5} more)" if len(missing) > 5 else "")
            raise ApiError(f"No credentials profile is attached to: {shown}. Assign a profile on the Home page "
                           "first - the topology is read from the switches with it.")
        cm = self._creds_map(hosts)
        with self.lock:
            self._poll_all()
            busy = {h for h in hosts if self._upgrade_active(h)}
        with self.topo_lock:
            if self.topo_run and self.topo_run.get("site") == self.site_name and not self.topo_run["finished"]:
                raise ApiError("A topology scan is already running - wait for it to finish.", 409)
            # The diagram shows only the switches selected for this discovery: drop stored scans of the rest.
            store = self._topology_load()
            keep = set(hosts)
            stale = [h for h in list(store["devices"]) if h not in keep]
            if stale:
                for h in stale:
                    store["devices"].pop(h, None)
                self._topology_save(store)
            run = {"site": self.site_name, "started": now_str(), "finished": None, "total": len(hosts), "done": 0,
                   "hosts": {h: {"state": "queued", "detail": ""} for h in hosts}}
            for h in busy:
                run["hosts"][h] = {"state": "skipped", "detail": "Upgrade in progress - kept the previous result"}
                run["done"] += 1
            if run["done"] >= run["total"]:
                run["finished"] = now_str()
            self.topo_run = run
        todo = [h for h in hosts if h not in busy]
        for h in todo:
            self.pool.submit(self._topology_one, h, self.site_name, *cm[h])
        self.events.add("-", "", f"Topology scan started on {len(todo)} device(s)", "info")
        msg = f"Reading CDP neighbours from {len(todo)} switch(es) ..."
        if busy:
            msg += f" {len(busy)} skipped because an upgrade is running on them."
        return {"message": msg}

    @activity("Topology scan")
    def _topology_one(self, host, site, username, password, secret):
        with self.topo_lock:
            if self.topo_run and self.topo_run.get("site") == site and host in self.topo_run["hosts"]:
                self.topo_run["hosts"][host] = {"state": "running", "detail": "Connecting ..."}
        info = self.info.get(host) or {}
        device = {"device_type": BACKUP_DEVICE_TYPES.get(info.get("type"), "cisco_xe"), "host": host,
                  "username": username, "password": password, "secret": secret,
                  "conn_timeout": 15, "auth_timeout": 15, "banner_timeout": 15}
        ok, detail, neighbors, hostname = False, "", [], ""
        try:
            with ConnectHandler(**device) as conn:
                if secret:
                    conn.enable()
                hostname = re.sub(r"\(.*?\)|[#>\s]+$", "", conn.find_prompt()).strip()
                raw = conn.send_command("show cdp neighbors detail", read_timeout=tmo('show_cmd'))
            if re.search(r"CDP is not enabled|Invalid input", raw or "", re.I):
                detail = "CDP is not enabled on this switch (cdp run)"
            else:
                neighbors = parse_cdp_neighbors_detail(raw)
                ok, detail = True, f"{len(neighbors)} CDP neighbour(s)"
        except NetmikoAuthenticationException:
            detail = "Authentication failed"
        except NetmikoTimeoutException:
            detail = "Connection timed out"
        except Exception as exc:  # noqa: BLE001
            detail = f"{type(exc).__name__}: {exc}"[:160]
        with self.topo_lock:
            if self.site_name == site:                       # the site was switched meanwhile: drop the result
                store = self._topology_load()
                old = store["devices"].get(host) if isinstance(store["devices"].get(host), dict) else {}
                if ok:
                    store["devices"][host] = {"hostname": hostname, "time": now_str(), "ok": True, "error": "",
                                              "neighbors": neighbors}
                else:                                        # keep what we knew; only note why the scan failed
                    store["devices"][host] = {"hostname": old.get("hostname", ""), "time": old.get("time", ""),
                                              "ok": False, "error": detail, "neighbors": old.get("neighbors") or []}
                self._topology_save(store)
            run = self.topo_run
            if run and run.get("site") == site and host in run["hosts"]:
                run["hosts"][host] = {"state": "ok" if ok else "failed", "detail": detail}
                run["done"] += 1
                if run["done"] >= run["total"] and not run["finished"]:
                    run["finished"] = now_str()
        self.events.add(host, self.label(host) if host in self.info else host, f"Topology scan: {detail}",
                        "ok" if ok else "bad")

    def topology_set_root(self, host, on):
        """Choose which switch is the top of its tree (the distribution / core), overriding the automatic guess."""
        self._require_site()
        if not isinstance(host, str) or host not in self.info:
            raise ApiError("Unknown switch.")
        with self.topo_lock:
            store = self._topology_load()
            with self.lock:
                order, info = list(self.order), {h: dict(self.info[h]) for h in self.order}
            graph = build_topology_graph(info, order, store["devices"], set())
            comp = {n["id"]: n["comp"] for n in graph["nodes"]}
            if host not in comp:
                raise ApiError("Scan this switch first - it is not on the diagram yet.")
            store["roots"] = [r for r in store["roots"] if comp.get(r) != comp[host] and r != host]
            if on:
                store["roots"].append(host)
            self._topology_save(store)
        return {"message": f"{self.label(host)} is now the top of its tree." if on else "Top of tree cleared - back to automatic."}

    # ---- web terminal (device drawer > Console) ------------------------------ #
    def _term_reap(self):
        """Drop finished sessions and close the ones nobody has used for a while. Called with term_lock held."""
        now = time.time()
        for sid, s in list(self.terms.items()):
            if not s.closed and now - s.last_used > TERM_IDLE_SECONDS:
                s.close()
            if s.closed and now - s.last_used > 60:
                del self.terms[sid]

    def _term(self, host, sid):
        s = self.terms.get(str(sid or ""))
        if s is None or s.host != str(host or ""):
            raise ApiError("That console session no longer exists - click Connect.", 404)
        return s

    def term_open(self, host, user):
        host = str(host or "")
        self._require_site()
        with self.lock:
            if host not in self.info:
                raise ApiError("Unknown switch.", 404)
            self._poll_job(host)
            if self._upgrade_active(host):
                raise ApiError(f"{host} is being upgraded right now - the console is closed until it finishes.", 409)
            username, password, secret = self._creds_map([host])[host]
            label = self.info[host]["label"]
        with self.term_lock:
            self._term_reap()
            for s in self.terms.values():                       # one live console per switch: reconnect to it
                if s.host == host and not s.closed:
                    s.last_used = time.time()
                    return {"sid": s.id, "offset": 0, "reused": True}
            if sum(1 for s in self.terms.values() if not s.closed) >= TERM_MAX_SESSIONS:
                raise ApiError(f"Too many consoles are open ({TERM_MAX_SESSIONS}) - close one first.", 429)
            s = TerminalSession(host, label, username, password, secret, user, self._protocol_of(host), self._port_of(host))
            self.terms[s.id] = s
        self.events.add(host, label, f"Web console opened by {user}", "info")
        return {"sid": s.id, "offset": 0, "reused": False}

    def term_send(self, host, sid, line=None, raw=None, secret=False, user="local"):
        s = self._term(host, sid)
        if raw is not None:
            keys = {"ctrl_c": "\x03", "ctrl_z": "\x1a"}
            if raw not in keys:
                raise ApiError("Unknown key.")
            s.send(keys[raw])
            return {"ok": True}
        line = str(line if line is not None else "")
        if len(line) > TERM_MAX_LINE or any(ord(c) < 32 and c != "\t" for c in line):
            raise ApiError("That line is too long or contains control characters.")
        if not secret and line.strip():
            with activity("Web console"):
                record_command(s.host, line.strip(), True, 0)          # the command log shows what was typed, never passwords
        s.send(line if line.endswith("?") else line + "\n")             # '?' shows help at once, without Enter
        return {"ok": True}

    def term_read(self, host, sid, offset):
        s = self._term(host, sid)
        try:
            offset = max(0, int(offset))
        except (TypeError, ValueError):
            offset = 0
        s.last_used = time.time()
        return s.read(offset)

    def term_close(self, host, sid, user="local"):
        s = self._term(host, sid)
        s.close()
        self.events.add(s.host, s.label, f"Web console closed ({user})", "info")
        return {"ok": True}

    def term_close_host(self, host):
        """Close the consoles of a switch (an upgrade is starting, or it was deleted). Never blocks on self.lock."""
        with self.term_lock:
            for s in self.terms.values():
                if s.host == host and not s.closed:
                    s.close()

    def term_close_all(self):
        with self.term_lock:
            for s in self.terms.values():
                s.close()

    # ---- device drawer ------------------------------------------------------ #
    def _device_dirs(self, host):
        """(backup folder, runs folder) of a device; refuses unknown devices / no active site."""
        if host not in self.info:
            raise ApiError("Unknown switch.", 404)
        if not self.site_name or self.backup_dir is None or self.runs_dir is None:
            raise ApiError("No site is selected.", 409)
        return self.backup_dir / device_folder_name(self.info[host]["label"], host), self.runs_dir

    def device_detail(self, host, image_name="", shared_target=False):
        """Everything for the device drawer (see the comment above read_device_facts)."""
        host = str(host or "")
        image_name = os.path.basename(str(image_name or "").replace("\\", "/"))[:200]
        with self.lock:
            device_dir, runs_dir = self._device_dirs(host)
            self._poll_job(host)
            ops = json.loads(json.dumps(self.ops[host], default=str))     # a private copy to work on
            if image_name:
                self.target_image = image_name          # remembered so a Viewer (no image picker) sees it too
            elif shared_target:
                # A Viewer has no image picker: show the image an Operator/Administrator picked, else the
                # one scheduled for / used by the latest upgrade of this switch.
                sched_img = os.path.basename(str((self.schedule.get(host) or {}).get("image_path") or ""))
                up_img = os.path.basename(str((ops.get("upgrade") or {}).get("image") or ""))
                image_name = self.target_image or sched_img or up_img
            info = dict(self.info[host])
            profile = self.assign.get(host)
            sched = self.schedule.get(host)
            scheduled = None if sched is None else datetime.fromtimestamp(sched["run_at"]).strftime("%Y-%m-%d %H:%M")
            job_id = (self.jobs.get(host) or {}).get("id")
            site = self.site_name
        with self.monitor.lock:
            s = self.monitor.stats.get(host) or {}
            mon = {"status": s.get("status") or "PENDING", "last_checked": s.get("last_checked"),
                   "last_change": s.get("last_change"), "last_rtt": s.get("last_rtt"),
                   "checks": s.get("checks"), "successes": s.get("successes")}
        if not self.monitor.running:                  # up/down is only known while monitoring is on
            mon["status"], mon["monitoring_off"] = "DOWN", True

        facts = read_device_facts(device_dir)
        backups = list_device_backups(device_dir)
        run_files = list_run_files(runs_dir, host)
        model = ops.get("model") or facts.get("members", [{}])[0].get("model") or info.get("platform") or None
        version = ops.get("version")

        target = None
        if image_name:
            target = {"image": image_name, "version": guess_version_from_filename(image_name),
                      "problems": check_image_for_device(image_name, model, version)}

        pre = None
        status = latest_upgrade_status(runs_dir, host)
        if status and isinstance(status.get("pre_snapshot"), dict):
            try:
                pre = {"summary": summarize_snapshot(status["pre_snapshot"]), "time": _stamp_to_time(status["_stamp"])}
            except (KeyError, TypeError):
                pre = None

        checklist = self._device_checklist(mon, profile, ops, model, version, backups, target)
        missing = sum(1 for c in checklist if c["state"] in ("todo", "bad"))
        warnings = sum(1 for c in checklist if c["state"] == "warn")

        events = [e for e in self.events.items(300) if e.get("host") == host][:80]
        history = [r for r in reversed(load_history(site)) if r.get("ip_address") == host][:15]
        return {
            "host": host, "label": info.get("label") or "", "type": info.get("type"),
            "model": model, "version": version, "profile": profile, "scheduled": scheduled, "job_id": job_id,
            "monitor": mon, "facts": facts, "ops": ops, "target": target, "pre_snapshot": pre,
            "checklist": checklist, "missing": missing, "warnings": warnings,
            "backups": {"folder": device_dir.name, "runs": backups}, "check_files": run_files,
            "commands": read_command_log(runs_dir, host), "events": events, "history": history,
        }

    @staticmethod
    def _device_checklist(mon, profile, ops, model, version, backups, target):
        """The steps that should be done before an upgrade, each ok / warn / bad / todo / run, with a reason."""
        items = []

        def add(key, title, state, detail):
            items.append({"key": key, "title": title, "state": state, "detail": detail})

        st = mon.get("status")
        if mon.get("monitoring_off"):
            add("reach", "Reachable", "bad", "Monitoring is stopped - the switch is shown as down until monitoring is started")
        elif st == "UP":
            rtt = mon.get("last_rtt")
            add("reach", "Reachable", "ok", "Answers ping" + (f" ({rtt:.1f} ms)" if isinstance(rtt, (int, float)) else ""))
        elif st == "DOWN":
            add("reach", "Reachable", "bad", "Not answering ping - check the switch and the network path")
        else:
            add("reach", "Reachable", "todo", "Not checked yet - start monitoring")

        creds = ops.get("creds") or {}
        if not profile:
            add("profile", "Credential profile attached", "bad", "No profile is attached - nothing can be read from or done on this switch")
        else:
            add("profile", "Credential profile attached", "ok", f"Profile '{profile}'")
        cs = creds.get("state")
        if cs == "ok":
            add("creds", "Credentials tested", "ok", creds.get("detail") or "Login OK")
        elif cs == "failed":
            add("creds", "Credentials tested", "bad", creds.get("detail") or "Login failed")
        elif cs == "testing":
            add("creds", "Credentials tested", "run", "Testing ...")
        else:
            add("creds", "Credentials tested", "todo", "Not tested yet - use 'Test credentials'")

        if model and version:
            add("ident", "Platform and software version known", "ok", f"{model}, running {version}")
        else:
            add("ident", "Platform and software version known", "todo", "Run a backup or 'Refresh version' to read them from the switch")

        bk, newest = ops.get("backup") or {}, (backups[0] if backups else None)
        age_h = None
        if newest:
            try:
                age_h = (datetime.now() - datetime.strptime(newest["stamp"], "%Y%m%d_%H%M%S")).total_seconds() / 3600
            except ValueError:
                pass
        if bk.get("state") in ("queued", "running"):
            add("backup", "Configuration backup", "run", bk.get("detail") or "Backing up ...")
        elif bk.get("state") == "failed":
            add("backup", "Configuration backup", "bad", bk.get("detail") or "The last backup failed")
        elif newest:
            when = newest["time"]
            if age_h is not None and age_h > 24:
                add("backup", "Configuration backup", "warn", f"Last backup {when} is {int(age_h // 24)} day(s) old - take a fresh one before upgrading")
            else:
                add("backup", "Configuration backup", "ok", f"Last backup {when} ({len(newest['files'])} file(s))")
        else:
            add("backup", "Configuration backup", "todo", "No backup yet - back up the configuration first")

        rd = ops.get("readiness") or {}
        rs, ch = rd.get("state"), rd.get("checks") or {}
        rmap = {"ready": "ok", "warning": "warn", "not_ready": "bad", "checking": "run"}
        if rs in rmap:
            add("ready", "Readiness check", rmap[rs], rd.get("detail") or "")
        else:
            add("ready", "Readiness check", "todo", "Not run yet - use 'Readiness check'")

        boot = ops.get("boot") or {}
        if boot.get("mode"):
            if boot.get("mode") == "INSTALL" and boot.get("boot_ok") is True:
                add("boot", "Boot mode", "ok", boot.get("detail") or "INSTALL mode, boot variable correct")
            else:
                add("boot", "Boot mode", "warn", (boot.get("detail") or f"{boot['mode']} mode") + " - the upgrade converts / fixes this")
        else:
            add("boot", "Boot mode", "todo", "Not read yet - run a backup, readiness check or 'Refresh version'")

        lvl = ch.get("space_level")
        if lvl == "ok":
            add("space", "Flash space", "ok", ch.get("space_detail") or "Enough free flash")
        elif lvl == "low":
            add("space", "Flash space", "warn", ch.get("space_detail") or "Flash is tight")
        elif lvl == "fail" or ch.get("space_ok") is False:
            add("space", "Flash space", "bad", ch.get("space_detail") or "Not enough free flash")
        else:
            add("space", "Flash space", "todo", "Not checked - run the readiness check (select the image first so it is counted)")

        if not target:
            add("image", "Firmware image selected", "todo", "Pick an image under 'Firmware upgrade' to see whether it suits this switch")
        elif target["problems"]:
            add("image", "Firmware image suits this switch", "bad", "; ".join(target["problems"]))
        elif target["version"] and version and target["version"] == version:
            last_run = ops.get("upgrade") or {}
            if last_run.get("state") == "completed" and last_run.get("action") != "distribute":
                # the switch is on this version because the last run just upgraded it - nothing to warn about
                add("image", "Firmware image suits this switch", "ok", f"Already upgraded to {version} by the last run")
            else:
                add("image", "Firmware image suits this switch", "warn", f"The switch already runs {version} - the upgrade would be skipped")
        else:
            add("image", "Firmware image suits this switch", "ok", target["image"])

        up = ops.get("upgrade") or {}
        if up.get("state") in ACTIVE:
            add("busy", "No upgrade running", "run", up.get("message") or "An upgrade is running")
        if up.get("state") in ("completed", "failed", "attention"):
            pc = ops.get("postcheck") or {}
            pmap = {"passed": "ok", "warning": "warn", "failed": "bad", "checking": "run"}
            if up.get("action") != "distribute":
                add("post", "Post-upgrade check", pmap.get(pc.get("state"), "todo"),
                    pc.get("detail") or "Not run")
        return items

    def device_file(self, host, kind, name, tail_wanted=False):
        """Text of one saved file for the drawer's 'View' button. Only files that belong to this
        device are served (a name is matched against the device's own listing / file-name pattern)."""
        host = str(host or "")
        with self.lock:
            device_dir, runs_dir = self._device_dirs(host)
        name = os.path.basename(str(name or "").replace("\\", "/"))
        if kind == "backup":
            directory = device_dir
            ok = bool(STAMP_RE.search(name))
        elif kind == "run":
            directory = runs_dir
            ok = bool(run_file_pattern(host).match(name))
        else:
            raise ApiError("Unknown file type.")
        path = directory / name
        try:
            real = path.resolve()
            real.relative_to(directory.resolve())
            is_file = ok and real.is_file() and not path.is_symlink()
        except (ValueError, OSError):
            is_file = False
        if not is_file:
            raise ApiError("File not found.", 404)
        size = real.stat().st_size
        tail = bool(tail_wanted) or name.endswith((".log", ".console.txt"))   # long logs: the end is what matters
        cut = tail and size > DEVICE_FILE_MAX_BYTES
        with open(real, "rb") as fh:
            if cut:
                fh.seek(size - DEVICE_FILE_MAX_BYTES)
            data = fh.read(DEVICE_FILE_MAX_BYTES)
        text = data.decode("utf-8", "replace")
        if cut and "\n" in text:
            text = text.split("\n", 1)[1]              # start at a whole line
        return {"name": name, "size": size, "truncated": size > DEVICE_FILE_MAX_BYTES, "tail": cut,
                "limit": DEVICE_FILE_MAX_BYTES, "text": text}

    UPGRADE_REPORT_FIELDS = ["hostname", "ip_address", "previous_image", "current_image",
                             "upgrade_started_at", "upgrade_duration", "status", "finished_at",
                             "postcheck_result", "postcheck_reason", "action", "image_file", "started_by"]

    POSTCHECK_LABELS = {"passed": "Pass", "warning": "Differences found", "failed": "Fail",
                        "checking": "In progress", "idle": "Not run"}

    def _postcheck_report(self, upgrade_state, pc):
        """(result, reason) for the report. The reason is only filled when the result is not
        a Pass, so a reader can see at a glance why a switch needs attention."""
        pc = pc or {}
        state = pc.get("state") or "idle"
        result = self.POSTCHECK_LABELS.get(state, state)
        if state == "passed":
            return result, ""
        detail = str(pc.get("detail") or "").strip()
        if state == "warning":
            reason = ("Switch is on the new version but something differs from before the upgrade: "
                      + detail) if detail else "Differences found compared with the pre-upgrade snapshot"
        elif state == "failed":
            reason = detail or "Post-check failed"
        elif state == "checking":
            reason = detail or "Post-check is still running"
        else:  # idle - never ran
            if upgrade_state == "idle":
                reason = "No upgrade attempted, so no post-check was run"
            elif upgrade_state in ("completed", "skipped"):
                reason = ("Upgrade not needed (already on target version), so no post-check was run"
                          if upgrade_state == "skipped" else "Post-check did not run after the upgrade")
            else:
                reason = f"Post-check not run - upgrade status is '{upgrade_state}'"
        return result, reason

    @staticmethod
    def _format_duration(seconds):
        if seconds is None or seconds < 0:
            return ""
        seconds = int(round(seconds))
        h, rem = divmod(seconds, 3600)
        m, s = divmod(rem, 60)
        return (f"{h}h {m:02d}m {s:02d}s" if h else f"{m}m {s:02d}s" if m else f"{s}s")

    def upgrade_report_rows(self):
        """One row per device: hostname, IP, previous/current software version, and how long
        its most recent upgrade attempt took. Devices never upgraded still appear, blank."""
        with self.lock:
            self._poll_all()
            rows = []
            for host in self.order:
                o = self.ops[host]
                up = o["upgrade"]
                had_attempt = up["state"] != "idle"
                pc_result, pc_reason = self._postcheck_report(up["state"], o.get("postcheck"))
                duration = None
                if up.get("started_ts") and up.get("finished_ts"):
                    duration = up["finished_ts"] - up["started_ts"]
                rows.append({
                    "hostname": self.info[host]["label"] or host,
                    "ip_address": host,
                    "previous_image": up.get("old_version") or "",
                    "current_image": up.get("new_version") or o["version"] or "",
                    "upgrade_started_at": up.get("started") or "",
                    "upgrade_duration": self._format_duration(duration),
                    "status": up["state"] if had_attempt else "not attempted",
                    "finished_at": up.get("finished") or "",
                    "postcheck_result": pc_result,
                    "postcheck_reason": pc_reason,
                    "action": self.ACTION_VERB.get(up.get("action") or "full", "Upgrade") if up.get("action") else "",
                    "image_file": up.get("image") or "",
                    "started_by": up.get("started_by") or "",
                })
            return rows

    REPORT_LABELS = {"hostname": "Hostname", "ip_address": "IP address", "previous_image": "Previous version",
                     "current_image": "Current version", "upgrade_started_at": "Started",
                     "upgrade_duration": "Duration", "status": "Status", "finished_at": "Finished",
                     "postcheck_result": "Post-check", "postcheck_reason": "Post-check detail",
                     "action": "Action", "image_file": "Image file", "started_by": "Started by"}
    REPORT_STATUS_FILLS = {"completed": "C6EFCE", "failed": "FFC7CE", "attention": "FFEB9C",
                           "running": "DDEBF7", "starting": "DDEBF7", "scheduled": "EDEDED"}

    def _report_counts(self, rows):
        counts = {}
        for r in rows:
            counts[r["status"]] = counts.get(r["status"], 0) + 1
        return counts

    def upgrade_report_xlsx(self, generated_by="local"):
        """Excel report: Summary, current status of every switch, and the full run history."""
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        rows = self.upgrade_report_rows()
        history = list(reversed(load_history(self.site_name))) if self.site_name else []
        head_fill, head_font = PatternFill("solid", fgColor="1F4E79"), Font(bold=True, color="FFFFFF")

        def fill_sheet(ws, headers, data, status_col=None):
            ws.append(headers)
            for cell in ws[1]:
                cell.fill, cell.font = head_fill, head_font
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            for r in data:
                ws.append(r)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, h in enumerate(headers, 1):
                width = max([len(str(h))] + [len(str(r[i - 1])) for r in data[:300]]) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(max(width, 10), 60)
            if status_col is not None:
                for row in ws.iter_rows(min_row=2, min_col=status_col + 1, max_col=status_col + 1):
                    for cell in row:
                        color = self.REPORT_STATUS_FILLS.get(str(cell.value).lower())
                        if color:
                            cell.fill = PatternFill("solid", fgColor=color)

        wb = Workbook()
        ws = wb.active
        ws.title = "Summary"
        counts = self._report_counts(rows)
        ws.append([APP_NAME + " - upgrade report"])
        ws["A1"].font = Font(bold=True, size=14)
        for line in [("Site", self.site_name or ""), ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                     ("Generated by", generated_by), ("Application version", APP_VERSION), ("", ""),
                     ("Switches", len(rows)), ("Completed", counts.get("completed", 0)),
                     ("Failed", counts.get("failed", 0)), ("Needs attention", counts.get("attention", 0)),
                     ("Running / starting", counts.get("running", 0) + counts.get("starting", 0)),
                     ("Scheduled", counts.get("scheduled", 0)), ("Not attempted", counts.get("not attempted", 0)),
                     ("Runs in history", len(history))]:
            ws.append(list(line))
        for r in range(3, ws.max_row + 1):
            ws.cell(r, 1).font = Font(bold=True)
        ws.append([])
        ws.append(["Developed by " + APP_AUTHOR])
        ws.append([APP_REPO_URL])
        ws.cell(ws.max_row, 1).hyperlink = APP_REPO_URL
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 34

        fields = self.UPGRADE_REPORT_FIELDS
        status_idx = fields.index("status")
        fill_sheet(wb.create_sheet("Current status"), [self.REPORT_LABELS[f] for f in fields],
                   [[r.get(f, "") for f in fields] for r in rows], status_idx)
        hist_fields = ["finished", "hostname", "ip_address", "action", "image_file", "previous_image",
                       "current_image", "status", "upgrade_duration", "started_by", "postcheck_result",
                       "postcheck_reason", "message"]
        hist_labels = ["Finished", "Hostname", "IP address", "Action", "Image file", "Previous version",
                       "Current version", "Result", "Duration", "Started by", "Post-check", "Post-check detail",
                       "Message"]
        fill_sheet(wb.create_sheet("History"), hist_labels,
                   [[h.get(f, "") for f in hist_fields] for h in history], hist_fields.index("status"))
        buf = io.BytesIO()
        wb.save(buf)
        data = buf.getvalue()
        self._save_report(data, "upgrade_report", "xlsx")
        return data

    def upgrade_report_html(self, generated_by="local"):
        """A clean, printable page - use the browser's Print > Save as PDF."""
        e = html_escape
        rows = self.upgrade_report_rows()
        history = list(reversed(load_history(self.site_name)))[:100] if self.site_name else []
        counts = self._report_counts(rows)

        def cls(text):
            return "s-" + re.sub(r"[^a-z]", "", str(text).lower())

        def table(headers, data, status_col):
            head = "".join("<th>{}</th>".format(e(h)) for h in headers)
            body = "".join(
                "<tr>" + "".join(
                    "<td class=\"{}\">{}</td>".format(cls(v) if i == status_col else "", e(str(v)))
                    for i, v in enumerate(r)) + "</tr>" for r in data)
            return "<table><thead><tr>{}</tr></thead><tbody>{}</tbody></table>".format(head, body or
                   "<tr><td colspan=\"{}\">Nothing to show yet.</td></tr>".format(len(headers)))

        fields = self.UPGRADE_REPORT_FIELDS
        tiles = "".join("<div class=\"tile\"><b>{}</b><span>{}</span></div>".format(n, e(t)) for t, n in [
            ("Switches", len(rows)), ("Completed", counts.get("completed", 0)), ("Failed", counts.get("failed", 0)),
            ("Needs attention", counts.get("attention", 0)), ("Scheduled", counts.get("scheduled", 0))])
        hist_fields = ["finished", "hostname", "action", "previous_image", "current_image", "status",
                       "upgrade_duration", "started_by", "postcheck_result"]
        hist_labels = ["Finished", "Hostname", "Action", "From", "To", "Result", "Duration", "Started by",
                       "Post-check"]
        return (
            "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>{app} - upgrade report</title>"
            "<style>body{{font:13px/1.45 system-ui,Segoe UI,sans-serif;margin:24px;color:#1b2430}}"
            "h1{{font-size:20px;margin:0}}h2{{font-size:15px;margin:22px 0 6px}}.sub{{color:#51606f;margin:2px 0 14px}}"
            ".tiles{{display:flex;gap:10px;flex-wrap:wrap}}.tile{{border:1px solid #d5dce4;border-radius:8px;"
            "padding:8px 14px;min-width:92px}}.tile b{{display:block;font-size:20px}}.tile span{{color:#51606f}}"
            "table{{border-collapse:collapse;width:100%;font-size:12px}}th,td{{border:1px solid #d5dce4;"
            "padding:4px 7px;text-align:left;vertical-align:top}}th{{background:#1f4e79;color:#fff}}"
            ".s-completed{{background:#c6efce}}.s-failed{{background:#ffc7ce}}.s-attention{{background:#ffeb9c}}"
            ".s-running,.s-starting{{background:#ddebf7}}.foot{{margin-top:22px;color:#51606f;font-size:12px;"
            "border-top:1px solid #d5dce4;padding-top:8px}}button{{padding:6px 14px;margin:0 0 12px}}"
            "@media print{{button{{display:none}}body{{margin:8mm}}}}</style></head><body>"
            "<button onclick=\"window.print()\">Print / Save as PDF</button>"
            "<h1>{app} - Upgrade Report</h1>"
            "<p class=\"sub\">Site: <b>{site}</b> &nbsp;|&nbsp; Generated {now} by {who} &nbsp;|&nbsp; v{ver}</p>"
            "<div class=\"tiles\">{tiles}</div>"
            "<h2>Current Status</h2>{t1}<h2>History (Latest 100 Runs)</h2>{t2}"
            "<p class=\"foot\">Developed by <b>{author}</b> | <a href=\"{repo}\">GitHub</a></p></body></html>"
        ).format(app=e(APP_NAME), site=e(self.site_name or ""), now=datetime.now().strftime("%Y-%m-%d %H:%M"),
                 who=e(generated_by), ver=e(APP_VERSION), tiles=tiles,
                 t1=table([self.REPORT_LABELS[f] for f in fields], [[r.get(f, "") for f in fields] for r in rows],
                          fields.index("status")),
                 t2=table(hist_labels, [[h.get(f, "") for f in hist_fields] for h in history],
                          hist_fields.index("status")),
                 author=e(APP_AUTHOR), repo=e(APP_REPO_URL, quote=True))

    def upgrade_report_csv(self):
        buf = io.StringIO()
        w = csv.DictWriter(buf, fieldnames=self.UPGRADE_REPORT_FIELDS)
        w.writeheader()
        w.writerows(self.upgrade_report_rows())
        text = buf.getvalue()
        self._save_report(text, "upgrade_report")
        return text

    def _save_report(self, data, prefix, ext="csv"):
        """Every generated report is also saved under this site's Reports folder, not just
        streamed to the browser - best-effort, a failure here never blocks the download."""
        try:
            self.reports_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = self.reports_dir / f"{prefix}_{stamp}.{ext}"
            if isinstance(data, bytes):
                path.write_bytes(data)
            else:
                path.write_text(data, encoding="utf-8")
        except (OSError, AttributeError):
            pass

    def _record_history(self, host):
        """Append the run that just finished to this site's permanent history (best effort:
        a problem here must never affect the upgrade itself)."""
        try:
            o = self.ops[host]
            up = o["upgrade"]
            pc_result, pc_reason = self._postcheck_report(up["state"], o.get("postcheck"))
            ts_start, ts_end = up.get("started_ts"), up.get("finished_ts") or time.time()
            fmt = "%Y-%m-%d %H:%M:%S"
            append_history(self.site_name, {
                "finished": datetime.fromtimestamp(ts_end).strftime(fmt),
                "started": datetime.fromtimestamp(ts_start).strftime(fmt) if ts_start else "",
                "hostname": self.label(host), "ip_address": host,
                "action": self.ACTION_VERB.get(up.get("action") or "full", "Upgrade"),
                "image_file": up.get("image") or "",
                "previous_image": up.get("old_version") or "", "current_image": up.get("new_version") or "",
                "status": up["state"], "message": up.get("message") or "",
                "upgrade_duration": self._format_duration(ts_end - ts_start) if ts_start else "",
                "started_by": up.get("started_by") or "",
                "postcheck_result": pc_result, "postcheck_reason": pc_reason})
        except Exception:  # noqa: BLE001
            pass

    def _enforce_image_check(self, hosts, image, skip, user):
        """Refuse (HTTP 409) when the chosen image looks wrong for a selected switch, unless
        the person ticked 'Ignore image check warnings'. Called with self.lock held."""
        name = Path(image).name
        problems = []
        for h in hosts:
            o = self.ops[h]
            model = o.get("model") or self.info[h].get("platform") or ""
            for p in check_image_for_device(name, model, o.get("version")):
                problems.append(f"{self.label(h) or h} ({h}): {p}")
        if not problems:
            return
        if skip:
            self.events.add("-", "", f"Image check warnings ignored by {user or 'local'}: "
                            + "; ".join(problems[:3]) + (" ..." if len(problems) > 3 else ""), "warn")
            return
        shown = "; ".join(problems[:4]) + (f"; and {len(problems) - 4} more" if len(problems) > 4 else "")
        raise ApiError("Image check: " + shown + ". Fix the selection, or tick 'Ignore image check "
                       "warnings' to continue anyway.", 409)

    def active_upgrades(self):
        with self.lock:
            self._poll_all()
            return [h for h in self.order if self._upgrade_active(h)]

    def _save_schedule(self):
        """Pending scheduled upgrades are saved with the site, so a restart can tell which ones
        were missed. They contain no credentials (those come from the profiles when they run)."""
        if not self.site_name:
            return
        try:
            tmp = site_schedule_path(self.site_name).with_suffix(".json.tmp")
            tmp.write_text(json.dumps(self.schedule), encoding="utf-8")
            tmp.replace(site_schedule_path(self.site_name))
        except OSError as exc:
            print(f"Could not save the scheduled upgrades: {exc}", flush=True)

    def _need_awake(self):
        """True while an upgrade is running or a scheduled upgrade is waiting."""
        with self.lock:
            return bool(self.schedule) or any(self.ops[h]["upgrade"]["state"] in ACTIVE for h in self.order)

    def _awake_changed(self, on, ok):
        if on:
            self.events.add("-", "", "Windows is being kept awake (an upgrade is running or a scheduled "
                                     "upgrade is waiting)" + ("" if ok else " - WARNING: the request failed"), "info")
        else:
            self.events.add("-", "", "No upgrade running or scheduled - Windows may sleep again", "info")

    def close(self):
        self.term_close_all()
        self.keep_awake.stop()
        self._stop.set()
        self.pool.shutdown(wait=False, cancel_futures=True)
        shutil.rmtree(self.job_tmp, ignore_errors=True)


# --------------------------------------------------------------------------- #
# Web GUI
# --------------------------------------------------------------------------- #
DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NetUpgrade</title>
<style>
  /* Cisco Cloud Control style. Light by default, dark follows the OS/browser or the toggle. */
  :root {
    --bg:#eef1f8; --bg2:#e4e9f5; --tb:linear-gradient(90deg,#fff,#eef0fb); --tbd:#d6dcec;
    --pn:#fff; --pn2:#f5f7fc; --bd:#dbe1ee; --tx:#14213d; --mu:#5d6b88;
    --ac:#2f6fe4; --acs:#e3ecfd; --sel:#e3ecfd;
    --t-b:#dcecf9; --t-r:#fde3e6; --t-a:#fcefc9; --t-g:#e9edf4; --t-gn:#dff3e6;
    --c-b:#0b6aa8; --c-r:#c62a3a; --c-a:#9a6a00;
    --gr:#12834f; --am:#b97500; --rd:#c62a3a; --sh:0 1px 3px #1a2a5014; color-scheme:light;
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      --bg:#080d1a; --bg2:#0d1428; --tb:linear-gradient(90deg,#1b2552,#10173a 55%,#0c1230); --tbd:#232d57;
      --pn:#0f1626; --pn2:#0b1120; --bd:#1d2742; --tx:#e8edf8; --mu:#93a1c0;
      --ac:#4f86f7; --acs:#16254a; --sel:#1b2a52;
      --t-b:#1e3b52; --t-r:#3a2129; --t-a:#3b3320; --t-g:#2a3140; --t-gn:#17382a;
      --c-b:#6cc0f5; --c-r:#ff7b8a; --c-a:#f4c542;
      --gr:#4ade80; --am:#f4c542; --rd:#ff7b8a; --sh:none; color-scheme:dark;
    }
  }
  :root[data-theme="dark"] {
    --bg:#080d1a; --bg2:#0d1428; --tb:linear-gradient(90deg,#1b2552,#10173a 55%,#0c1230); --tbd:#232d57;
    --pn:#0f1626; --pn2:#0b1120; --bd:#1d2742; --tx:#e8edf8; --mu:#93a1c0;
    --ac:#4f86f7; --acs:#16254a; --sel:#1b2a52;
    --t-b:#1e3b52; --t-r:#3a2129; --t-a:#3b3320; --t-g:#2a3140; --t-gn:#17382a;
    --c-b:#6cc0f5; --c-r:#ff7b8a; --c-a:#f4c542;
    --gr:#4ade80; --am:#f4c542; --rd:#ff7b8a; --sh:none; color-scheme:dark;
  }
  * { box-sizing:border-box; }
  html, body { height:100%; margin:0; }
  body { background:radial-gradient(1200px 500px at 70% -10%, var(--bg2), var(--bg)); color:var(--tx);
         font:13px/1.45 "Inter",system-ui,"Segoe UI",Arial,sans-serif;
         display:flex; flex-direction:column; overflow:hidden; }
  svg.i { width:18px; height:18px; fill:none; stroke:currentColor; stroke-width:1.7; stroke-linecap:round;
          stroke-linejoin:round; flex:none; }
  .muted { color:var(--mu); }
  [data-i] { display:grid; place-items:center; }
  .sp { flex:1; }

  /* ---- header ---- */
  header { background:var(--tb); border-bottom:1px solid var(--tbd); display:flex; align-items:center; gap:10px;
           padding:0 16px; height:56px; flex:none; }
  .appver { font-size:11px; font-weight:600; opacity:.65; align-self:flex-end; margin-bottom:3px; }
  .brand { font-size:18px; font-weight:300; display:flex; align-items:center; gap:8px; padding-right:14px;
           border-right:1px solid var(--tbd); height:32px; white-space:nowrap; }
  .brand img { width:26px; height:26px; object-fit:contain; flex:none; }
  .lost { color:var(--rd); font-weight:600; font-size:12px; }
  .ib { background:none; border:0; color:var(--tx); border-radius:8px; width:36px; height:36px; display:grid;
        place-items:center; cursor:pointer; }
  .ib:hover { background:var(--acs); }
  .mstate { font-size:12px; font-weight:600; padding:5px 13px; border-radius:999px; background:var(--t-a);
            color:var(--c-a); border:1px solid var(--c-a); white-space:nowrap; }
  .mdot { width:11px; height:11px; border-radius:50%; background:var(--mu); flex:none; margin-right:-4px; }
  .mdot.running { background:var(--gr); animation:mon-blink 1.3s ease-in-out infinite; }
  @keyframes mon-blink { 0% { box-shadow:0 0 0 0 rgba(34,170,90,.65); opacity:1; } 50% { opacity:.45; }
                         100% { box-shadow:0 0 0 9px rgba(34,170,90,0); opacity:1; } }
  @media (prefers-reduced-motion:reduce) { .mdot.running { animation:none; } }
  /* status note: floats over the page (no layout space) and fades out by itself */
  .mainmsg { position:fixed; right:18px; bottom:18px; z-index:1200; max-width:min(520px,calc(100vw - 36px));
             margin:0; padding:9px 14px; border-radius:8px; background:var(--pn); border:1px solid var(--bd);
             box-shadow:0 6px 20px rgba(0,0,0,.25); pointer-events:none; }
  .mainmsg:has(#msg:empty) { display:none; } .mainmsg .msg { margin:0; }
  .mstate.running { background:var(--t-gn); color:var(--gr); border-color:var(--gr); }

  /* site badge + dropdown */
  .badge { display:inline-block; padding:6px 14px; border-radius:999px; font-size:13px; font-weight:500;
           background:var(--t-g); color:var(--mu); border:1px solid var(--tbd); }
  .badge.accent { background:var(--acs); color:var(--tx); border-color:var(--ac); }
  .badge.warn { background:var(--t-a); color:var(--c-a); border-color:var(--c-a); }
  .site-btn { cursor:pointer; font-family:inherit; }
  .site-btn::after { content:" \25BE"; }
  .site-btn:disabled { opacity:.55; cursor:not-allowed; }
  .site-btn.needed { animation:pulse 1.4s ease-in-out infinite; }
  .sitemenu { position:relative; display:inline-block; }
  .site-pop { position:absolute; left:0; top:calc(100% + 8px); z-index:950; min-width:230px; max-width:320px;
              max-height:300px; overflow-y:auto; background:var(--pn); color:var(--tx); border:1px solid var(--bd);
              border-radius:10px; padding:4px; box-shadow:0 10px 30px rgba(0,0,0,.25); font-size:13px; }
  .site-pop[hidden] { display:none; }
  .notemenu { position:relative; display:inline-block; }
  .ib .nbadge { position:absolute; top:2px; right:1px; min-width:16px; height:16px; padding:0 4px; border-radius:8px;
                background:var(--rd); color:#fff; font-size:10px; font-weight:700; line-height:16px; text-align:center; }
  .ib .nbadge[hidden] { display:none; }
  .note-pop { position:absolute; right:0; top:calc(100% + 8px); z-index:1300; width:360px; max-width:calc(100vw - 24px);
              background:var(--pn); color:var(--tx); border:1px solid var(--bd); border-radius:10px;
              box-shadow:0 10px 30px rgba(0,0,0,.25); font-size:13px; overflow:hidden; }
  .note-pop[hidden] { display:none; }
  .note-pop .nh { display:flex; align-items:center; justify-content:space-between; padding:9px 12px; font-weight:700;
                  border-bottom:1px solid var(--bd); }
  .note-pop .nh button { background:none; border:0; color:var(--ac); cursor:pointer; font:inherit; font-size:12px; font-weight:600; }
  .note-pop .nl { max-height:380px; overflow-y:auto; }
  .note-pop .ni { display:flex; gap:10px; padding:9px 12px; border-bottom:1px solid var(--bd); cursor:pointer; }
  .note-pop .ni:last-child { border-bottom:0; }
  .note-pop .ni:hover { background:var(--acs); }
  .note-pop .ni.unread { background:var(--sel); }
  .note-pop .ni .ic { flex:0 0 auto; font-size:15px; line-height:1.3; }
  .note-pop .ni .nt { font-weight:600; }
  .note-pop .ni .nm { color:var(--mu); font-size:12px; word-break:break-word; margin-top:1px; }
  .note-pop .ni .tm { color:var(--mu); font-size:11px; margin-top:2px; }
  .note-pop .none { padding:16px 12px; color:var(--mu); text-align:center; }
  .site-pop button { display:block; width:100%; text-align:left; background:transparent; color:var(--tx); border:0;
                     padding:7px 10px; border-radius:7px; font:inherit; cursor:pointer; white-space:nowrap;
                     overflow:hidden; text-overflow:ellipsis; }
  .site-pop button:hover { background:var(--acs); }
  .site-pop button.cur { font-weight:700; color:var(--ac); }
  .site-pop .none { padding:7px 10px; color:var(--mu); }
  .site-pop hr { border:0; border-top:1px solid var(--bd); margin:4px 0; }
  .site-pop .newsite { color:var(--ac); font-weight:600; }
  .modal { position:fixed; inset:0; z-index:1300; background:rgba(0,0,0,.45); display:flex; align-items:center;
           justify-content:center; padding:16px; }
  .modal[hidden] { display:none; }
  .modal-card { background:var(--pn); border:1px solid var(--bd); border-radius:12px; padding:16px 18px; width:100%;
                max-width:340px; box-shadow:0 14px 50px rgba(0,0,0,.4); }
  .modal-card h2 { font-size:15px; margin:0 0 6px; }

  /* ---- buttons and inputs ---- */
  .btn { background:var(--pn2); border:1px solid var(--bd); color:var(--tx); border-radius:7px; padding:6px 12px;
         cursor:pointer; font:inherit; text-decoration:none; display:inline-block; text-align:center; }
  .btn:hover:not(:disabled) { border-color:var(--ac); }
  .btn.p { background:var(--ac); border-color:var(--ac); color:#fff; font-weight:600; }
  .btn.sm { padding:4px 10px; font-size:12px; }
  .btn.on { border-color:var(--ac); color:var(--ac); }
  .btn:disabled { opacity:.45; cursor:not-allowed; }
  label { display:block; font-size:11px; color:var(--mu); margin:10px 0 4px; }
  input[type=text], input[type=password], input[type=datetime-local], textarea, select {
    width:100%; background:var(--pn); border:1px solid var(--bd); color:var(--tx); border-radius:6px;
    padding:7px 9px; font:inherit; }
  input[type=text] + input[type=text] { margin-top:6px; }
  input:focus, select:focus, textarea:focus { outline:none; border-color:var(--ac); box-shadow:0 0 0 2px var(--acs); }
  input[type=file] { width:100%; font-size:11px; color:var(--mu); }
  input[type=checkbox] { accent-color:var(--ac); width:15px; height:15px; cursor:pointer; margin:0; }
  textarea { height:80px; resize:vertical; font-family:ui-monospace,Consolas,monospace; font-size:11.5px; }
  a.lnk { color:var(--ac); text-decoration:underline; cursor:pointer; font-size:11px; }
  a.lnk:hover { opacity:.8; }
  .hint, .note { font-size:11px; color:var(--mu); margin-top:6px; }
  .hint.bad { color:var(--rd); }
  .msg { font-size:12px; font-weight:600; margin:2px 0 6px; }
  .msg:empty { display:none; }
  .side .body > .msg { margin:12px 0 0; overflow-wrap:anywhere; }          /* status lines sit at the bottom of each section */
  .protobox { display:flex; align-items:center; gap:12px; flex-wrap:nowrap; margin-top:2px; }
  .protobox .ropt { display:flex; align-items:center; gap:6px; margin:0; font-size:12px; color:var(--tx); cursor:pointer; }
  input[type=radio] { accent-color:var(--ac); width:15px; height:15px; cursor:pointer; margin:0; }
  .protobox .portwrap { display:flex; align-items:center; gap:6px; margin-left:auto; flex:none; white-space:nowrap; }
  .protobox .portwrap label { margin:0; }
  .protobox .portwrap input[type=text] { width:56px; text-align:center; padding:5px 6px; }
  .msg.ok { color:var(--gr); } .msg.err { color:var(--rd); }
  .row { display:flex; gap:6px; margin-top:10px; }
  .row .btn { flex:1; padding:6px 8px; white-space:nowrap; }
  .chk { display:flex; align-items:center; gap:8px; font-size:12px; color:var(--tx); margin:12px 0 0; cursor:pointer; }
  hr.sep { border:0; border-top:1px solid var(--bd); margin:14px 0 2px; }
  /* help tips: a small ? (how it works) or ! (important) icon; the text shows on hover, focus or click */
  .tip { display:inline-grid; place-items:center; vertical-align:middle; width:16px; height:16px; margin-left:6px; flex:none; padding:0;
         border-radius:50%; border:1.5px solid var(--ac); color:var(--ac); background:transparent; cursor:help; user-select:none;
         font:700 11px/1 system-ui,Segoe UI,sans-serif; text-transform:none; letter-spacing:0; }
  .tip.w { border-color:var(--am); color:var(--am); }
  .tip.bad { border-color:var(--rd); color:var(--rd); }
  .tip:hover, .tip:focus-visible, .tip.on { background:var(--ac); color:var(--pn); }
  .tip.w:hover, .tip.w:focus-visible, .tip.w.on { background:var(--am); }
  .tip.bad:hover, .tip.bad:focus-visible, .tip.bad.on { background:var(--rd); }
  .row .tip { align-self:center; margin-left:2px; }
  .dr-term .th .tip { border-color:#8fb0ff; color:#8fb0ff; margin-left:0; }
  .dr-term .th .tip:hover, .dr-term .th .tip.on { background:#8fb0ff; color:#151c33; }
  #tip-pop { position:fixed; z-index:2000; max-width:280px; padding:9px 12px; border-radius:8px; background:var(--pn); color:var(--tx);
             border:1px solid var(--bd); border-left:3px solid var(--ac); box-shadow:0 8px 28px #00000040; font-size:12px; line-height:1.45;
             font-weight:400; pointer-events:none; white-space:pre-line; }
  #tip-pop.w { border-left-color:var(--am); } #tip-pop.bad { border-left-color:var(--rd); }
  #tip-pop[hidden] { display:none; }
  .dr-term .tn:empty { display:none; }
  :focus-visible { outline:2px solid var(--ac); outline-offset:1px; }

  /* ---- layout: sidebar + main ---- */
  .app { flex:1; display:flex; min-height:0; }
  .side { flex:none; display:flex; flex-direction:column; overflow:auto; border-right:1px solid var(--bd);
          background:var(--pn2); transition:width .15s; }
  .side.wide { width:266px; }
  .side.narrow { width:62px; align-items:center; padding:8px 0; gap:4px; }
  .sh { display:flex; align-items:center; gap:12px; padding:14px 14px 10px; font-size:17px; font-weight:600; }
  .cb { background:none; border:0; color:var(--tx); cursor:pointer; width:34px; height:34px; border-radius:8px;
        display:grid; place-items:center; flex:none; }
  .cb:hover { background:var(--acs); }
  .sec { margin:2px 8px; }
  .sec > h3 { margin:0; padding:11px 12px; font-size:14px; font-weight:500; cursor:pointer; display:flex;
              align-items:center; gap:12px; border-radius:8px; color:var(--mu); user-select:none; }
  .sec > h3:hover { background:var(--acs); color:var(--tx); }
  .sec.o > h3 { background:var(--sel); color:var(--tx); box-shadow:3px 0 0 var(--ac) inset; }
  .sec > h3 .ch { margin-left:auto; transition:transform .15s; display:grid; }
  .sec.o > h3 .ch { transform:rotate(180deg); }
  .sec .body { display:none; padding:4px 12px 14px; }
  .sec.o .body { display:block; }
  .side.narrow .sh { padding:0; }
  .side.narrow .shl, .side.narrow .sec > h3 .lbl, .side.narrow .sec > h3 .ch, .side.narrow .sec .body { display:none !important; }
  .side.narrow .sec { margin:0; }
  .side.narrow .sec > h3 { width:42px; height:42px; padding:0; justify-content:center; border-radius:10px; }
  .side.narrow .sec.o > h3 { box-shadow:none; }
  #dashboard.nosite .side .sec { opacity:.45; pointer-events:none; }

  main { flex:1; min-width:0; display:flex; flex-direction:column; gap:12px; padding:12px 16px; overflow:auto; }
  .card { background:var(--pn); border:1px solid var(--bd); border-radius:10px; display:flex; flex-direction:column;
          min-height:0; box-shadow:var(--sh); overflow:hidden; }
  .card > h4 { margin:0; padding:9px 14px; font-size:14px; font-weight:600; border-bottom:1px solid var(--bd);
               display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
  .card > h4 .a { margin-left:auto; display:flex; align-items:center; gap:6px; flex-wrap:wrap; font-weight:400;
                  font-size:12px; color:var(--mu); }
  .fl { display:flex; gap:6px; margin-left:6px; }
  .col-pop { position:fixed; z-index:1250; background:var(--pn); border:1px solid var(--bd); border-radius:10px;
             box-shadow:var(--sh); padding:8px 6px; min-width:190px; font-size:12.5px; }
  .col-pop[hidden] { display:none; }
  .col-pop .cp-h { padding:2px 8px 6px; color:var(--mu); font-size:11px; text-transform:uppercase; letter-spacing:.04em; }
  .col-pop label.ci { display:flex; align-items:center; gap:8px; margin:0; padding:5px 8px; border-radius:6px;
                      font-size:12.5px; color:var(--tx); cursor:pointer; }
  .col-pop label.ci:hover { background:var(--acs); }
  .col-pop .cp-f { padding:6px 8px 2px; border-top:1px solid var(--bd); margin-top:4px; }
  table.hc-platform .c-platform, table.hc-status .c-status, table.hc-creds .c-creds, table.hc-ready .c-ready,
  table.hc-version .c-version, table.hc-boot .c-boot, table.hc-backup .c-backup, table.hc-stage .c-stage, table.hc-upgrade .c-upgrade,
  table.hc-post .c-post, table.hc-recent .c-recent { display:none; }

  /* status bar (replaces the old progress card + toolbar link) */
  .fbar { position:relative; display:flex; align-items:center; gap:14px; background:var(--pn); border:1px solid var(--bd);
          border-radius:10px; padding:10px 14px; box-shadow:var(--sh); overflow:hidden; flex:none; }
  .fbar .ftext { font-weight:600; white-space:nowrap; }
  .fbar #p-cycles { color:var(--mu); font-weight:400; margin-left:10px; }
  .fbar #meta { color:var(--mu); font-size:12px; white-space:nowrap; }
  .fbar .msg { margin:0; max-width:40%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .pbar { position:absolute; left:0; right:0; bottom:0; height:3px; background:transparent; }
  .pbar > div { height:100%; width:0; background:var(--ac); transition:width .5s; }

  /* summary cards */
  .wg { display:grid; grid-template-columns:repeat(2, 1fr); gap:12px; flex:none; }
  .tl { display:grid; grid-template-columns:repeat(auto-fit, minmax(70px, 1fr)); gap:10px; padding:12px 14px; }
  .t { border-radius:6px; padding:8px 6px; text-align:center; background:var(--t-g); }
  .t b { display:block; font-size:22px; font-weight:600; line-height:1.2; }
  .t span { font-size:12px; color:var(--mu); }
  .t.b { background:var(--t-b); } .t.r { background:var(--t-r); } .t.a { background:var(--t-a); }
  .t.gn { background:var(--t-gn); } .t.g { background:var(--t-g); }

  /* devices table: rows scroll inside the card */
  .work { flex:1; min-height:280px; display:flex; flex-direction:column; gap:12px; }
  .card.tb { flex:1; }
  .card.max { position:fixed; inset:10px; z-index:1000; height:auto; box-shadow:0 14px 60px rgba(0,0,0,.45); }
  .card.max .sc { max-height:none; }
  body.card-max { overflow:hidden; }
  .card > h4 .ib.mx { width:26px; height:26px; margin:-3px -6px -3px 0; flex:0 0 auto; }
  .card.e > h4 .ib.mx { margin-left:auto; }
  .card > h4 .ib.mx .i { width:16px; height:16px; }
  .card.e { height:128px; flex:none; }
  .card.e.max { height:auto; }
  .sc { overflow:auto; flex:1; min-height:0; }
  table { width:100%; border-collapse:collapse; min-width:900px; }
  th { font-size:11px; color:var(--mu); text-align:left; padding:9px 6px; position:sticky; top:0; z-index:1;
       background:var(--pn2); font-weight:600; white-space:normal; line-height:1.25; }
  td { padding:9px 6px; border-top:1px solid var(--bd); white-space:nowrap; vertical-align:middle; }
  td.ck, th.ck { width:34px; padding-right:0; padding-left:14px; }
  tr:hover td { background:var(--pn2); }
  tr.sel td { background:var(--sel); }
  .host { font-weight:600; }
  th.rc, td.rc { width:196px; min-width:196px; max-width:196px; }
  .hist { display:flex; gap:2px; align-items:center; width:178px; }
  .hist i { display:block; width:3px; height:16px; border-radius:1px; background:var(--gr); flex:none; }
  .hist i.f { background:var(--rd); }
  .hist i.e { background:var(--t-g); }
  .label { color:var(--mu); font-size:12px; }
  .clip { display:block; max-width:112px; overflow:hidden; text-overflow:ellipsis; }
  .clink { font-size:11px; color:var(--ac); text-decoration:none; margin-left:6px; font-weight:600; }
  .clink:hover { text-decoration:underline; }

  .pill { display:inline-block; min-width:56px; text-align:center; padding:2px 8px; border-radius:999px;
          font-size:11px; font-weight:700; }
  .pill.UP { background:var(--t-b); color:var(--c-b); }
  .pill.OK { background:var(--t-gn); color:var(--gr); }
  .pill.DOWN { background:var(--t-r); color:var(--c-r); }
  .pill.PENDING { background:var(--t-g); color:var(--mu); }
  .pill.RUN { background:var(--t-a); color:var(--c-a); animation:pulse 1.4s ease-in-out infinite; }
  .pill.WARN, .pill.CHECK { background:var(--t-a); color:var(--c-a); box-shadow:inset 0 0 0 1px var(--c-a); }
  @keyframes pulse { 50% { opacity:.55; } }
  @media (prefers-reduced-motion: reduce) { .pill.RUN, .site-btn.needed { animation:none; } }

  /* events */
  .ev { padding:6px 14px; overflow:auto; flex:1; font-size:12px; }
  .ev > div { padding:3px 0; color:var(--mu); }
  .ev b { color:var(--tx); }
  .dot { display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:6px; background:var(--mu); }
  .dot.ok { background:var(--gr); } .dot.bad { background:var(--rd); } .dot.warn { background:var(--am); }

  /* ---- device drawer: opens on the right when a switch name is clicked ---- */
  :root { --dw: 470px; }
  .devlink { display:block; color:inherit; text-decoration:none; cursor:pointer; }
  .devlink:hover .host { color:var(--ac); text-decoration:underline; }
  .devlink:focus-visible { outline:2px solid var(--ac); outline-offset:2px; border-radius:4px; }
  tr.dopen td { background:var(--acs); }
  tr.dopen td:first-child + td { box-shadow:inset 3px 0 0 var(--ac); }
  body.drawer-open #dashboard { margin-right:var(--dw); }
  body.drawer-open #toasts { right:calc(var(--dw) + 14px); }
  .drawer { position:fixed; top:56px; right:0; bottom:0; width:var(--dw); z-index:1000; background:var(--pn);
            border-left:1px solid var(--bd); box-shadow:-8px 0 28px rgba(0,0,0,.14); display:flex; flex-direction:column; }
  .drawer[hidden] { display:none; }
  @media (max-width:1399px) { body.drawer-open #dashboard { margin-right:0; } }   /* narrower screens: the drawer overlays instead of squeezing the table */
  @media (max-width:1000px) {
    body.drawer-open #toasts { right:14px; }
    .drawer { width:100vw; max-width:100vw; }
  }
  .dr-head { display:flex; align-items:center; gap:11px; padding:14px 12px 10px 14px; flex:none; }
  .dr-ic { width:42px; height:42px; border-radius:50%; background:var(--acs); color:var(--ac); display:grid; place-items:center; flex:none; }
  .dr-title { min-width:0; flex:1; }
  .dr-name { font-size:16px; font-weight:700; display:flex; align-items:center; gap:8px; flex-wrap:wrap; overflow-wrap:anywhere; }
  .dr-sub { color:var(--mu); font-size:12px; margin-top:2px; overflow-wrap:anywhere; }
  .dr-tabs { display:flex; flex-wrap:wrap; justify-content:space-between; gap:0; padding:0 8px; border-bottom:1px solid var(--bd); flex:none; }
  .dr-tabs button { background:none; border:0; border-bottom:2px solid transparent; color:var(--mu); padding:9px 5px; margin-bottom:-1px;
                    font:inherit; font-size:12.5px; cursor:pointer; white-space:nowrap; }
  .dr-tabs button:hover { color:var(--tx); }
  .dr-tabs button.on { color:var(--ac); border-bottom-color:var(--ac); font-weight:600; }
  .dr-tabs .n { display:inline-block; min-width:17px; padding:0 4px; margin-left:3px; border-radius:9px; background:var(--t-g); color:var(--mu);
                font-size:10.5px; line-height:17px; text-align:center; font-weight:600; }
  .dr-body { flex:1; overflow-y:auto; padding:12px 14px 22px; }
  .dr-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:14px; }
  .dr-card { border:1px solid var(--bd); border-radius:8px; padding:9px 11px; background:var(--pn2); min-width:0; }
  .dr-card .k { font-size:11px; color:var(--mu); }
  .dr-card .v { font-size:15px; font-weight:700; margin-top:3px; overflow-wrap:anywhere; }
  .dr-card .s { font-size:11px; color:var(--mu); margin-top:3px; overflow-wrap:anywhere; }
  .dr-card .s.bad { color:var(--rd); }
  .dr-sec { margin:0 0 16px; }
  .dr-fold { border:1px solid var(--bd); border-radius:8px; padding:0 11px; }
  .dr-fold > h5 { margin:0; padding:9px 0; }
  .dr-fold.open > h5 { border-bottom:1px solid var(--bd); margin-bottom:8px; }
  .fold-btn { flex:1; min-width:0; display:flex; align-items:center; gap:6px; background:none; border:0; padding:0; color:inherit;
              font:inherit; font-weight:700; text-align:left; cursor:pointer; }
  .fold-btn:focus-visible { outline:2px solid var(--ac); outline-offset:3px; border-radius:4px; }
  .fold-btn .chev { display:inline-grid; transition:transform .15s; color:var(--mu); order:2; margin-left:auto; }
  .dr-fold.open .chev { transform:rotate(180deg); }
  .dr-fold > h5 .sp { display:none; }
  .dr-fold > h5 .pill { order:1; }
  .dr-fold > h5 { display:flex; }
  .fold-body { padding-bottom:8px; }
  .dr-sec > h5 { margin:0 0 6px; font-size:13px; font-weight:700; display:flex; align-items:center; gap:8px; }
  .dr-sec > h5 .sp { flex:1; }
  .dr-sec > h5 .btn { font-weight:500; }
  .dr-kv, .cm { min-width:0; }      /* the page's global table { min-width:900px } must not apply here */
  .dr-kv { width:100%; border-collapse:collapse; font-size:12.5px; }
  .dr-kv td { padding:5px 0; border-top:0; border-bottom:1px solid var(--bd); vertical-align:top; white-space:normal; }
  .dr-kv td:first-child { color:var(--mu); width:40%; padding-right:10px; }
  .dr-kv td:last-child { text-align:right; overflow-wrap:anywhere; }
  .dr-note { font-size:11.5px; color:var(--mu); margin-top:6px; }
  .dr-empty { color:var(--mu); font-size:12.5px; padding:6px 0; }
  .dr-banner { border-radius:8px; padding:8px 11px; font-weight:600; font-size:12.5px; margin-bottom:4px; }
  .dr-banner.ok { background:var(--t-gn); color:var(--gr); }
  .dr-banner.warn { background:var(--t-a); color:var(--c-a); }
  .dr-banner.bad { background:var(--t-r); color:var(--c-r); }
  .steps { display:flex; align-items:flex-start; margin:4px 0 12px; }
  .step { flex:1; display:flex; flex-direction:column; align-items:center; position:relative; font-size:10.5px; color:var(--mu); text-align:center; }
  .step .c { width:24px; height:24px; border-radius:50%; display:grid; place-items:center; font-size:11.5px; font-weight:700;
             background:var(--t-g); color:var(--mu); z-index:1; margin-bottom:4px; }
  .step::before { content:""; position:absolute; top:11px; left:-50%; right:50%; height:2px; background:var(--bd); }
  .step:first-child::before { display:none; }
  .step.done .c { background:var(--gr); color:#fff; }
  .step.done::before { background:var(--gr); }
  .step.warn .c { background:var(--am); color:#fff; }
  .step.warn::before { background:var(--am); }
  .step.failed .c { background:var(--rd); color:#fff; }
  .step.active .c { background:var(--ac); color:#fff; animation:pulse 1.4s ease-in-out infinite; }
  .step.done, .step.warn, .step.failed, .step.active { color:var(--tx); }
  .dr-bar { height:6px; border-radius:3px; background:var(--t-g); overflow:hidden; flex:1; }
  .dr-bar > i { display:block; height:100%; background:var(--ac); transition:width .4s; }
  .dr-barrow { display:flex; align-items:center; gap:10px; font-size:12px; }
  .dr-ev { display:flex; align-items:flex-start; gap:6px; padding:5px 0; border-bottom:1px solid var(--bd); font-size:12.5px; }
  .dr-ev .t { margin-left:auto; color:var(--mu); font-size:11px; white-space:nowrap; padding-left:8px; }
  .dr-ev .x { min-width:0; overflow-wrap:anywhere; }
  .dr-ev .dot { margin-top:5px; flex:none; }
  .cm { width:100%; border-collapse:collapse; font-size:12px; }
  .cm td { padding:6px 4px; border-top:0; border-bottom:1px solid var(--bd); vertical-align:top; white-space:normal; }
  .cm code { font:11.5px ui-monospace,Consolas,"Courier New",monospace; overflow-wrap:anywhere; }
  .cm .w { color:var(--mu); font-size:11px; white-space:nowrap; }
  .cm td:first-child { width:88px; } .cm td:last-child { width:22px; text-align:right; }
  .cm .ok { color:var(--gr); font-weight:700; } .cm .no { color:var(--rd); font-weight:700; }
  .src { display:inline-block; padding:1px 7px; border-radius:999px; background:var(--acs); color:var(--ac); font-size:10.5px; font-weight:600; white-space:nowrap; }
  .chips { display:flex; flex-wrap:wrap; gap:5px; margin-bottom:8px; }
  .chips button { background:var(--pn2); border:1px solid var(--bd); color:var(--tx); border-radius:999px; padding:3px 10px; font:inherit; font-size:11.5px; cursor:pointer; }
  .chips button.on { border-color:var(--ac); color:var(--ac); font-weight:600; background:var(--acs); }
  .fl-row { display:flex; align-items:center; gap:8px; padding:6px 0; border-bottom:1px solid var(--bd); font-size:12px; }
  .fl-row .nm { flex:1; min-width:0; overflow-wrap:anywhere; }
  .fl-row .nm small { display:block; color:var(--mu); font-size:11px; }
  .fl-row .sz { color:var(--mu); font-size:11px; white-space:nowrap; }
  .dr-run { border:1px solid var(--bd); border-radius:8px; margin-bottom:8px; overflow:hidden; }
  .dr-run > .rh { display:flex; gap:8px; align-items:center; padding:7px 10px; background:var(--pn2); font-weight:600; font-size:12.5px; }
  .dr-run > .rh span.m { margin-left:auto; color:var(--mu); font-weight:400; font-size:11px; }
  .dr-run .fl-row { padding:6px 10px; }
  .dr-run .fl-row:last-child { border-bottom:0; }
  .hrow { padding:7px 0; border-bottom:1px solid var(--bd); font-size:12px; }
  .hrow .a { display:flex; align-items:center; gap:7px; flex-wrap:wrap; }
  .hrow .b { color:var(--mu); font-size:11.5px; margin-top:2px; overflow-wrap:anywhere; }
  .dr-view { position:absolute; inset:0; background:var(--pn); display:flex; flex-direction:column; z-index:3; }
  .dr-view[hidden] { display:none; }
  .dr-view .vh { display:flex; align-items:center; gap:8px; padding:10px 12px; border-bottom:1px solid var(--bd); font-weight:600; font-size:12.5px; }
  .dr-view .vh .nm { flex:1; min-width:0; overflow-wrap:anywhere; }
  .dr-view pre { margin:0; padding:10px 12px 56px; flex:1; min-height:0; overflow:auto; font:11.5px/1.45 ui-monospace,Consolas,"Courier New",monospace; background:var(--pn2); white-space:pre; }
  .dr-view .vn { padding:6px 12px; font-size:11.5px; color:var(--c-a); background:var(--t-a); }

  /* ---- device drawer > Console (interactive SSH) ---- */
  .dr-term { flex:1; min-height:0; display:flex; flex-direction:column; margin:10px 14px 14px; background:#0b1020; color:#d6e2ff;
             border:1px solid #2a3558; border-radius:10px; overflow:hidden; }
  .dr-term[hidden] { display:none; }
  .dr-term .th { display:flex; align-items:center; gap:6px; padding:6px 8px 6px 11px; background:#151c33; flex:none; flex-wrap:wrap; }
  .dr-term .tt { font-weight:600; font-size:12px; white-space:nowrap; display:none; }
  .dr-term.max .tt { display:inline; }
  .dr-term .tst { font-size:11px; padding:1px 8px; border-radius:999px; background:#2a3558; color:#d6e2ff; white-space:nowrap; }
  .dr-term .tst.on { background:#17382a; color:#4ade80; }
  .dr-term .tst.busy { background:#3b3320; color:#f4c542; }
  .dr-term .sp { flex:1; }
  .dr-term .tb { background:transparent; border:1px solid #3b4a7a; color:#d6e2ff; border-radius:6px; padding:3px 9px; font:inherit; font-size:11.5px; cursor:pointer; }
  .dr-term .tb:hover:not(:disabled) { background:rgba(255,255,255,.12); }
  .dr-term .tb:disabled { opacity:.4; cursor:not-allowed; }
  .dr-term .tb.go { background:#2f6fe4; border-color:#2f6fe4; color:#fff; font-weight:600; }
  .dr-term .tout { flex:1; min-height:0; margin:0; padding:8px 11px; overflow:auto; white-space:pre-wrap; word-break:break-word;
                   font:12px/1.45 ui-monospace,Consolas,"Courier New",monospace; }
  .dr-term .tin { display:flex; align-items:center; gap:7px; padding:6px 11px; border-top:1px solid #2a3558; background:#0e1429; flex:none; }
  .dr-term .tin span { color:#4ade80; font:12px ui-monospace,Consolas,monospace; }
  .dr-term .tin input { flex:1; min-width:0; background:transparent; border:0; outline:0; color:#f3f6ff; caret-color:#4ade80;
                        font:12.5px ui-monospace,Consolas,"Courier New",monospace; padding:3px 0; }
  .dr-term .tin input:disabled { opacity:.45; }
  .dr-term .tn { padding:5px 11px; font-size:10.5px; color:#93a1c0; border-top:1px solid #2a3558; flex:none; }
  .dr-term.max { position:fixed; inset:28px; margin:0; z-index:1150; box-shadow:0 14px 60px rgba(0,0,0,.55); }
  .dr-term.max .tout { font-size:13.5px; }
  #dr-term-bd { position:fixed; inset:0; background:rgba(0,0,0,.55); z-index:1140; }
  #dr-term-bd[hidden] { display:none; }


  .dr-term .tin[hidden], .dr-term .tb[hidden] { display:none; }
  .dr-term .tn { white-space:pre-wrap; word-break:break-word; }

  /* ---- device list: search + filters ---- */
  .dev-tools { display:flex; flex-wrap:wrap; align-items:center; gap:8px; padding:8px 12px; border-bottom:1px solid var(--bd); }
  .dev-tools .st-search { flex:1 1 220px; min-width:180px; }
  .dev-tools select { width:auto; max-width:170px; padding:5px 8px; font-size:12px; }
  .dev-tools .dev-shown { font-size:11.5px; color:var(--mu); white-space:nowrap; }

  /* ---- completion pop-ups ---- */
  #toasts { position:fixed; top:66px; right:14px; z-index:1200; display:flex; flex-direction:column; gap:8px;
            width:min(370px, calc(100vw - 28px)); pointer-events:none; }
  .toast { pointer-events:auto; cursor:pointer; background:var(--pn); color:var(--tx); border:1px solid var(--bd);
           border-left:5px solid var(--gr); border-radius:10px; padding:10px 34px 10px 12px; position:relative;
           box-shadow:0 10px 30px rgba(0,0,0,.28); animation:toast-in .25s ease-out; }
  .toast.warn { border-left-color:var(--am); } .toast.bad { border-left-color:var(--rd); }
  .toast .tt { font-weight:700; } .toast .tmsg { color:var(--mu); font-size:12px; margin-top:2px; word-break:break-word; }
  .toast .tx { position:absolute; top:6px; right:6px; border:0; background:transparent; color:var(--mu);
               cursor:pointer; font-size:14px; width:22px; height:22px; border-radius:6px; }
  .toast .tx:hover { background:var(--t-g); }
  @keyframes toast-in { from { transform:translateX(40px); opacity:0; } }
  @media (prefers-reduced-motion: reduce) { .toast { animation:none; } }

  /* ---- Home / Storage tabs (header) ---- */
  .app[hidden], .btn[hidden], .chk[hidden] { display:none; }
  .tabs { display:flex; align-self:stretch; gap:2px; margin-left:8px; }
  .tab { background:none; border:0; border-bottom:3px solid transparent; border-top:3px solid transparent;
         color:var(--mu); font:inherit; font-size:14px; font-weight:500; padding:0 16px; cursor:pointer;
         display:flex; align-items:center; gap:8px; white-space:nowrap; }
  .tab:hover { color:var(--tx); background:var(--acs); }
  .tab.on { color:var(--tx); font-weight:600; border-bottom-color:var(--ac); }

  /* ---- Diagram page ---- */
  .dg-side { width:300px; }
  .dg-sec { padding:0 14px 14px; }
  .dg-actions { position:sticky; top:0; z-index:3; display:flex; gap:8px; margin:8px -14px 0; padding:8px 14px; background:var(--pn2); border-bottom:1px solid var(--bd); }
  .dg-actions .btn { flex:1; min-width:0; justify-content:center; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  #diagram main { overflow:hidden; }
  .dg-list { display:flex; flex-direction:column; gap:2px; margin:10px 0 4px; }
  .dg-dev { display:flex; align-items:flex-start; gap:9px; padding:7px 8px; border-radius:8px; cursor:pointer; margin:0; font-size:13px; color:var(--tx); }
  .dg-dev:hover { background:var(--acs); }
  .dg-dev.off { opacity:.62; cursor:not-allowed; }
  .dg-dev input { flex:none; margin-top:3px; }
  .dg-dev .meta { min-width:0; flex:1; }
  .dg-dev .t1 { font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .dg-dev .t2 { font-size:11.5px; color:var(--mu); word-break:break-word; }
  .dg-dev .t2.bad { color:var(--rd); } .dg-dev .t2.ok { color:var(--gr); } .dg-dev .t2.warn { color:var(--c-a); }
  .dg-card { flex:1; min-height:0; }
  .dg-banner { margin:10px 14px 0; padding:9px 12px; border-radius:8px; background:var(--t-a); color:var(--c-a); border:1px solid var(--c-a); font-size:13px; font-weight:600; }
  .dg-banner[hidden] { display:none; }
  .dg-wrap { position:relative; flex:1; min-height:380px; overflow:hidden; background:var(--pn); cursor:grab; touch-action:none; user-select:none; -webkit-user-select:none; }
  .dg-wrap.panning { cursor:grabbing; }
  #dg-svg { position:absolute; inset:0; width:100%; height:100%; display:block; }
  .dg-ico { display:inline-grid; place-items:center; padding:5px 8px; }
  .dg-ico .i { width:15px; height:15px; }
  .dg-hint { position:absolute; left:12px; bottom:10px; font-size:11px; color:var(--mu); background:var(--pn); border:1px solid var(--bd);
             border-radius:999px; padding:2px 10px; pointer-events:none; opacity:.9; z-index:1; }
  .dg-hint:empty { display:none; }
  /* Full screen: the whole Topology card (toolbar, canvas, legend, details panel) fills the screen */
  .dg-card:fullscreen { width:100vw; height:100vh; border-radius:0; border:0; background:var(--pn); }
  .dg-card:fullscreen .dg-wrap { min-height:0; }
  .dg-card.dg-full { position:fixed; inset:0; z-index:500; border-radius:0; border:0; background:var(--pn); }
  .dg-card.dg-full .dg-wrap { min-height:0; }
  .dg-empty { position:absolute; inset:0; display:flex; align-items:center; justify-content:center; text-align:center; padding:24px; color:var(--mu); font-size:14px; pointer-events:none; }
  .dg-empty:empty { display:none; }
  .dg-legend { display:flex; flex-wrap:wrap; gap:6px 16px; padding:8px 14px; border-top:1px solid var(--bd); font-size:11.5px; color:var(--mu); align-items:center; }
  .dg-legend .dot { margin-right:5px; }
  .dg-legend .lg-run, .dg-legend .lg-done, .dg-legend .lg-fail, .dg-legend .lg-q, .dg-legend .lg-hold { padding:1px 8px; border-radius:999px; font-weight:600; }
  .dg-legend .lg-run { background:var(--t-a); color:var(--c-a); } .dg-legend .lg-done { background:var(--t-gn); color:var(--gr); }
  .dg-legend .lg-fail { background:var(--t-r); color:var(--c-r); } .dg-legend .lg-q { background:var(--t-b); color:var(--c-b); }
  .dg-legend .lg-hold { border:1px dashed var(--am); color:var(--c-a); }
  .dg-n { cursor:grab; outline:none; }
  .dg-n:active { cursor:grabbing; }
  .dg-n .bx { fill:var(--pn2); stroke:var(--bd); stroke-width:2; }
  .dg-n.s-UP .bx { stroke:var(--gr); } .dg-n.s-DOWN .bx { stroke:var(--rd); } .dg-n.s-PENDING .bx { stroke:var(--mu); }
  .dg-n.ext .bx { stroke:var(--mu); stroke-dasharray:5 4; }
  .dg-n .sd { fill:var(--mu); } .dg-n.s-UP .sd { fill:var(--gr); } .dg-n.s-DOWN .sd { fill:var(--rd); }
  .dg-n text { fill:var(--tx); font-size:12px; pointer-events:none; }
  .dg-n .nm { font-weight:700; }
  .dg-n .ipt { fill:var(--mu); font-size:10.5px; }
  .dg-n .stt { font-size:10.5px; font-weight:700; fill:var(--mu); }
  .dg-n.s-UP .stt { fill:var(--gr); } .dg-n.s-DOWN .stt { fill:var(--rd); }
  .dg-n .strip rect { fill:var(--t-g); }
  .dg-n .strip text { font-size:10.5px; font-weight:600; fill:var(--mu); }
  .dg-n.u-none .strip { display:none; }
  .dg-n.u-running .strip rect, .dg-n.u-attention .strip rect { fill:var(--t-a); }
  .dg-n.u-running .strip text, .dg-n.u-attention .strip text { fill:var(--c-a); }
  .dg-n.u-completed .strip rect { fill:var(--t-gn); } .dg-n.u-completed .strip text { fill:var(--gr); }
  .dg-n.u-failed .strip rect { fill:var(--t-r); } .dg-n.u-failed .strip text { fill:var(--c-r); }
  .dg-n.u-queued .strip rect { fill:var(--t-b); } .dg-n.u-queued .strip text { fill:var(--c-b); }
  .dg-n.u-running .bx { stroke-width:3; animation:pulse 1.4s ease-in-out infinite; }
  .dg-n.hold .bx { stroke:var(--am); stroke-dasharray:6 4; stroke-width:3; }
  .dg-n.u-hold .strip rect { fill:var(--t-a); stroke:var(--am); stroke-dasharray:4 3; } .dg-n.u-hold .strip text { fill:var(--c-a); }
  .dg-n.sel .bx { filter:drop-shadow(0 0 5px var(--ac)); }
  .dg-n:focus-visible .bx { stroke:var(--ac); }
  .dg-e { fill:none; stroke:var(--mu); opacity:.85; }
  .dg-e.down { stroke:var(--rd); stroke-dasharray:6 5; }
  .dg-e.upg { stroke:var(--am); }
  .dg-e.hold { stroke:var(--am); stroke-dasharray:6 4; }
  .dg-lbl { fill:var(--mu); font-size:9.5px; paint-order:stroke; stroke:var(--pn); stroke-width:4px; stroke-linejoin:round; pointer-events:none; }
  .dg-cap { fill:var(--mu); font-size:12px; font-weight:600; }
  .dg-detail { position:absolute; top:10px; right:10px; width:290px; max-height:calc(100% - 20px); overflow:auto; background:var(--pn);
               border:1px solid var(--bd); border-radius:10px; box-shadow:0 8px 28px rgba(0,0,0,.28); padding:12px 14px; font-size:13px; z-index:2; }
  .dg-detail[hidden] { display:none; }
  .dg-detail h5 { margin:0 0 8px; font-size:14px; display:flex; align-items:center; gap:8px; }
  .dg-detail h5 span { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .dg-detail h5 button { background:none; border:0; color:var(--mu); font-size:18px; cursor:pointer; line-height:1; }
  .dg-detail dl { display:grid; grid-template-columns:auto 1fr; gap:4px 10px; margin:0 0 10px; }
  .dg-detail dt { color:var(--mu); } .dg-detail dd { margin:0; word-break:break-word; }
  .dg-detail .nbs { margin:0 0 10px; padding:0; list-style:none; font-size:12px; }
  .dg-detail .nbs li { padding:4px 0; border-top:1px solid var(--bd); }
  .dg-detail .nbs .pt { color:var(--mu); }
  @media (prefers-reduced-motion: reduce) { .dg-n.u-running .bx { animation:none; } }

  /* ---- Storage page ---- */
  .st-side { width:266px; }
  .st-side .sh { font-size:15px; padding:14px 14px 8px; }
  .st-sites { padding:0 8px 12px; display:flex; flex-direction:column; gap:2px; }
  .st-site { display:flex; align-items:center; gap:10px; width:100%; text-align:left; background:none; border:0;
             color:var(--mu); font:inherit; font-size:14px; padding:10px 12px; border-radius:8px; cursor:pointer; }
  .st-site:hover { background:var(--acs); color:var(--tx); }
  .st-site.on { background:var(--sel); color:var(--tx); font-weight:600; box-shadow:3px 0 0 var(--ac) inset; }
  .st-site .nm { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .st-site .tag { font-size:10px; font-weight:700; padding:1px 7px; border-radius:999px; background:var(--t-gn);
                  color:var(--gr); }
  .st-empty { padding:10px 14px; color:var(--mu); font-size:12px; }
  #storage main { overflow:hidden; }
  .st-card { flex:1; }
  .st-bar { display:flex; align-items:center; gap:8px; padding:10px 14px; border-bottom:1px solid var(--bd);
            flex-wrap:wrap; }
  .crumbs { display:flex; align-items:center; flex-wrap:wrap; gap:2px; font-size:13px; min-width:0; }
  .crumbs button { background:none; border:0; color:var(--ac); font:inherit; cursor:pointer; padding:3px 6px;
                   border-radius:6px; max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .crumbs button:hover { background:var(--acs); }
  .crumbs button.cur { color:var(--tx); font-weight:600; cursor:default; }
  .crumbs button.cur:hover { background:none; }
  .crumbs .sepc { color:var(--mu); }
  .st-search { position:relative; width:min(340px, 100%); display:flex; align-items:center; }
  .st-search [data-i], .st-search .si { position:absolute; left:9px; color:var(--mu); pointer-events:none; }
  .st-search input { padding-left:32px; padding-right:28px; }
  .st-clear { position:absolute; right:4px; background:none; border:0; color:var(--mu); cursor:pointer; font-size:16px;
              width:24px; height:24px; border-radius:6px; line-height:1; }
  .st-clear:hover { background:var(--t-g); }
  .st-clear[hidden] { display:none; }
  .st-table { min-width:640px; }
  .st-table th.sortable { cursor:pointer; user-select:none; }
  .st-table th.sortable:hover { color:var(--tx); }
  #dev-table th[data-sort] { cursor:pointer; user-select:none; }
  #dev-table th[data-sort]:hover { color:var(--tx); }
  #dev-table th[data-sort].sorted { color:var(--tx); }
  .st-table td { white-space:nowrap; }
  .st-table td.nm { width:100%; max-width:0; overflow:hidden; text-overflow:ellipsis; }
  .st-table td.act { text-align:right; }
  .st-table td.act .btn { margin-left:4px; }
  .st-table .ico { display:inline-flex; vertical-align:middle; margin-right:8px; color:var(--mu); }
  .st-table .ico.dir { color:var(--ac); }
  .st-name { background:none; border:0; font:inherit; color:var(--tx); cursor:pointer; padding:0; text-align:left;
             max-width:100%; overflow:hidden; text-overflow:ellipsis; vertical-align:middle; }
  .st-name:hover { color:var(--ac); text-decoration:underline; }
  .st-loc { color:var(--mu); font-size:12px; }
  .st-foot { padding:7px 14px; border-top:1px solid var(--bd); color:var(--mu); font-size:12px; flex:none; }
  .st-none { padding:34px 14px; text-align:center; color:var(--mu); }
  .modal-card.wide { max-width:min(1200px, 100%); height:calc(100vh - 32px); max-height:calc(100vh - 32px); display:flex; flex-direction:column; min-height:0; }
  #st-pre { margin:8px 0 0; padding:10px 12px 48px; overflow:auto; flex:1; min-height:0; background:var(--pn2);
            border:1px solid var(--bd); border-radius:8px; white-space:pre; font:12px/1.5 ui-monospace,Consolas,monospace; }

  @media (max-width:1000px) {
    .wg { grid-template-columns:1fr; }
    .side.wide, .st-side { width:250px; }
    .fbar #meta { display:none; }
  }
  .repbtns { display:inline-flex; gap:6px; }
  /* ---- signed-in user menu ---- */
  .usermenu { position:relative; display:inline-block; margin-left:4px; }
  .userbtn { display:flex; align-items:center; gap:7px; height:36px; padding:0 10px; border:0; border-radius:18px;
             background:none; color:var(--tx); font:inherit; cursor:pointer; }
  .userbtn:hover, .userbtn[aria-expanded="true"] { background:rgba(128,128,128,.18); }
  .userbtn:focus-visible { outline:2px solid var(--ac); outline-offset:1px; }
  .userbtn svg { width:20px; height:20px; }
  .userbtn .uname { font-size:13px; font-weight:600; max-width:140px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .user-pop { position:absolute; right:0; top:calc(100% + 8px); z-index:1300; min-width:230px; background:var(--pn);
              color:var(--tx); border:1px solid var(--bd); border-radius:10px; box-shadow:0 10px 30px rgba(0,0,0,.25);
              font-size:13px; overflow:hidden; }
  .user-pop[hidden] { display:none; }
  .user-pop .uh { padding:10px 12px; border-bottom:1px solid var(--bd); }
  .user-pop .uh b { display:block; font-size:14px; }
  .user-pop .uh span { display:block; color:var(--mu); font-size:12px; margin-top:2px; }
  .user-pop a { display:block; padding:10px 12px; color:var(--tx); text-decoration:none; font-weight:600; }
  .user-pop a:hover, .user-pop a:focus-visible { background:rgba(128,128,128,.18); }
  .user-pop a[hidden] { display:none; }
  @media (max-width:900px) { .userbtn .uname { display:none; } }
  /* ---- roles: the server enforces them; this only hides what a role cannot use ---- */
  body.role-viewer #side, body.role-viewer #storage, body.role-viewer #tab-storage, body.role-viewer #diagram, body.role-viewer #tab-diagram, body.role-viewer #btn-mon,
  body.role-viewer #sel-count, body.role-viewer #btn-version, body.role-viewer #btn-delete-devices,
  body.role-viewer .ck, body.role-viewer #note-clear { display:none !important; }
  body.role-viewer #site-badge { pointer-events:none; cursor:default; }
  body.role-operator #btn-mon[data-action="stop"] { display:none !important; }
  /* ---- credit footer ---- */
  footer.credit { flex:none; display:flex; align-items:center; justify-content:center; gap:10px; padding:7px 16px;
                  background:var(--tb); border-top:1px solid var(--tbd); font-size:12px; color:var(--mu); }
  footer.credit strong { color:var(--tx); font-weight:600; }
  footer.credit .sep { opacity:.45; }
  footer.credit a { display:inline-flex; align-items:center; gap:6px; color:var(--ac); font-weight:600;
                    text-decoration:none; }
  footer.credit a:hover, footer.credit a:focus-visible { text-decoration:underline; }
  footer.credit svg { width:14px; height:14px; fill:currentColor; }

  /* ---- v-next: settings page, events collapse, image list, hover tips ---- */
  .ib[hidden], .card[hidden] { display:none; }
  .ib.on { background:var(--acs); color:var(--ac); }
  input[type=number] { background:var(--pn); border:1px solid var(--bd); color:var(--tx); border-radius:6px; padding:6px 8px; font:inherit; }
  .set-nav { display:flex; flex-direction:column; gap:2px; padding:4px 8px 12px; }
  .set-nav button { text-align:left; background:none; border:0; border-radius:8px; padding:10px 12px; font:inherit;
                    font-size:14px; color:var(--mu); cursor:pointer; }
  .set-nav button:hover { background:var(--acs); color:var(--tx); }
  .set-nav button.on { background:var(--sel); color:var(--tx); font-weight:600; box-shadow:3px 0 0 var(--ac) inset; }
  .set-card { flex:none; }
  .set-body { padding:2px 18px 16px; }
  .set-intro { color:var(--mu); font-size:12.5px; margin:12px 0 2px; max-width:780px; line-height:1.5; }
  .ab-card { border:1px solid var(--bd); border-radius:10px; padding:12px 14px; margin-top:12px; max-width:780px; }
  .ab-kv { display:grid; grid-template-columns:150px minmax(0,1fr); gap:6px 12px; font-size:13px; }
  .ab-kv span:nth-child(odd) { color:var(--mu); }
  .ab-ver { margin-top:16px; max-width:780px; border-top:1px solid var(--bd); padding-top:10px; }
  .ab-ver h4 { margin:0 0 4px; font-size:14px; display:flex; gap:8px; align-items:center; }
  .ab-ver small { color:var(--mu); font-weight:400; }
  .ab-cur { font-size:10.5px; font-weight:700; letter-spacing:.04em; text-transform:uppercase; padding:1px 8px; border-radius:999px; background:var(--sel); color:var(--ac); }
  .ab-ver ul { margin:4px 0 0; padding-left:20px; font-size:12.5px; line-height:1.55; }
  .set-sec { margin:20px 0 2px; font-size:11px; letter-spacing:.06em; text-transform:uppercase; color:var(--mu); font-weight:700; }
  .set-row { display:flex; gap:24px; align-items:center; justify-content:space-between; padding:11px 0; border-top:1px solid var(--bd); }
  .set-sec + .set-row { margin-top:4px; }
  .set-info { flex:1; min-width:0; }
  .set-label { font-weight:600; font-size:13px; }
  .set-row.chg .set-label::after { content:"  \2022 changed"; color:var(--am); font-size:11px; font-weight:600; }
  .set-help { color:var(--mu); font-size:12px; margin-top:2px; max-width:600px; line-height:1.45; }
  .set-ctl { flex:none; display:flex; align-items:center; gap:8px; width:360px; justify-content:flex-end; }
  .set-ctl input[type=number] { width:92px; text-align:right; }
  .set-unit { color:var(--mu); font-size:12px; min-width:18px; }
  .set-hint { color:var(--mu); font-size:11px; width:64px; white-space:nowrap; }
  .set-def { background:none; border:1px dashed var(--bd); color:var(--mu); border-radius:6px; font:inherit; font-size:11px;
             padding:3px 8px; cursor:pointer; white-space:nowrap; }
  .set-def:hover:not(:disabled) { border-color:var(--ac); color:var(--ac); }
  .set-def:disabled { cursor:default; }
  .set-chk { margin:0; }
  .set-note { padding:8px 18px 0; color:var(--mu); font-size:12px; }
  .nt-form { display:grid; grid-template-columns:180px minmax(220px,420px); gap:10px 14px; align-items:center; max-width:720px; margin-top:16px; }
  .nt-form label { margin:0; font-size:13px; color:var(--tx); font-weight:600; }
  .nt-form .nt-inline { display:flex; gap:8px; align-items:center; }
  .nt-form .nt-inline input[type=text] { width:90px; }
  .nt-form .nt-inline select { flex:1; width:auto; }
  .nt-form .nt-hint { grid-column:2; color:var(--mu); font-size:11.5px; margin-top:-6px; }
  .nt-kinds { display:flex; flex-wrap:wrap; gap:6px 18px; margin:6px 0 8px; max-width:720px; }
  .nt-kinds label { display:flex; align-items:center; gap:6px; margin:0; font-size:13px; color:var(--tx); }
  .nt-actions { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-top:14px; }
  .nt-last { color:var(--mu); font-size:12px; margin-top:10px; }
  .nt-last.ok { color:var(--gr); } .nt-last.bad { color:var(--rd); }
  .card.e { cursor:default; }
  .card.e > h4 { cursor:pointer; user-select:none; flex-wrap:nowrap; }
  .card.e.collapsed { height:auto; }
  .card.e.collapsed > h4 { border-bottom:0; }
  .card.e.collapsed .ev { display:none; }
  .evn { font-size:11px; font-weight:600; color:var(--mu); background:var(--pn2); border:1px solid var(--bd); border-radius:999px; padding:0 8px; }
  .evn:empty { display:none; }
  .evlast { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:12px; font-weight:400; color:var(--mu); }
  .card.e:not(.collapsed) .evlast { display:none; }
  .img-card { flex:none; }
  .img-body { padding:10px 14px 12px; max-height:300px; overflow:auto; }
  .img-form { display:grid; grid-template-columns:minmax(200px,1fr) minmax(260px,1.2fr) auto; gap:12px; align-items:end; }
  .img-form label { margin-top:0; }
  .img-table { min-width:0; margin-top:10px; }
  .img-table td, .img-table th { padding:7px 8px; }

  /* ---- command duration profiles + command time test (Settings > Command Outputs) ---- */
  .tp-bar, .ct-card { border:1px solid var(--bd); border-radius:10px; background:var(--pn2); margin:14px 0 4px; padding:12px 14px; max-width:1280px; }
  .tp-row { display:flex; flex-wrap:wrap; gap:8px; align-items:center; }
  .tp-row + .tp-row { margin-top:10px; }
  .tp-row > label, .ct-row > label { font-weight:600; font-size:13px; }
  .tp-bar select, .ct-card select, .tp-bar input[type=text], .ct-card input[type=text], .ct-card input[type=number] {
    background:var(--pn); border:1px solid var(--bd); color:var(--tx); border-radius:6px; padding:6px 8px; font:inherit; font-size:13px; max-width:100%; }
  .tp-bar select { width:auto; min-width:260px; flex:0 1 380px; }
  .tp-row input[type=text] { width:auto; flex:1 1 200px; max-width:380px; }
  .tp-row input[type=text] + input[type=text] { margin-top:0; }
  .ct-actions input[type=text] { width:auto; flex:1 1 220px; max-width:360px; }
  .tp-state { font-size:12px; color:var(--mu); }
  .tp-state b { color:var(--tx); }
  .tp-info { font-size:12px; color:var(--mu); margin-top:8px; line-height:1.5; }
  .tp-vals { margin-top:6px; font-size:12px; color:var(--mu); }
  .tp-vals summary { cursor:pointer; }
  .tp-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(290px, 1fr)); gap:2px 18px; margin-top:6px; }
  .tp-grid span:nth-child(odd) { color:var(--tx); }
  /* ---- command time test: step list | step detail + console | results ---- */
  .ct-toggle { display:flex; align-items:center; gap:8px; width:100%; background:none; border:0; color:var(--tx); font:inherit; font-weight:600; font-size:14px; cursor:pointer; padding:0; text-align:left; }
  .ct-toggle .sub { font-weight:400; font-size:12px; color:var(--mu); }
  .ct-toggle .car { margin-left:auto; transition:transform .15s; display:inline-grid; }
  .ct-card.open .ct-toggle .car { transform:rotate(180deg); }
  .ct-bar { display:flex; flex-wrap:wrap; gap:10px 14px; align-items:flex-end; margin-top:14px; }
  .ct-f { flex:1 1 260px; min-width:0; }
  .ct-f > label { display:block; font-weight:600; font-size:12px; color:var(--mu); margin-bottom:4px; }
  .ct-f select { width:100%; }
  .ct-bar-r { display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-left:auto; }
  .ct-wb { display:grid; grid-template-columns:292px minmax(0,1fr); gap:16px; margin-top:14px; align-items:start; }
  .ct-rail { border:1px solid var(--bd); border-radius:10px; background:var(--pn); padding:6px; }
  .ct-rs + .ct-rs { margin-top:4px; border-top:1px solid var(--bd); padding-top:4px; }
  .ct-rs-h { display:flex; align-items:center; gap:7px; width:100%; background:none; border:0; color:var(--mu); font:inherit; font-size:11px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; padding:8px 8px 5px; cursor:pointer; text-align:left; }
  .ct-rs-h:hover { color:var(--tx); }
  .ct-rs-c { margin-left:auto; font-weight:600; letter-spacing:0; }
  .ct-car { display:inline-grid; transition:transform .15s; transform:rotate(-90deg); }
  .ct-rs.open .ct-car { transform:none; }
  .ct-r { display:flex; align-items:center; gap:9px; width:100%; background:none; border:0; border-radius:7px; padding:7px 8px; font:inherit; font-size:13px; color:var(--tx); cursor:pointer; text-align:left; }
  .ct-r:hover { background:var(--pn2); }
  .ct-rs.off .ct-r, .ct-rs.off .ct-rs-h { opacity:.5; }
  .ct-r.sel { background:var(--acs); box-shadow:inset 3px 0 0 var(--ac); font-weight:600; }
  .ct-r-t { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .ct-r-v { font-size:11.5px; font-weight:400; color:var(--mu); font-variant-numeric:tabular-nums; white-space:nowrap; }
  .ct-n { flex:none; width:22px; height:22px; border-radius:50%; background:var(--pn2); border:1px solid var(--bd); display:grid; place-items:center; font-size:11px; font-weight:700; color:var(--mu); }
  .ct-r.done .ct-n { background:var(--gr); border-color:var(--gr); color:#fff; }
  .ct-r.bad .ct-n { background:var(--rd); border-color:var(--rd); color:#fff; }
  .ct-r.run .ct-n { background:var(--ac); border-color:var(--ac); color:#fff; }
  .ct-main { min-width:0; display:flex; flex-direction:column; gap:14px; }
  .ct-d { border:1px solid var(--bd); border-radius:10px; background:var(--pn); padding:14px 16px; }
  .ct-d-h { display:flex; flex-wrap:wrap; gap:12px 16px; align-items:flex-start; }
  .ct-d-id { flex:1 1 280px; min-width:0; }
  .ct-d-k { font-size:11px; font-weight:700; letter-spacing:.06em; text-transform:uppercase; color:var(--mu); }
  .ct-d-t { font-size:17px; font-weight:700; display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-top:2px; }
  .ct-cmd { display:block; margin-top:5px; font-size:12.5px; color:var(--mu); overflow-wrap:anywhere; }
  .ct-d-a { display:flex; flex-wrap:wrap; gap:8px; margin-left:auto; }
  .ct-d-a .btn { padding:8px 18px; font-size:13px; }
  .ct-d-desc { margin:10px 0 0; font-size:13px; line-height:1.55; }
  .ct-chips { display:flex; flex-wrap:wrap; gap:6px; align-items:center; margin-top:10px; font-size:12px; color:var(--mu); }
  .ct-chip { padding:2px 9px; border-radius:999px; background:var(--pn2); border:1px solid var(--bd); color:var(--tx); font-size:11.5px; }
  .ct-tiles { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-top:12px; }
  .ct-tile { border:1px solid var(--bd); border-radius:8px; padding:8px 11px; background:var(--pn2); }
  .ct-tile span { display:block; font-size:10.5px; letter-spacing:.05em; text-transform:uppercase; color:var(--mu); }
  .ct-tile b { font-size:18px; font-variant-numeric:tabular-nums; }
  .ct-status { margin-top:10px; font-size:12.5px; color:var(--mu); }
  .ct-status .dot { margin-right:5px; }
  .ct-badge { font-size:10.5px; font-weight:600; padding:1px 8px; border-radius:999px; background:var(--t-a); color:var(--c-a); }
  .ct-badge.d { background:var(--t-r); color:var(--c-r); }
  .ct-badge.n { background:var(--t-g); color:var(--mu); }
  .ct-panel { margin-top:12px; border:1px solid var(--bd); border-radius:8px; background:var(--pn2); }
  .ct-panel-h { padding:7px 10px; font-size:12px; border-bottom:1px solid var(--bd); }
  .ct-panel pre { margin:0; padding:8px 10px; max-height:240px; overflow:auto; font-size:11.5px; line-height:1.4; white-space:pre; }
  .ct-panel table { width:100%; border-collapse:collapse; font-size:12.5px; min-width:0; }
  .ct-panel th { text-align:left; font-size:11px; letter-spacing:.04em; text-transform:uppercase; color:var(--mu); padding:5px 10px; }
  .ct-panel td { padding:5px 10px; border-top:1px solid var(--bd); }
  .ct-panel td.n, .ct-panel th.n, .ct-table td.n, .ct-table th.n { text-align:right; font-variant-numeric:tabular-nums; }
  .ct-empty { margin-top:12px; padding:14px; border:1px dashed var(--bd); border-radius:8px; font-size:12.5px; color:var(--mu); text-align:center; }
  .ct-result { margin-top:16px; border-top:1px solid var(--bd); padding-top:14px; }
  .ct-result-t { font-weight:700; font-size:14px; }
  .ct-result-g { display:grid; grid-template-columns:minmax(0,1.5fr) minmax(280px,1fr); gap:18px; align-items:start; margin-top:8px; }
  .ct-result-g > * { min-width:0; }
  #ct-sugg-t { overflow-x:auto; }
  .ct-sg-h label { font-size:13px; font-weight:600; color:var(--tx); display:inline-flex; align-items:center; gap:6px; }
  .ct-side { border:1px solid var(--bd); border-radius:10px; background:var(--pn); padding:12px 14px; }
  .ct-sg-h { display:flex; flex-wrap:wrap; gap:12px; align-items:center; }
  .ct-sg-h input[type=number] { width:72px; }
  .ct-table { width:100%; border-collapse:collapse; font-size:12.5px; min-width:0; }
  .ct-table td, .ct-table th, .ct-panel td, .ct-panel th { white-space:normal; }
  .ct-table th { text-align:left; font-size:11px; letter-spacing:.04em; text-transform:uppercase; color:var(--mu); padding:5px 8px; border-bottom:1px solid var(--bd); }
  .ct-table td { padding:6px 8px; border-bottom:1px solid var(--bd); }
  .ct-actions { display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-top:10px; }
  .ct-note { font-size:12px; color:var(--mu); margin-top:8px; line-height:1.5; }
  .btn.dg { border-color:var(--rd); color:var(--rd); }
  /* Colour-coded function buttons: green = run / save / add, blue = check / read / refresh, amber = schedule / reset,
     red = delete / cancel / stop. They reuse the theme colours, so light and dark mode both match. */
  .btn.ok, .btn.in, .btn.wn, .btn.dg, #btn-backup, #btn-upgrade, #btn-add-device, #btn-upload-devices, #btn-prof-save,
  #btn-upload-image, #btn-site-create, #set-save, #nt-save, #rep-go, [data-tp="create"], [data-ct="run"], [data-ct="save"],
  [data-ct="saveform"], #ck-save, #btn-prof-assign, #btn-prof-test, #btn-readiness, #btn-version, #dg-run, #st-refresh,
  #img-refresh, #dg-refresh, [data-verify], [data-tp="apply"], #cmp-go, #st-compare, #nt-test, #btn-schedule, #set-reset,
  [data-tp="update"], #btn-delete-devices, #btn-prof-delete, #btn-unschedule, [data-tp="delete"], [data-ct="stop"],
  #ck-add-pre, #ck-add-post { font-weight:600; }
  .btn.ok, #btn-backup, #btn-upgrade, #btn-add-device, #btn-upload-devices, #btn-prof-save, #btn-upload-image,
  #btn-site-create, #set-save, #nt-save, #rep-go, [data-tp="create"], [data-ct="run"], [data-ct="save"],
  [data-ct="saveform"], #ck-save, #ck-add-pre, #ck-add-post { background:var(--t-gn); border-color:var(--gr); color:var(--gr); }
  .btn.in, #btn-prof-assign, #btn-prof-test, #btn-readiness, #btn-version, #dg-run, #st-refresh, #img-refresh, #dg-refresh,
  [data-verify], [data-tp="apply"], #cmp-go, #st-compare, #nt-test { background:var(--acs); border-color:var(--ac); color:var(--ac); }
  .btn.wn, #btn-schedule, #set-reset, [data-tp="update"] { background:var(--t-a); border-color:var(--am); color:var(--c-a); }
  .btn.dg, #btn-delete-devices, #btn-prof-delete, #btn-unschedule, [data-tp="delete"], [data-ct="stop"] { background:var(--t-r); border-color:var(--rd); color:var(--rd); }
  .btn.ok:hover:not(:disabled), #btn-backup:hover:not(:disabled), #btn-upgrade:hover:not(:disabled), #btn-add-device:hover:not(:disabled),
  #btn-upload-devices:hover:not(:disabled), #btn-prof-save:hover:not(:disabled), #btn-upload-image:hover:not(:disabled),
  #btn-site-create:hover:not(:disabled), #set-save:hover:not(:disabled), #nt-save:hover:not(:disabled), #rep-go:hover:not(:disabled),
  [data-tp="create"]:hover:not(:disabled), [data-ct="run"]:hover:not(:disabled), [data-ct="save"]:hover:not(:disabled),
  [data-ct="saveform"]:hover:not(:disabled), #ck-save:hover:not(:disabled), #ck-add-pre:hover:not(:disabled), #ck-add-post:hover:not(:disabled)
    { border-color:var(--gr); box-shadow:0 0 0 3px color-mix(in srgb, var(--gr) 22%, transparent); }
  .btn.in:hover:not(:disabled), #btn-prof-assign:hover:not(:disabled), #btn-prof-test:hover:not(:disabled), #btn-readiness:hover:not(:disabled),
  #btn-version:hover:not(:disabled), #dg-run:hover:not(:disabled), #st-refresh:hover:not(:disabled), #img-refresh:hover:not(:disabled),
  #dg-refresh:hover:not(:disabled), [data-verify]:hover:not(:disabled), [data-tp="apply"]:hover:not(:disabled),
  #cmp-go, #st-compare:hover:not(:disabled), #nt-test:hover:not(:disabled) { box-shadow:0 0 0 3px color-mix(in srgb, var(--ac) 22%, transparent); }
  .btn.wn:hover:not(:disabled), #btn-schedule:hover:not(:disabled), #set-reset:hover:not(:disabled), [data-tp="update"]:hover:not(:disabled)
    { border-color:var(--am); box-shadow:0 0 0 3px color-mix(in srgb, var(--am) 22%, transparent); }
  .btn.dg:hover:not(:disabled), #btn-delete-devices:hover:not(:disabled), #btn-prof-delete:hover:not(:disabled), #btn-unschedule:hover:not(:disabled),
  [data-tp="delete"]:hover:not(:disabled), [data-ct="stop"]:hover:not(:disabled)
    { border-color:var(--rd); box-shadow:0 0 0 3px color-mix(in srgb, var(--rd) 22%, transparent); }
  /* backup comparison */
  .cmp-box { border:1px solid var(--bd); background:var(--pn2); border-radius:8px; padding:8px 10px; margin-bottom:8px; }
  .cmp-box .cmp-row { display:flex; align-items:center; gap:6px; margin:4px 0; }
  .cmp-box .cmp-row > span { width:42px; flex:none; font-size:11.5px; color:var(--mu); }
  .cmp-box select { flex:1; min-width:0; }
  .cmp-box .cmp-opt { display:flex; align-items:center; gap:6px; margin:6px 0; font-size:12px; color:var(--tx); }
  .cmp-bar { display:flex; flex-wrap:wrap; align-items:center; gap:6px 10px; margin:10px 0 6px; }
  .cmp-bar label { margin:0; }
  .cmp-bar select { width:auto; min-width:170px; }
  .cmp-opt { display:flex; align-items:center; gap:6px; font-size:12px; color:var(--tx); }
  .cmp-diff { flex:1; overflow:auto; min-height:160px; padding:8px 10px; background:var(--pn2); border:1px solid var(--bd); border-radius:8px; }
  .df-sum { display:flex; flex-wrap:wrap; gap:6px; align-items:center; margin-bottom:8px; font-size:12px; }
  .df-sum .pill { margin:0; }
  .df-cmd { border:1px solid var(--bd); border-radius:8px; margin:6px 0; background:var(--pn); }
  .df-cmd > summary { cursor:pointer; padding:6px 10px; font:600 12px ui-monospace,Consolas,monospace; display:flex; gap:8px; align-items:center; }
  .df-cmd > summary .st { margin-left:auto; font:600 11px system-ui,sans-serif; }
  .df-cmd pre { margin:0; padding:0; background:transparent; overflow:auto; font:11.5px/1.45 ui-monospace,Consolas,"Courier New",monospace; }
  .df-cmd pre div { padding:0 10px; white-space:pre; }
  .df-a { background:var(--t-gn); color:var(--gr); }
  .df-r { background:var(--t-r); color:var(--rd); }
  .df-h { background:var(--t-b); color:var(--c-b); }
  .df-same { color:var(--mu); font-size:12px; padding:2px 4px; }

  /* backup comparison: full screen, left (older) / right (newer) */
  .modal.cmp-full { padding:0; align-items:stretch; background:var(--pn); }
  .cmp-card { max-width:none; width:100%; height:100%; border-radius:0; border:0; box-shadow:none; display:flex; flex-direction:column; min-height:0; padding:10px 14px 12px; }
  .cmp-top { display:flex; align-items:center; gap:10px; flex:none; }
  .cmp-top h2 { flex:1; min-width:0; margin:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .cmp-seg { display:inline-flex; border:1px solid var(--bd); border-radius:7px; overflow:hidden; flex:none; }
  .cmp-seg button { background:var(--pn2); border:0; color:var(--tx); padding:5px 12px; font:inherit; cursor:pointer; }
  .cmp-seg button.on { background:var(--acs); color:var(--ac); font-weight:600; }
  .cmp-bar { flex:none; }
  .cmp-main { flex:1; min-height:0; display:flex; gap:10px; }
  .cmp-list { width:280px; flex:none; display:flex; flex-direction:column; min-height:0; border:1px solid var(--bd); border-radius:8px; background:var(--pn2); }
  .cmp-list .lh { padding:7px 10px; border-bottom:1px solid var(--bd); font-size:12px; flex:none; }
  .cmp-list .lb { flex:1; min-height:0; overflow:auto; }
  .cmp-item { display:block; width:100%; text-align:left; background:none; border:0; border-bottom:1px solid var(--bd); color:var(--tx); padding:7px 10px; cursor:pointer; font:inherit; }
  .cmp-item:hover { background:var(--sel); }
  .cmp-item.on { background:var(--acs); box-shadow:inset 3px 0 0 var(--ac); }
  .cmp-item .c { font:600 12px ui-monospace,Consolas,"Courier New",monospace; overflow-wrap:anywhere; }
  .cmp-item .s { font-size:11px; margin-top:2px; }
  .cmp-pane { flex:1; min-width:0; display:flex; flex-direction:column; min-height:0; }
  .cmp-pane .df-sum { flex:none; }
  .cmp-diff { min-height:0; padding:0; }
  .cmp-diff > .cmp-msg { padding:12px 14px; color:var(--mu); }
  table.sbs { width:100%; border-collapse:collapse; table-layout:fixed; font:12px/1.45 ui-monospace,Consolas,"Courier New",monospace; }
  table.sbs th { position:sticky; top:0; z-index:2; background:var(--pn); padding:7px 10px; font:600 12px system-ui,sans-serif; color:var(--tx);
                 text-transform:none; letter-spacing:0; border-bottom:1px solid var(--bd); text-align:left; }
  table.sbs th + th { border-left:2px solid var(--bd); }
  table.sbs td { padding:0 8px; vertical-align:top; border:0; }
  table.sbs td.ln { text-align:right; color:var(--mu); user-select:none; -webkit-user-select:none; font-size:11px; padding-right:6px; background:var(--pn2); }
  table.sbs td.ln + td.tx + td.ln { border-left:2px solid var(--bd); }
  table.sbs td.tx { white-space:pre-wrap; overflow-wrap:anywhere; }
  table.sbs td.l-r { background:var(--t-r); } table.sbs td.l-a { background:var(--t-gn); } table.sbs td.emp { background:var(--pn2); }
  table.sbs td.ln.l-r { background:var(--t-r); color:var(--rd); } table.sbs td.ln.l-a { background:var(--t-gn); color:var(--gr); }
  table.sbs mark { background:color-mix(in srgb, currentColor 28%, transparent); color:inherit; border-radius:2px; }
  table.sbs td.vol { color:var(--mu); }
  table.sbs tr.gap td { text-align:center; padding:3px 8px; background:var(--t-b); color:var(--c-b); cursor:pointer; font:11.5px system-ui,sans-serif; }
  table.sbs tr.gap:hover td { filter:brightness(1.15); }
  pre.uni { margin:0; padding:8px 0 40px; font:12px/1.45 ui-monospace,Consolas,"Courier New",monospace; }
  pre.uni div { padding:0 12px; white-space:pre-wrap; overflow-wrap:anywhere; }
  @media (max-width:860px) { .cmp-main { flex-direction:column; } .cmp-list { width:auto; max-height:32%; } }
  /* console under the step detail */
  .cc { height:340px; min-height:220px; max-height:85vh; resize:vertical; display:flex; flex-direction:column; background:#0b1020; color:#d6e2ff; border:1px solid #2a3558; border-radius:10px; overflow:hidden; }
  .cc-h { display:flex; align-items:center; gap:6px; padding:6px 8px 6px 11px; background:#151c33; flex:none; flex-wrap:wrap; }
  .cc-t { font-weight:600; font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:100%; }
  .cc-st { font-size:11px; padding:1px 8px; border-radius:999px; background:#2a3558; color:#d6e2ff; white-space:nowrap; }
  .cc-st.on { background:#17382a; color:#4ade80; }
  .cc-st.busy { background:#3b3320; color:#f4c542; }
  .cc-sp { flex:1; }
  .cc-b { background:transparent; border:1px solid #3b4a7a; color:#d6e2ff; border-radius:6px; padding:3px 9px; font:inherit; font-size:11.5px; cursor:pointer; }
  .cc-b:hover:not(:disabled) { background:rgba(255,255,255,.12); }
  .cc-b:disabled { opacity:.4; cursor:not-allowed; }
  .cc-b.go { background:#2f6fe4; border-color:#2f6fe4; color:#fff; font-weight:600; }
  .cc-out { flex:1; min-height:0; margin:0; padding:8px 11px; overflow:auto; white-space:pre-wrap; word-break:break-word; font:12px/1.45 ui-monospace,Consolas,"Courier New",monospace; }
  .cc-in { display:flex; align-items:center; gap:7px; padding:6px 11px; border-top:1px solid #2a3558; background:#0e1429; flex:none; }
  .cc-in span { color:#4ade80; font:12px ui-monospace,Consolas,monospace; }
  .cc-in input { flex:1; min-width:0; background:transparent; border:0; outline:0; color:#f3f6ff; caret-color:#4ade80; font:12.5px ui-monospace,Consolas,"Courier New",monospace; padding:3px 0; }
  .cc-in input:disabled { opacity:.45; }
  .cc-n { padding:5px 11px; font-size:10.5px; color:#93a1c0; border-top:1px solid #2a3558; flex:none; }
  .cc.max { position:fixed; inset:28px; margin:0; z-index:1150; height:auto; max-height:none; resize:none; box-shadow:0 14px 60px rgba(0,0,0,.55); }
  .cc.max .cc-out { font-size:13.5px; }
  @media (max-width:1000px) { .ct-wb, .ct-result-g { grid-template-columns:minmax(0,1fr); } .ct-tiles { grid-template-columns:repeat(2,minmax(0,1fr)); } }
  .pill[data-rtip] { cursor:help; }
  [data-rtip] { cursor:help; }
</style>
</head>
<body class="role-__ROLE__">

<header>
  <div class="brand">
    <img src="__APP_ICON__" alt="" aria-hidden="true">
    <span>NetUpgrade</span>
    <span class="appver" title="Application version">v__APP_VERSION__</span>
  </div>
  <span class="sitemenu"><button type="button" id="site-badge" class="badge warn site-btn" aria-haspopup="true" aria-expanded="false" title="Select or create a site">Site: Not Assigned</button><div id="site-pop" class="site-pop" role="menu" hidden></div></span>
  <nav class="tabs" aria-label="Sections" role="tablist">
    <button type="button" class="tab on" id="tab-home" role="tab" aria-selected="true" data-view="home"><span data-i="home"></span>Home</button>
    <button type="button" class="tab" id="tab-diagram" role="tab" aria-selected="false" data-view="diagram"><span data-i="diagram"></span>Diagram</button>
    <button type="button" class="tab" id="tab-storage" role="tab" aria-selected="false" data-view="storage"><span data-i="storage"></span>Storage</button>
  </nav>
  <span class="sp"></span>
  <span class="lost" id="lost"></span>
  <span class="mdot" id="mon-dot" role="img" aria-label="Monitoring status" title="Monitoring status"></span>
  <span id="state" class="mstate">Monitor starting</span>
  <button class="btn p ok" id="btn-mon" style="padding:7px 14px">Start Monitoring</button>
  <span class="notemenu"><button class="ib" id="btn-bell" title="Notifications - upgrades started and completed" aria-haspopup="true" aria-expanded="false" style="position:relative"><span data-i="bell"></span><span class="nbadge" id="bell-count" hidden>0</span></button><div class="note-pop" id="note-pop" role="dialog" aria-label="Notifications" hidden><div class="nh"><span>Notifications</span><span><button type="button" id="note-sound" title="Play a sound for new alerts (not for logins)">Sound: On</button> <button type="button" id="note-clear">Clear All</button></span></div><div class="nl" id="note-list"></div></div></span>
  <button class="ib" id="btn-settings" title="Settings" aria-label="Settings"><span data-i="gear"></span></button>
  <button class="ib" id="btn-theme" title="Switch between light and dark mode"></button>
  <span class="usermenu"><button type="button" class="userbtn" id="btn-user" aria-haspopup="true" aria-expanded="false" title="Signed-in user"><span data-i="user"></span><span class="uname" id="uname"></span></button><div class="user-pop" id="user-pop" role="dialog" aria-label="Signed-in user" hidden><div class="uh"><b id="uname2"></b><span id="urole2"></span></div><a id="btn-logout" href="/__guard/logout" target="_top">Sign out</a></div></span>
</header>

<div class="app nosite" id="dashboard">
  <aside class="side wide" id="side">
    <div class="sh"><button class="cb" id="btn-collapse" title="Collapse panel"><span data-i="menu"></span></button><span class="shl">NetUpgrade</span></div>

    <div class="sec" data-sec="1">
      <h3 tabindex="0" role="button" title="Device List"><span data-i="1"></span><span class="lbl">Device List</span><span class="ch" data-i="dn"></span></h3>
      <div class="body">
        <label for="device-file" style="margin-top:0">Upload device list (.xlsx)<span class="tip" data-tip="Excel columns: ip (required), hostname and platform. Use Download template for a ready-made file."></span></label>
        <input type="file" id="device-file" accept=".xlsx">
        <div class="note" id="device-file-last"></div>
        <div class="row"><button class="btn p" id="btn-upload-devices" disabled>Upload List</button></div>
        <div class="note"><a class="lnk" id="btn-device-template" href="/api/devices/template.xlsx" download="device_list_template.xlsx" title="Download an empty Excel file with the right columns (ip, hostname, platform) to fill in and upload">Download template</a></div>
        <hr class="sep">
        <label for="new-hostname">Add device manually</label>
        <input type="text" id="new-hostname" placeholder="Hostname (required)" autocomplete="off" spellcheck="false">
        <input type="text" id="new-ip" placeholder="IP address (required)" autocomplete="off" spellcheck="false">
        <input type="text" id="new-platform" list="platform-options" placeholder="Platform (e.g. C9200L-48P-4G)" autocomplete="off" spellcheck="false" maxlength="40">
        <datalist id="platform-options"><option value="C9200"><option value="C9200L"><option value="C9300"><option value="C9300L"><option value="C9500"><option value="2960X"></datalist>
        <div class="row"><button class="btn p" id="btn-add-device">Add Device</button></div>
        <div class="msg" id="msg-1" role="status" aria-live="polite"></div>
      </div>
    </div>

    <div class="sec" data-sec="2">
      <h3 tabindex="0" role="button" title="Credential Profiles"><span data-i="2"></span><span class="lbl">Credential Profiles</span><span class="ch" data-i="dn"></span></h3>
      <div class="body">
        <label for="prof-edit" style="margin-top:0">Create or edit a profile<span class="tip w" data-tip="Every device needs a profile before it can be backed up or upgraded (credential tests can use any saved profile without assigning it). Passwords are saved encrypted (only your Windows user can read them), never written to logs and never shown again. When editing, leave a password empty to keep the saved one.\n\nTo test: select a switch in the Devices table, press Test Credentials, and save the profile once the login works. The test never saves or assigns anything."></span></label>
        <select id="prof-edit"></select>
        <input type="text" id="prof-name" placeholder="Profile name (e.g. Core switches)" maxlength="40" autocomplete="off" spellcheck="false">
        <input type="text" id="prof-user" placeholder="Username" autocomplete="off" spellcheck="false">
        <input type="password" id="prof-pass" placeholder="Password" autocomplete="new-password">
        <input type="password" id="prof-secret" placeholder="Enable secret (optional)" autocomplete="new-password">
        <label class="chk"><input type="checkbox" id="prof-show"><span>Show password</span></label>
        <label class="chk" id="prof-clear-wrap" hidden><input type="checkbox" id="prof-clear"><span>Remove the saved enable secret</span></label>
        <label id="prof-proto-label">Connection<span class="tip w" data-tip="Telnet sends the username, password and every command in clear text. Use it only on a trusted management network. Upgrades need SSH (the image is copied with SCP); backups, credential tests, readiness checks, the console and topology discovery work over Telnet."></span></label>
        <div class="protobox" role="radiogroup" aria-labelledby="prof-proto-label">
          <label class="ropt"><input type="radio" name="prof-proto" value="ssh" checked><span>SSH</span></label>
          <label class="ropt"><input type="radio" name="prof-proto" value="telnet"><span>Telnet</span></label>
          <span class="portwrap"><label for="prof-port">Port</label><input type="text" id="prof-port" value="22" placeholder="22" inputmode="numeric" maxlength="5" autocomplete="off" spellcheck="false" title="TCP port used to reach the switch. SSH is normally 22, Telnet 23."></span>
        </div>
        <div class="row">
          <button class="btn" id="btn-prof-test" disabled
              title="Log in (and enable, if a secret is typed) on the switches selected in the Devices table using the details above - nothing is saved or assigned">Test Credentials</button>
          <button class="btn p" id="btn-prof-save">Save Profile</button>
        </div>
        <div class="row"><button class="btn" id="btn-prof-delete" disabled>Delete Profile</button></div>
        <hr class="sep">
        <label for="prof-assign">Use a saved profile on the selected devices</label>
        <select id="prof-assign"></select>
        <div class="row" style="flex-direction:column;align-items:stretch;gap:8px">
          <button class="btn p" id="btn-prof-assign" disabled>Assign To Selected</button>
        </div>
        <div class="msg" id="msg-2" role="status" aria-live="polite"></div>
      </div>
    </div>

    <div class="sec" data-sec="3">
      <h3 tabindex="0" role="button" title="Configuration Backup"><span data-i="3"></span><span class="lbl">Configuration Backup</span><span class="ch" data-i="dn"></span></h3>
      <div class="body">
        <label for="custom-commands" style="margin-top:0">Additional commands (optional)</label>
        <textarea id="custom-commands" placeholder="show vlan brief&#10;show interfaces status"></textarea>
        <div class="row"><button class="btn p" id="btn-backup" disabled>Run Backup</button><span class="tip" data-tip="Saves show running-config, show version, show boot, show inventory and the additional commands (all outputs in ONE file per switch), then shows the current software version in the table. Additional commands: one per line."></span></div>
        <div class="hint" id="backup-status">No backup run yet.</div>
        <div class="hint" id="backup-commands-log" style="max-height:110px; overflow-y:auto; display:none;"></div>
        <div class="msg" id="msg-3" role="status" aria-live="polite"></div>
      </div>
    </div>

    <div class="sec o" data-sec="4">
      <h3 tabindex="0" role="button" title="Firmware Upgrade"><span data-i="4"></span><span class="lbl">Firmware Upgrade</span><span class="ch" data-i="dn"></span></h3>
      <div class="body">
        <label for="image-select" id="image-select-label" style="margin-top:0">Select image to install<span class="tip" data-tip="To add an image, open Storage and use IOS-XE images (upload + MD5 check). Rollout options (switches at the same time, stop on failure) and command timeouts are in Settings (gear icon)."></span></label>
        <select id="image-select"><option value="">No images uploaded yet</option></select>
        <label for="upgrade-action">Action<span class="tip" id="action-tip" data-tip="Full upgrade: uploads the image, activates it and reloads the switch in one step."></span></label>
        <select id="upgrade-action">
          <option value="full">Full upgrade</option>
          <option value="distribute">Distribute only</option>
          <option value="activate">Activate only</option>
        </select>
        <div class="row"><button class="btn" id="btn-readiness" disabled
              title="Log in and check version, free flash space, CPU, memory and stack health - read-only, no upload or reload">Readiness Check</button></div>
        <label class="chk"><input type="checkbox" id="backup-before-upgrade" checked><span>Back up configuration first</span></label>
        <label class="chk"><input type="checkbox" id="skip-image-check"><span>Ignore image check warnings</span><span class="tip w" data-tip="Tick only if you are sure: the image does not match the switch model, or it is older than the running version."></span></label>
        <label class="chk"><input type="checkbox" id="ignore-space"><span>Ignore flash space check</span><span class="tip w" data-tip="Tick only if you are sure: continues even when free flash is below the image size + 1 GB that Cisco asks for. Free some flash first if you can (delete old .bin / .pkg files). If flash really runs out, the install can fail part-way."></span></label>
        <label for="sched-time">Schedule for (optional)<span class="tip" id="slot-tip" data-tip="A live web console opens at the bottom of this page for every switch, immediately or at its scheduled time. The switch reloads by itself; the table shows the new version once it is back."></span></label>
        <input type="datetime-local" id="sched-time">
        <div class="row">
          <button class="btn p" id="btn-upgrade" disabled>Start Now</button>
          <button class="btn" id="btn-schedule" disabled>Schedule</button>
          <button class="btn" id="btn-unschedule" disabled title="Cancel the scheduled upgrade on the selected switches">Cancel</button>
        </div>
        <div class="msg" id="msg-4" role="status" aria-live="polite"></div>
      </div>
    </div>
    <div class="sec" data-sec="5">
      <h3 tabindex="0" role="button" title="Reports"><span data-i="5"></span><span class="lbl">Reports</span><span class="ch" data-i="dn"></span></h3>
      <div class="body">
        <label for="rep-type">Report Type</label>
        <select id="rep-type">
          <option value="xlsx">Excel (.xlsx)</option>
          <option value="pdf">Print / PDF</option>
          <option value="csv">CSV (.csv)</option>
        </select>
        <p class="muted" id="rep-desc" style="margin:6px 0 10px"></p>
        <button type="button" class="btn p" id="rep-go" style="width:100%">Generate</button>
      </div>
    </div>
  </aside>

  <main>
    <div class="mainmsg"><span class="msg" id="msg" role="status" aria-live="polite"></span></div>
    <div hidden>
      <span id="p-text"></span><span id="p-cycles"></span><span id="meta"></span><div id="p-bar"></div>
    </div>

    <div class="wg">
      <div class="card"><h4>Inventory</h4><div class="tl">
        <div class="t b"><b id="c-total">-</b><span>Switches</span></div>
        <div class="t b"><b id="c-up">-</b><span>Up</span></div>
        <div class="t r"><b id="c-down">-</b><span>Down</span></div>
      </div></div>
      <div class="card"><h4>Upgrades</h4><div class="tl">
        <div class="t a"><b id="c-active">-</b><span>Running</span></div>
        <div class="t g"><b id="c-scheduled">-</b><span>Scheduled</span></div>
        <div class="t gn"><b id="c-completed">-</b><span>Completed</span></div>
      </div></div>
    </div>

    <div class="work">
      <div class="card tb">
        <h4>Devices
          <span class="a">
            <span id="sel-count">0 selected</span>
            <button class="btn sm" id="btn-version" disabled title="Log in to the selected switches and read the current software version">Refresh Version</button>
            <button class="btn sm" id="btn-delete-devices" disabled title="Remove the selected devices from the list (existing backups are not affected)">Delete Selected</button>
            <button class="btn sm" id="btn-cols" type="button" title="Choose which columns the table shows">Columns</button>
            <span class="fl">
              <button class="btn sm on" data-f="all">All</button>
              <button class="btn sm" data-f="UP">Up</button>
              <button class="btn sm" data-f="DOWN">Down</button>
            </span>
          </span>
          <button type="button" class="ib mx" data-max="Devices" aria-pressed="false"><span data-i="expand"></span></button>
        </h4>
        <div class="dev-tools" id="dev-tools">
          <div class="st-search">
            <span class="si" data-i="search"></span>
            <input type="text" id="dev-search" placeholder="Search name, IP, label, platform, version..." autocomplete="off" spellcheck="false" aria-label="Search devices">
            <button type="button" class="st-clear" id="dev-clear" title="Clear search" hidden>&times;</button>
          </div>
          <select id="dev-f-upgrade" aria-label="Filter by upgrade state">
            <option value="all">Upgrade: any</option>
            <option value="running">Running</option>
            <option value="scheduled">Scheduled</option>
            <option value="completed">Completed</option>
            <option value="failed">Failed</option>
            <option value="attention">Needs attention</option>
            <option value="none">Not started</option>
          </select>
          <select id="dev-f-ready" aria-label="Filter by readiness">
            <option value="all">Readiness: any</option>
            <option value="ready">Ready</option>
            <option value="warning">Warning</option>
            <option value="not_ready">Not ready</option>
            <option value="checking">Checking</option>
            <option value="none">Not checked</option>
          </select>
          <select id="dev-f-platform" aria-label="Filter by platform"><option value="all">Platform: any</option></select>
          <span class="dev-shown" id="dev-shown"></span>
          <button type="button" class="btn sm" id="dev-reset" hidden>Reset Filters</button>
        </div>
        <div class="sc">
          <table id="dev-table">
            <thead><tr>
              <th class="ck"><input type="checkbox" id="sel-all" title="Select all shown"></th>
              <th data-sort="host">Switch</th><th class="c-platform" data-sort="platform">Platform</th><th class="c-status" data-sort="status">Status</th><th class="c-creds" data-sort="creds">Credentials</th><th class="c-ready" data-sort="ready">Readiness</th>
              <th class="c-version" data-sort="version">Software Version</th><th class="c-boot" data-sort="boot" title="Running mode (INSTALL / BUNDLE) and the boot variable">Boot Mode</th><th class="c-backup" data-sort="backup">Backup</th><th class="c-stage" data-sort="stage" title="Where the firmware image is on this switch: distributed to flash, being activated, or upgraded">Image Stage</th><th class="c-upgrade" data-sort="upgrade">Firmware Upgrade</th><th class="c-post" data-sort="post">Post-Check</th><th class="rc c-recent">Recent Checks</th>
            </tr></thead>
            <tbody id="rows"></tbody>
          </table>
        </div>
      </div>
      <div class="card e collapsed" id="ev-card"><h4 id="ev-head"><span>Events</span><span class="evn" id="ev-count"></span><span class="evlast" id="ev-last"></span><button type="button" class="ib mx" id="ev-toggle" data-max="Events" aria-pressed="false" aria-expanded="false" aria-controls="events"><span data-i="expand"></span></button></h4><div class="ev" id="events"><span class="muted">Nothing yet.</span></div></div>
    </div>
  </main>
</div>

<div class="app" id="diagram" hidden>
  <aside class="side dg-side" aria-label="Diagram devices">
    <div class="sh"><span data-i="diagram"></span><span class="shl">Diagram</span></div>
    <div class="dg-sec">
      <div class="note" id="dg-site" style="margin-top:0"></div>
      <div class="dg-actions"><button type="button" class="btn" id="dg-all" disabled>Select All</button><button type="button" class="btn p" id="dg-run" title="Discover topology: read the CDP neighbours of the ticked switches" disabled>Discover</button></div>
      <div class="dg-list" id="dg-list"></div>
      <div class="msg" id="dg-msg" role="status" aria-live="polite" style="margin-top:8px"></div>
    </div>
  </aside>
  <main>
    <div class="card dg-card" id="dg-card">
      <h4>Topology<span class="tip" data-tip="Built from CDP: each selected switch is asked for its CDP neighbours, so the links shown are the ones the switches report. Devices need a credentials profile. The top of each tree is picked automatically (the switch in the middle of the network). Only the switches ticked for the last Discover are shown. Drag a switch to place it anywhere; drag empty space to move the whole diagram; scroll the mouse wheel to zoom."></span>
        <span class="a">
          <label class="chk" style="margin:0"><input type="checkbox" id="dg-ports" checked>Port labels</label>
          <label class="chk" style="margin:0"><input type="checkbox" id="dg-other">APs, phones &amp; hosts</label>
          <button type="button" class="btn sm" id="dg-zout" title="Zoom out">&minus;</button>
          <button type="button" class="btn sm" id="dg-zin" title="Zoom in">+</button>
          <button type="button" class="btn sm" id="dg-reset" title="Forget dragged positions and arrange the switches again">Reset Layout</button>
          <button type="button" class="btn sm" id="dg-fit" title="Fit the whole diagram in the window">Fit</button>
          <button type="button" class="btn sm" id="dg-png" title="Save the diagram as a PNG picture, with the switches where you placed them" disabled>Export PNG</button>
          <button type="button" class="btn sm" id="dg-refresh">Refresh</button>
          <button type="button" class="btn sm dg-ico" id="dg-full" title="Full screen (Esc to exit)" aria-label="Full screen" aria-pressed="false"><span data-i="expand"></span></button>
        </span></h4>
      <div class="dg-banner" id="dg-banner" hidden></div>
      <div class="dg-wrap" id="dg-wrap">
        <svg id="dg-svg" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Network topology"></svg>
        <div class="dg-empty" id="dg-empty"></div>
        <div class="dg-hint" id="dg-hint">Drag a switch to move it &middot; drag the background to pan &middot; mouse wheel to zoom</div>
        <aside class="dg-detail" id="dg-detail" hidden></aside>
      </div>
      <div class="dg-legend">
        <span><i class="dot ok"></i>Up</span><span><i class="dot bad"></i>Down</span><span><i class="dot"></i>Not checked yet</span>
        <span class="lg-run">Upgrading</span><span class="lg-done">Upgrade completed</span><span class="lg-fail">Failed</span><span class="lg-q">Scheduled / queued</span>
        <span class="lg-hold">Hold: uplink of a switch being upgraded</span><span class="lg-ext">Dashed grey: not onboarded</span>
      </div>
    </div>
  </main>
</div>

<div class="app" id="storage" hidden>
  <aside class="side st-side" aria-label="Sites">
    <div class="sh"><span data-i="site"></span><span class="shl">Sites</span></div>
    <div class="st-sites" id="st-sites"></div>
  </aside>
  <main>
    <div class="card img-card" id="img-card">
      <h4>IOS-XE Images<span class="tip" data-tip="Images are saved in the folder selected during installation and used by Firmware upgrade. Paste the MD5 checksum from Cisco's download page: the upload is rejected if the file does not match it."></span>
        <span class="a"><button type="button" class="btn sm" id="img-refresh">Refresh</button></span></h4>
      <div class="img-body">
        <div class="img-form">
          <div><label for="image-file">Image file (.bin)</label><input type="file" id="image-file" accept=".bin"></div>
          <div><label for="image-md5">Cisco MD5 checksum</label>
            <input type="text" id="image-md5" maxlength="32" spellcheck="false" autocomplete="off" autocapitalize="off" placeholder="32 hexadecimal characters"></div>
          <div><button class="btn p" id="btn-upload-image" disabled>Upload And Verify</button></div>
        </div>
        <div class="hint" id="image-upload-status"></div>
        <div class="msg" id="img-msg" role="status" aria-live="polite"></div>
        <table class="st-table img-table">
          <thead><tr><th>Image</th><th>Size</th><th>Checksum</th><th></th></tr></thead>
          <tbody id="img-rows"></tbody>
        </table>
      </div>
    </div>
    <div class="card st-card">
      <div class="st-bar">
        <nav class="crumbs" id="st-crumbs" aria-label="Folder path"></nav>
        <span class="sp"></span>
        <div class="st-search">
          <span class="si" data-i="search"></span>
          <input type="text" id="st-search" placeholder="Search backups, reports and runs in this site..." autocomplete="off" spellcheck="false" aria-label="Search files and folders">
          <button type="button" class="st-clear" id="st-clear" title="Clear search" hidden>&times;</button>
        </div>
        <button type="button" class="btn sm in" id="st-compare" title="Compare two backups of this switch" hidden>Compare Backups</button>
        <a class="btn sm" id="st-zip" href="#" title="Download this folder as a ZIP file">Download Folder (.zip)</a>
        <button type="button" class="btn sm" id="st-refresh">Refresh</button>
      </div>
      <div class="msg" id="st-msg" role="status" aria-live="polite" style="margin:6px 14px 0"></div>
      <div class="sc">
        <table class="st-table">
          <thead id="st-head"></thead>
          <tbody id="st-rows"></tbody>
        </table>
        <div class="st-none" id="st-none" hidden></div>
      </div>
      <div class="st-foot" id="st-foot"></div>
    </div>
  </main>
</div>

<div class="app" id="settings" hidden>
  <aside class="side st-side" aria-label="Settings menu">
    <div class="sh"><span data-i="gear"></span><span class="shl">Settings</span></div>
    <div class="set-nav" id="set-nav" role="tablist" aria-orientation="vertical"></div>
  </aside>
  <main>
    <div class="card set-card">
      <h4><span id="set-title">Settings</span>
        <span class="a"><span id="set-dirty"></span>
          <button type="button" class="btn sm" id="set-reset" title="Put every value on this tab back to its default (not saved until you click Save)">Restore Defaults</button>
          <button type="button" class="btn p sm" id="set-save">Save Settings</button></span></h4>
      <div class="msg" id="set-msg" role="status" aria-live="polite" style="margin:8px 18px 0"></div>
      <div class="set-note" id="set-note" hidden>Only an Administrator can change these settings.</div>
      <div class="set-body" id="set-body"></div>
    </div>
  </main>
</div>

<footer class="credit">
  <span>Developed by <strong>__APP_AUTHOR__</strong></span>
  <span class="sep" aria-hidden="true">|</span>
  <a href="__APP_REPO_URL__" target="_blank" rel="noopener noreferrer" title="View this project on GitHub">
    <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>GitHub</a>
</footer>

<div id="st-modal" class="modal" hidden>
  <div class="modal-card wide" role="dialog" aria-labelledby="st-modal-title">
    <div style="display:flex; align-items:center; gap:10px;">
      <h2 id="st-modal-title" style="flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"></h2>
      <button type="button" class="btn sm" id="st-modal-top" title="Jump to the start of the file">Top</button>
      <button type="button" class="btn sm" id="st-modal-end" title="Jump to the end of the file">End</button>
      <a class="btn sm p" id="st-modal-dl" href="#">Download</a>
      <button type="button" class="btn sm" id="st-modal-close">Close</button>
    </div>
    <div class="hint" id="st-modal-meta" style="margin-top:2px"></div>
    <pre id="st-pre"></pre>
  </div>
</div>

<div id="cmp-modal" class="modal cmp-full" hidden>
  <div class="modal-card cmp-card" role="dialog" aria-labelledby="cmp-title">
    <div class="cmp-top">
      <h2 id="cmp-title">Compare backups</h2>
      <div class="cmp-seg" id="cmp-view" role="group" aria-label="Comparison layout">
        <button type="button" data-v="side" class="on">Side by side</button><button type="button" data-v="uni">Unified</button>
      </div>
      <button type="button" class="btn sm" id="cmp-close">Close</button>
    </div>
    <div class="cmp-bar">
      <label for="cmp-a">Left (older)</label><select id="cmp-a"></select>
      <label for="cmp-b">Right (newer)</label><select id="cmp-b"></select>
      <label class="cmp-opt"><input type="checkbox" id="cmp-ign" checked> Ignore lines that change every time (timestamps, uptime, byte counts)</label>
      <label class="cmp-opt"><input type="checkbox" id="cmp-all"> Show all lines</label>
      <button type="button" class="btn sm" id="cmp-prev" title="Previous change">&#9650; Prev</button>
      <button type="button" class="btn sm" id="cmp-next" title="Next change">&#9660; Next</button>
      <button type="button" class="btn sm in" id="cmp-go">Compare</button>
    </div>
    <div class="cmp-main">
      <div class="cmp-list">
        <div class="lh"><label class="cmp-opt"><input type="checkbox" id="cmp-same"> Show unchanged commands</label></div>
        <div class="lb" id="cmp-list"></div>
      </div>
      <div class="cmp-pane">
        <div class="df-sum" id="cmp-sum"></div>
        <div class="cmp-diff" id="cmp-diff"></div>
      </div>
    </div>
  </div>
</div>

<div id="site-modal" class="modal" hidden>
  <div class="modal-card" role="dialog" aria-labelledby="site-modal-title">
    <h2 id="site-modal-title">Create New Site</h2>
    <label for="site-new">Site name</label>
    <input type="text" id="site-new" placeholder="e.g. Chicago-HQ, Customer 12 ..." maxlength="100" autocomplete="off">
    <div class="msg" id="site-msg" role="status" aria-live="polite" style="margin-top:8px"></div>
    <div style="display:flex; gap:8px; justify-content:flex-end; margin-top:8px;">
      <button class="btn" id="btn-site-cancel" type="button">Cancel</button>
      <button class="btn p" id="btn-site-create" type="button">Create Site</button>
    </div>
  </div>
</div>
<div id="toasts" role="status" aria-live="polite"></div>

<aside id="drawer" class="drawer" role="dialog" aria-label="Switch details" hidden>
  <div class="dr-head">
    <div class="dr-ic" id="dr-ic"></div>
    <div class="dr-title"><div class="dr-name" id="dr-name"></div><div class="dr-sub" id="dr-sub"></div></div>
    <button type="button" class="ib" id="dr-close" title="Close (Esc)" aria-label="Close"></button>
  </div>
  <nav class="dr-tabs" id="dr-tabs" role="tablist" aria-label="Switch details"></nav>
  <div class="dr-body" id="dr-body"></div>
  <div class="dr-term" id="dr-term" hidden>
    <div class="th">
      <span class="tt" id="dr-term-title">SSH console</span><span class="tst" id="dr-term-state">Not connected</span><span class="tip" data-tip="Runs directly on the switch with its credential profile. Every line you type is recorded in the Commands tab. Closes after 15 minutes without use."></span><span class="sp"></span>
      <button type="button" class="tb" id="dr-term-mode" hidden></button>
      <button type="button" class="tb go" id="dr-term-connect">Connect</button>
      <button type="button" class="tb" id="dr-term-ctrlc" title="Send Ctrl+C to the switch">Ctrl+C</button>
      <button type="button" class="tb" id="dr-term-clear" title="Clear this window (the session stays open)">Clear</button>
      <button type="button" class="tb" id="dr-term-max" title="Maximise / restore (Esc restores)">Maximise</button>
    </div>
    <pre class="tout" id="dr-term-out"></pre>
    <div class="tin" id="dr-term-tin"><span>&gt;</span><input type="text" id="dr-term-in" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" placeholder="Type a command and press Enter" disabled></div>
    <div class="tn" id="dr-term-note"></div>
  </div>
  <div id="dr-term-bd" hidden></div>
  <div class="dr-view" id="dr-view" hidden>
    <div class="vh"><button type="button" class="btn sm" id="dr-view-back">&larr; Back</button><span class="nm" id="dr-view-name"></span><button type="button" class="btn sm" id="dr-view-top" title="Jump to the start of the file">Top</button><button type="button" class="btn sm" id="dr-view-end" title="Jump to the end of the file">End</button></div>
    <div class="vn" id="dr-view-note" hidden></div>
    <pre id="dr-view-text"></pre>
  </div>
</aside>

<script>
const $ = id => document.getElementById(id);
const ME = __ME__;     // signed-in user and role, filled in by the server
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g,
  c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmtDur = sec => { sec = Math.max(0, Math.round(sec));
  return Math.floor(sec / 60) + "m " + String(sec % 60).padStart(2, "0") + "s"; };
const fmtMs = v => v == null ? "-" : v.toFixed(1) + " ms";

// ---- Icons (same set as the Cloud Control theme) ----
const I = p => '<svg class="i" viewBox="0 0 24 24">' + p + "</svg>";
const IC = {
  1: I('<rect x="3" y="4" width="18" height="6" rx="1.5"/><rect x="3" y="14" width="18" height="6" rx="1.5"/><path d="M7 7h.01M7 17h.01"/>'),
  2: I('<circle cx="8" cy="15" r="4"/><path d="M10.8 12.2 20 3M16 7l3 3M14 9l2 2"/>'),
  3: I('<ellipse cx="12" cy="5.5" rx="8" ry="3"/><path d="M4 5.5v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6M4 11.5v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/>'),
  5: I('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>'),
  4: I('<path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>'),
  sun: I('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
  moon: I('<path d="M21 13A9 9 0 1 1 11 3a7 7 0 0 0 10 10z"/>'),
  menu: I('<path d="M4 6h16M4 12h16M4 18h16"/>'),
  dn: I('<path d="m6 9 6 6 6-6"/>'),
  home: I('<path d="M3 11 12 3l9 8"/><path d="M5 10v10h5v-6h4v6h5V10"/>'),
  diagram: I('<rect x="9" y="3" width="6" height="5" rx="1"/><rect x="3" y="16" width="6" height="5" rx="1"/><rect x="15" y="16" width="6" height="5" rx="1"/><path d="M12 8v4M6 16v-4h12v4"/>'),
  storage: I('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
  folder: I('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
  file: I('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/>'),
  search: I('<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>'),
  expand: I('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'),
  shrink: I('<path d="M9 4v5H4M15 4v5h5M9 20v-5H4M15 20v-5h5"/>'),
  user: I('<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>'),
  bell: I('<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.9 1.9 0 0 0 3.4 0"/>'),
  site: I('<path d="M12 21s7-6.2 7-11a7 7 0 1 0-14 0c0 4.8 7 11 7 11z"/><circle cx="12" cy="10" r="2.5"/>'),
  gear: I('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>')
};
document.querySelectorAll("[data-i]").forEach(el => { el.innerHTML = IC[el.dataset.i]; });

// ---- Help tips: a small ? (how it works) or ! (important) icon; its text shows on hover, focus or click ----
const tipPop = document.createElement("div");
tipPop.id = "tip-pop"; tipPop.setAttribute("role", "tooltip"); tipPop.hidden = true;
document.body.appendChild(tipPop);
let tipEl = null, tipPinned = false;
function tipHide() { if (tipEl) tipEl.classList.remove("on"); tipEl = null; tipPinned = false; tipPop.hidden = true; }
function tipShow(el) {
  if (!el.dataset.tip) return;
  if (tipEl && tipEl !== el) tipEl.classList.remove("on");
  tipEl = el; el.classList.add("on");
  tipPop.textContent = el.dataset.tip;
  tipPop.className = el.classList.contains("bad") ? "bad" : el.classList.contains("w") ? "w" : "";
  tipPop.hidden = false;
  const r = el.getBoundingClientRect(), pw = tipPop.offsetWidth, ph = tipPop.offsetHeight, m = 8;
  let x, y;
  if (el.closest(".side") && r.right + 10 + pw < innerWidth - m) {      // sidebar: open to the right so no control is covered
    x = r.right + 10; y = Math.min(Math.max(m, r.top + r.height / 2 - 16), Math.max(m, innerHeight - ph - m));
  } else {
    x = Math.min(Math.max(m, r.left + r.width / 2 - pw / 2), Math.max(m, innerWidth - pw - m));
    y = r.bottom + 8;
    if (y + ph > innerHeight - m) y = Math.max(m, r.top - ph - 8);
  }
  tipPop.style.left = x + "px"; tipPop.style.top = y + "px";
}
// Change a tip's text and kind: "" = ? (information), "w" = ! (important), "bad" = ! in red (a problem).
function tipSet(el, text, kind) {
  el.dataset.tip = text;
  el.className = "tip" + (kind ? " " + kind : "") + (tipEl === el ? " on" : "");
  el.textContent = kind ? "!" : "?";
  el.setAttribute("aria-label", text);
  if (tipEl === el) tipShow(el);
}
function tipInit(el) {
  if (el.dataset.ready) return;
  el.dataset.ready = "1";
  el.tabIndex = 0; el.setAttribute("role", "button");
  tipSet(el, el.dataset.tip || "", el.classList.contains("bad") ? "bad" : el.classList.contains("w") ? "w" : "");
  el.addEventListener("mouseenter", () => tipShow(el));
  el.addEventListener("mouseleave", () => { if (!tipPinned) tipHide(); });
  el.addEventListener("focus", () => tipShow(el));
  el.addEventListener("blur", tipHide);
  el.addEventListener("click", ev => {                 // also works inside labels / headings without toggling them
    ev.preventDefault(); ev.stopPropagation();
    if (tipEl === el && tipPinned) tipHide(); else { tipShow(el); tipPinned = true; }
  });
  el.addEventListener("keydown", ev => {
    if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); el.click(); }
    else if (ev.key === "Escape") tipHide();
  });
}
document.querySelectorAll(".tip").forEach(tipInit);
document.addEventListener("click", ev => { if (tipEl && !ev.target.closest(".tip")) tipHide(); });
document.addEventListener("keydown", ev => { if (ev.key === "Escape") tipHide(); });
window.addEventListener("resize", tipHide);
window.addEventListener("scroll", tipHide, true);

// ---- Signed-in user menu (the role itself is enforced by the server) ----
(function () {
  $("uname").textContent = ME.user;
  $("uname2").textContent = ME.user;
  $("urole2").textContent = "Permission: " + ME.role_label;
  $("btn-logout").hidden = !ME.sign_out;           // no sign-in when used locally without the launcher
  const pop = $("user-pop"), btn = $("btn-user");
  const closeUser = () => { pop.hidden = true; btn.setAttribute("aria-expanded", "false"); };
  btn.addEventListener("click", e => {
    e.stopPropagation(); pop.hidden = !pop.hidden; btn.setAttribute("aria-expanded", String(!pop.hidden));
  });
  document.addEventListener("click", e => { if (!e.target.closest(".usermenu")) closeUser(); });
  document.addEventListener("keydown", e => { if (e.key === "Escape") closeUser(); });
  if (ME.role === "viewer") $("site-badge").disabled = true;    // a viewer cannot switch sites
})();

let filter = "all", last = null;
const selected = new Set();

// ---- column chooser (remembered in this browser) ----------------------------------------------
const COLS = [["platform", "Platform"], ["status", "Status"], ["creds", "Credentials"], ["ready", "Readiness"],
              ["version", "Software version"], ["boot", "Boot mode"], ["backup", "Backup"],
              ["stage", "Image stage"], ["upgrade", "Firmware upgrade"], ["post", "Post-check"], ["recent", "Recent checks"]];
let hiddenCols = [];
try { hiddenCols = JSON.parse(localStorage.getItem("hiddenCols") || "[]").filter(k => COLS.some(c => c[0] === k)); } catch (e) {}
const colPop = document.createElement("div");
colPop.className = "col-pop"; colPop.hidden = true;
colPop.innerHTML = '<div class="cp-h">Show columns</div>' +
  COLS.map(c => '<label class="ci"><input type="checkbox" data-col="' + c[0] + '"> ' + c[1] + "</label>").join("") +
  '<div class="cp-f"><button type="button" class="btn sm" id="cols-all">Show All</button></div>';
document.body.appendChild(colPop);
function syncColChecks() { colPop.querySelectorAll("input[data-col]").forEach(i => { i.checked = !hiddenCols.includes(i.dataset.col); }); }
function applyCols() {
  COLS.forEach(c => $("dev-table").classList.toggle("hc-" + c[0], hiddenCols.includes(c[0])));
  try { localStorage.setItem("hiddenCols", JSON.stringify(hiddenCols)); } catch (e) {}
  $("btn-cols").textContent = hiddenCols.length ? "Columns (" + (COLS.length - hiddenCols.length) + "/" + COLS.length + ")" : "Columns";
}
function closeColPop() { colPop.hidden = true; }
$("btn-cols").addEventListener("click", e => {
  e.stopPropagation();
  if (!colPop.hidden) { closeColPop(); return; }
  syncColChecks();
  const r = $("btn-cols").getBoundingClientRect();
  colPop.style.top = (r.bottom + 4) + "px";
  colPop.style.right = Math.max(8, window.innerWidth - r.right) + "px";
  colPop.hidden = false;
});
colPop.addEventListener("click", e => e.stopPropagation());
colPop.addEventListener("change", e => {
  const k = e.target.dataset.col; if (!k) return;
  hiddenCols = e.target.checked ? hiddenCols.filter(x => x !== k) : hiddenCols.concat([k]);
  applyCols();
});
colPop.querySelector("#cols-all").addEventListener("click", () => { hiddenCols = []; syncColChecks(); applyCols(); });
document.addEventListener("click", closeColPop);
document.addEventListener("keydown", e => { if (e.key === "Escape") closeColPop(); });
applyCols();

let sayTimer = null;
function say(kind, text) {
  const m = $("msg"); m.className = "msg " + kind; m.textContent = text;
  clearTimeout(sayTimer);
  if (text) sayTimer = setTimeout(() => { m.textContent = ""; }, kind === "err" ? 12000 : 6000);
}
function sayIn(tile, kind, text) {
  const m = tile && $("msg-" + tile);
  if (!m) { say(kind, text); return; }
  m.className = "msg " + kind;
  m.textContent = text;
}

async function post(url, body) {
  const r = await fetch(url, {method: "POST", headers: {"Content-Type": "application/json"},
                              body: JSON.stringify(body || {})});
  let d = {}; try { d = await r.json(); } catch (e) {}
  if (!r.ok) throw new Error(d.error || ("Request failed (HTTP " + r.status + ")"));
  return d;
}

async function postFile(url, file, headers = {}) {
  // No Content-Type header here on purpose - the browser sets the multipart boundary itself.
  const fd = new FormData();
  fd.append("file", file, file.name);
  const r = await fetch(url, {method: "POST", headers, body: fd});
  let d = {}; try { d = await r.json(); } catch (e) {}
  if (!r.ok) throw new Error(d.error || ("Upload failed (HTTP " + r.status + ")"));
  return d;
}

// ---- Hover details for the result tags of the device list (the long grey texts used to sit under every tag) ----
const RT = {el: null, key: ""};
const rtAttr = text => text ? ' data-rtip="' + esc(text) + '"' : "";
function pillT(text, cls, tip) { return '<span class="pill ' + cls + '"' + rtAttr(tip) + ">" + esc(text) + "</span>"; }
function rtipPlace(el) {
  tipPop.textContent = el.dataset.rtip; tipPop.className = ""; tipPop.hidden = false;
  const r = el.getBoundingClientRect(), pw = tipPop.offsetWidth, ph = tipPop.offsetHeight, m = 8;
  let x = Math.min(Math.max(m, r.left + r.width / 2 - pw / 2), Math.max(m, innerWidth - pw - m)), y = r.bottom + 6;
  if (y + ph > innerHeight - m) y = Math.max(m, r.top - ph - 6);
  tipPop.style.left = x + "px"; tipPop.style.top = y + "px";
}
function rtipShow(el) { RT.el = el; RT.key = el.dataset.rk || ""; rtipPlace(el); }
function rtipHide() { RT.el = null; RT.key = ""; if (!tipEl) tipPop.hidden = true; }
// The table is rebuilt on every poll: give each tag a stable key and move an open tip onto its new copy.
function rtipIndex() {
  document.querySelectorAll("#rows tr[data-host]").forEach(tr => {
    tr.querySelectorAll("[data-rtip]").forEach((el, i) => { el.dataset.rk = tr.dataset.host + "|" + i; });
  });
  if (!RT.el) return;
  const el = RT.key ? document.querySelector('#rows [data-rk="' + RT.key.replace(/"/g, "") + '"]') : null;
  if (el) { RT.el = el; rtipPlace(el); } else rtipHide();
}
$("rows").addEventListener("mouseover", e => {
  const el = e.target.closest("[data-rtip]");
  if (el && el !== RT.el) rtipShow(el); else if (!el && RT.el) rtipHide();
});
$("rows").addEventListener("mouseleave", rtipHide);
window.addEventListener("scroll", rtipHide, true);

function pill(text, cls) { return '<span class="pill ' + cls + '">' + esc(text) + "</span>"; }

function backupCell(b) {
  const map = {queued: ["QUEUED", "PENDING"], running: ["RUNNING", "RUN"],
               completed: ["COMPLETED", "OK"], failed: ["FAILED", "DOWN"]};
  const m = map[b.state];
  if (!m) return '<span class="muted">-</span>';
  const tip = [m[0] + (b.time ? " " + b.time : ""), b.detail].filter(Boolean).join("\n");
  return pillT(m[0], m[1], tip) + (b.state === "completed" && b.time ? '<span class="label clip">' + esc(b.time) + "</span>" : "");
}

// Where the firmware image is on a switch: [pill text, pill class, image name, detail] or null.
// Live run first, then what the history remembers (the image distributed and not yet activated, or the last upgrade).
const baseName = n => String(n || "").split(/[\\/]/).pop();
const DIST_PHASES = ["waiting_backup", "starting", "connecting", "connected", "prechecking", "boot", "cleaning", "space_check", "uploading", "md5"];
function stageInfo(h) {
  const u = h.upgrade, st = h.stage || {}, act = u.action;
  if (ACTIVE_STATES.includes(u.state)) {
    const dist = act === "distribute" || (act === "full" && DIST_PHASES.includes(u.phase));
    return [dist ? "DISTRIBUTING" : act === "activate" ? "ACTIVATING" : "UPGRADING", "RUN", baseName(u.image) || (st.staged && st.staged.image) || "", u.message || ""];
  }
  if (u.state === "completed" && act === "distribute" && /^DISTRIBUTED/.test(u.message || "")) {
    return ["DISTRIBUTED", "OK", baseName(u.image), "Ready to activate" + (u.finished ? " \u00B7 " + u.finished : "")];
  }
  if (st.staged) return ["DISTRIBUTED", "OK", baseName(st.staged.image), "Ready to activate" + (st.staged.when ? " \u00B7 " + st.staged.when : "")];
  if (u.state === "completed" && act !== "distribute") {
    const nv = u.new_version || h.version;
    return ["UPGRADED", "OK", baseName(u.image), nv ? "Now on " + nv : ""];
  }
  if (st.upgraded) return ["UPGRADED", "OK", baseName(st.upgraded.image), (st.upgraded.version ? "Now on " + st.upgraded.version : "") + (st.upgraded.when ? " \u00B7 " + st.upgraded.when : "")];
  return null;
}
function stageCell(h) {
  const s = stageInfo(h);
  if (!s) return '<span class="muted">-</span>';
  return pillT(s[0], s[1], [s[2], s[3]].filter(Boolean).join("\n"));
}
// Warn (never block) when an image was already distributed to a selected switch.
function stagedWarning(action, image) {
  if (action === "activate" || !last) return "";
  const want = baseName(image), lines = [];
  for (const host of selected) {
    const h = last.hosts.find(x => x.host === host), s = h && h.stage && h.stage.staged;
    if (!s) continue;
    lines.push(" - " + (h.label ? h.label + " (" + host + ")" : host) + ": " + baseName(s.image) +
               (s.when ? " (distributed " + s.when + ")" : "") + (want ? (baseName(s.image) === want ? " - the same image" : " - a different image") : ""));
  }
  if (!lines.length) return "";
  return "\n\nWARNING: an image was already distributed to " + lines.length + " of these switch(es):\n" + lines.join("\n") +
    "\nRunning this again may remove it and copy an image to flash again (time, and flash space while it is copied). " +
    "To use the image that is already there, choose 'Activate only' instead.";
}

function upgradeCell(h) {
  const u = h.upgrade;
  if (h.scheduled && ["idle", "completed", "failed", "attention"].includes(u.state)) {
    if (h.scheduled.queued) return pillT("QUEUED", "PENDING", "Waiting for a free upgrade slot");
    return pillT("SCHEDULED", "PENDING", "Scheduled for " + h.scheduled.run_at_str) + '<span class="label clip">' + esc(h.scheduled.run_at_str) + "</span>";
  }
  const map = {starting: ["STARTING", "RUN"], running: ["RUNNING", "RUN"],
               completed: [u.action === "distribute" ? "DISTRIBUTED" : "UPGRADED", "OK"],
               failed: ["FAILED", "DOWN"], attention: ["CHECK", "WARN"]};
  const m = map[u.state];
  if (!m) return '<span class="muted">-</span>';
  const link = u.job_id ? ' <a href="#" class="clink" data-console="' + esc(h.host) + '">console</a>' : "";
  const upgraded = u.state === "completed" && u.action !== "distribute";
  const nv = u.new_version || h.version;
  return pillT(m[0], m[1], [u.message, u.finished ? "Finished " + u.finished : ""].filter(Boolean).join("\n")) + link +
    (upgraded && nv ? '<span class="label" style="display:block">Now on <b>' + esc(nv) + "</b></span>" : "");
}

function credsCell(c, profile) {
  const map = {testing: ["TESTING", "RUN"], ok: ["OK", "OK"], failed: ["FAILED", "DOWN"]};
  const m = map[c.state];
  const who = profile ? '<div class="label" title="Credential profile"><b>' + esc(profile) + "</b></div>"
                      : '<div class="label bad">No profile</div>';
  if (!m) return who + '<span class="muted">Not tested</span>';
  return who + pillT(m[0], m[1], [m[0] + (c.time ? " " + c.time : ""), c.detail].filter(Boolean).join("\n"));
}

function readinessCell(r) {
  const map = {checking: ["CHECKING", "RUN"], ready: ["READY", "OK"],
               warning: ["WARNING", "CHECK"], not_ready: ["NOT READY", "DOWN"]};
  const m = map[r.state];
  if (!m) return '<span class="muted">Not checked</span>';
  return pillT(m[0], m[1], [m[0] + (r.time ? " " + r.time : ""), r.detail].filter(Boolean).join("\n")) +
    (r.state === "ready" && r.time ? '<span class="label clip">' + esc(r.time) + "</span>" : "");
}

function postcheckCell(p) {
  const map = {checking: ["CHECKING", "RUN"], passed: ["PASSED", "OK"],
               warning: ["DIFFERENCES", "CHECK"], failed: ["FAILED", "DOWN"]};
  const m = map[p && p.state];
  if (!m) return '<span class="muted">-</span>';
  return pillT(m[0], m[1], [m[0] + (p.time ? " " + p.time : ""), p.detail].filter(Boolean).join("\n")) +
    (p.state === "passed" && p.time ? '<span class="label clip">' + esc(p.time) + "</span>" : "");
}

function versionCell(h) {
  const u = h.upgrade;
  if (u.state === "completed" && u.action !== "distribute" && u.old_version) {
    // upgrade finished: this column keeps the version from BEFORE the upgrade; the new one is in Firmware upgrade
    return "<b" + rtAttr("Version before the upgrade") + ">" + esc(u.old_version) + "</b>";
  }
  if (!h.version) return '<span class="muted">-</span>';
  const tip = [h.prev_version ? "Was " + h.prev_version : "", h.version_note].filter(Boolean).join("\n");
  return "<b" + rtAttr(tip) + ">" + esc(h.version) + "</b>";
}

function bootCell(b) {
  // Running mode + whether the boot variable is what Cisco prescribes for install mode. Follows the
  // upgrade live: the value read before it, after the boot variable was fixed, and after the reload.
  if (!b || (!b.mode && b.boot_ok == null))
    return '<span class="muted" title="Run a backup or click Refresh version to read it">-</span>';
  const good = b.mode === "INSTALL" && b.boot_ok === true;
  const tip = [b.detail, b.prev ? "Was " + b.prev : "", b.time ? "Read " + b.time : ""].filter(Boolean).join("\n");
  return pillT(b.mode || "UNKNOWN", good ? "OK" : "CHECK", tip);
}

// Search box + filters above the device list. All of them combine with the Up/Down buttons.
const DF = {q: "", upg: "all", ready: "all", plat: "all", platKey: null};
function upgradeKey(h) {
  const st = h.upgrade.state;
  if (st === "starting" || st === "running") return "running";
  if (h.scheduled) return "scheduled";
  return {completed: "completed", failed: "failed", attention: "attention"}[st] || "none";
}
function devFiltersActive() { return !!(DF.q || DF.upg !== "all" || DF.ready !== "all" || DF.plat !== "all" || filter !== "all"); }
function hostMatches(h) {
  if (filter !== "all" && h.status !== filter) return false;
  if (DF.upg !== "all" && upgradeKey(h) !== DF.upg) return false;
  if (DF.ready !== "all" && ((h.readiness && h.readiness.state) || "none") !== DF.ready) return false;
  if (DF.plat !== "all" && (h.platform || "") !== DF.plat) return false;
  if (DF.q) {
    const hay = [h.host, h.label, h.platform, h.version, h.profile, h.status, h.boot && h.boot.mode,
                 h.upgrade && h.upgrade.new_version, h.upgrade && h.upgrade.message]
                .map(v => String(v == null ? "" : v)).join(" \n").toLowerCase();
    if (!DF.q.split(/\s+/).every(t => hay.includes(t))) return false;     // every word must appear somewhere
  }
  return true;
}
const DSORT = {key: null, dir: 1};
const SORT_ORDER = {status: {UP: 0, DOWN: 1}};
function sortValue(h, k) {
  const st = o => String((o && o.state) || "");
  switch (k) {
    case "host": return String(h.label || h.host || "");
    case "platform": return String(h.platform || "");
    case "status": return String(h.status || "");
    case "creds": return st(h.creds);
    case "ready": return st(h.readiness);
    case "version": return String(h.version || "");
    case "boot": return String((h.boot && h.boot.mode) || "");
    case "backup": return st(h.backup);
    case "stage": return String(h.stage || "");
    case "upgrade": return upgradeKey(h);
    case "post": return st(h.postcheck);
  }
  return "";
}
function sortHosts(list) {
  if (!DSORT.key) return list;
  const k = DSORT.key, dir = DSORT.dir;
  const cmp = new Intl.Collator(undefined, {numeric: true, sensitivity: "base"});
  return list.map((h, i) => [h, i, sortValue(h, k)]).sort((a, b) => {
    const ae = a[2] === "", be = b[2] === "";
    if (ae !== be) return ae ? 1 : -1;               // empty values always last
    return cmp.compare(a[2], b[2]) * dir || a[1] - b[1];
  }).map(x => x[0]);
}
function syncSortHeaders() {
  document.querySelectorAll("#dev-table th[data-sort]").forEach(th => {
    th.querySelectorAll(".sort-ind").forEach(e => e.remove());
    const on = th.dataset.sort === DSORT.key;
    th.classList.toggle("sorted", on);
    th.setAttribute("aria-sort", on ? (DSORT.dir > 0 ? "ascending" : "descending") : "none");
    if (on) {
      const sp = document.createElement("span");
      sp.className = "sort-ind";
      sp.textContent = DSORT.dir > 0 ? " \u25B2" : " \u25BC";
      th.appendChild(sp);
    }
  });
}
document.querySelectorAll("#dev-table th[data-sort]").forEach(th => {
  th.addEventListener("click", () => {
    const k = th.dataset.sort;
    if (DSORT.key !== k) { DSORT.key = k; DSORT.dir = 1; }
    else if (DSORT.dir === 1) DSORT.dir = -1;
    else { DSORT.key = null; DSORT.dir = 1; }     // third click clears the sort
    syncSortHeaders();
    if (last) render(last);
  });
});
// Download the device-list template through the page's own connection. A plain <a download> makes the
// browser open a SECOND connection from its download manager, which can fail with "Network issue"
// (e.g. over HTTPS with a self-signed certificate). If the in-page fetch fails, fall back to the plain link.
$("btn-device-template").addEventListener("click", async e => {
  e.preventDefault();
  const a = e.currentTarget;
  try {
    let r = null;
    for (let i = 0; i < 3 && !(r && r.ok); i++) {
      try { r = await fetch(a.getAttribute("href"), {cache: "no-store", credentials: "same-origin"}); } catch (_) { r = null; }
      if (!(r && r.ok)) await new Promise(res => setTimeout(res, 400));
    }
    if (!(r && r.ok)) throw new Error("download failed");
    const url = URL.createObjectURL(await r.blob());
    const t = document.createElement("a");
    t.href = url; t.download = "device_list_template.xlsx";
    document.body.appendChild(t); t.click(); t.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  } catch (_) {
    window.location.href = a.getAttribute("href");
  }
});
function visibleHosts() {
  return last ? sortHosts(last.hosts.filter(hostMatches)) : [];
}
function syncDevFilters(d) {
  const plats = [...new Set(d.hosts.map(h => h.platform).filter(Boolean))].sort();
  const key = plats.join("|");
  const sel = $("dev-f-platform");
  if (key !== DF.platKey) {
    DF.platKey = key;
    if (DF.plat !== "all" && !plats.includes(DF.plat)) DF.plat = "all";
    sel.innerHTML = "";
    for (const v of ["all"].concat(plats)) {
      const o = document.createElement("option");
      o.value = v; o.textContent = v === "all" ? "Platform: any" : v;
      sel.appendChild(o);
    }
    sel.value = DF.plat;
  }
  const active = devFiltersActive();
  $("dev-clear").hidden = !DF.q;
  $("dev-reset").hidden = !active;
  $("dev-shown").textContent = active ? visibleHosts().length + " of " + d.hosts.length + " shown" : "";
}
function devFilterChanged() { if (last) render(last); }
$("dev-search").addEventListener("input", () => { DF.q = $("dev-search").value.trim().toLowerCase(); devFilterChanged(); });
$("dev-search").addEventListener("keydown", e => { if (e.key === "Escape" && $("dev-search").value) { e.stopPropagation(); $("dev-clear").click(); } });
$("dev-clear").addEventListener("click", () => { $("dev-search").value = ""; DF.q = ""; devFilterChanged(); $("dev-search").focus(); });
$("dev-f-upgrade").addEventListener("change", e => { DF.upg = e.target.value; devFilterChanged(); });
$("dev-f-ready").addEventListener("change", e => { DF.ready = e.target.value; devFilterChanged(); });
$("dev-f-platform").addEventListener("change", e => { DF.plat = e.target.value; devFilterChanged(); });
$("dev-reset").addEventListener("click", () => {
  DF.q = ""; DF.upg = "all"; DF.ready = "all"; DF.plat = "all"; filter = "all";
  $("dev-search").value = ""; $("dev-f-upgrade").value = "all"; $("dev-f-ready").value = "all"; $("dev-f-platform").value = "all";
  document.querySelectorAll(".fl [data-f]").forEach(x => x.classList.toggle("on", x.dataset.f === "all"));
  devFilterChanged();
});

// ---- Settings > Notifications (e-mail copies of the notifications; moved here from the desktop launcher) ----
const NT = {data: null, draft: null, pw: "", busy: false, msg: null};
async function ntLoad() {
  try {
    const r = await fetch("/api/notifications/email", {cache: "no-store"});
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "Could not load the notification settings.");
    NT.data = d; NT.draft = JSON.parse(JSON.stringify(d.values)); NT.pw = "";
  } catch (e) { NT.data = null; NT.msg = {kind: "err", text: e.message}; }
  if (SET.tab === "notifications") renderSettings();
}
function ntSay(kind, text) { NT.msg = text ? {kind: kind, text: text} : null; const m = $("nt-msg"); if (m) { m.className = "msg " + (kind || ""); m.textContent = text || ""; } }
function ntEnabledState() {
  const root = $("nt-root"); if (!root) return;
  const on = !!(NT.draft && NT.draft.enabled);
  root.querySelectorAll("[data-nt-dep]").forEach(el => { el.disabled = !SET_CAN_EDIT || !on; });
}
function ntRender() {
  const root = $("nt-root"); if (!root) return;
  if (!NT.data) { root.innerHTML = '<div class="msg ' + (NT.msg ? NT.msg.kind : "") + '" style="margin-top:14px">' + esc(NT.msg ? NT.msg.text : "Loading ...") + "</div>"; return; }
  const d = NT.draft, dis = SET_CAN_EDIT ? "" : " disabled";
  const secOpts = NT.data.security_options.map(o => '<option value="' + esc(o.key) + '"' + (o.key === d.security ? " selected" : "") + ">" + esc(o.label) + "</option>").join("");
  const kinds = NT.data.kinds.map(k => '<label><input type="checkbox" data-nt-kind="' + esc(k.key) + '" data-nt-dep' + (d.kinds.includes(k.key) ? " checked" : "") + ">" + esc(k.label) + "</label>").join("");
  const last = NT.data.last && NT.data.last.time ? '<div class="nt-last ' + (NT.data.last.ok ? "ok" : "bad") + '">Last e-mail ' + esc(NT.data.last.time) + ": " + esc(NT.data.last.detail || "") + "</div>" : "";
  root.innerHTML =
    '<div class="nt-form">' +
    '<label for="nt-enabled">Send e-mail notifications</label><label class="chk set-chk" style="font-weight:400"><input type="checkbox" id="nt-enabled"' + (d.enabled ? " checked" : "") + dis + "><span>" + (d.enabled ? "On" : "Off") + "</span></label>" +
    '<label for="nt-host">SMTP server</label><input type="text" id="nt-host" data-nt-dep autocomplete="off" spellcheck="false" maxlength="200" value="' + esc(d.host) + '">' +
    '<label for="nt-port">Port / security</label><div class="nt-inline"><input type="text" id="nt-port" data-nt-dep inputmode="numeric" maxlength="5" value="' + esc(String(d.port)) + '"><select id="nt-sec" data-nt-dep>' + secOpts + "</select></div>" +
    '<label for="nt-user">Gmail address</label><input type="text" id="nt-user" data-nt-dep autocomplete="off" spellcheck="false" maxlength="200" value="' + esc(d.username) + '">' +
    '<label for="nt-pw">App password</label><input type="password" id="nt-pw" data-nt-dep autocomplete="new-password" maxlength="200" value="' + esc(NT.pw) + '">' +
    '<div class="nt-hint">' + (NT.data.password_saved ? "Password saved - leave empty to keep it. " : "") + "</div>" +
    '<label for="nt-to">Send to</label><input type="text" id="nt-to" data-nt-dep autocomplete="off" spellcheck="false" maxlength="2000" value="' + esc(d.to) + '">' +
    '<div class="nt-hint">Several recipients: separate them with commas.</div></div>' +
    '<div class="set-sec">E-mail me about</div><div class="nt-kinds">' + kinds + "</div>" +
    '<div class="nt-actions"><button type="button" class="btn sm" data-nt-q="all" data-nt-dep>Select All</button>' +
    '<button type="button" class="btn sm" data-nt-q="problems" data-nt-dep>Problems Only</button>' +
    '<button type="button" class="btn sm" data-nt-q="none" data-nt-dep>Clear</button></div>' +
    '<div class="nt-actions"><button type="button" class="btn sm" id="nt-test" data-nt-dep>Send Test E-Mail</button>' +
    '<button type="button" class="btn p sm" id="nt-save"' + dis + ">Save Notification Settings</button></div>" +
    '<div class="msg ' + (NT.msg ? NT.msg.kind : "") + '" id="nt-msg" role="status" aria-live="polite" style="margin:10px 0 0">' + esc(NT.msg ? NT.msg.text : "") + "</div>" + last;
  ntEnabledState();
}
function ntValues() {
  const d = NT.draft;
  return {enabled: d.enabled, host: d.host.trim(), port: d.port, security: d.security, username: d.username.trim(),
          to: d.to.trim(), kinds: d.kinds.slice(), password: NT.pw};
}
function ntCollect() {
  const d = NT.draft, g = id => { const el = $(id); return el ? el.value : null; };
  if (g("nt-host") !== null) { d.host = g("nt-host"); d.port = g("nt-port").trim(); d.security = g("nt-sec"); d.username = g("nt-user"); d.to = g("nt-to"); NT.pw = g("nt-pw"); }
}
document.addEventListener("click", async e => {
  const root = $("nt-root"); if (!root || !root.contains(e.target) || !SET_CAN_EDIT || !NT.data) return;
  const q = e.target.closest("[data-nt-q]");
  if (q) {
    ntCollect();
    NT.draft.kinds = q.dataset.ntQ === "all" ? NT.data.kinds.map(k => k.key) : q.dataset.ntQ === "problems" ? NT.data.problem_kinds.slice() : [];
    ntRender(); return;
  }
  if (e.target.id === "nt-test" || e.target.id === "nt-save") {
    if (NT.busy) return;
    ntCollect();
    const test = e.target.id === "nt-test", btn = e.target;
    NT.busy = true; btn.disabled = true; ntSay("", test ? "Sending test e-mail ..." : "Saving ...");
    try {
      const body = ntValues(); if (test) body.enabled = true;
      const r = await post(test ? "/api/notifications/email/test" : "/api/notifications/email", {values: body});
      if (!test) { NT.data = r; NT.draft = JSON.parse(JSON.stringify(r.values)); NT.pw = ""; }
      NT.busy = false; ntRender(); ntSay("ok", r.message);
    } catch (err) { NT.busy = false; btn.disabled = false; ntSay("err", err.message); }
  }
});
document.addEventListener("change", e => {
  const root = $("nt-root"); if (!root || !root.contains(e.target) || !SET_CAN_EDIT || !NT.data) return;
  if (e.target.id === "nt-enabled") { ntCollect(); NT.draft.enabled = e.target.checked; e.target.parentNode.querySelector("span").textContent = e.target.checked ? "On" : "Off"; ntEnabledState(); }
  else if (e.target.dataset.ntKind) { const k = e.target.dataset.ntKind, set = new Set(NT.draft.kinds); e.target.checked ? set.add(k) : set.delete(k); NT.draft.kinds = NT.data.kinds.map(x => x.key).filter(x => set.has(x)); }
  else if (e.target.id === "nt-sec") {
    ntCollect();
    if (["", "587", "465"].includes(String(NT.draft.port))) NT.draft.port = e.target.value === "ssl" ? "465" : "587";
    ntRender();
  }
});

let ROLLOUT = {max_parallel: 1, stop_on_failure: true, limit: 10};   // kept up to date from the status poll; edited on the Settings page
function applyRolloutSettings(s) { ROLLOUT = s; }

function render(d) {
  if (d.me && d.me.role !== ME.role) { location.reload(); return; }   // role changed in the launcher
  last = d;
  if (d.upgrade && d.upgrade.settings) applyRolloutSettings(d.upgrade.settings);
  for (const h of [...selected]) if (!d.hosts.some(x => x.host === h)) selected.delete(h);

  const m = d.monitor;
  const st = $("state"); st.textContent = "Monitor " + m.state; st.className = "mstate " + m.state;
  const md = $("mon-dot"); md.className = "mdot " + m.state;
  $("meta").textContent = "Check: " + m.mode + " every " + m.interval + "s" + (m.started_at ? "  |  started " + m.started_at : "");
  $("btn-mon").textContent = m.state === "running" ? "Stop Monitoring" : "Start Monitoring";
  $("btn-mon").className = "btn " + (m.state === "running" ? "dg" : "p ok");
  $("btn-mon").dataset.action = m.state === "running" ? "stop" : "start";
  if ((d.site || null) !== siteName) { applySite(d.site); loadSites(); }
  const upBusy = d.upgrade.active > 0;
  $("btn-mon").disabled = (m.state !== "running" && !d.site) || (m.state === "running" && upBusy);
  $("btn-mon").title = (m.state !== "running" && !d.site) ? "Assign this dashboard to a site first"
    : (m.state === "running" && upBusy) ? "Locked while an upgrade is running" : "";
  $("site-badge").disabled = upBusy;
  $("site-badge").title = upBusy ? "Locked while an upgrade is running" : "Select or create a site";
  if (upBusy) closeSiteMenu();

  $("c-total").textContent = d.summary.total;
  $("c-up").textContent = d.summary.up;
  $("c-down").textContent = d.summary.down;
  $("c-active").textContent = d.upgrade.active;
  $("c-scheduled").textContent = d.upgrade.scheduled;
  $("c-completed").textContent = d.upgrade.completed;

  if (m.duration > 0) {
    $("p-bar").style.width = Math.min(100, m.elapsed / m.duration * 100) + "%";
    $("p-text").textContent = m.state === "running"
      ? "Monitoring: " + fmtDur(m.elapsed) + " of " + fmtDur(m.duration)
      : "Monitoring " + m.state + " after " + fmtDur(m.elapsed);
  } else {
    $("p-bar").style.width = m.state === "running" ? "100%" : "0";
    $("p-text").textContent = m.state === "running"
      ? "Monitoring for " + fmtDur(m.elapsed) + " (no time limit)"
      : "Monitoring is " + m.state;
  }
  $("p-cycles").textContent = m.cycles + " check cycles";
  { // static text only (no timers or counters): a title that changes every poll makes the browser tooltip flicker
    const tip = "Monitoring is " + m.state + " - Check: " + m.mode + " every " + m.interval + "s";
    const md2 = $("mon-dot");
    if (md2.title !== tip) { md2.title = tip; md2.setAttribute("aria-label", tip); $("state").title = tip; } }

  fillProfileLists(d);
  syncNotes(d);
  const HN = 36;   // fixed number of ping results shown per switch (newest on the right)
  const rows = visibleHosts().map(h => {
    const recent = h.history.slice(-HN), okc = recent.filter(Boolean).length;
    const tip = recent.length ? "Last " + recent.length + " pings: " + okc + " ok, " + (recent.length - okc) + " failed" +
                (h.last_rtt != null ? " (latest " + fmtMs(h.last_rtt) + ")" : "") : "No checks yet";
    const hist = '<i class="e"></i>'.repeat(HN - recent.length) +
                 recent.map(v => '<i class="' + (v ? "" : "f") + '"></i>').join("");
    return '<tr class="' + (selected.has(h.host) ? "sel" : "") + (h.host === DR.host ? " dopen" : "") + '" data-host="' + esc(h.host) + '">' +
      '<td class="ck"><input type="checkbox" data-host="' + esc(h.host) + '"' + (selected.has(h.host) ? " checked" : "") + "></td>" +
      '<td><a href="#" class="devlink" data-detail="' + esc(h.host) + '" title="Show details of this switch"><div class="host">' + esc(h.host) + "</div>" +
        (h.label ? '<div class="label">' + esc(h.label) + "</div>" : "") + "</a></td>" +
      '<td class="c-platform">' + (h.platform ? esc(h.platform) : '<span class="muted">-</span>') + "</td>" +
      '<td class="c-status">' + pill(h.status, h.status) + "</td>" +
      '<td class="c-creds">' + credsCell(h.creds, h.profile) + "</td>" +
      '<td class="c-ready">' + readinessCell(h.readiness) + "</td>" +
      '<td class="c-version">' + versionCell(h) + "</td>" +
      '<td class="c-boot">' + bootCell(h.boot) + "</td>" +
      '<td class="c-backup">' + backupCell(h.backup) + "</td>" +
      '<td class="c-stage">' + stageCell(h) + "</td>" +
      '<td class="c-upgrade">' + upgradeCell(h) + "</td>" +
      '<td class="c-post">' + postcheckCell(h.postcheck) + "</td>" +
      '<td class="rc c-recent"><div class="hist" title="' + esc(tip) + '">' + hist + "</div></td></tr>";
  });
  $("rows").innerHTML = rows.join("") || '<tr><td colspan="' + (COLS.length + 2) + '" class="muted">' +
    (d.hosts.length ? "No switches match the search or filters." : "No devices loaded yet. Open Device list to upload or add switches.") + "</td></tr>";

  rtipIndex();
  syncDevFilters(d);
  const vis = visibleHosts();
  $("sel-all").checked = vis.length > 0 && vis.every(h => selected.has(h.host));

  const b = d.backup_run;
  const bs = $("backup-status");
  if (!b) { bs.textContent = "No backup run yet."; bs.className = "hint"; }
  else if (b.running) { bs.textContent = "Backup in progress: " + b.done + " of " + b.total + " done ..."; bs.className = "hint"; }
  else {
    bs.textContent = "Backup completed at " + b.finished + ": " + b.ok + " of " + b.total + " OK" + (b.failed ? ", " + b.failed + " failed" : "") + ".";
    bs.className = "hint " + (b.failed ? "bad" : "");
  }

  $("ev-count").textContent = d.events.length ? String(d.events.length) : "";
  { const le = d.events[0];
    $("ev-last").textContent = le ? le.time + "   " + (le.host && le.host !== "-" ? le.host + "   " : "") + le.text : "No events yet"; }
  $("events").innerHTML = d.events.length
    ? d.events.map(e => '<div><span class="muted">' + esc(e.time) + '</span> &nbsp;<span class="dot ' + esc(e.level) + '"></span><b>' +
        esc(e.host) + "</b>" + (e.label ? " (" + esc(e.label) + ")" : "") + " &nbsp;" + esc(e.text) + "</div>").join("")
    : '<span class="muted">Nothing yet.</span>';

  notifyFinished(d);
  syncConsoles(d);
  updateButtons();
  drawerSync();
}

function updateButtons() {
  const n = selected.size;
  const rows = last ? last.hosts.filter(h => selected.has(h.host)) : [];
  const anyActive = rows.some(h => ["starting", "running"].includes(h.upgrade.state));
  const anyScheduled = rows.some(h => h.scheduled);
  const anyBusy = rows.some(h => ["queued", "running"].includes(h.backup.state)) || anyActive;
  $("sel-count").textContent = n + " selected";
  $("btn-backup").disabled = n === 0;
  $("btn-version").disabled = n === 0;
  $("btn-prof-test").disabled = n === 0 || profTestBusy;
  $("btn-prof-test").title = n === 0 ? "Select one or more switches in the Devices table first - the test logs in to them with the details typed above"
    : "Log in on the " + n + " selected switch" + (n === 1 ? "" : "es") + " with the details above - nothing is saved or assigned";
  $("btn-prof-assign").disabled = n === 0 || !$("prof-assign").value;
  $("btn-readiness").disabled = n === 0 || anyActive;
  $("btn-delete-devices").disabled = n === 0 || anyBusy;
  $("btn-delete-devices").title = anyBusy
    ? "One or more selected devices have a backup or upgrade in progress"
    : "Remove the selected devices from the list (existing backups are not affected)";
  $("btn-backup").textContent = "Run Backup" + (n ? " (" + n + " Selected)" : "");
  $("btn-upgrade").disabled = n === 0 || anyActive;
  $("btn-schedule").disabled = n === 0 || anyActive;
  $("btn-unschedule").disabled = n === 0 || !anyScheduled;
  const slot = $("slot-tip"), SLOT_NOTE = "A live web console opens at the bottom of this page for every switch, immediately or at its scheduled time. The switch reloads by itself; the table shows the new version once it is back.";
  if (n === 0) tipSet(slot, "Select switches in the Devices table to enable these buttons.\n\n" + SLOT_NOTE, "");
  else if (anyActive) tipSet(slot, "One or more selected switches are already upgrading.\n\n" + SLOT_NOTE, "bad");
  else if (anyScheduled) tipSet(slot, "One or more selected switches already have a scheduled upgrade - starting or scheduling again will replace it.\n\n" + SLOT_NOTE, "w");
  else tipSet(slot, "Any number of switches can upgrade at the same time.\n\n" + SLOT_NOTE, "");
}

function creds() {
  // Credentials now come from each device's profile (on the server). Just make sure every
  // selected device has one, so the person gets a clear message before anything starts.
  const missing = last ? last.hosts.filter(h => selected.has(h.host) && !h.profile).map(h => h.host) : [];
  if (missing.length) {
    openSection(2);
    sayIn(2, "err", "No credentials profile is attached to: " + missing.slice(0, 5).join(", ") +
          (missing.length > 5 ? " (+" + (missing.length - 5) + " more)" : "") +
          ". Create a credentials profile and attach it to the switch first: pick a profile, select the " +
          "switch and press 'Assign to selected'.");
    return null;
  }
  return {};
}

// ---- credential profiles ----
let profEditing = null;           // name of the profile being edited, or null for a new one
var profTestBusy = false;         // a "Test Credentials" run is in progress
function profProto() { const r = document.querySelector('input[name="prof-proto"]:checked'); return r && r.value === "telnet" ? "telnet" : "ssh"; }
function setProfProto(v) { document.querySelectorAll('input[name="prof-proto"]').forEach(r => { r.checked = r.value === (v === "telnet" ? "telnet" : "ssh"); }); }
function profPort() { return $("prof-port").value.trim(); }
function portValid() { const v = profPort(); return v === "" || (/^\d{1,5}$/.test(v) && +v >= 1 && +v <= 65535); }
function portBad() { sayIn(2, "err", "Enter a port number between 1 and 65535."); $("prof-port").focus(); }
function fillProfileLists(d) {
  const names = (d.profiles || []).map(p => p.name);
  const sig = JSON.stringify(d.profiles || []);
  if (sig === fillProfileLists.sig) return;
  fillProfileLists.sig = sig;
  const a = $("prof-assign"), e = $("prof-edit");
  const keepA = a.value, keepE = e.value;
  a.innerHTML = names.length ? names.map(n => '<option value="' + esc(n) + '">' + esc(n) + "</option>").join("")
                             : '<option value="">No profiles yet - create one below</option>';
  e.innerHTML = '<option value="">New profile...</option>' + (d.profiles || []).map(p =>
    '<option value="' + esc(p.name) + '">' + esc(p.name) + " (" + esc(p.username) + ", " + (p.protocol === "telnet" ? "Telnet" : "SSH") + " " + (p.port || (p.protocol === "telnet" ? 23 : 22)) + ", " + p.devices + " device" +
    (p.devices === 1 ? "" : "s") + ")</option>").join("");
  if (names.includes(keepA)) a.value = keepA;
  if (names.includes(keepE)) e.value = keepE;
  updateButtons();
}
function loadProfileForm() {
  const name = $("prof-edit").value;
  const p = last && (last.profiles || []).find(x => x.name === name);
  profEditing = p ? p.name : null;
  $("prof-name").value = p ? p.name : "";
  $("prof-user").value = p ? p.username : "";
  setProfProto(p ? p.protocol : "ssh");
  $("prof-port").value = p ? String(p.port || (p.protocol === "telnet" ? 23 : 22)) : "22"; protoNote();
  $("prof-pass").value = ""; $("prof-secret").value = ""; $("prof-clear").checked = false;
  $("prof-pass").placeholder = p ? "Password (leave empty to keep the saved one)" : "Password";
  $("prof-secret").placeholder = p ? (p.has_secret ? "Enable secret (saved - leave empty to keep)" : "Enable secret (optional)")
                                   : "Enable secret (optional)";
  $("prof-clear-wrap").hidden = !(p && p.has_secret);
  $("btn-prof-delete").disabled = !p;
}
function protoNote() {
  $("prof-port").placeholder = profProto() === "telnet" ? "23" : "22";
}
document.querySelectorAll('input[name="prof-proto"]').forEach(r => r.addEventListener("change", () => {
  const port = profPort();                                   // follow the protocol only while the port is still a default
  if (port === "" || port === "22" || port === "23") $("prof-port").value = profProto() === "telnet" ? "23" : "22";
  protoNote();
}));
$("prof-edit").addEventListener("change", loadProfileForm);
$("prof-assign").addEventListener("change", updateButtons);
$("prof-show").addEventListener("change", e => {
  const t = e.target.checked ? "text" : "password";
  $("prof-pass").type = t; $("prof-secret").type = t;
});
function sleep(ms) { return new Promise(res => setTimeout(res, ms)); }
async function showCredsTestResult(hosts) {
  // The login tests run in the background on the server - poll until every switch has an answer, then summarise.
  for (let i = 0; i < 60; i++) {
    await sleep(1500);
    try { await refresh(); } catch (e) {}
    const rows = last ? last.hosts.filter(h => hosts.includes(h.host)) : [];
    if (rows.length && rows.some(h => h.creds.state === "testing")) continue;
    const bad = rows.filter(h => h.creds.state !== "ok");
    if (!rows.length) return;
    if (!bad.length) {
      sayIn(2, "ok", "Login works on " + rows.length + " of " + rows.length + " switch" + (rows.length === 1 ? "" : "es") +
                     ". You can save this profile now.");
    } else {
      sayIn(2, "err", "Login failed on " + bad.length + " of " + rows.length + ": " +
            bad.slice(0, 3).map(h => h.host + " - " + (h.creds.detail || "no answer")).join("; ") +
            (bad.length > 3 ? " (+" + (bad.length - 3) + " more, see the Credentials column)" : ""));
    }
    return;
  }
  sayIn(2, "ok", "The test is still running - the results appear in the Credentials column of the Devices table.");
}
$("btn-prof-test").addEventListener("click", async () => {
  if (profTestBusy) return;
  if (!selected.size) { sayIn(2, "err", "Select one or more switches in the Devices table first."); return; }
  if (!$("prof-user").value.trim()) { sayIn(2, "err", "Enter the username first."); $("prof-user").focus(); return; }
  if (!$("prof-pass").value && !profEditing) { sayIn(2, "err", "Enter the password first."); $("prof-pass").focus(); return; }
  if (!portValid()) { portBad(); return; }
  const hosts = [...selected], btn = $("btn-prof-test");
  profTestBusy = true; btn.disabled = true; btn.textContent = "Testing ...";
  try {
    const r = await post("/api/profile/test", {hosts: hosts, original: profEditing, username: $("prof-user").value,
      password: $("prof-pass").value, secret: $("prof-secret").value, clear_secret: $("prof-clear").checked,
      protocol: profProto(), port: profPort()});
    sayIn(2, "ok", r.message);
    await showCredsTestResult(hosts);
  } catch (e) { sayIn(2, "err", e.message); }
  finally { profTestBusy = false; btn.textContent = "Test Credentials"; updateButtons(); }
});
$("btn-prof-save").addEventListener("click", async () => {
  if (!portValid()) { portBad(); return; }
  try {
    const r = await post("/api/profile/save", {original: profEditing, name: $("prof-name").value,
      username: $("prof-user").value, password: $("prof-pass").value, secret: $("prof-secret").value,
      clear_secret: $("prof-clear").checked, protocol: profProto(), port: profPort()});
    sayIn(2, "ok", r.message);
    $("prof-pass").value = ""; $("prof-secret").value = "";
    fillProfileLists.sig = null; await refresh();
    $("prof-edit").value = r.name; loadProfileForm();
  } catch (e) { sayIn(2, "err", e.message); }
});
$("btn-prof-delete").addEventListener("click", async () => {
  if (!profEditing || !confirm("Delete the profile '" + profEditing + "'?")) return;
  try {
    const r = await post("/api/profile/delete", {name: profEditing});
    sayIn(2, "ok", r.message);
    fillProfileLists.sig = null; await refresh();
    $("prof-edit").value = ""; loadProfileForm();
  } catch (e) { sayIn(2, "err", e.message); }
});
$("btn-prof-assign").addEventListener("click", async () => {
  try {
    const r = await post("/api/profile/assign", {hosts: [...selected], profile: $("prof-assign").value});
    sayIn(2, "ok", r.message); refresh();
  } catch (e) { sayIn(2, "err", e.message); }
});

document.querySelectorAll(".fl [data-f]").forEach(b => b.addEventListener("click", () => {
  filter = b.dataset.f;
  document.querySelectorAll(".fl [data-f]").forEach(x => x.classList.toggle("on", x === b));
  if (last) render(last);
}));

$("rows").addEventListener("change", e => {
  const h = e.target.dataset && e.target.dataset.host;
  if (!h) return;
  e.target.checked ? selected.add(h) : selected.delete(h);
  const tr = e.target.closest("tr"); if (tr) tr.classList.toggle("sel", e.target.checked);
  $("sel-all").checked = visibleHosts().every(x => selected.has(x.host));
  updateButtons();
});

$("sel-all").addEventListener("change", e => {
  visibleHosts().forEach(h => e.target.checked ? selected.add(h.host) : selected.delete(h.host));
  if (last) render(last);
});

$("btn-mon").addEventListener("click", async () => {
  try { await post("/api/monitor/" + $("btn-mon").dataset.action); refresh(); }
  catch (e) { say("err", e.message); }
});

$("btn-backup").addEventListener("click", async () => {
  const c = creds(3); if (!c) return;
  const customCommands = $("custom-commands").value;
  try {
    const r = await post("/api/backup", Object.assign({hosts: [...selected], custom_commands: customCommands}, c));
    sayIn(3, "ok", r.message); refresh();
  } catch (e) { sayIn(3, "err", e.message); }
});

$("btn-version").addEventListener("click", async () => {
  const c = creds(); if (!c) return;
  try { const r = await post("/api/version", Object.assign({hosts: [...selected]}, c)); say("ok", r.message); setTimeout(refresh, 3000); }
  catch (e) { say("err", e.message); }
});

$("btn-readiness").addEventListener("click", async () => {
  const c = creds(4); if (!c) return;
  const image = $("image-select").value || null;
  try {
    const r = await post("/api/readiness-check", Object.assign({hosts: [...selected], image_path: image}, c));
    sayIn(4, "ok", r.message); setTimeout(refresh, 1500);
  } catch (e) { sayIn(4, "err", e.message); }
});

$("btn-delete-devices").addEventListener("click", async () => {
  const list = [...selected];
  if (!confirm("Remove " + list.length + " device(s) from the list?\n\n" + list.join(", ") +
               "\n\nThis only removes them from this list - backups already taken are not affected.")) return;
  try {
    const r = await post("/api/devices/delete", {hosts: list});
    say("ok", r.message);
    list.forEach(h => selected.delete(h));
    refresh();
  } catch (e) { say("err", e.message); }
});

const ACTION_LABEL = {full: "firmware upgrade", distribute: "image distribution", activate: "image activation"};
const ACTION_HINT = {
  full: "Full upgrade: uploads the image, activates it and reloads the switch in one step.",
  distribute: "Distribute only: uploads the image to flash - no activation, no reload. Safe to run any time.",
  activate: "Activate only: activates an image already distributed to this switch's flash, then reloads. No image upload here.",
};
function syncActionUI() {
  const action = $("upgrade-action").value;
  tipSet($("action-tip"), ACTION_HINT[action], "");
  const needsImage = action !== "activate";
  $("image-select").disabled = !needsImage;
  $("image-select-label").textContent = needsImage ? "Select image to install" : "Select image (not used)";
  const verb = action === "full" ? "Start Now" : action === "distribute" ? "Distribute" : "Activate";
  $("btn-upgrade").textContent = verb;
  $("btn-schedule").textContent = "Schedule";
}
$("upgrade-action").addEventListener("change", syncActionUI);
syncActionUI();

$("btn-upgrade").addEventListener("click", async () => {
  const c = creds(4); if (!c) return;
  const action = $("upgrade-action").value;
  const image = $("image-select").value;
  if (action !== "activate" && !image) { sayIn(4, "err", "Choose an image from the dropdown (upload one above if the list is empty)."); return; }
  const backupFirst = $("backup-before-upgrade").checked;
  const list = [...selected].join(", ");
  const reloadNote = action === "distribute" ? "" : "\n\nEach switch will reload and be offline for several minutes.";
  const qNote = (action !== "distribute" && selected.size > ROLLOUT.max_parallel)
    ? "\n\nOnly " + ROLLOUT.max_parallel + " switch(es) run at a time; the other " + (selected.size - ROLLOUT.max_parallel) +
      " wait in a queue" + (ROLLOUT.stop_on_failure ? " and are not started if one fails." : ".") : "";
  if (!confirm("Start the " + ACTION_LABEL[action] + " now on " + selected.size + " switch(es)?\n\n" + list +
               reloadNote + qNote + stagedWarning(action, image) + (backupFirst ? "\n\nEach switch's configuration will be backed up first." : ""))) return;
  try {
    const skipCheck = $("skip-image-check").checked; $("skip-image-check").checked = false;   // applies to this one action only
    const ignoreSpace = $("ignore-space").checked; $("ignore-space").checked = false;
    const r = await post("/api/upgrade", Object.assign({hosts: [...selected], image_path: image || null, backup_first: backupFirst, skip_image_check: skipCheck, ignore_space: ignoreSpace, action}, c));
    sayIn(4, "ok", r.message); refresh();
  } catch (e) { sayIn(4, "err", e.message); }
});

$("btn-schedule").addEventListener("click", async () => {
  const c = creds(4); if (!c) return;
  const action = $("upgrade-action").value;
  const image = $("image-select").value;
  if (action !== "activate" && !image) { sayIn(4, "err", "Choose an image from the dropdown (upload one above if the list is empty)."); return; }
  const when = $("sched-time").value;
  if (!when) { sayIn(4, "err", "Pick a date and time to schedule this."); return; }
  const runAt = new Date(when).getTime() / 1000;
  if (!Number.isFinite(runAt) || runAt < Date.now() / 1000 + 30) { sayIn(4, "err", "Pick a time at least a minute in the future."); return; }
  const backupFirst = $("backup-before-upgrade").checked;
  const list = [...selected].join(", ");
  if (!confirm("Schedule the " + ACTION_LABEL[action] + " for " + new Date(when).toLocaleString() + " on " +
               selected.size + " switch(es)?\n\n" + list + stagedWarning(action, image) +
               (backupFirst ? "\n\nEach switch's configuration will be backed up first." : ""))) return;
  try {
    const skipCheck = $("skip-image-check").checked; $("skip-image-check").checked = false;
    const ignoreSpace = $("ignore-space").checked; $("ignore-space").checked = false;
    const r = await post("/api/schedule", Object.assign({hosts: [...selected], image_path: image || null, run_at: runAt, backup_first: backupFirst, skip_image_check: skipCheck, ignore_space: ignoreSpace, action}, c));
    sayIn(4, "ok", r.message); refresh();
  } catch (e) { sayIn(4, "err", e.message); }
});

$("btn-unschedule").addEventListener("click", async () => {
  try { const r = await post("/api/schedule/cancel", {hosts: [...selected]}); sayIn(4, "ok", r.message); refresh(); }
  catch (e) { sayIn(4, "err", e.message); }
});

async function refreshImages() {
  if (ME.role === "viewer") return;               // image list is only used by the (hidden) side panel
  try {
    const r = await fetch("/api/images", {cache: "no-store"});
    const d = await r.json();
    const sel = $("image-select");
    const current = sel.value;
    if (!d.images.length) {
      sel.innerHTML = '<option value="">No images uploaded yet</option>';
      return;
    }
    sel.innerHTML = d.images.map(name => '<option value="' + esc(name) + '">' + esc(name) + "</option>").join("");
    if (d.images.includes(current)) sel.value = current;
  } catch (e) { /* dashboard poll will report the lost connection */ }
}

$("device-file").addEventListener("change", () => {
  $("btn-upload-devices").disabled = !$("device-file").files.length;
});
$("btn-upload-devices").addEventListener("click", async () => {
  const file = $("device-file").files[0];
  if (!file) return;
  $("btn-upload-devices").disabled = true;
  try {
    const r = await postFile("/api/devices/upload", file);
    sayIn(1, "ok", r.message);
    // The native file input always reverts to "No file chosen" once cleared - keep the
    // name visible ourselves so it doesn't look like nothing happened.
    $("device-file-last").textContent = "Last uploaded: " + file.name;
    $("device-file").value = "";
    refresh();
  } catch (e) { sayIn(1, "err", e.message); }
  finally { $("btn-upload-devices").disabled = !$("device-file").files.length; }
});

function isIpAddress(v) {
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(v);
  if (m) return m.slice(1).every(o => +o <= 255 && String(+o) === o) && v !== "0.0.0.0";
  return /^[0-9a-f:]+$/i.test(v) && (v.match(/:/g) || []).length >= 2;      // IPv6: the server checks it exactly
}
$("new-ip").addEventListener("input", e => {
  const v = e.target.value.trim();
  e.target.style.borderColor = v && !isIpAddress(v) ? "var(--rd)" : "";
  e.target.title = v && !isIpAddress(v) ? "This is not a valid IP address" : "";
});
$("btn-add-device").addEventListener("click", async () => {
  const hostname = $("new-hostname").value.trim(), ip = $("new-ip").value.trim();
  const platform = $("new-platform").value.trim();
  if (!hostname || !ip) { sayIn(1, "err", "Both hostname and IP address are required."); return; }
  if (!isIpAddress(ip)) { sayIn(1, "err", "'" + ip + "' is not a valid IP address (for example 10.10.0.11)."); $("new-ip").focus(); return; }
  try {
    const r = await post("/api/devices/add", {hostname, ip, platform});
    sayIn(1, "ok", r.message);
    $("new-hostname").value = ""; $("new-ip").value = ""; $("new-platform").value = "";
    refresh();
  } catch (e) { sayIn(1, "err", e.message); }
});

// ---- IOS-XE images (Storage page): upload with its Cisco MD5, list, re-verify ----
const IMG = {list: []};
function imgSay(kind, text) { const m = $("img-msg"); m.className = "msg " + kind; m.textContent = text || ""; }
const imageMd5 = () => $("image-md5").value.replace(/\s+/g, "").toLowerCase();
function updateImageUploadButton() {
  $("btn-upload-image").disabled = !$("image-file").files.length || !/^[a-f0-9]{32}$/.test(imageMd5());
}
$("image-file").addEventListener("change", updateImageUploadButton);
$("image-md5").addEventListener("input", updateImageUploadButton);
function renderImageList() {
  const body = $("img-rows");
  if (!IMG.list.length) { body.innerHTML = '<tr><td colspan="4" class="muted">No images uploaded yet.</td></tr>'; return; }
  body.innerHTML = IMG.list.map(i => {
    const tag = i.checksum
      ? '<span class="pill OK" title="Saved when the image was uploaded. Use Verify to check the file again.">' + esc(i.checksum) + "</span>"
      : '<span class="pill WARN" title="Upgrades refuse an image without a checksum. Upload it again with its MD5.">No checksum</span>';
    return "<tr><td>" + esc(i.name) + "</td><td>" + fmtSize(i.size) + "</td><td>" + tag + '</td><td class="act">' +
      (i.checksum ? '<button type="button" class="btn sm" data-verify="' + esc(i.name) + '">Verify</button>' : "") + "</td></tr>";
  }).join("");
}
async function loadImageList() {
  if (ME.role === "viewer") return;
  try {
    const r = await fetch("/api/images/details", {cache: "no-store"});
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "Could not read the image list.");
    IMG.list = d.images; renderImageList();
  } catch (e) { imgSay("err", e.message); }
}
$("img-refresh").addEventListener("click", loadImageList);
$("img-rows").addEventListener("click", async ev => {
  const b = ev.target.closest("[data-verify]");
  if (!b) return;
  b.disabled = true; const was = b.textContent; b.textContent = "Verifying ...";
  imgSay("", "");
  try { const r = await post("/api/images/verify", {name: b.dataset.verify}); imgSay("ok", r.message); }
  catch (e) { imgSay("err", e.message); }
  finally { b.disabled = false; b.textContent = was; }
});
$("btn-upload-image").addEventListener("click", async () => {
  const file = $("image-file").files[0];
  const md5 = imageMd5();
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".bin")) { imgSay("err", "Only .bin image files can be uploaded."); return; }
  if (!/^[a-f0-9]{32}$/.test(md5)) { imgSay("err", "Paste the 32-character MD5 checksum shown on Cisco's download page."); return; }
  const status = $("image-upload-status");
  $("btn-upload-image").disabled = true;
  imgSay("", "");
  status.textContent = "Uploading " + file.name + " and checking its MD5 checksum ... this can take a while.";
  status.className = "hint";
  try {
    const r = await postFile("/api/images/upload", file, {"X-Image-MD5": md5});
    status.textContent = "";
    imgSay("ok", r.message);
    $("image-file").value = ""; $("image-md5").value = "";
    await loadImageList();
    await refreshImages();
    if (r.images.includes(r.filename)) $("image-select").value = r.filename;
  } catch (e) { status.textContent = ""; imgSay("err", e.message); }
  finally { updateImageUploadButton(); }
});

// ---- Upgrade output: shown in the right-hand drawer (Upgrade tab) instead of separate windows ----
const consoles = new Map();          // host -> run record {host, jobId, node, offset, busy, fetched, done, ...}
const dismissedRuns = new Set();     // "host|jobId" the user chose not to see - never auto-opened
const knownRuns = new Set();         // "host|jobId" this page has already seen
let consolesPrimed = false;          // false until the first poll has been processed
const ACTIVE_STATES = ["starting", "running"];
const TERMINAL_STATES = ["completed", "failed", "attention"];
const MAX_CONSOLE_CHARS = 200000;
const UP = {rec: null, pref: {}};    // the run shown in the Console tab; pref: host -> "upgrade" | "ssh" (which view the user last chose)

const consoleLabel = h => h.label ? h.label + " (" + h.host + ")" : h.host;

// One record per switch run. It owns a detached text node that collects the output, so the text keeps
// growing while another switch is open in the drawer and appears instantly when you come back.
function runRec(h) {
  const jobId = h.upgrade.job_id;
  let c = consoles.get(h.host);
  if (c && c.jobId === jobId) return c;
  c = {host: h.host, jobId, node: document.createTextNode(""), offset: 0, busy: false, fetched: false,
       done: false, terminal: false, termAt: 0, after: 0};
  consoles.set(h.host, c);
  return c;
}

function scrollEnd(c, force) {
  if (UP.rec !== c) return;
  const o = $("dr-term-out");
  if (force || o.scrollHeight - o.scrollTop - o.clientHeight < 40) o.scrollTop = o.scrollHeight;
}

// Show a switch's upgrade output: open the drawer on that switch, Console tab, upgrade view.
function openConsole(h) {
  if (!h || !h.upgrade || !h.upgrade.job_id) return;
  runRec(h);
  UP.pref[h.host] = "upgrade";
  openDrawer(h.host);
  drSetTab("console");
}

async function pollConsole(c) {
  if (c.busy || c.done) return;
  c.busy = true; c.fetched = true;
  try {
    const url = "/api/console?host=" + encodeURIComponent(c.host) + "&job=" + encodeURIComponent(c.jobId) +
                "&offset=" + c.offset;
    const r = await fetch(url, {cache: "no-store"});
    if (!r.ok) return;
    const d = await r.json();
    if (consoles.get(c.host) !== c) return;       // replaced while the request was running
    if (d.reset) { c.node.data = ""; }
    if (d.text) {
      c.node.appendData(d.text);
      if (c.node.length > MAX_CONSOLE_CHARS) c.node.deleteData(0, c.node.length - MAX_CONSOLE_CHARS / 2);
      scrollEnd(c, false);
    }
    c.offset = d.offset;
    // After the run ends, keep reading for a few seconds to catch the worker's last lines, then stop.
    if (c.terminal && ++c.after > 3 && Date.now() - c.termAt > 6000) c.done = true;
  } catch (e) { /* the dashboard poll reports a lost connection */ }
  finally { c.busy = false; }
}

const PILL = {starting: ["STARTING", "RUN"], running: ["RUNNING", "RUN"], completed: ["DONE", "OK"],
              failed: ["FAILED", "DOWN"], attention: ["CHECK", "WARN"]};

function syncConsoles(d) {
  const hostsWithRun = new Set();
  for (const h of d.hosts) {
    const u = h.upgrade;
    if (!u.job_id) continue;
    hostsWithRun.add(h.host);
    const runKey = h.host + "|" + u.job_id;
    const isNewRun = !knownRuns.has(runKey);
    knownRuns.add(runKey);
    let c = consoles.get(h.host);
    if (!c || c.jobId !== u.job_id) {
      c = runRec(h);
      // The drawer is never opened automatically (several switches can upgrade at once): the user clicks the
      // switch name or its "console" link in the device list to watch a run. The output is still collected.
    }
    const term = TERMINAL_STATES.includes(u.state);
    if (term && !c.terminal) { c.termAt = Date.now(); c.after = 0; c.done = false; }
    if (!term) { c.after = 0; c.done = false; }
    c.terminal = term;
    // Old finished runs are only fetched when somebody opens them.
    if (UP.rec === c || ACTIVE_STATES.includes(u.state) || c.fetched) pollConsole(c);
  }
  for (const c of [...consoles.values()]) if (!hostsWithRun.has(c.host)) {   // device deleted / site changed
    if (UP.rec === c) UP.rec = null;
    consoles.delete(c.host);
  }
  consolesPrimed = true;
}

// "SSH" or "Telnet": the protocol of the profile attached to this switch (the console follows it).
function hostProto(h) {
  const r = last && (last.hosts || []).find(x => x.host === h);
  const p = r && (last.profiles || []).find(x => x.name === r.profile);
  return p && p.protocol === "telnet" ? "Telnet" : "SSH";
}
// The Console tab has two views: the interactive SSH session, and the read-only upgrade output. While an
// upgrade runs on the switch it is always the upgrade output (the SSH console cannot be used then).
function termUpgradeView(h) {
  if (!h || !h.upgrade || !h.upgrade.job_id) return false;
  if (ACTIVE_STATES.includes(h.upgrade.state)) { UP.pref[h.host] = "upgrade"; return true; }
  if (ME.role === "viewer") return true;                      // viewers only ever see the output
  return UP.pref[h.host] === "upgrade" && !termGet(h.host).sid;
}
const TERM_NOTE = $("dr-term-note").textContent;
const TERM_BTNS = ["dr-term-connect", "dr-term-ctrlc", "dr-term-clear", "dr-term-tin"];

function upPaint(h) {
  const c = runRec(h), out = $("dr-term-out"), u = h.upgrade, active = ACTIVE_STATES.includes(u.state);
  if (UP.rec !== c) { out.replaceChildren(c.node); UP.rec = c; out.scrollTop = out.scrollHeight; }
  const m = PILL[u.state], st = $("dr-term-state");
  st.textContent = "Upgrade " + (m ? m[0].toLowerCase() : "");
  st.className = "tst" + (active ? " busy" : u.state === "completed" ? " on" : "");
  $("dr-term-title").textContent = consoleLabel(h);
  for (const id of TERM_BTNS) $(id).hidden = true;
  $("dr-term-in").disabled = true;
  const mode = $("dr-term-mode");
  mode.hidden = active || ME.role === "viewer"; mode.textContent = hostProto(h && h.host) + " Session";
  $("dr-term-note").textContent = active
    ? "Upgrade in progress - live output (read-only). The console is available again when it finishes."
    : (u.message || "");
  pollConsole(c);
  DR.tabsHtml = ""; drTabsRender(DR.data);
}
function termSshChrome(h) {
  if (UP.rec) { UP.rec = null; $("dr-term-out").textContent = "\u200b"; }   // drop the upgrade text, force a repaint
  for (const id of TERM_BTNS) $(id).hidden = false;
  $("dr-term-note").textContent = TERM_NOTE;
  const mode = $("dr-term-mode");
  mode.hidden = !(h && h.upgrade && h.upgrade.job_id) || ME.role === "viewer"; mode.textContent = "Upgrade Output";
}
const upHas = () => { const h = drHost(); return !!(h && h.upgrade && h.upgrade.job_id); };
const upActive = () => { const h = drHost(); return !!(h && h.upgrade && ACTIVE_STATES.includes(h.upgrade.state)); };

$("dr-term-mode").addEventListener("click", () => {
  const h = drHost(); if (!h) return;
  UP.pref[h.host] = termUpgradeView(h) ? "ssh" : "upgrade";
  termPaint();
});

$("rows").addEventListener("click", e => {
  const a = e.target.closest("a[data-console]");
  if (!a) return;
  e.preventDefault();
  const h = last && last.hosts.find(x => x.host === a.dataset.console);
  if (h) openConsole(h);
});

// ---- Pop-up notifications when a switch's upgrade finishes ----
const seenFinish = new Set();
let toastsPrimed = false;       // the first poll only records what already finished before this page opened

function showToast(h) {
  const u = h.upgrade;
  const kind = u.state === "completed" ? "" : u.state === "failed" ? "bad" : "warn";
  const title = u.state === "completed" ? (u.action === "distribute" ? "Image distributed" :
                                           u.action === "activate" ? "Image activated - upgrade completed" : "Upgrade completed")
              : u.state === "failed" ? "Upgrade failed" : "Upgrade needs attention";
  const t = document.createElement("div");
  t.className = "toast " + kind;
  t.innerHTML = '<div class="tt"></div><div class="tmsg"></div><button class="tx" title="Dismiss">&#10005;</button>';
  t.querySelector(".tt").textContent = (kind === "" ? "\u2705 " : kind === "bad" ? "\u274C " : "\u26A0\uFE0F ") + title + " \u2014 " + consoleLabel(h);
  t.querySelector(".tmsg").textContent = u.message || "";
  const close = () => t.remove();
  t.querySelector(".tx").addEventListener("click", e => { e.stopPropagation(); close(); });
  t.addEventListener("click", () => {                       // click = look at that switch's console
    const row = last && last.hosts.find(x => x.host === h.host);
    if (row && row.upgrade.job_id) {
      dismissedRuns.delete(h.host + "|" + row.upgrade.job_id);
      openConsole(row);
    }
    close();
  });
  $("toasts").appendChild(t);
  setTimeout(close, 20000);
}

function notifyFinished(d) {
  for (const h of d.hosts) {
    const u = h.upgrade;
    if (!TERMINAL_STATES.includes(u.state) || !u.finished_ts) continue;
    const key = h.host + "|" + u.finished_ts;
    if (seenFinish.has(key)) continue;
    seenFinish.add(key);
    if (toastsPrimed) showToast(h);
  }
  toastsPrimed = true;
}

// ---- Notification bell: upgrades started / completed (kept on the server, per site) ----
const NOTE_ICON = {login: "\uD83D\uDD11", logout: "\uD83D\uDEAA", missed: "\u23F0", started: "\uD83D\uDE80", completed: "\u2705", failed: "\u274C", attention: "\u26A0\uFE0F"};
let notes = [];
function notesSeenKey() { return "notesSeen:" + (siteName || ""); }
function notesSeen() { try { return parseInt(localStorage.getItem(notesSeenKey()) || "0", 10) || 0; } catch (e) { return 0; } }
function markNotesSeen() {
  const top = notes.length ? notes[0].id : 0;
  try { localStorage.setItem(notesSeenKey(), String(top)); } catch (e) {}
  updateBell();
}
function updateBell() {
  const seen = notesSeen();
  const n = notes.filter(x => x.id > seen).length;
  const b = $("bell-count");
  b.hidden = n === 0 || !$("note-pop").hidden;
  b.textContent = n > 99 ? "99+" : String(n);
}
function renderNotes() {
  const seen = notesSeen(), list = $("note-list");
  list.innerHTML = "";
  if (!notes.length) { const e = document.createElement("div"); e.className = "none"; e.textContent = "No notifications yet."; list.appendChild(e); return; }
  for (const n of notes) {
    const row = document.createElement("div");
    row.className = "ni" + (n.id > seen ? " unread" : "");
    row.innerHTML = '<span class="ic"></span><div><div class="nt"></div><div class="nm"></div><div class="tm"></div></div>';
    row.querySelector(".ic").textContent = NOTE_ICON[n.kind] || "\u2139\uFE0F";
    row.querySelector(".nt").textContent = n.title + " \u2014 " + (n.label ? n.label + " (" + n.host + ")" : n.host);
    row.querySelector(".nm").textContent = n.text || "";
    row.querySelector(".tm").textContent = n.time;
    row.addEventListener("click", () => {                       // click = look at that switch's console
      const h = last && last.hosts.find(x => x.host === n.host);
      if (h && h.upgrade.job_id) {
        dismissedRuns.delete(h.host + "|" + h.upgrade.job_id);
        openConsole(h);
      }
      closeNotes();
    });
    list.appendChild(row);
  }
}
function closeNotes() { $("note-pop").hidden = true; $("btn-bell").setAttribute("aria-expanded", "false"); markNotesSeen(); }
$("btn-bell").addEventListener("click", e => {
  e.stopPropagation();
  if (!$("note-pop").hidden) { closeNotes(); return; }
  renderNotes();
  $("note-pop").hidden = false; $("btn-bell").setAttribute("aria-expanded", "true");
  $("bell-count").hidden = true;
});
document.addEventListener("click", e => { if (!$("note-pop").hidden && !e.target.closest(".notemenu")) closeNotes(); });
document.addEventListener("keydown", e => { if (e.key === "Escape" && !$("note-pop").hidden) closeNotes(); });
$("note-clear").addEventListener("click", async () => {
  try { await post("/api/notifications/clear"); notes = []; renderNotes(); updateBell(); refresh(); } catch (e) {}
});
// ---- Alert sounds: a different tone per kind (none for login / logout) ----
// Generated in the browser with the Web Audio API, so no sound files are needed. Browsers only
// allow audio after the page has been clicked once, so the audio engine is unlocked on first click.
let audioCtx = null, notesSoundPrimed = false, notesSoundLastId = 0;
function soundOn() { try { return localStorage.getItem("noteSound") !== "off"; } catch (e) { return true; } }
function paintSoundBtn() { const b = $("note-sound"); if (b) b.textContent = soundOn() ? "Sound: On" : "Sound: Off"; }
function unlockAudio() {
  try {
    if (!audioCtx) { const AC = window.AudioContext || window.webkitAudioContext; if (AC) audioCtx = new AC(); }
    if (audioCtx && audioCtx.state === "suspended") audioCtx.resume();
  } catch (e) {}
}
document.addEventListener("click", unlockAudio);
document.addEventListener("keydown", unlockAudio);
// [frequency Hz, start s, length s] notes per kind, plus the oscillator wave and loudness
const NOTE_SOUND = {
  started:   {wave: "sine",     vol: 0.16, seq: [[660, 0, 0.12], [880, 0.13, 0.18]]},                      // soft rising blip
  completed: {wave: "sine",     vol: 0.20, seq: [[523, 0, 0.14], [659, 0.15, 0.14], [784, 0.30, 0.34]]},   // happy C-E-G chime
  failed:    {wave: "sawtooth", vol: 0.10, seq: [[220, 0, 0.28], [165, 0.32, 0.45]]},                      // low falling buzz
  attention: {wave: "triangle", vol: 0.20, seq: [[740, 0, 0.16], [554, 0.20, 0.16], [740, 0.40, 0.16]]},   // two-tone warning
  missed:    {wave: "square",   vol: 0.07, seq: [[900, 0, 0.09], [900, 0.16, 0.09], [900, 0.32, 0.09]]}    // three short beeps
};
function playNoteSound(kind) {
  if (!soundOn() || !audioCtx || audioCtx.state !== "running") return;
  const def = NOTE_SOUND[kind] || {wave: "sine", vol: 0.15, seq: [[800, 0, 0.15]]};
  const t0 = audioCtx.currentTime + 0.02;
  for (const [freq, at, len] of def.seq) {
    const o = audioCtx.createOscillator(), g = audioCtx.createGain();
    o.type = def.wave; o.frequency.value = freq;
    g.gain.setValueAtTime(0.0001, t0 + at);
    g.gain.exponentialRampToValueAtTime(def.vol, t0 + at + 0.015);
    g.gain.exponentialRampToValueAtTime(0.0001, t0 + at + len);
    o.connect(g); g.connect(audioCtx.destination);
    o.start(t0 + at); o.stop(t0 + at + len + 0.05);
  }
}
function soundNewNotes() {
  const top = notes.length ? notes[0].id : 0;
  if (!notesSoundPrimed) { notesSoundPrimed = true; notesSoundLastId = top; return; }   // page just opened: stay quiet
  const fresh = notes.filter(n => n.id > notesSoundLastId && n.kind !== "login" && n.kind !== "logout");
  notesSoundLastId = Math.max(notesSoundLastId, top);
  if (!fresh.length) return;
  // Several alerts at once: play only the most important one, so it never turns into noise.
  const order = ["failed", "attention", "missed", "completed", "started"];
  fresh.sort((a, b) => order.indexOf(a.kind) - order.indexOf(b.kind));
  playNoteSound(fresh[0].kind);
}
paintSoundBtn();
$("note-sound").addEventListener("click", e => {
  e.stopPropagation();
  try { localStorage.setItem("noteSound", soundOn() ? "off" : "on"); } catch (err) {}
  paintSoundBtn(); unlockAudio(); if (soundOn()) playNoteSound("started");   // short preview
});
function syncNotes(d) {
  notes = d.notifications || [];
  soundNewNotes();
  if (!$("note-pop").hidden) { renderNotes(); markNotesSeen(); }   // panel is open: keep it live and read
  else updateBell();
}

// ---- Maximize / restore a section (Devices, Events) with the icon at its top right ----
function setCardMax(card, on) {
  card.classList.toggle("max", on);
  document.body.classList.toggle("card-max", !!document.querySelector(".card.max"));
  const b = card.querySelector("[data-max]"), name = b.dataset.max;
  b.innerHTML = on ? IC.shrink : IC.expand;
  b.title = on ? "Back to the normal layout (Esc)" : "Maximize " + name + " (Esc to restore)";
  b.setAttribute("aria-pressed", on ? "true" : "false");
  if (card.id === "ev-card") {              // Events: the one icon switches between minimised and full view
    card.classList.toggle("collapsed", !on);
    b.setAttribute("aria-expanded", on ? "true" : "false");
    b.title = on ? "Minimise the events (Esc)" : "Show the events in full view";
  }
}
document.querySelectorAll("[data-max]").forEach(b => {
  b.title = "Maximize " + b.dataset.max + " (Esc to restore)";
  b.addEventListener("click", () => {
    const card = b.closest(".card"), on = !card.classList.contains("max");
    document.querySelectorAll(".card.max").forEach(c => { if (c !== card) setCardMax(c, false); });
    setCardMax(card, on);
  });
});
// ---- Events: minimised by default (the header shows the count and the latest event). ONE icon at the right of the
// header switches between minimised and full view; clicking the header does the same. Esc minimises again. ----
setCardMax($("ev-card"), false);
$("ev-head").addEventListener("click", e => {
  // the icon has its own handler (above). It swaps its own SVG, so e.target is already detached from the page when this
  // runs - test the event's original path instead of e.target.closest()
  if (e.composedPath().some(n => n.classList && n.classList.contains("mx"))) return;
  const card = $("ev-card"), on = !card.classList.contains("max");
  document.querySelectorAll(".card.max").forEach(c => { if (c !== card) setCardMax(c, false); });
  setCardMax(card, on);
});
document.addEventListener("keydown", e => {
  if (e.key !== "Escape" || !document.querySelector(".card.max")) return;
  if (!$("note-pop").hidden) return;      // Esc closes those first
  document.querySelectorAll(".card.max").forEach(c => setCardMax(c, false));
});

async function refresh() {
  try {
    const r = await fetch("/api/status", {cache: "no-store"});
    render(await r.json());
    $("lost").textContent = "";
  } catch (e) { $("lost").textContent = "Lost connection to the script"; }
  refreshImages();
}

// Theme toggle: defaults to the OS/browser setting, remembered per browser after that.
function applyTheme(mode) {
  if (mode) document.documentElement.setAttribute("data-theme", mode);
  else document.documentElement.removeAttribute("data-theme");
  $("btn-theme").innerHTML = (mode === "dark" ||
    (!mode && window.matchMedia("(prefers-color-scheme: dark)").matches)) ? IC.sun : IC.moon;
}
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem("theme"); } catch (e) {}
  applyTheme(saved);
})();
$("btn-theme").addEventListener("click", () => {
  const dark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  const current = document.documentElement.getAttribute("data-theme") || (dark ? "dark" : "light");
  const next = current === "dark" ? "light" : "dark";
  applyTheme(next);
  try { localStorage.setItem("theme", next); } catch (e) {}
});

// ---- Site profiles: small dropdown on the site badge + "create new site" popup ----
let siteName = null, siteList = [];

function siteSay(kind, text) { const m = $("site-msg"); m.className = "msg " + kind; m.textContent = text; }
function closeSiteMenu() { $("site-pop").hidden = true; $("site-badge").setAttribute("aria-expanded", "false"); }

function renderSiteMenu() {
  const pop = $("site-pop");
  pop.innerHTML = "";
  if (!siteList.length) {
    const n = document.createElement("div"); n.className = "none"; n.textContent = "No saved sites yet"; pop.appendChild(n);
  }
  for (const name of siteList) {
    const b = document.createElement("button");
    b.type = "button"; b.setAttribute("role", "menuitem");
    b.className = name === siteName ? "cur" : "";
    b.textContent = (name === siteName ? "\u2713 " : "") + name;
    b.title = name;
    b.addEventListener("click", () => { closeSiteMenu(); if (name !== siteName) changeSite("/api/site/select", name); });
    pop.appendChild(b);
  }
  pop.appendChild(document.createElement("hr"));
  const add = document.createElement("button");
  add.type = "button"; add.className = "newsite"; add.setAttribute("role", "menuitem");
  add.textContent = "+ Create New Site\u2026";
  add.addEventListener("click", () => { closeSiteMenu(); openSiteModal(); });
  pop.appendChild(add);
}

async function loadSites() {
  if (ME.role === "viewer") return;               // viewers cannot switch or create sites
  try {
    const r = await fetch("/api/sites", {cache: "no-store"});
    const d = await r.json();
    siteList = d.sites;
  } catch (e) { /* the dashboard poll reports a lost connection */ }
  renderSiteMenu();
}

function applySite(name) {
  siteName = name || null;
  const badge = $("site-badge");
  badge.textContent = "Site: " + (siteName || "Not Assigned");
  badge.className = "badge site-btn " + (siteName ? "accent" : "warn needed");
  $("dashboard").classList.toggle("nosite", !siteName);
  renderSiteMenu();
  if (typeof renderStorageSites === "function" && !$("storage").hidden) renderStorageSites();
}

async function changeSite(url, name) {
  if (siteName && last && last.monitor.state === "running" &&
      !confirm("Changing the site stops monitoring and loads the other site's device list.\n\nContinue?")) return false;
  try {
    const r = await post(url, {name});
    selected.clear();
    applySite(r.site);
    await loadSites();
    say("ok", r.message);
    refresh();
    return true;
  } catch (e) {
    if ($("site-modal").hidden) say("err", e.message); else siteSay("err", e.message);
    return false;
  }
}

function openSiteModal() {
  $("site-new").value = ""; siteSay("", "");
  $("site-modal").hidden = false;
  $("site-new").focus();
}
function closeSiteModal() { $("site-modal").hidden = true; }

$("site-badge").addEventListener("click", e => {
  e.stopPropagation();
  const open = $("site-pop").hidden;
  if (open) { loadSites(); }
  $("site-pop").hidden = !open;
  $("site-badge").setAttribute("aria-expanded", String(open));
});
document.addEventListener("click", e => { if (!e.target.closest(".sitemenu")) closeSiteMenu(); });
document.addEventListener("keydown", e => {
  if (e.key !== "Escape") return;
  closeSiteMenu(); closeSiteModal();
});
$("btn-site-cancel").addEventListener("click", closeSiteModal);
$("site-modal").addEventListener("click", e => { if (e.target === $("site-modal")) closeSiteModal(); });
$("btn-site-create").addEventListener("click", async () => {
  const name = $("site-new").value.trim();
  if (!name) { siteSay("err", "Enter a name for the new site."); return; }
  $("btn-site-create").disabled = true;
  try { if (await changeSite("/api/site/create", name)) closeSiteModal(); }
  finally { $("btn-site-create").disabled = false; }
});
$("site-new").addEventListener("keydown", e => { if (e.key === "Enter") $("btn-site-create").click(); });

// ---- Sidebar: accordion sections + 62px icon rail ----
let openSec = 4, collapsed = false;
try {
  if (localStorage.getItem("col") === "1") collapsed = true;
  const so = localStorage.getItem("open");
  if (so === "none") openSec = null; else if (so !== null && [1, 2, 3, 4, 5].includes(+so)) openSec = +so;
} catch (e) {}
function layoutSide() {
  const side = $("side");
  side.classList.toggle("narrow", collapsed);
  side.classList.toggle("wide", !collapsed);
  document.querySelectorAll(".sec").forEach(s => s.classList.toggle("o", +s.dataset.sec === openSec));
  $("btn-collapse").title = collapsed ? "Expand panel" : "Collapse panel";
  try {
    localStorage.setItem("col", collapsed ? "1" : "0");
    localStorage.setItem("open", openSec === null ? "none" : String(openSec));
  } catch (e) {}
}
function openSection(n) { collapsed = false; openSec = n; layoutSide(); }
$("btn-collapse").addEventListener("click", () => { collapsed = !collapsed; layoutSide(); });
const REP_TYPES = {
  xlsx: {url: "/api/upgrade-report.xlsx", desc: "Excel workbook: summary, current status and the full run history."},
  pdf: {url: "/api/upgrade-report/print", blank: true, desc: "Printable page - use Print, then Save as PDF."},
  csv: {url: "/api/upgrade-report.csv", desc: "Plain CSV file of the current status."}};
function repSync() { $("rep-desc").textContent = REP_TYPES[$("rep-type").value].desc; }
$("rep-type").addEventListener("change", repSync); repSync();
$("rep-go").addEventListener("click", () => {
  const t = REP_TYPES[$("rep-type").value];
  if (t.blank) window.open(t.url, "_blank", "noopener"); else window.location.href = t.url;
});
document.querySelectorAll(".sec > h3").forEach(h => {
  const n = +h.parentNode.dataset.sec;
  const act = () => { if (collapsed) openSection(n); else { openSec = openSec === n ? null : n; layoutSide(); } };
  h.addEventListener("click", act);
  h.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); act(); } });
});
layoutSide();

// ---- Home / Storage tabs + Storage page (browse, search, view and download a site's backups, reports and run logs) ----
const stg = {site: null, path: "", query: "", mode: "browse", entries: [], sortKey: null, sortDir: 1, seq: 0, timer: null};

function fmtSize(n) {
  if (n == null) return "-";
  if (n < 1024) return n + " B";
  if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
  if (n < 1073741824) return (n / 1048576).toFixed(1) + " MB";
  return (n / 1073741824).toFixed(2) + " GB";
}
const fmtDate = t => t ? new Date(t * 1000).toLocaleString() : "-";
const stUrl = (kind, rel) => "/api/storage/" + kind + "?site=" + encodeURIComponent(stg.site) + "&path=" + encodeURIComponent(rel || "");

async function getJSON(url) {
  const r = await fetch(url, {cache: "no-store"});
  let d = {}; try { d = await r.json(); } catch (e) {}
  if (!r.ok) throw new Error(d.error || ("Request failed (HTTP " + r.status + ")"));
  return d;
}
function stSay(kind, text) { const m = $("st-msg"); m.className = "msg " + kind; m.textContent = text || ""; }

function showView(view, push) {
  if (view === "settings" && ME.role === "viewer") view = "home";
  if (view === "diagram" && ME.role === "viewer") view = "home";
  if (view !== "storage" && view !== "settings" && view !== "diagram") view = "home";
  const storage = view === "storage", settings = view === "settings", diagram = view === "diagram";
  if (storage || settings || diagram) { try { closeDrawer(); } catch (e) {} }
  $("dashboard").hidden = storage || settings || diagram;
  $("storage").hidden = !storage;
  $("diagram").hidden = !diagram;
  $("settings").hidden = !settings;
  for (const [id, on] of [["tab-home", view === "home"], ["tab-diagram", diagram], ["tab-storage", storage]]) {
    $(id).classList.toggle("on", on);
    $(id).setAttribute("aria-selected", String(on));
  }
  $("btn-settings").classList.toggle("on", settings);
  if (push !== false) { try { history.replaceState(null, "", "#" + view); } catch (e) {} }
  if (storage) openStorage();
  if (settings) openSettings();
  if (diagram) openDiagram(); else dgStop();
}
document.querySelectorAll(".tab").forEach(b => b.addEventListener("click", () => showView(b.dataset.view)));

async function openStorage() {
  loadImageList();
  await loadSites();                       // refreshes siteList from the server
  renderStorageSites();
  if (stg.site) stRefresh(); else stRender();
}

function renderStorageSites() {
  const box = $("st-sites");
  if (!stg.site || !siteList.includes(stg.site)) {
    const pick = siteList.includes(siteName) ? siteName : (siteList[0] || null);
    if (pick !== stg.site) { stg.site = pick; stg.path = ""; stg.query = ""; $("st-search").value = ""; }
  }
  box.innerHTML = "";
  if (!siteList.length) {
    const d = document.createElement("div"); d.className = "st-empty";
    d.textContent = "No sites yet. Create one from the Site menu on the Home page.";
    box.appendChild(d); return;
  }
  for (const name of siteList) {
    const b = document.createElement("button");
    b.type = "button"; b.className = "st-site" + (name === stg.site ? " on" : ""); b.title = name;
    const ic = document.createElement("span"); ic.innerHTML = IC.site; b.appendChild(ic);
    const nm = document.createElement("span"); nm.className = "nm"; nm.textContent = name; b.appendChild(nm);
    if (name === siteName) { const t = document.createElement("span"); t.className = "tag"; t.textContent = "ACTIVE"; b.appendChild(t); }
    b.addEventListener("click", () => {
      if (name === stg.site) return;
      stg.site = name; stg.path = ""; stg.query = ""; stg.sortKey = null; $("st-search").value = "";
      renderStorageSites(); stRefresh();
    });
    box.appendChild(b);
  }
}

async function stRefresh() {
  if (!stg.site) { stRender(); return; }
  const mine = ++stg.seq;
  stSay("", "");
  try {
    let d;
    if (stg.query) { stg.mode = "search"; d = await getJSON("/api/storage/search?site=" + encodeURIComponent(stg.site) + "&q=" + encodeURIComponent(stg.query)); }
    else { stg.mode = "browse"; d = await getJSON(stUrl("list", stg.path)); }
    if (mine !== stg.seq) return;           // a newer request replaced this one
    stg.entries = d.entries || [];
    stg.truncated = !!d.truncated;
    stg.exists = d.exists !== false;
  } catch (e) {
    if (mine !== stg.seq) return;
    stg.entries = []; stSay("err", e.message);
  }
  stRender();
}

function stSorted() {
  const list = stg.entries.slice();
  if (!stg.sortKey) return list;
  const k = stg.sortKey, dir = stg.sortDir;
  const val = e => k === "name" ? e.name.toLowerCase() : k === "size" ? (e.size == null ? -1 : e.size)
                 : k === "loc" ? (e.parent || "").toLowerCase() : e.mtime;
  list.sort((a, b) => {
    if ((a.type === "dir") !== (b.type === "dir")) return a.type === "dir" ? -1 : 1;   // folders stay on top
    const x = val(a), y = val(b);
    return (x < y ? -1 : x > y ? 1 : 0) * dir;
  });
  return list;
}

function stRenderCrumbs() {
  const nav = $("st-crumbs"); nav.innerHTML = "";
  const add = (label, path, cur) => {
    const b = document.createElement("button"); b.type = "button"; b.textContent = label; b.title = label;
    if (cur) b.className = "cur"; else b.addEventListener("click", () => stOpen(path));
    nav.appendChild(b);
  };
  const sep = () => { const s = document.createElement("span"); s.className = "sepc"; s.textContent = "/"; nav.appendChild(s); };
  if (!stg.site) return;
  if (stg.query) { add(stg.site, "", false); sep(); add('Search: "' + stg.query + '"', "", true); return; }
  const parts = stg.path ? stg.path.split("/") : [];
  add(stg.site, "", parts.length === 0);
  let acc = "";
  parts.forEach((p, i) => { sep(); acc = acc ? acc + "/" + p : p; add(p, acc, i === parts.length - 1); });
}

// ---- Compare two backups of one switch (shown while a switch's folder under Switch Backups is open) ----
// Full-screen, command list on the left, the selected command shown LEFT (older) / RIGHT (newer), or unified.
const CMP = {cmd: "", view: "side", show: false, all: false, last: null, cache: {}, open: new Set(), tok: 0, folder: ""};
const CMP_ST = {changed: ["CHANGED", "var(--c-a)"], added: ["ONLY IN NEWER", "var(--gr)"], removed: ["ONLY IN OLDER", "var(--rd)"], same: ["no changes", "var(--mu)"]};
const stCompareFolder = () => { const p = (stg.path || "").split("/").filter(Boolean); return p.length === 2 && p[0] === "Switch Backups"; };
function cmpMsg(text) { $("cmp-diff").innerHTML = '<div class="cmp-msg">' + esc(text) + "</div>"; }
function cmpSumHtml(d) {
  const s = d.summary;
  return "<b>" + esc(d.a.time) + "</b> &rarr; <b>" + esc(d.b.time) + "</b> " +
    (s.changed + s.added + s.removed === 0 ? pill("NO DIFFERENCES", "OK") :
      pill(s.changed + " CHANGED", "WARN") + (s.added ? pill(s.added + " ADDED", "OK") : "") + (s.removed ? pill(s.removed + " REMOVED", "DOWN") : "") +
      '<span class="muted">+' + s.lines_added + " / &minus;" + s.lines_removed + " lines</span>");
}
function cmpListHtml() {
  if (!CMP.last) return "";
  const list = CMP.last.commands.filter(c => c.status !== "same" || CMP.show);
  if (!list.length) return '<div class="cmp-msg" style="padding:10px;color:var(--mu)">No changed commands.</div>';
  return list.map(c => {
    const st = CMP_ST[c.status];
    return '<button type="button" class="cmp-item' + (c.command === CMP.cmd ? " on" : "") + '" data-cmp-cmd="' + esc(c.command) + '"><div class="c">' + esc(c.command) +
      '</div><div class="s" style="color:' + st[1] + '">' + st[0] + (c.status === "same" ? "" : " &middot; +" + c.added + " &minus;" + c.removed) + "</div></button>";
  }).join("");
}
function cmpHl(a, b) {                 // wrap the part of each line that differs (common start / end are left plain)
  if (a.length > 2000 || b.length > 2000) return [esc(a), esc(b)];
  let p = 0; while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let q = 0; while (q < a.length - p && q < b.length - p && a[a.length - 1 - q] === b[b.length - 1 - q]) q++;
  const w = (t) => { const mid = t.slice(p, t.length - q); return esc(t.slice(0, p)) + (mid ? "<mark>" + esc(mid) + "</mark>" : "") + esc(t.slice(t.length - q)); };
  return [w(a), w(b)];
}
function cmpRowHtml(r, first) {
  const k = r[4], at = first && k !== "=" ? " data-chg" : "";
  const t = x => esc(x) || " ";
  if (k === "=") return "<tr><td class=\"ln\">" + r[0] + '</td><td class="tx' + (r[1] !== r[3] ? " vol" : "") + '">' + t(r[1]) + '</td><td class="ln">' + r[2] + '</td><td class="tx' + (r[1] !== r[3] ? " vol" : "") + '">' + t(r[3]) + "</td></tr>";
  if (k === "-") return "<tr" + at + '><td class="ln l-r">' + r[0] + '</td><td class="tx l-r">' + t(r[1]) + '</td><td class="ln emp"></td><td class="tx emp"> </td></tr>';
  if (k === "+") return "<tr" + at + '><td class="ln emp"></td><td class="tx emp"> </td><td class="ln l-a">' + r[2] + '</td><td class="tx l-a">' + t(r[3]) + "</td></tr>";
  const h = cmpHl(r[1], r[3]);
  return "<tr" + at + '><td class="ln l-r">' + r[0] + '</td><td class="tx l-r">' + (h[0] || " ") + '</td><td class="ln l-a">' + r[2] + '</td><td class="tx l-a">' + (h[1] || " ") + "</td></tr>";
}
function cmpSideHtml(d) {
  const rows = d.rows, n = rows.length;
  let keep = null;
  if (!CMP.all) {                      // changed lines with 3 lines of context; the rest folds into a clickable gap
    keep = new Uint8Array(n);
    for (let i = 0; i < n; i++) if (rows[i][4] !== "=") for (let k = Math.max(0, i - 3); k <= Math.min(n - 1, i + 3); k++) keep[k] = 1;
  }
  const na = d.in_a ? "" : " \u2014 command not in this backup", nb = d.in_b ? "" : " \u2014 command not in this backup";
  let out = '<table class="sbs"><colgroup><col style="width:56px"><col><col style="width:56px"><col></colgroup><thead><tr><th colspan="2">Left &middot; older &middot; ' +
    esc(d.a.time + na) + '</th><th colspan="2">Right &middot; newer &middot; ' + esc(d.b.time + nb) + "</th></tr></thead><tbody>";
  let i = 0;
  while (i < n) {
    if (keep && !keep[i]) {
      let j = i; while (j < n && !keep[j]) j++;
      if (CMP.open.has(i)) { for (let k = i; k < j; k++) out += cmpRowHtml(rows[k], false); }
      else out += '<tr class="gap" data-gap="' + i + '"><td colspan="4">&hellip; ' + (j - i) + " unchanged line" + (j - i === 1 ? "" : "s") + " &mdash; click to show</td></tr>";
      i = j; continue;
    }
    out += cmpRowHtml(rows[i], i === 0 || rows[i - 1][4] === "=");
    i++;
  }
  if (!n) out += '<tr><td colspan="4" class="cmp-msg" style="padding:12px 14px;color:var(--mu)">Both outputs are empty.</td></tr>';
  if (d.truncated) out += '<tr class="gap"><td colspan="4">Only the first ' + BACKUP_SIDE_MAX_LINES_JS + " lines of each side are compared.</td></tr>";
  return out + "</tbody></table>";
}
const BACKUP_SIDE_MAX_LINES_JS = 30000;
function cmpUniHtml(c) {
  if (!c || c.status === "same") return '<div class="cmp-msg">No changes in this command.</div>';
  return '<pre class="uni">' + c.diff.map(l => '<div class="' + (l[0] === "+" ? "df-a" : l[0] === "-" ? "df-r" : l[0] === "@" ? "df-h" : "") + '">' + esc(l || " ") + "</div>").join("") +
    (c.truncated ? '<div class="df-h">... difference is longer than shown ...</div>' : "") + "</pre>";
}
async function cmpPaint(keepScroll) {
  const box = $("cmp-diff"), d = CMP.last;
  if (!d) return;
  $("cmp-list").innerHTML = cmpListHtml();
  const c = d.commands.find(x => x.command === CMP.cmd);
  if (!c) { cmpMsg("Select a command on the left."); return; }
  const top = keepScroll ? box.scrollTop : 0;
  if (CMP.view === "uni") { box.innerHTML = cmpUniHtml(c); box.scrollTop = top; return; }
  const key = [CMP.folder, d.a.stamp, d.b.stamp, d.ignore_volatile ? 1 : 0, c.command].join("|"), tok = ++CMP.tok;
  let sd = CMP.cache[key];
  if (!sd) {
    cmpMsg("Loading " + c.command + " ...");
    try {
      sd = await getJSON(stUrl("compare", CMP.folder) + "&a=" + encodeURIComponent(d.a.stamp) + "&b=" + encodeURIComponent(d.b.stamp) +
        "&ignore=" + (d.ignore_volatile ? "1" : "0") + "&cmd=" + encodeURIComponent(c.command));
    } catch (e) { if (tok === CMP.tok) cmpMsg("Could not load this command: " + e.message); return; }
    CMP.cache[key] = sd;
  }
  if (tok !== CMP.tok || $("cmp-modal").hidden) return;
  box.innerHTML = cmpSideHtml(sd); box.scrollTop = top;
}
function cmpPickDefault() {
  const cs = CMP.last.commands;
  if (cs.some(c => c.command === CMP.cmd)) return;
  const ch = cs.filter(c => c.status !== "same");
  CMP.cmd = (ch.find(c => /running-config/i.test(c.command)) || ch[0] || cs.find(c => /running-config/i.test(c.command)) || cs[0] || {command: ""}).command;
}
async function cmpOpen() {
  const folder = stg.path;
  CMP.folder = folder; CMP.last = null; CMP.cache = {}; CMP.open = new Set(); CMP.cmd = "";
  $("cmp-title").textContent = "Compare backups \u2014 " + folder.split("/").pop();
  $("cmp-modal").hidden = false; cmpMsg("Loading ..."); $("cmp-list").innerHTML = ""; $("cmp-sum").innerHTML = "";
  $("cmp-a").innerHTML = $("cmp-b").innerHTML = ""; $("cmp-go").disabled = true;
  try {
    const d = await getJSON(stUrl("backups", folder));
    if ($("cmp-modal").hidden || folder !== stg.path) return;
    if (d.runs.length < 2) { cmpMsg("This switch has fewer than two backups, so there is nothing to compare yet."); return; }
    const opts = d.runs.map(r => '<option value="' + esc(r.stamp) + '">' + esc(r.time) + "</option>").join("");
    $("cmp-a").innerHTML = opts; $("cmp-b").innerHTML = opts;
    $("cmp-b").value = d.runs[0].stamp; $("cmp-a").value = d.runs[1].stamp;     // newest against the one before it
    $("cmp-go").disabled = false;
    cmpRun();
  } catch (e) { cmpMsg("Could not load the backups: " + e.message); }
}
async function cmpRun() {
  const folder = stg.path, a = $("cmp-a").value, b = $("cmp-b").value;
  cmpMsg("Comparing ...");
  try {
    const d = await getJSON(stUrl("compare", folder) + "&a=" + encodeURIComponent(a) + "&b=" + encodeURIComponent(b) + "&ignore=" + ($("cmp-ign").checked ? "1" : "0"));
    if ($("cmp-modal").hidden) return;
    CMP.last = d; CMP.cache = {}; CMP.open = new Set(); CMP.folder = folder;
    $("cmp-sum").innerHTML = cmpSumHtml(d);
    cmpPickDefault(); cmpPaint(false);
  } catch (e) { CMP.last = null; $("cmp-list").innerHTML = ""; $("cmp-sum").innerHTML = ""; cmpMsg("Could not compare: " + e.message); }
}
function cmpJump(dir) {                // next / previous block of changed lines in the side-by-side view
  const box = $("cmp-diff"), bl = box.querySelectorAll("tr[data-chg]");
  if (!bl.length) return;
  const base = box.getBoundingClientRect().top + 40;
  let t = null;
  if (dir > 0) { for (const x of bl) if (x.getBoundingClientRect().top > base + 4) { t = x; break; } t = t || bl[0]; }
  else { for (let i = bl.length - 1; i >= 0; i--) if (bl[i].getBoundingClientRect().top < base - 4) { t = bl[i]; break; } t = t || bl[bl.length - 1]; }
  box.scrollTop += t.getBoundingClientRect().top - base;
}
$("st-compare").addEventListener("click", cmpOpen);
$("cmp-go").addEventListener("click", cmpRun);
$("cmp-a").addEventListener("change", () => { if (!$("cmp-go").disabled) cmpRun(); });
$("cmp-b").addEventListener("change", () => { if (!$("cmp-go").disabled) cmpRun(); });
$("cmp-ign").addEventListener("change", () => { if (!$("cmp-go").disabled) cmpRun(); });
$("cmp-all").addEventListener("change", e => { CMP.all = e.target.checked; cmpPaint(true); });
$("cmp-same").addEventListener("change", e => { CMP.show = e.target.checked; if (CMP.last) $("cmp-list").innerHTML = cmpListHtml(); });
$("cmp-prev").addEventListener("click", () => cmpJump(-1));
$("cmp-next").addEventListener("click", () => cmpJump(1));
$("cmp-view").addEventListener("click", e => {
  const b = e.target.closest("button[data-v]"); if (!b) return;
  CMP.view = b.dataset.v;
  $("cmp-view").querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b));
  $("cmp-all").disabled = $("cmp-prev").disabled = $("cmp-next").disabled = CMP.view === "uni";
  cmpPaint(false);
});
$("cmp-list").addEventListener("click", e => {
  const it = e.target.closest(".cmp-item"); if (!it) return;
  CMP.cmd = it.dataset.cmpCmd; CMP.open = new Set(); cmpPaint(false);
});
$("cmp-diff").addEventListener("click", e => {
  const g = e.target.closest("tr.gap[data-gap]"); if (!g) return;
  CMP.open.add(Number(g.dataset.gap)); cmpPaint(true);
});
$("cmp-close").addEventListener("click", () => { $("cmp-modal").hidden = true; $("cmp-diff").textContent = ""; $("cmp-list").innerHTML = ""; CMP.last = null; CMP.cache = {}; });
document.addEventListener("keydown", e => { if (e.key === "Escape" && !$("cmp-modal").hidden) $("cmp-close").click(); });

function stOpen(path) {
  stg.path = path || ""; stg.query = ""; $("st-search").value = ""; $("st-clear").hidden = true;
  stRefresh();
}

function stRender() {
  stRenderCrumbs();
  const searching = stg.mode === "search" && !!stg.query;
  const head = $("st-head"), rows = $("st-rows"), none = $("st-none");
  head.innerHTML = ""; rows.innerHTML = "";

  // column headers (click to sort)
  const tr = document.createElement("tr");
  const cols = [["name", "Name"]].concat(searching ? [["loc", "Folder"]] : []).concat([["size", "Size"], ["mtime", "Modified"]]);
  for (const [key, label] of cols) {
    const th = document.createElement("th"); th.className = "sortable";
    th.textContent = label + (stg.sortKey === key ? (stg.sortDir > 0 ? " \u25B2" : " \u25BC") : "");
    th.addEventListener("click", () => {
      if (stg.sortKey === key) stg.sortDir = -stg.sortDir; else { stg.sortKey = key; stg.sortDir = key === "mtime" ? -1 : 1; }
      stRender();
    });
    tr.appendChild(th);
  }
  const thA = document.createElement("th"); thA.style.textAlign = "right"; thA.textContent = "Actions"; tr.appendChild(thA);
  head.appendChild(tr);

  // buttons that only make sense in some states
  const zip = $("st-zip");
  zip.hidden = !stg.site || searching;
  zip.href = stg.site ? stUrl("download", stg.path) : "#";
  zip.textContent = stg.path ? "Download Folder (.zip)" : "Download Site Storage (.zip)";
  $("st-refresh").disabled = !stg.site;
  $("st-compare").hidden = !stg.site || searching || !stCompareFolder() || ME.role === "viewer";
  $("st-search").disabled = !stg.site;
  $("st-clear").hidden = !stg.query;

  const list = stSorted();
  const foot = $("st-foot");
  if (!stg.site) {
    none.hidden = false; none.textContent = "Select or create a site to see its files."; foot.textContent = ""; return;
  }
  if (!list.length) {
    none.hidden = false;
    none.textContent = searching ? 'No files or folders match "' + stg.query + '".'
      : stg.exists === false ? "Nothing has been saved for this site yet."
      : "This folder is empty.";
  } else none.hidden = true;

  for (const e of list) {
    const row = document.createElement("tr");
    const isDir = e.type === "dir";
    const nameTd = document.createElement("td"); nameTd.className = "nm";
    const ico = document.createElement("span"); ico.className = "ico" + (isDir ? " dir" : ""); ico.innerHTML = isDir ? IC.folder : IC.file;
    const nm = document.createElement("button"); nm.type = "button"; nm.className = "st-name"; nm.textContent = e.name; nm.title = e.name;
    nm.addEventListener("click", () => isDir ? stOpen(e.rel) : stView(e));
    nameTd.appendChild(ico); nameTd.appendChild(nm); row.appendChild(nameTd);
    if (searching) {
      const loc = document.createElement("td"); loc.className = "st-loc"; loc.textContent = e.parent || "/"; loc.title = e.parent || "";
      row.appendChild(loc);
    }
    const sz = document.createElement("td");
    sz.textContent = isDir ? (e.items == null ? "-" : e.items + (e.items === 1 ? " item" : " items")) : fmtSize(e.size);
    row.appendChild(sz);
    const md = document.createElement("td"); md.textContent = fmtDate(e.mtime); row.appendChild(md);
    const act = document.createElement("td"); act.className = "act";
    if (isDir) {
      const o = document.createElement("button"); o.type = "button"; o.className = "btn sm"; o.textContent = "Open";
      o.addEventListener("click", () => stOpen(e.rel)); act.appendChild(o);
    } else {
      const v = document.createElement("button"); v.type = "button"; v.className = "btn sm"; v.textContent = "View";
      v.addEventListener("click", () => stView(e)); act.appendChild(v);
    }
    const dl = document.createElement("a"); dl.className = "btn sm"; dl.href = stUrl("download", e.rel);
    dl.textContent = isDir ? "ZIP" : "Download"; dl.title = isDir ? "Download this folder as a ZIP" : "Download " + e.name;
    act.appendChild(dl);
    row.appendChild(act);
    rows.appendChild(row);
  }
  const nDirs = list.filter(e => e.type === "dir").length, nFiles = list.length - nDirs;
  foot.textContent = list.length
    ? nDirs + (nDirs === 1 ? " folder" : " folders") + ", " + nFiles + (nFiles === 1 ? " file" : " files")
      + (stg.truncated ? " - showing the first " + list.length + " matches, refine your search to narrow it down" : "")
    : "";
}

const STV = {e: null, tail: false, truncated: false};
async function stLoad(tail) {
  const e = STV.e, pre = $("st-pre");
  $("st-modal-meta").textContent = "Loading...";
  try {
    const d = await getJSON(stUrl("read", e.rel) + (tail ? "&tail=1" : ""));
    if ($("st-modal").hidden || STV.e !== e) return;
    STV.tail = !!d.tail; STV.truncated = !!d.truncated;
    $("st-modal-meta").textContent = fmtSize(d.size) + " - " + (e.parent || e.rel.split("/").slice(0, -1).join("/") || "/")
      + (d.truncated ? " - showing the " + (d.tail ? "last " : "first ") + fmtSize(d.limit || 0) + " of the file; use " + (d.tail ? "Top" : "End") + " for the other part, or download the full file" : "");
    pre.textContent = d.text;
    pre.scrollTop = tail ? pre.scrollHeight : 0;
  } catch (err) {
    $("st-modal-meta").textContent = "";
    pre.textContent = err.message;
  }
}
function stView(e) {
  $("st-modal-title").textContent = e.name;
  $("st-modal-title").title = e.name;
  $("st-modal-dl").href = stUrl("download", e.rel);
  $("st-pre").textContent = "";
  STV.e = e; STV.tail = false; STV.truncated = false;
  $("st-modal").hidden = false;
  return stLoad(false);
}
$("st-modal-top").addEventListener("click", () => { if (STV.tail) stLoad(false); else $("st-pre").scrollTop = 0; });
$("st-modal-end").addEventListener("click", () => { if (STV.truncated && !STV.tail) stLoad(true); else $("st-pre").scrollTop = $("st-pre").scrollHeight; });
function closeStModal() { $("st-modal").hidden = true; $("st-pre").textContent = ""; }
$("st-modal-close").addEventListener("click", closeStModal);
$("st-modal").addEventListener("click", e => { if (e.target === $("st-modal")) closeStModal(); });
document.addEventListener("keydown", e => { if (e.key === "Escape" && !$("st-modal").hidden) closeStModal(); });

$("st-search").addEventListener("input", () => {
  clearTimeout(stg.timer);
  stg.timer = setTimeout(() => {
    const q = $("st-search").value.trim();
    if (q === stg.query) return;
    stg.query = q; stRefresh();
  }, 250);
});
$("st-search").addEventListener("keydown", e => { if (e.key === "Escape") $("st-clear").click(); });
$("st-clear").addEventListener("click", () => { $("st-search").value = ""; stg.query = ""; stRefresh(); $("st-search").focus(); });
$("st-refresh").addEventListener("click", stRefresh);

// ---- Settings page (gear icon): menu on the left, one panel per group; the values are global ----
const SET = {data: null, draft: {}, tab: "rollout"};
const SET_CAN_EDIT = ME.role !== "operator" && ME.role !== "viewer";
if (ME.role === "viewer") $("btn-settings").hidden = true;
$("btn-settings").addEventListener("click", () => showView($("settings").hidden ? "settings" : "home"));
function setSay(kind, text) { const m = $("set-msg"); m.className = "msg " + kind; m.textContent = text || ""; }
const setUnit = (item, v) => item.type === "bool" ? (v ? "On" : "Off") : v + " " + item.unit;
function setMinutes(item, v) {
  if (item.unit !== "s" || !(v >= 120)) return "";
  const m = v / 60;
  return "\u2248 " + (Number.isInteger(m) ? m : m.toFixed(1)) + " min";
}
function setDirtyKeys() {
  return SET.data ? SET.data.items.filter(i => SET.draft[i.key] !== SET.data.values[i.key]).map(i => i.key) : [];
}
function setUpdateDirty() {
  const n = setDirtyKeys().length;
  $("set-dirty").textContent = n ? n + " unsaved change" + (n === 1 ? "" : "s") : "";
  $("set-save").disabled = !SET_CAN_EDIT || !n;
}
function setRow(item) {
  const v = SET.draft[item.key], def = SET.data.defaults[item.key], k = esc(item.key), dis = SET_CAN_EDIT ? "" : " disabled";
  const ctl = item.type === "bool"
    ? '<label class="chk set-chk"><input type="checkbox" data-sk="' + k + '"' + (v ? " checked" : "") + dis + "><span>" + (v ? "On" : "Off") + "</span></label>"
    : '<input type="number" data-sk="' + k + '" min="' + item.min + '" max="' + item.max + '" step="1" value="' + (Number.isNaN(v) ? "" : v) + '"' + dis + ">" +
      '<span class="set-unit">' + esc(item.unit) + '</span><span class="set-hint">' + esc(setMinutes(item, v)) + "</span>";
  return '<div class="set-row' + (v !== def ? " chg" : "") + '" data-row="' + k + '"><div class="set-info"><div class="set-label">' + esc(item.label) +
    '</div><div class="set-help">' + esc(item.help) + '</div></div><div class="set-ctl">' + ctl +
    '<button type="button" class="set-def" data-sdef="' + k + '" title="Use the default value"' + dis + ">Default: " + esc(setUnit(item, def)) + "</button></div></div>";
}
function aboutRender() {
  const root = $("ab-root"), a = SET.data && SET.data.about; if (!root || !a) return;
  const link = /^https:\/\//.test(a.repo) ? '<a href="' + esc(a.repo) + '" target="_blank" rel="noopener noreferrer">' + esc(a.repo) + "</a>" : esc(a.repo);
  root.innerHTML = '<div class="ab-card"><div class="ab-kv"><span>Application</span><span>' + esc(a.name) + "</span>" +
    "<span>Version</span><span><b>" + esc(a.version) + "</b></span><span>Developed by</span><span>" + esc(a.author) + "</span>" +
    "<span>Project page</span><span>" + link + "</span></div></div>" +
    '<div class="set-sec">New features in each version</div>' +
    a.changelog.map(v => '<div class="ab-ver"><h4>' + (/^\d/.test(v.version) && !v.version.includes(" ") ? "Version " : "") + esc(v.version) +
      (v.date ? " <small>" + esc(v.date) + "</small>" : "") + (v.version === a.version ? ' <span class="ab-cur">Installed</span>' : "") + "</h4><ul>" +
      v.features.map(f => "<li>" + esc(f) + "</li>").join("") + "</ul></div>").join("");
}
function renderSettings() {
  if (!SET.data) return;
  if (!SET.data.groups.some(g => g.id === SET.tab)) SET.tab = SET.data.groups[0].id;
  $("set-nav").innerHTML = SET.data.groups.map(g => '<button type="button" role="tab" data-stab="' + esc(g.id) + '" class="' +
    (g.id === SET.tab ? "on" : "") + '" aria-selected="' + (g.id === SET.tab) + '">' + esc(g.title) + "</button>").join("");
  const g = SET.data.groups.find(x => x.id === SET.tab);
  $("set-title").textContent = g.title;
  $("set-note").hidden = SET_CAN_EDIT;
  let html = '<div class="set-intro">' + esc(g.intro) + "</div>" +
    (g.id === "timeouts" ? '<div id="tp-root"></div><div id="ct-root"></div>' : "") +
    (g.id === "notifications" ? '<div id="nt-root"></div>' : "") +
    (g.id === "about" ? '<div id="ab-root"></div>' : ""), sec = null;
  for (const item of SET.data.items.filter(i => i.group === g.id)) {
    if (item.section && item.section !== sec) { html += '<div class="set-sec">' + esc(item.section) + "</div>"; }
    sec = item.section;
    html += setRow(item);
  }
  $("set-body").innerHTML = html;
  if (g.id === "timeouts") { tpRender(); ctRender(); }
  if (g.id === "notifications") { ntRender(); if (!NT.data && !NT.loading) { NT.loading = true; ntLoad().finally(() => { NT.loading = false; }); } }
  if (g.id === "about") aboutRender();
  const own = g.id === "notifications" || g.id === "about";            // these tabs have their own Save button
  $("set-reset").hidden = own; $("set-save").hidden = own; $("set-dirty").hidden = own;
  $("set-reset").disabled = !SET_CAN_EDIT;
  setUpdateDirty();
}
async function openSettings() {
  setSay("", "");
  try {
    const r = await fetch("/api/settings", {cache: "no-store"});
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "Could not load the settings.");
    SET.data = d; SET.draft = {...d.values}; renderSettings();
    tpLoad(); NT.data = null; if (SET.tab === "notifications") renderSettings(); else ntLoad();
  } catch (e) { setSay("err", e.message); }
}
function setRowRefresh(key) {
  const item = SET.data.items.find(i => i.key === key), row = $("set-body").querySelector('[data-row="' + key + '"]');
  if (!item || !row) return;
  const v = SET.draft[key];
  row.classList.toggle("chg", v !== SET.data.defaults[key]);
  const hint = row.querySelector(".set-hint"); if (hint) hint.textContent = setMinutes(item, v);
  const lab = row.querySelector(".set-chk span"); if (lab) lab.textContent = v ? "On" : "Off";
  setUpdateDirty();
}
$("set-nav").addEventListener("click", e => {
  const b = e.target.closest("[data-stab]");
  if (b) { SET.tab = b.dataset.stab; setSay("", ""); renderSettings(); }
});
$("set-body").addEventListener("input", e => {
  const k = e.target.dataset.sk;
  if (!k || e.target.type !== "number") return;
  SET.draft[k] = e.target.value.trim() === "" ? NaN : Number(e.target.value);
  setRowRefresh(k);
});
$("set-body").addEventListener("change", e => {
  const k = e.target.dataset.sk;
  if (!k || e.target.type !== "checkbox") return;
  SET.draft[k] = e.target.checked; setRowRefresh(k);
});
$("set-body").addEventListener("click", e => {
  const b = e.target.closest("[data-sdef]");
  if (!b || !SET_CAN_EDIT) return;
  SET.draft[b.dataset.sdef] = SET.data.defaults[b.dataset.sdef]; renderSettings();
});
$("set-reset").addEventListener("click", () => {
  if (!SET.data) return;
  for (const i of SET.data.items) if (i.group === SET.tab) SET.draft[i.key] = SET.data.defaults[i.key];
  renderSettings();
});
$("set-save").addEventListener("click", async () => {
  if (!SET.data || !SET_CAN_EDIT) return;
  for (const i of SET.data.items) {
    if (i.type === "bool") continue;
    const v = SET.draft[i.key];
    if (!Number.isInteger(v) || v < i.min || v > i.max) {
      setSay("err", i.label + " must be a whole number from " + i.min + " to " + i.max + "."); return;
    }
  }
  try {
    const r = await post("/api/settings", {values: SET.draft});
    SET.data = r; SET.draft = {...r.values};
    applyRolloutSettings({max_parallel: r.values.max_parallel, stop_on_failure: r.values.stop_on_failure, limit: r.limit});
    renderSettings(); setSay("ok", r.message);
  } catch (e) { setSay("err", e.message); }
});

// ============================================================================================
// Command duration profiles + Command time test (Settings > Command Outputs).
//  - Profiles: named sets of the command timeouts. "Default" is built in. Apply copies a profile into the settings.
//  - Command time test: run the preset commands on one switch with the Execute buttons, see each result and note
//    the time it took. The numbers can become a report (CSV / printable page) or a new profile.
// ============================================================================================
const TP = {data: null, sel: "Default", form: false};
const CT = {open: false, host: null, image: "", results: {}, outs: {}, running: null, margin: 3, keepDefault: true, form: false, sel: null, railClosed: {}, path: "three"};
try {
  const saved = JSON.parse(localStorage.getItem("ct") || "null");
  if (saved && typeof saved === "object") {
    CT.results = saved.results && typeof saved.results === "object" ? saved.results : {};
    CT.host = saved.host || null;
    CT.image = saved.image || "";
    if (saved.margin >= 1 && saved.margin <= 20) CT.margin = saved.margin;
  }
} catch (e) {}
function ctPersist() { try { localStorage.setItem("ct", JSON.stringify({results: CT.results, host: CT.host, image: CT.image, margin: CT.margin})); } catch (e) {} }
const tmoItems = () => SET.data ? SET.data.items.filter(i => i.group === "timeouts") : [];
const fmtS = n => typeof n === "number" ? n.toFixed(2) : "";
function tipsIn(id) { const r = $(id); if (r) r.querySelectorAll(".tip").forEach(tipInit); }

async function tpLoad() {
  try {
    const r = await fetch("/api/timeout-profiles", {cache: "no-store"});
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "Could not load the profiles.");
    TP.data = d;
    if (!d.profiles.some(p => p.name === TP.sel)) TP.sel = d.active || "Default";
    tpRender(); ctRender();
  } catch (e) { setSay("err", e.message); }
}
function tpDraftValues() {
  const out = {};
  for (const i of tmoItems()) {
    const v = SET.draft[i.key];
    if (!Number.isInteger(v) || v < i.min || v > i.max) throw new Error(i.label + " must be a whole number from " + i.min + " to " + i.max + ".");
    out[i.key] = v;
  }
  return out;
}
function tpRender() {
  const root = $("tp-root"); if (!root || !SET.data) return;
  if (!TP.data) { root.innerHTML = '<div class="tp-bar tp-state">Loading the profiles ...</div>'; return; }
  const list = TP.data.profiles, active = TP.data.active;
  if (!list.some(p => p.name === TP.sel)) TP.sel = active || "Default";
  const cur = list.find(p => p.name === TP.sel), dis = SET_CAN_EDIT ? "" : " disabled";
  const edited = tmoItems().some(i => SET.draft[i.key] !== SET.data.values[i.key]);
  const state = (active ? "In use: <b>" + esc(active) + "</b>" : "In use: <b>Custom values</b> (not the same as any profile)") +
    (edited ? " &nbsp;|&nbsp; unsaved edits on this page" : "");
  const opts = list.map(p => '<option value="' + esc(p.name) + '"' + (p.name === TP.sel ? " selected" : "") + ">" + esc(p.name) +
    (p.builtin ? " (built-in)" : "") + (p.name === active ? " - in use" : "") + "</option>").join("");
  const meta = cur.builtin ? "The timeouts the tool ships with. It cannot be changed or deleted."
    : [cur.note, "Saved " + cur.saved_at + (cur.saved_by ? " by " + cur.saved_by : ""), cur.site ? "site " + cur.site : "",
       cur.source === "command test" ? "built from a command time test" + (cur.switch ? " on " + cur.switch : "") : "typed in by hand"]
      .filter(Boolean).join(" | ");
  const grid = tmoItems().map(i => "<span>" + esc(i.label) + "</span><span>" + cur.values[i.key] + " s</span>").join("");
  root.innerHTML = '<div class="tp-bar">' +
    '<div class="tp-row"><label for="tp-sel">Command duration profile<span class="tip" data-tip="A profile is a named set of the command timeouts below. Apply puts its values into the settings and saves them. Create profiles from the values on this page, or from a Command time test measured on a real switch (for example one profile per site)."></span></label>' +
    '<select id="tp-sel">' + opts + "</select>" +
    '<button type="button" class="btn p sm" data-tp="apply"' + (SET_CAN_EDIT && cur.name !== active ? "" : " disabled") + ' title="Use the values of this profile now">Apply</button>' +
    '<button type="button" class="btn sm" data-tp="new"' + dis + ' title="Save the values shown below as a new profile">New From Page Values...</button>' +
    '<button type="button" class="btn sm" data-tp="update"' + (SET_CAN_EDIT && !cur.builtin ? "" : " disabled") + ' title="Replace the values of this profile with the values shown below">Update From Page Values</button>' +
    '<button type="button" class="btn sm" data-tp="delete"' + (SET_CAN_EDIT && !cur.builtin ? "" : " disabled") + ">Delete</button>" +
    '<span class="tp-state">' + state + "</span></div>" +
    (TP.form ? '<div class="tp-row"><input type="text" id="tp-name" maxlength="40" placeholder="Profile name, e.g. Colombo site" aria-label="New profile name">' +
      '<input type="text" id="tp-note" maxlength="300" placeholder="Note (optional)" aria-label="Note" style="min-width:260px">' +
      '<button type="button" class="btn p sm" data-tp="create">Save Profile</button><button type="button" class="btn sm" data-tp="cancel">Cancel</button></div>' : "") +
    '<div class="tp-info">' + esc(meta) + "</div>" +
    '<details class="tp-vals"><summary>Values in this profile</summary><div class="tp-grid">' + grid + "</div></details></div>";
  tipsIn("tp-root");
}
async function tpAction(act) {
  try {
    if (!TP.data) return;
    if (act === "new") { TP.form = true; tpRender(); const n = $("tp-name"); if (n) n.focus(); return; }
    if (act === "cancel") { TP.form = false; tpRender(); return; }
    if (act === "apply") {
      if (setDirtyKeys().length && !confirm("There are unsaved changes on this page. Applying the profile replaces the timeout values and saves them. Continue?")) return;
      const r = await post("/api/timeout-profiles/apply", {name: TP.sel});
      SET.data = r.settings; SET.draft = {...r.settings.values}; TP.data = r;
      renderSettings(); setSay("ok", r.message); return;
    }
    if (act === "delete") {
      if (!confirm("Delete the profile '" + TP.sel + "'? The timeouts now in use do not change.")) return;
      const r = await post("/api/timeout-profiles/delete", {name: TP.sel});
      TP.data = r; TP.sel = "Default"; tpRender(); setSay("ok", r.message); return;
    }
    if (act === "update" || act === "create") {
      const values = tpDraftValues();
      let name = TP.sel, note = "", overwrite = true;
      if (act === "create") {
        name = ($("tp-name").value || "").trim(); note = ($("tp-note").value || "").trim();
        const ex = TP.data.profiles.find(p => p.name.toLowerCase() === name.toLowerCase());
        overwrite = !!ex;
        if (ex && !ex.builtin && !confirm("A profile named '" + ex.name + "' already exists. Replace its values?")) return;
      } else {
        if (!confirm("Replace the values of '" + name + "' with the values shown on this page?")) return;
        note = (TP.data.profiles.find(p => p.name === name) || {}).note || "";
      }
      const r = await post("/api/timeout-profiles/save", {name, values, note, source: "manual", site: siteName || "", overwrite});
      TP.data = r; TP.sel = r.saved; TP.form = false; tpRender(); setSay("ok", r.message);
    }
  } catch (e) { setSay("err", e.message); }
}

// ---- Command time test: a manual upgrade of ONE switch, step by step, every step timed ----
const ctHosts = () => (last && last.hosts) ? last.hosts : [];
const ctHostInfo = () => ctHosts().find(h => h.host === CT.host) || null;
const ctRuns = (host, id) => ((CT.results[host] || {})[id] || {}).runs || [];
const ctGood = r => r.st === "ok" || r.st === "warning" || r.st === "timeout";      // the time of these runs counts
const STATUS_TXT = {ok: ["ok", "OK"], warning: ["warn", "Check the result"], timeout: ["warn", "Timed out"], skipped: ["", "Nothing to do"],
  stopped: ["warn", "Stopped"], unsupported: ["warn", "Not supported"], error: ["bad", "Failed"], connect_failed: ["bad", "Login failed"]};
function ctStats(runs) {
  const good = runs.filter(ctGood).map(r => r.s).filter(v => typeof v === "number");
  const lastRun = runs.length ? runs[runs.length - 1] : null;
  return {n: runs.length, last: lastRun, lastS: lastRun && ctGood(lastRun) ? lastRun.s : null,
          max: good.length ? Math.max(...good) : null, avg: good.length ? good.reduce((a, b) => a + b, 0) / good.length : null};
}
function ctFeeds(keys) { const m = {}; for (const i of tmoItems()) m[i.key] = i.label; return keys.map(k => m[k] || k).join(" | "); }
function ctStatusHtml(r) {
  if (!r) return "not run yet";
  const t = STATUS_TXT[r.st] || ["", r.st];
  return '<span class="dot ' + t[0] + '"></span>' + esc(t[1]) + (r.d ? " - " + esc(r.d) : "") + (r.at ? ' <span class="muted">(' + esc(r.at) + ")</span>" : "");
}
function ctPartStats(p) {                 // steps that time several things (bundles, activate): per part, over all runs
  const order = [], by = {};
  for (const r of ctRuns(CT.host, p.id)) for (const x of (r.parts || [])) {
    if (!(x.c in by)) { by[x.c] = []; order.push(x.c); }
    by[x.c].push({s: x.s, st: x.st});
  }
  return order.map(c => ({command: c, ...ctStats(by[c])}));
}
function ctMeasured() {
  const R = CT.results[CT.host] || {}, per = {};
  const add = (k, id, v) => { if (typeof v !== "number") return; per[k] = per[k] || {}; per[k][id] = Math.max(per[k][id] || 0, v); };
  for (const p of (TP.data ? TP.data.presets : [])) for (const r of ctRuns(CT.host, p.id)) {
    if (r.parts && r.parts.length) { for (const x of r.parts) if (ctGood(x)) add(x.k, p.id, x.s); }
    else if (ctGood(r)) for (const k of p.keys) add(k, p.id, r.s);
  }
  for (const v of (R._connect || [])) add("ssh_connect", "_login", v);
  const m = {};
  for (const k of Object.keys(per)) {
    if (k === "install_total") {            // 'Whole install' = Add + Activate (two steps), or the one-command step; never the sum of both ways
      const two = (per[k].install_add || 0) + (per[k].install_activate || 0), one = per[k].install_full || 0;
      m[k] = Math.max(two, one);
    } else m[k] = Math.max(...Object.values(per[k]));
  }
  return m;
}
function ctSuggested(m) {
  const values = {}, rows = [];
  for (const i of tmoItems()) {
    const def = SET.data.defaults[i.key];
    let v = def;
    if (i.key in m) {
      v = Math.ceil(m[i.key] * CT.margin / 5) * 5;
      if (CT.keepDefault) v = Math.max(v, def);
      v = Math.min(i.max, Math.max(i.min, v));
    }
    values[i.key] = v;
    if (i.key in m) rows.push({item: i, measured: m[i.key], def, value: v});
  }
  return {values, rows};
}
function ctPartsHtml(parts) {
  return "<table><thead><tr><th>Measured</th><th class=\"n\">Took (s)</th><th>Result</th></tr></thead><tbody>" +
    parts.map(x => "<tr><td>" + esc(x.command || x.c) + '</td><td class="n">' + fmtS(x.seconds != null ? x.seconds : x.s) + "</td><td>" +
      esc((STATUS_TXT[x.status || x.st] || ["", x.status || x.st])[1]) + (x.detail ? " - " + esc(x.detail) : "") + "</td></tr>").join("") + "</tbody></table>";
}
function ctOutHtml(r) {
  return '<div class="ct-panel-h"><b>' + (r.seconds != null ? "Took " + fmtS(r.seconds) + " s" : "No time") + "</b>" +
    (r.kind !== "bundle" && r.kind !== "activate" ? ", " + r.lines + " lines" : "") + (r.detail ? " - " + esc(r.detail) : "") +
    (r.truncated ? ' <span class="muted">(first ' + r.output.length + " of " + r.chars + " characters)</span>" : "") + "</div>" +
    (r.parts && r.parts.length ? ctPartsHtml(r.parts) : "") +
    (r.output && !(r.kind === "bundle") ? "<pre>" + esc(r.output) + "</pre>" : r.parts && r.parts.length ? "" : "<pre>(no output)</pre>");
}
function ctLiveHtml() {
  const r = CT.running; if (!r) return "";
  return '<div class="ct-panel-h"><b>Running ... ' + Math.round(r.elapsed) + " s</b>" + (r.stopping ? " - stopping" : "") +
    ' <span class="muted">- live output. If nothing appears for minutes the switch is waiting: press Stop to see what it printed.</span></div>' +
    "<pre>" + esc(r.output || "(no output yet)") + "</pre>";
}
function ctRenderLive() {
  const box = $("ct-live"); if (!box || !CT.running) return;
  const old = box.querySelector("pre"), atEnd = !old || old.scrollTop + old.clientHeight >= old.scrollHeight - 8;
  box.innerHTML = ctLiveHtml();
  const nu = box.querySelector("pre"); if (nu && atEnd) nu.scrollTop = nu.scrollHeight;
}
function ctStepDone(p) { const st = ctStats(ctRuns(CT.host, p.id)); return !!(st.last && (ctGood(st.last) || st.last.st === "skipped")); }
function ctSelected(presets) {            // the step shown on the right: your choice, else the first unfinished one
  if (!presets.some(p => p.id === CT.sel)) CT.sel = (presets.find(p => !ctStepDone(p)) || presets[0]).id;
  return presets.find(p => p.id === CT.sel);
}
const CT_PATH_GROUP = {"3. Install - Activate/Distribute": "three", "3. Install - Full Upgrade": "one"};
function ctPathOff(g) { return !!CT_PATH_GROUP[g] && CT_PATH_GROUP[g] !== CT.path; }   // the install path you did not choose
function ctRailHtml(presets) {
  const busy = !!CT.running && CT.running.host === CT.host;
  let html = "", num = 0;
  for (const g of [...new Set(presets.map(p => p.group))]) {
    const steps = presets.filter(p => p.group === g), open = !CT.railClosed[g], done = steps.filter(ctStepDone).length;
    const off = ctPathOff(g);
    html += '<div class="ct-rs' + (open ? " open" : "") + (off ? " off" : "") + '"><button type="button" class="ct-rs-h" data-ct="stage" data-g="' + esc(g) + '" aria-expanded="' + open + '">' +
      '<span class="ct-car">' + IC.dn + '</span><span>' + esc(g) + '</span><span class="ct-rs-c">' + done + "/" + steps.length + "</span></button>";
    for (const p of steps) {
      num++;
      if (!open) continue;
      const st = ctStats(ctRuns(CT.host, p.id)), running = busy && CT.running.id === p.id;
      const cls = (off ? " off" : "") + (p.id === CT.sel ? " sel" : "") + (running ? " run" : st.last ? (ctStepDone(p) ? " done" : " bad") : "");
      html += '<button type="button" class="ct-r' + cls + '" data-ct="sel" data-id="' + esc(p.id) + '" title="' + esc(p.title) + '"><span class="ct-n">' + (ctStepDone(p) && !running ? "&#10003;" : num) + '</span>' +
        '<span class="ct-r-t">' + esc(p.title) + '</span><span class="ct-r-v">' + (running ? Math.round(CT.running.elapsed) + " s" : st.max != null ? fmtS(st.max) + " s" : "") + "</span></button>";
    }
    html += "</div>";
  }
  return html;
}
function ctDetailHtml(p, num, total) {
  const host = CT.host, runs = ctRuns(host, p.id), st = ctStats(runs), key = host + "|" + p.id, busy = !!CT.running;
  const running = busy && CT.running.host === host && CT.running.id === p.id;
  const badges = (p.warn ? (p.kind === "activate" ? '<span class="ct-badge d">reloads the switch</span>' : '<span class="ct-badge">changes the switch</span>') : "") +
    (p.needs_image && !CT.image ? '<span class="ct-badge n">choose an image above</span>' : "");
  const act = running ? '<button type="button" class="btn dg" data-ct="stop"' + (CT.running.stopping ? " disabled" : "") + ">" + (CT.running.stopping ? "Stopping ..." : "Stop") + "</button>"
    : '<button type="button" class="btn p" data-ct="run" data-id="' + esc(p.id) + '"' + (busy ? " disabled" : "") + ">Execute</button>";
  const idx = TP.data.presets.findIndex(x => x.id === p.id), nxt = TP.data.presets.slice(idx + 1).find(x => !ctPathOff(x.group));
  const chips = p.keys.length ? p.keys.map(k => '<span class="ct-chip">' + esc(ctFeeds([k])) + "</span>").join("") : '<span class="ct-chip">no timeout setting - for your information</span>';
  const body = running ? '<div class="ct-panel" id="ct-live">' + ctLiveHtml() + "</div>"
    : CT.outs[key] ? '<div class="ct-panel">' + ctOutHtml(CT.outs[key]) + "</div>"
    : '<div class="ct-empty">Not run yet. Press Execute: the output appears here and in the console below.</div>';
  return '<div class="ct-d"><div class="ct-d-h"><div class="ct-d-id"><div class="ct-d-k">Step ' + num + " of " + total + " &middot; " + esc(p.group.replace(/^\d+\.\s*/, "")) + "</div>" +
    '<div class="ct-d-t">' + esc(p.title) + badges + '</div><code class="ct-cmd">' + esc(p.command) + "</code></div>" +
    '<div class="ct-d-a">' + act + (nxt ? '<button type="button" class="btn" data-ct="next">Next Step</button>' : "") + "</div></div>" +
    '<p class="ct-d-desc">' + esc(p.desc) + "</p>" +
    '<div class="ct-chips">Feeds: ' + chips + "</div>" +
    '<div class="ct-tiles"><div class="ct-tile"><span>Last</span><b>' + (st.lastS != null ? fmtS(st.lastS) + " s" : "-") + '</b></div><div class="ct-tile"><span>Slowest</span><b>' + (st.max != null ? fmtS(st.max) + " s" : "-") +
    '</b></div><div class="ct-tile"><span>Average</span><b>' + (st.avg != null ? fmtS(st.avg) + " s" : "-") + '</b></div><div class="ct-tile"><span>Runs</span><b>' + st.n + "</b></div></div>" +
    (running || st.last ? '<div class="ct-status">' + (running ? '<span class="dot warn"></span>Running on ' + esc(host) : ctStatusHtml(st.last)) + "</div>" : "") + body + "</div>";
}
function ctRender() {
  const root = $("ct-root"); if (!root || !SET.data) return;
  const presets = TP.data ? TP.data.presets : [], hosts = ctHosts(), images = TP.data ? TP.data.images : [];
  if (!hosts.some(h => h.host === CT.host)) CT.host = hosts.length ? hosts[0].host : null;
  if (CT.image && !images.includes(CT.image)) CT.image = "";
  let html = '<div class="ct-card' + (CT.open ? " open" : "") + '"><button type="button" class="ct-toggle" data-ct="toggle" aria-expanded="' + CT.open + '">' +
    '<span>Command time test</span><span class="sub">upgrade ONE switch by hand, step by step, and time every step</span><span class="car">' + IC.dn + "</span></button>";
  if (CT.open) {
    if (!hosts.length || !presets.length) {
      html += '<div class="ct-note">There are no switches in the device list yet. Add the test switch in the Device list and give it a credentials profile, then come back.</div>';
    } else {
      const busy = !!CT.running, p = ctSelected(presets), num = presets.indexOf(p) + 1;
      const connect = (CT.results[CT.host] || {})._connect || [];
      html += '<div class="ct-bar"><div class="ct-f"><label for="ct-host">Switch</label><select id="ct-host"' + (busy ? " disabled" : "") + ">" + hosts.map(h => '<option value="' + esc(h.host) + '"' + (h.host === CT.host ? " selected" : "") + ">" +
          esc(h.host + (h.label ? " - " + h.label : "") + (h.platform ? " (" + h.platform + ")" : "")) + "</option>").join("") + "</select></div>" +
        '<div class="ct-f"><label for="ct-image">Image</label><select id="ct-image"' + (busy ? " disabled" : "") + '><option value="">- choose the image to test with -</option>' +
        images.map(n => '<option value="' + esc(n) + '"' + (n === CT.image ? " selected" : "") + ">" + esc(n) + "</option>").join("") + "</select></div>" +
        '<div class="ct-bar-r"><span class="tip" data-tip="Run the steps from top to bottom on a TEST switch: it is a real upgrade done by hand. Pick a step on the left, press Execute, read the output here and in the console, then go on with Next step. The image is needed for the copy, the MD5 check and the add step. Steps marked as changing the switch ask first; Activate RELOADS the switch. Run a step again to get a second sample - the slowest time of each step is the one that counts."></span>' +
        '<button type="button" class="btn sm" data-ct="stages" data-v="1">Expand All</button><button type="button" class="btn sm" data-ct="stages" data-v="0">Collapse All</button>' +
        '<button type="button" class="btn sm" data-ct="clear"' + (busy ? " disabled" : "") + ">Clear Results</button></div></div>";
      html += '<div class="ct-wb"><nav class="ct-rail" aria-label="Steps">' + ctRailHtml(presets) + '</nav><div class="ct-main">' + ctDetailHtml(p, num, presets.length) + '<div id="cc-slot"></div></div></div>';
      html += '<div class="ct-result"><div class="ct-result-t">Results And Profile</div><div class="ct-result-g"><div id="ct-sugg-t"></div><div class="ct-side">' +
        '<div class="ct-sg-h"><label for="ct-margin">Safety margin &times;<span class="tip" data-tip="The suggested timeout is the slowest measured time multiplied by this number, rounded up to 5 seconds. A switch is slower during a busy upgrade than during your test, so keep some room (3 is a good start)."></span></label>' +
        '<input type="number" id="ct-margin" min="1" max="20" step="0.5" value="' + CT.margin + '">' +
        '<label class="chk"><input type="checkbox" id="ct-keep"' + (CT.keepDefault ? " checked" : "") + '><span>Never go below the default</span></label></div>' +
        '<div class="ct-note">Every Execute also times the SSH login' + (connect.length ? " (" + connect.length + " sample" + (connect.length > 1 ? "s" : "") + ", slowest " + fmtS(Math.max(...connect)) + " s)" : "") + ".</div>" +
        '<div id="ct-sugg-a"></div></div></div></div>';
    }
  }
  root.innerHTML = html + "</div>";
  ctRenderSuggest();
  tipsIn("ct-root");
  const slot = $("cc-slot"); if (slot) slot.appendChild(CC_EL);
  ccSync();
}
function ctRenderSuggest() {
  const tbl = $("ct-sugg-t"), act = $("ct-sugg-a"); if (!tbl || !act) return;
  const m = ctMeasured(), sg = ctSuggested(m);
  if (!sg.rows.length) { tbl.innerHTML = '<div class="ct-empty" style="margin-top:0">Execute at least one step to see the suggested timeouts.</div>'; act.innerHTML = ""; return; }
  const dis = SET_CAN_EDIT ? "" : " disabled";
  tbl.innerHTML = '<table class="ct-table"><thead><tr><th>Timeout setting</th><th class="n">Slowest (s)</th><th class="n">Default (s)</th><th class="n">Suggested (s)</th></tr></thead><tbody>' +
    sg.rows.map(r => "<tr><td>" + esc(r.item.label) + '</td><td class="n">' + fmtS(r.measured) + '</td><td class="n">' + r.def + '</td><td class="n"><b>' + r.value + "</b></td></tr>").join("") + "</tbody></table>" +
    '<div class="ct-note">' + (tmoItems().length - sg.rows.length) + " timeout(s) have no measurement and keep their default value in the profile (for example the backup commands, which this test does not run).</div>";
  act.innerHTML = '<div class="ct-actions"><button type="button" class="btn p" data-ct="saveform"' + dis + (SET_CAN_EDIT ? "" : ' title="Only an Administrator can save profiles"') + ">Save As Profile...</button>" +
    '<button type="button" class="btn sm" data-ct="csv">Report (CSV)</button><button type="button" class="btn sm" data-ct="print">Printable Report</button></div>' +
    (CT.form ? '<div class="ct-actions"><input type="text" id="ct-pname" maxlength="40" placeholder="Profile name, e.g. Production" aria-label="Profile name">' +
      '<button type="button" class="btn p sm" data-ct="save"' + dis + '>Save</button><button type="button" class="btn sm" data-ct="cancelform">Cancel</button></div>' : "");
}
async function ctRun(id) {
  if (CT.running || !CT.host || !TP.data) return;
  const p = TP.data.presets.find(x => x.id === id); if (!p) return;
  const host = CT.host;
  if (p.needs_image && !CT.image) { setSay("err", "Choose the image to test with first."); return; }
  if (p.warn && !confirm("'" + p.title + "' " + p.warn.replace("{host}", host) + "\n\nContinue?")) return;
  CT.running = {host, id, elapsed: 0, output: "", stopping: false, job: null}; CT.sel = id; setSay("", ""); ctRender();
  try {
    const job = await post("/api/command-test/start", {host, id, image: CT.image});
    CT.running.job = job.job;
    let r = null, offset = 0;
    while (!r) {
      await new Promise(res => setTimeout(res, 1000));
      const d = await post("/api/command-test/poll", {job: job.job, offset});
      offset = d.offset; ccAppend(host, d.text);
      if (d.state === "done") r = d.result;
      else { CT.running.elapsed = d.elapsed; CT.running.output = d.output || ""; ctRenderLive(); }
    }
    const R = CT.results[host] = CT.results[host] || {};
    const e = R[id] = R[id] || {runs: []};
    e.runs.push({s: r.seconds, st: r.status, d: r.detail || "", at: r.time,
                 parts: (r.parts || []).map(x => ({c: x.command, k: x.key, s: x.seconds, st: x.status}))});
    if (e.runs.length > 20) e.runs.shift();
    if (typeof r.connect_seconds === "number") { (R._connect = R._connect || []).push(r.connect_seconds); if (R._connect.length > 40) R._connect.shift(); }
    CT.outs[host + "|" + id] = r; ctPersist();
  } catch (e) { setSay("err", e.message); }
  CT.running = null; ctRender();
}
async function ctStop() {
  if (!CT.running || !CT.running.job) return;
  CT.running.stopping = true; ctRender();
  try { await post("/api/command-test/stop", {job: CT.running.job}); } catch (e) { setSay("err", e.message); }
}
function ctReportData() {
  const h = ctHostInfo() || {host: CT.host}, m = ctMeasured(), sg = ctSuggested(m), cmds = [], details = [];
  const R = CT.results[CT.host] || {};
  const label = k => (STATUS_TXT[k] || ["", k])[1];
  const connect = ctStats((R._connect || []).map(v => ({s: v, st: "ok"})));
  if (connect.n) cmds.push({group: "", command: "SSH login (measured on every Execute)", feeds: ctFeeds(["ssh_connect"]), n: connect.n, last: connect.last.s, max: connect.max, avg: connect.avg, status: "OK"});
  for (const p of TP.data.presets) {
    const runs = ctRuns(CT.host, p.id); if (!runs.length) continue;
    const st = ctStats(runs);
    cmds.push({group: p.group, command: p.title, feeds: ctFeeds(p.keys), n: st.n, last: st.lastS, max: st.max, avg: st.avg, status: label(st.last.st)});
    const parts = ctPartStats(p); if (parts.length) details.push({name: p.title, parts});
  }
  return {host: h, cmds, details, sg, image: CT.image, when: new Date().toLocaleString(), site: siteName || "",
          ver: (document.querySelector(".appver") || {}).textContent || ""};
}
const csvCell = v => { let t = v == null ? "" : String(v); if (typeof v === "string" && /^[=+\-@]/.test(t)) t = "'" + t; return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };
function ctDownload(name, text, type) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["\ufeff" + text], {type})); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
const ctStamp = () => { const d = new Date(), z = n => String(n).padStart(2, "0"); return d.getFullYear() + z(d.getMonth() + 1) + z(d.getDate()) + "_" + z(d.getHours()) + z(d.getMinutes()); };
function ctCsv() {
  const d = ctReportData(), L = [];
  const row = (...c) => L.push(c.map(csvCell).join(","));
  row("NetUpgrade command time test report"); row("Generated", d.when); row("Application", d.ver); row("Site", d.site);
  row("Switch", d.host.host, d.host.label || "", d.host.platform || "", d.host.version || ""); row("Image", d.image || "");
  row("Safety margin (x)", CT.margin); row("Never below the default", CT.keepDefault ? "yes" : "no"); row();
  row("Stage", "Step", "Feeds this timeout", "Runs", "Last (s)", "Slowest (s)", "Average (s)", "Last result");
  for (const c of d.cmds) row(c.group, c.command, c.feeds, c.n, c.last == null ? "" : c.last, c.max == null ? "" : c.max, c.avg == null ? "" : +c.avg.toFixed(2), c.status);
  for (const b of d.details) { row(); row(b.name + " - detail"); row("Measured", "Runs", "Last (s)", "Slowest (s)", "Average (s)");
    for (const x of b.parts) row(x.command, x.n, x.lastS == null ? "" : x.lastS, x.max == null ? "" : x.max, x.avg == null ? "" : +x.avg.toFixed(2)); }
  row(); row("Timeout setting", "Setting key", "Slowest measured (s)", "Default (s)", "Suggested (s)");
  for (const r of d.sg.rows) row(r.item.label, r.item.key, r.measured, r.def, r.value);
  ctDownload("command_time_test_" + String(d.host.host).replace(/[^A-Za-z0-9_.-]/g, "_") + "_" + ctStamp() + ".csv", L.join("\r\n") + "\r\n", "text/csv;charset=utf-8");
}
function ctPrint() {
  const d = ctReportData();
  const tr = cells => "<tr>" + cells.map(c => "<td>" + esc(c) + "</td>").join("") + "</tr>";
  const th = cells => "<tr>" + cells.map(c => "<th>" + esc(c) + "</th>").join("") + "</tr>";
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Command time test - ' + esc(d.host.host) + "</title><style>" +
    "body{font:14px/1.45 Segoe UI,Arial,sans-serif;margin:28px;color:#111}h1{font-size:21px;margin:0 0 4px}h2{font-size:15px;margin:22px 0 4px}" +
    "table{border-collapse:collapse;width:100%;margin:8px 0}th,td{border:1px solid #bbb;padding:6px 9px;text-align:left}th{background:#eef1f5}" +
    ".m{color:#555;font-size:13px}button{padding:7px 14px;margin-bottom:14px}@media print{.np{display:none}}</style></head><body>" +
    '<button class="np" onclick="window.print()">Print / Save As PDF</button><h1>Command Time Test Report</h1>' +
    '<div class="m">' + esc(d.when) + " | NetUpgrade " + esc(d.ver) + (d.site ? " | Site: " + esc(d.site) : "") + "</div>" +
    "<h2>Switch</h2><table>" + tr(["Address", d.host.host]) + tr(["Name", d.host.label || ""]) + tr(["Platform", d.host.platform || ""]) + tr(["Software version", d.host.version || ""]) + tr(["Image tested with", d.image || ""]) + "</table>" +
    "<h2>Measured Step Times</h2><table>" + th(["Stage", "Step", "Feeds this timeout", "Runs", "Last (s)", "Slowest (s)", "Average (s)", "Last result"]) +
    d.cmds.map(c => tr([c.group, c.command, c.feeds, c.n, fmtS(c.last), fmtS(c.max), fmtS(c.avg), c.status])).join("") + "</table>" +
    d.details.map(b => "<h2>" + esc(b.name) + " - detail</h2><table>" + th(["Measured", "Runs", "Last (s)", "Slowest (s)", "Average (s)"]) +
      b.parts.map(x => tr([x.command, x.n, fmtS(x.lastS), fmtS(x.max), fmtS(x.avg)])).join("") + "</table>").join("") +
    "<h2>Suggested Timeouts</h2><div class=\"m\">Slowest measured time x " + CT.margin + ", rounded up to 5 seconds" + (CT.keepDefault ? ", never below the default" : "") + ". 'Whole install' is the add step plus the activation.</div><table>" +
    th(["Timeout setting", "Slowest measured (s)", "Default (s)", "Suggested (s)"]) + d.sg.rows.map(r => tr([r.item.label, fmtS(r.measured), r.def, r.value])).join("") + "</table></body></html>";
  const url = URL.createObjectURL(new Blob([html], {type: "text/html"}));
  if (!window.open(url, "_blank")) setSay("err", "The browser blocked the report window. Allow pop-ups for this page, or use the CSV download.");
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

// ---- Console next to the Command time test: the output of every step you Execute, live, plus an interactive SSH
// session to the selected switch (the same /api/term session as the device drawer's Console) ----
const CC = {};       // host -> {text, sid, offset, state: off|connecting|on|closed, hist, hi, polling}
const CC_EL = document.createElement("div"); CC_EL.id = "cc-root";     // the console lives here; ctRender moves it into the page
const ccGet = h => CC[h] || (CC[h] = {text: "", sid: null, offset: 0, state: "off", hist: [], hi: -1, polling: false});
function ccAppend(host, text) {
  if (!text) return;
  const c = ccGet(host);
  c.text += text; if (c.text.length > 400000) c.text = c.text.slice(-300000);
  if (host === CT.host) ccPaint();
}
const ccSecretPrompt = c => /(?:password|secret|passphrase)[^\n]*:\s*$/i.test(c.text.slice(c.text.lastIndexOf("\n") + 1));
function ccSync() {                        // (re)build the panel when the switch changed
  const want = CT.open && CT.host ? CT.host : "";
  if (CC_EL.dataset.host !== want || (want && !CC_EL.querySelector("#cc-out"))) ccRender();
}
function ccRender() {
  const root = CC_EL;
  const want = CT.open && CT.host ? CT.host : "";
  root.dataset.host = want;
  if (!want) { root.innerHTML = ""; return; }
  root.innerHTML = '<div class="cc" id="cc"><div class="cc-h"><span class="cc-t" id="cc-title"></span><span class="cc-st" id="cc-state"></span><span class="cc-sp"></span>' +
    '<button type="button" class="cc-b go" id="cc-connect">Connect</button><button type="button" class="cc-b" id="cc-ctrlc">Ctrl+C</button>' +
    '<button type="button" class="cc-b" id="cc-clear">Clear</button><button type="button" class="cc-b" id="cc-max">Maximise</button></div>' +
    '<pre class="cc-out" id="cc-out"></pre><div class="cc-in"><span>&gt;</span><input id="cc-input" type="text" autocomplete="off" spellcheck="false" aria-label="Command for the switch"></div>' +
    '<div class="cc-n">The commands and output of every step appear here as the switch prints them. Connect opens an interactive SSH session to this switch, so you can also type commands, for example show install summary.</div></div>';
  $("cc-connect").addEventListener("click", () => { const c = ccGet(CT.host); (c.state === "on" || c.state === "connecting") ? ccDisconnect(CT.host) : ccConnect(CT.host); });
  $("cc-ctrlc").addEventListener("click", () => { ccSend({raw: "ctrl_c"}); $("cc-input").focus(); });
  $("cc-clear").addEventListener("click", () => { ccGet(CT.host).text = ""; ccPaint(); });
  $("cc-max").addEventListener("click", () => { const m = $("cc").classList.toggle("max"); $("cc-max").textContent = m ? "Restore" : "Maximise"; const o = $("cc-out"); o.scrollTop = o.scrollHeight; });
  $("cc-out").addEventListener("click", () => { if (!String(getSelection()) && !$("cc-input").disabled) $("cc-input").focus(); });
  $("cc-input").addEventListener("keydown", e => {
    const c = ccGet(CT.host), inp = e.target;
    if (e.key === "Enter") {
      e.preventDefault();
      const line = inp.value, secret = inp.type === "password";
      inp.value = "";
      if (!secret && line.trim()) { c.hist.push(line); c.hist = c.hist.slice(-100); }
      c.hi = -1; ccSend({line, secret});
    } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
      if (inp.type === "password" || !c.hist.length) return;
      e.preventDefault();
      c.hi = e.key === "ArrowUp" ? Math.min(c.hist.length - 1, c.hi + 1) : Math.max(-1, c.hi - 1);
      inp.value = c.hi < 0 ? "" : c.hist[c.hist.length - 1 - c.hi];
    } else if (e.ctrlKey && (e.key === "c" || e.key === "C") && inp.selectionStart === inp.selectionEnd) {
      e.preventDefault(); inp.value = ""; ccSend({raw: "ctrl_c"});
    }
  });
  const c = ccGet(CT.host);
  if (c.sid && !c.polling) ccPoll(CT.host);
  ccPaint();
}
function ccPaint() {
  const out = $("cc-out"); if (!out) return;
  const host = CT.host, c = ccGet(host), h = ctHostInfo();
  const live = c.state === "on" || c.state === "connecting";
  const atEnd = out.scrollHeight - out.scrollTop - out.clientHeight < 40;
  const shown = c.text || "Nothing yet. Execute a step on the left and its commands and output appear here live.";
  if (out.textContent !== shown) { out.textContent = shown; if (atEnd || !c.text) out.scrollTop = out.scrollHeight; }
  $("cc-title").textContent = "Console - " + host + (h && h.label ? " - " + h.label : "");
  const st = $("cc-state");
  const px = hostProto(host); st.textContent = {off: "No " + px + " session", connecting: "Connecting ...", on: px + " session open", closed: "Session closed"}[c.state];
  st.className = "cc-st" + (c.state === "on" ? " on" : c.state === "connecting" ? " busy" : "");
  $("cc-connect").textContent = live ? "Disconnect" : "Connect";
  $("cc-connect").className = "cc-b" + (live ? "" : " go");
  $("cc-ctrlc").disabled = !live;
  const inp = $("cc-input"), wasOff = inp.disabled;
  inp.disabled = !live;
  if (wasOff && live) inp.focus();
  const pw = live && ccSecretPrompt(c);
  inp.type = pw ? "password" : "text";
  inp.placeholder = live ? (pw ? "Enter the password (not shown, not recorded)" : "Type a command and press Enter") : "Click Connect to type commands";
}
async function ccConnect(host) {
  const c = ccGet(host);
  c.state = "connecting"; c.offset = 0; c.text += "\n[Connecting " + (hostProto(host) === "Telnet" ? "a Telnet" : "an SSH") + " session to " + host + " ...]\n"; ccPaint();
  try {
    const d = await post("/api/term/open", {host});
    c.sid = d.sid; c.offset = 0; ccPoll(host);
  } catch (e) { c.state = "off"; c.sid = null; c.text += "[" + e.message + "]\n"; }
  ccPaint();
}
async function ccDisconnect(host) {
  const c = ccGet(host), sid = c.sid;
  c.sid = null; c.state = "off"; c.text += "\n[Disconnected]\n"; ccPaint();
  if (sid) { try { await post("/api/term/close", {host, sid}); } catch (e) {} }
}
async function ccPoll(host) {
  const c = ccGet(host);
  if (c.polling) return;
  c.polling = true;
  try {
    while (c.sid && document.body.contains(CC_EL)) {
      const sid = c.sid;
      let d = null;
      try {
        const r = await fetch("/api/term/read?host=" + encodeURIComponent(host) + "&sid=" + encodeURIComponent(sid) + "&offset=" + c.offset, {cache: "no-store"});
        d = await r.json();
        if (!r.ok) { if (r.status === 404) { c.sid = null; c.state = "closed"; c.text += "\n[The session no longer exists on the server]\n"; break; } throw new Error(d.error || r.status); }
      } catch (e) { await new Promise(res => setTimeout(res, 1500)); continue; }
      if (c.sid !== sid) break;
      if (d.text) { c.text += d.text; if (c.text.length > 400000) c.text = c.text.slice(-300000); }
      c.offset = d.offset;
      c.state = d.closed ? "closed" : (d.ready ? "on" : "connecting");
      if (host === CT.host) ccPaint();
      if (d.closed) { c.sid = null; break; }
      await new Promise(res => setTimeout(res, d.text ? 120 : 450));
    }
  } finally { c.polling = false; }
  if (host === CT.host) ccPaint();
}
async function ccSend(payload) {
  const host = CT.host, c = ccGet(host);
  if (!c.sid) return;
  try { await post("/api/term/send", Object.assign({host, sid: c.sid}, payload)); }
  catch (e) { c.text += "\n[" + e.message + "]\n"; ccPaint(); }
}
async function ctSaveProfile() {
  try {
    const name = (($("ct-pname") || {}).value || "").trim();
    if (!name) { setSay("err", "Type a name for the profile."); return; }
    const sg = ctSuggested(ctMeasured()), h = ctHostInfo();
    const ex = TP.data.profiles.find(p => p.name.toLowerCase() === name.toLowerCase());
    if (ex && !ex.builtin && !confirm("A profile named '" + ex.name + "' already exists. Replace its values?")) return;
    const note = "Measured on " + CT.host + (h && h.label ? " (" + h.label + ")" : "") + ", margin x" + CT.margin + (CT.keepDefault ? ", never below the default" : "");
    const r = await post("/api/timeout-profiles/save", {name, values: sg.values, note, source: "command test", site: siteName || "", switch: CT.host, overwrite: !!ex});
    TP.data = r; TP.sel = r.saved; CT.form = false; renderSettings(); setSay("ok", r.message);
  } catch (e) { setSay("err", e.message); }
}
$("set-body").addEventListener("click", e => {
  const t = e.target.closest("[data-tp]"); if (t) { tpAction(t.dataset.tp); return; }
  const c = e.target.closest("[data-ct]"); if (!c) return;
  switch (c.dataset.ct) {
    case "toggle": CT.open = !CT.open; ctRender(); break;
    case "run": ctRun(c.dataset.id); break;
    case "stop": ctStop(); break;
    case "sel": { CT.sel = c.dataset.id; const sp = TP.data.presets.find(x => x.id === CT.sel); if (sp && CT_PATH_GROUP[sp.group]) CT.path = CT_PATH_GROUP[sp.group]; ctRender(); break; }
    case "next": { const ps = TP.data.presets, i = ps.findIndex(x => x.id === CT.sel); let j = i + 1; while (ps[j] && ctPathOff(ps[j].group)) j++;
      if (ps[j]) { CT.sel = ps[j].id; CT.railClosed[ps[j].group] = false; ctRender(); } break; }
    case "stage": CT.railClosed[c.dataset.g] = !CT.railClosed[c.dataset.g]; ctRender(); break;
    case "stages": { for (const g of new Set(TP.data.presets.map(p => p.group))) CT.railClosed[g] = c.dataset.v !== "1"; ctRender(); break; }
    case "clear": if (confirm("Clear the measured times of " + CT.host + "?")) {
      delete CT.results[CT.host]; for (const k of Object.keys(CT.outs)) if (k.startsWith(CT.host + "|")) delete CT.outs[k];
      ctPersist(); ctRender(); } break;
    case "csv": ctCsv(); break;
    case "print": ctPrint(); break;
    case "saveform": CT.form = true; ctRenderSuggest(); { const n = $("ct-pname"); if (n) n.focus(); } break;
    case "cancelform": CT.form = false; ctRenderSuggest(); break;
    case "save": ctSaveProfile(); break;
  }
});
$("set-body").addEventListener("change", e => {
  if (e.target.id === "tp-sel") { TP.sel = e.target.value; tpRender(); }
  else if (e.target.id === "ct-host") { CT.host = e.target.value; CT.form = false; ctPersist(); ctRender(); }
  else if (e.target.id === "ct-image") { CT.image = e.target.value; ctPersist(); ctRender(); }
  else if (e.target.id === "ct-keep") { CT.keepDefault = e.target.checked; ctRenderSuggest(); }
});
$("set-body").addEventListener("input", e => {
  if (e.target.id !== "ct-margin") return;
  const v = Number(e.target.value);
  if (v >= 1 && v <= 20) { CT.margin = v; ctPersist(); ctRenderSuggest(); }
});

// ============================================================================================
// Diagram page: CDP topology of the active site, with live up/down and upgrade status on every switch.
// The links come from /api/topology (read from the switches); the colours come from the same
// /api/status poll the Home page uses, so they change while an upgrade runs.
// ============================================================================================
const DG_W = 240, DG_H = 76, DG_GX = 36, DG_GY = 120, DG_PAD = 40;
const dg = {data: null, site: undefined, graphKey: "", listKey: "", detailKey: "", sel: null, picked: new Set(),
            showOther: false, ports: true, zoom: 1, px: 0, py: 0, fitted: false, root: null, timer: null, tick: 0, busy: false, pos: {}, view: null, wasRunning: false};
const SVGNS = "http://www.w3.org/2000/svg";
function svgEl(tag, attrs, text) {
  const e = document.createElementNS(SVGNS, tag);
  for (const k in (attrs || {})) e.setAttribute(k, attrs[k]);
  if (text != null) e.textContent = text;
  return e;
}
// Fit text into maxPx by measuring it (SVG text does not wrap), cutting it with an ellipsis only when it really does not fit.
function dgSetText(el, text, maxPx) {
  text = String(text == null ? "" : text);
  const key = text + "|" + maxPx; if (el._k === key) return; el._k = key;
  el.textContent = text;
  let w = 0; try { w = el.getComputedTextLength(); } catch (e) {}
  if (!w || w <= maxPx) return;                       // fits (or the page is hidden and cannot be measured)
  let lo = 1, hi = text.length - 1, best = 1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; el.textContent = text.slice(0, mid) + "\u2026"; if (el.getComputedTextLength() <= maxPx) { best = mid; lo = mid + 1; } else hi = mid - 1; }
  el.textContent = text.slice(0, best) + "\u2026";
}
const dgClip = (s, n) => { s = String(s == null ? "" : s); return s.length > n ? s.slice(0, n - 1) + "\u2026" : s; };
const ifShort = s => String(s || "").replace(/^TenGigabitEthernet/i, "Te").replace(/^TwentyFiveGigE/i, "Twe")
  .replace(/^FortyGigabitEthernet/i, "Fo").replace(/^HundredGigE/i, "Hu").replace(/^GigabitEthernet/i, "Gi")
  .replace(/^FastEthernet/i, "Fa").replace(/^Port-channel/i, "Po").replace(/^Ethernet/i, "Eth");
function dgSay(kind, text) { const m = $("dg-msg"); m.className = "msg " + kind; m.textContent = text || ""; }
const dgPosKey = () => "dg-pos:" + (dg.site || "");
function dgLoadPos() { dg.pos = {}; try { dg.pos = JSON.parse(localStorage.getItem(dgPosKey()) || "{}") || {}; } catch (e) { dg.pos = {}; } }
function dgSavePos() { try { localStorage.setItem(dgPosKey(), JSON.stringify(dg.pos)); } catch (e) {} }
const dgLive = () => new Map(((last && last.hosts) || []).map(h => [h.host, h]));

function openDiagram() {
  dgStop();
  dgLoad(true);
  dg.timer = setInterval(dgTick, 2000);
}
function dgStop() { if (dg.timer) { clearInterval(dg.timer); dg.timer = null; } dgExitFull(); }
function dgTick() {
  if ($("diagram").hidden) { dgStop(); return; }
  dg.tick++;
  const running = dg.data && dg.data.run && !dg.data.run.finished;
  if (running || dg.tick % 5 === 0 || (siteName || null) !== (dg.site || null)) dgLoad(false);
  else { dgRenderList(); dgPaintLive(); }
}

async function dgLoad(force) {
  if (dg.busy) return;
  dg.busy = true;
  try {
    const d = await getJSON("/api/topology");
    if (d.site !== dg.site) { dg.site = d.site; dgLoadPos(); dg.picked = new Set(); dg.sel = null; dg.graphKey = ""; dg.listKey = ""; dg.fitted = false; force = true; }
    const running = !!(d.run && !d.run.finished);
    if (dg.wasRunning && !running && d.run) {
      const hs = Object.values(d.run.hosts), ok = hs.filter(x => x.state === "ok").length;
      dgSay(ok === hs.length ? "ok" : "err", "Scan finished: " + ok + " of " + hs.length + " switch(es) read." +
            (ok < hs.length ? " See the device list for the reason." : ""));
    }
    dg.wasRunning = running;
    dg.data = d;
    const key = JSON.stringify(d.graph) + "|" + dg.showOther;
    if (force || key !== dg.graphKey) { dg.graphKey = key; dgDraw(); }
    dgRenderList();
    dgPaintLive();
  } catch (e) { dgSay("err", e.message); }
  finally { dg.busy = false; }
}

// ---- left side: the site's device list ----
function dgRenderList() {
  const d = dg.data; if (!d) return;
  const live = dgLive(), run = d.run && !d.run.finished;
  $("dg-site").textContent = d.site ? "Site: " + d.site + " - tick the switches to read their CDP neighbours." : "";
  let html = "";
  if (!d.site) html = '<div class="st-empty">No site selected. Create or select one from the Site menu on the Home page.</div>';
  else if (!d.devices.length) html = '<div class="st-empty">This site has no devices yet. Add them on the Home page first.</div>';
  for (const v of d.devices) {
    const lv = live.get(v.host), st = lv ? lv.status : "PENDING", rs = d.run && d.run.hosts && d.run.hosts[v.host];
    let line, cls = "";
    if (!v.has_creds) { line = "No credentials profile - assign one on the Home page"; cls = "warn"; }
    else if (rs && (rs.state === "queued" || rs.state === "running")) { line = rs.state === "queued" ? "Waiting ..." : "Reading CDP ..."; }
    else if (rs && rs.state === "failed") { line = rs.detail; cls = "bad"; }
    else if (rs && rs.state === "skipped") { line = rs.detail; cls = "warn"; }
    else if (v.scan_ok === false) { line = "Last scan failed: " + v.scan_error; cls = "bad"; }
    else if (v.scan_time) { line = "Scanned " + v.scan_time + " - " + v.neighbors + " neighbour(s)"; cls = "ok"; }
    else { line = "Not scanned yet"; }
    const dis = !v.has_creds || run;
    html += '<label class="dg-dev' + (v.has_creds ? "" : " off") + '"><input type="checkbox" data-host="' + esc(v.host) + '"' +
      (dg.picked.has(v.host) && v.has_creds ? " checked" : "") + (dis ? " disabled" : "") + '><span class="meta"><div class="t1"><i class="dot ' +
      (st === "UP" ? "ok" : st === "DOWN" ? "bad" : "") + '"></i>' + esc(v.label) + '</div><div class="t2">' + esc(v.host) +
      (v.platform ? " - " + esc(v.platform) : "") + '</div><div class="t2 ' + cls + '">' + esc(line) + "</div></span></label>";
  }
  if (html !== dg.listKey) { dg.listKey = html; $("dg-list").innerHTML = html; }
  const n = [...dg.picked].filter(h => d.devices.some(v => v.host === h && v.has_creds)).length;
  $("dg-png").disabled = !dg.root;
  $("dg-run").disabled = !!run || n === 0;
  $("dg-run").title = "Discover topology: read the CDP neighbours of the ticked switches";
  $("dg-run").textContent = run ? "Scanning " + d.run.done + "/" + d.run.total : "Discover" + (n ? " (" + n + ")" : "");
  const pickable = d.devices.filter(v => v.has_creds).length, all = $("dg-all");
  const allOn = pickable > 0 && n === pickable;                 // everything ticked: the button becomes "Clear"
  all.textContent = allOn ? "Clear" : "Select All"; all.dataset.mode = allOn ? "clear" : "all";
  all.disabled = !!run || pickable === 0;
}
$("dg-list").addEventListener("change", e => {
  const h = e.target.dataset && e.target.dataset.host; if (!h) return;
  if (e.target.checked) dg.picked.add(h); else dg.picked.delete(h);
  dgRenderList();
});
$("dg-all").addEventListener("click", () => {
  if (!dg.data) return;
  if ($("dg-all").dataset.mode === "clear") dg.picked.clear();
  else for (const v of dg.data.devices) if (v.has_creds) dg.picked.add(v.host);
  dg.listKey = ""; dgRenderList();
});
$("dg-run").addEventListener("click", async () => {
  const hosts = [...dg.picked]; if (!hosts.length) return;
  dgSay("", "");
  try { const r = await post("/api/topology/discover", {hosts}); dgSay("ok", r.message); dg.wasRunning = true; await dgLoad(false); }
  catch (e) { dgSay("err", e.message); }
});

// ---- layout: one band per group of connected switches, level 0 (top) first ----
function dgLayout(g) {
  const vis = g.nodes.filter(n => dg.showOther || n.kind !== "other");
  const ids = new Set(vis.map(n => n.id)), nb = new Map(vis.map(n => [n.id, []]));
  for (const e of g.edges) if (ids.has(e.a) && ids.has(e.b)) { nb.get(e.a).push(e.b); nb.get(e.b).push(e.a); }
  const comps = new Map();
  for (const n of vis) { if (!comps.has(n.comp)) comps.set(n.comp, []); comps.get(n.comp).push(n); }
  const multi = [], lone = [];
  for (const list of comps.values()) (list.length > 1 ? multi : lone).push(list);
  multi.sort((a, b) => b.length - a.length);
  const pos = {}; let y0 = DG_PAD;
  const place = list => {
    const levels = new Map();
    for (const n of list) { if (!levels.has(n.level)) levels.set(n.level, []); levels.get(n.level).push(n); }
    const keys = [...levels.keys()].sort((a, b) => a - b), xs = new Map();
    const widest = Math.max(...keys.map(k => levels.get(k).length)), W = widest * (DG_W + DG_GX) - DG_GX;
    keys.forEach((k, ri) => {
      const row = levels.get(k);
      const bary = n => { const p = nb.get(n.id).filter(i => xs.has(i)).map(i => xs.get(i)); return p.length ? p.reduce((a, b) => a + b, 0) / p.length : null; };
      row.sort((a, b) => { const x = bary(a), y = bary(b);
        if (x != null && y != null && x !== y) return x - y; if (x != null && y == null) return -1; if (x == null && y != null) return 1;
        return a.name.localeCompare(b.name); });
      const off = (W - (row.length * (DG_W + DG_GX) - DG_GX)) / 2;
      row.forEach((n, i) => { const x = DG_PAD + off + i * (DG_W + DG_GX); xs.set(n.id, x + DG_W / 2); pos[n.id] = {x, y: y0 + ri * DG_GY}; });
    });
    y0 += keys.length * DG_GY + 24;
  };
  multi.forEach(place);
  let capY = null;
  if (lone.length) {
    capY = y0 - 6; y0 += 18;
    const flat = lone.map(l => l[0]).sort((a, b) => a.name.localeCompare(b.name));
    for (let i = 0; i < flat.length; i += 6) { flat.slice(i, i + 6).forEach((n, j) => { pos[n.id] = {x: DG_PAD + j * (DG_W + DG_GX), y: y0}; }); y0 += DG_GY; }
  }
  for (const n of vis) { const o = dg.pos[n.id]; if (o && isFinite(o.x) && isFinite(o.y)) pos[n.id] = {x: o.x, y: o.y}; }
  return {nodes: vis, pos, capY};
}

function dgDraw() {
  const svg = $("dg-svg"), d = dg.data, g = d && d.graph;
  svg.innerHTML = ""; dg.view = null;
  let msg = "";
  if (!d || !d.site) msg = "No site selected. Create or select a site from the Site menu on the Home page.";
  else if (!d.devices.length) msg = "This site has no devices yet.";
  else if (!g.nodes.length) msg = "No topology yet. Tick the switches on the left and press Discover topology.";
  $("dg-empty").textContent = msg;
  if (msg) { dg.root = null; dgPaintLive(); return; }
  const lay = dgLayout(g);
  // Everything that moves with the mouse (pan + zoom) lives in one group; its transform is the "camera".
  const root = svgEl("g", {id: "dg-root"});
  svg.appendChild(root); dg.root = root;
  const edgeLayer = svgEl("g"), nodeLayer = svgEl("g");
  root.appendChild(edgeLayer); root.appendChild(nodeLayer);
  if (lay.capY != null) root.appendChild(svgEl("text", {class: "dg-cap", x: DG_PAD, y: lay.capY}, "No CDP links found for these switches"));
  const nodes = new Map(), edges = [];
  for (const n of lay.nodes) {
    const p = {x: lay.pos[n.id].x, y: lay.pos[n.id].y};
    const el = svgEl("g", {class: "dg-n", "data-id": n.id, tabindex: 0, transform: "translate(" + p.x + "," + p.y + ")"});
    const title = svgEl("title", {}, "");
    el.appendChild(title);
    el.appendChild(svgEl("rect", {class: "bx", width: DG_W, height: DG_H, rx: 10}));
    el.appendChild(svgEl("circle", {class: "sd", cx: 15, cy: 18, r: 5}));
    const nm = svgEl("text", {class: "nm", x: 27, y: 22}, n.name);
    el.appendChild(nm);
    const stt = svgEl("text", {class: "stt", x: DG_W - 10, y: 22, "text-anchor": "end"}, "");
    el.appendChild(stt);
    const ipt = svgEl("text", {class: "ipt", x: 14, y: 40}, "");
    el.appendChild(ipt);
    const strip = svgEl("g", {class: "strip"});
    strip.appendChild(svgEl("rect", {x: 8, y: 49, width: DG_W - 16, height: 20, rx: 6}));
    const stripT = svgEl("text", {x: DG_W / 2, y: 63, "text-anchor": "middle"}, "");
    strip.appendChild(stripT); el.appendChild(strip);
    const V = {n, p, el, title, nm, stt, ipt, stripT};
    dgWireNode(V);
    nodeLayer.appendChild(el); nodes.set(n.id, V);
  }
  for (const e of g.edges) {
    if (!nodes.has(e.a) || !nodes.has(e.b)) continue;
    const tip = e.links.map(l => ifShort(l.a_if) + " \u2194 " + ifShort(l.b_if)).join("\n");
    const path = svgEl("path", {class: "dg-e"}); path.appendChild(svgEl("title", {}, tip));
    const la = svgEl("text", {class: "dg-lbl"}, ""), lb = svgEl("text", {class: "dg-lbl"}, ""), mid = svgEl("text", {class: "dg-lbl", "text-anchor": "middle"}, "");
    edgeLayer.appendChild(path); edgeLayer.appendChild(la); edgeLayer.appendChild(lb); edgeLayer.appendChild(mid);
    edges.push({e, path, la, lb, mid});
  }
  const meta = new Map(g.nodes.map(n => [n.id, n])), adj = new Map(g.nodes.map(n => [n.id, []]));
  for (const e of g.edges) { adj.get(e.a).push(e.b); adj.get(e.b).push(e.a); }
  dg.view = {nodes, edges, meta, adj};
  dgPlaceEdges();
  if (!dg.fitted && dgFitView()) dg.fitted = true; else dgApplyView();
  dgPaintLive();
}

// Screen pointer -> diagram coordinates (takes the current pan and zoom into account).
function dgPoint(ev) {
  const m = dg.root && dg.root.getScreenCTM();
  if (!m) return {x: ev.clientX, y: ev.clientY};
  const q = new DOMPoint(ev.clientX, ev.clientY).matrixTransform(m.inverse());
  return {x: q.x, y: q.y};
}
function dgWireNode(V) {
  const el = V.el;
  el.addEventListener("pointerdown", ev => {
    if (ev.button !== 0) return;
    ev.stopPropagation();                      // do not start panning the whole diagram
    const start = dgPoint(ev), ox = V.p.x, oy = V.p.y; let moved = false;
    try { el.setPointerCapture(ev.pointerId); } catch (e) {}
    const mv = e2 => {
      const q = dgPoint(e2), nx = ox + q.x - start.x, ny = oy + q.y - start.y;
      if (Math.abs(nx - ox) + Math.abs(ny - oy) > 4) moved = true;
      if (!moved) return;
      V.p.x = nx; V.p.y = ny; el.setAttribute("transform", "translate(" + nx + "," + ny + ")"); dgPlaceEdges();
    };
    const up = e2 => {
      el.removeEventListener("pointermove", mv); el.removeEventListener("pointerup", up); el.removeEventListener("pointercancel", up);
      try { el.releasePointerCapture(e2.pointerId); } catch (e) {}
      if (moved) { dg.pos[V.n.id] = {x: V.p.x, y: V.p.y}; dgSavePos(); } else dgSelect(V.n.id);
    };
    el.addEventListener("pointermove", mv); el.addEventListener("pointerup", up); el.addEventListener("pointercancel", up);
  });
  el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); dgSelect(V.n.id); } });
}

// Camera: apply the current pan (px, py) and zoom to the diagram group.
function dgApplyView() {
  if (dg.root) dg.root.setAttribute("transform", "translate(" + dg.px + "," + dg.py + ") scale(" + dg.zoom + ")");
}
// Scale and centre the whole diagram so every switch is visible in the window. Returns false if the window has no size yet.
function dgFitView() {
  const v = dg.view, wrap = $("dg-wrap"); if (!v || !dg.root) return false;
  const W = wrap.clientWidth, H = wrap.clientHeight; if (W < 50 || H < 50) return false;
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const V of v.nodes.values()) { x0 = Math.min(x0, V.p.x); y0 = Math.min(y0, V.p.y); x1 = Math.max(x1, V.p.x + DG_W); y1 = Math.max(y1, V.p.y + DG_H); }
  const m = 36, bw = Math.max(1, x1 - x0), bh = Math.max(1, y1 - y0);
  const z = Math.max(0.2, Math.min(1.25, (W - 2 * m) / bw, (H - 2 * m) / bh));
  dg.zoom = z; dg.px = (W - bw * z) / 2 - x0 * z; dg.py = (H - bh * z) / 2 - y0 * z;
  dgApplyView(); return true;
}
// Zoom by a factor, keeping the diagram point under (cx, cy) (window pixels) where it is.
function dgZoomAt(factor, cx, cy) {
  const z0 = dg.zoom, z1 = Math.max(0.2, Math.min(3, z0 * factor)); if (z1 === z0) return;
  dg.px = cx - (cx - dg.px) * (z1 / z0); dg.py = cy - (cy - dg.py) * (z1 / z0); dg.zoom = z1; dgApplyView();
}

// Draw each link from the bottom of the upper switch to the top of the lower one (side to side within a row).
function dgPlaceEdges() {
  const v = dg.view; if (!v) return;
  const att = new Map();
  for (const E of v.edges) {
    const A = v.nodes.get(E.e.a).p, B = v.nodes.get(E.e.b).p;
    if (Math.abs(A.y - B.y) < DG_H * 0.6) { E.mode = "side"; continue; }
    E.mode = "v"; E.top = A.y < B.y ? E.e.a : E.e.b; E.bot = A.y < B.y ? E.e.b : E.e.a;
    for (const [id, k, other] of [[E.top, "down", E.bot], [E.bot, "up", E.top]]) {
      if (!att.has(id)) att.set(id, {up: [], down: []});
      att.get(id)[k].push({E, other});
    }
  }
  for (const [id, a] of att) for (const k of ["up", "down"]) {
    const list = a[k], P = v.nodes.get(id).p;
    list.sort((p, q) => v.nodes.get(p.other).p.x - v.nodes.get(q.other).p.x);
    list.forEach((it, i) => { it.E[k === "down" ? "tx" : "bx"] = P.x + DG_W * (i + 1) / (list.length + 1); it.E[k === "down" ? "ti" : "bi"] = i; });
  }
  for (const E of v.edges) {
    const A = v.nodes.get(E.e.a).p, B = v.nodes.get(E.e.b).p, one = E.e.links.length === 1 && dg.ports;
    const l0 = E.e.links[0] || {a_if: "", b_if: ""};
    let d, mx, my, f = {};
    if (E.mode === "side") {
      const aLeft = A.x <= B.x, L = aLeft ? A : B, R = aLeft ? B : A;
      const x1 = L.x + DG_W, y1 = L.y + DG_H / 2, x2 = R.x, y2 = R.y + DG_H / 2;
      d = "M" + x1 + " " + y1 + " L" + x2 + " " + y2; mx = (x1 + x2) / 2; my = (y1 + y2) / 2 - 6;
      f = {la: [x1 + 4, y1 - 5, "start", aLeft ? l0.a_if : l0.b_if], lb: [x2 - 4, y2 - 5, "end", aLeft ? l0.b_if : l0.a_if]};
    } else {
      const T = v.nodes.get(E.top).p, Bn = v.nodes.get(E.bot).p, x1 = E.tx, y1 = T.y + DG_H, x2 = E.bx, y2 = Bn.y, m = (y1 + y2) / 2;
      d = "M" + x1 + " " + y1 + " C" + x1 + " " + m + " " + x2 + " " + m + " " + x2 + " " + y2; mx = (x1 + x2) / 2; my = m;
      const topIf = E.top === E.e.a ? l0.a_if : l0.b_if, botIf = E.top === E.e.a ? l0.b_if : l0.a_if;
      f = {la: [x1 + 4, y1 + 13 + ((E.ti || 0) % 2) * 11, "start", topIf], lb: [x2 + 4, y2 - 6 - ((E.bi || 0) % 2) * 11, "start", botIf]};
    }
    E.path.setAttribute("d", d);
    for (const k of ["la", "lb"]) {
      const t = E[k], s = f[k];
      t.setAttribute("x", s[0]); t.setAttribute("y", s[1]); t.setAttribute("text-anchor", s[2]); t.textContent = one ? ifShort(s[3]) : "";
    }
    E.mid.setAttribute("x", mx); E.mid.setAttribute("y", my);
    E.mid.textContent = E.e.links.length > 1 ? "\u00d7" + E.e.links.length : "";
  }
}

// ---- live part: up/down + upgrade progress, repainted every 2 seconds from the Home page status poll ----
function dgUpg(h) {
  if (!h) return {cls: "none", text: ""};
  const u = h.upgrade || {};
  if (u.state === "starting" || u.state === "running") return {cls: "running", text: "Upgrading" + (u.phase ? ": " + String(u.phase).replace(/[_-]+/g, " ") : "")};
  if (u.state === "completed") return {cls: "completed", text: u.action === "distribute" ? "Image distributed" : "Upgrade completed"};
  if (u.state === "failed") return {cls: "failed", text: "Upgrade failed"};
  if (u.state === "attention") return {cls: "attention", text: "Needs attention"};
  if (h.scheduled) return {cls: "queued", text: h.scheduled.queued ? "Queued for upgrade" : "Scheduled " + String(h.scheduled.run_at_str || "").slice(5)};
  return {cls: "none", text: ""};
}

function dgPaintLive() {
  const v = dg.view, banner = $("dg-banner");
  if (!v) { banner.hidden = true; dgRenderDetail(); return; }
  const live = dgLive(), info = new Map(), active = [];
  for (const [id, V] of v.nodes) {
    const h = V.n.managed ? live.get(V.n.host) : null;
    const status = !V.n.managed ? "EXT" : (h && h.status) || "PENDING", u = dgUpg(h);
    info.set(id, {h, status, u});
    if (u.cls === "running") active.push(id);
  }
  // Everything above a switch that is being upgraded would cut it off if it were upgraded at the same time.
  const holds = new Map();
  for (const a of active) {
    const seen = new Set([a]), queue = [a];
    while (queue.length) {
      const cur = queue.shift(), lv = v.meta.get(cur).level;
      for (const nx of v.adj.get(cur)) {
        const m = v.meta.get(nx);
        if (seen.has(nx) || !m.managed || m.kind !== "net" || m.level >= lv) continue;
        seen.add(nx); queue.push(nx);
        if (!holds.has(nx)) holds.set(nx, []);
        holds.get(nx).push(v.meta.get(a).name);
      }
    }
  }
  for (const [id, V] of v.nodes) {
    const {h, status, u} = info.get(id), n = V.n, hold = holds.get(id);
    const ucls = u.cls === "none" && hold ? "hold" : u.cls;
    V.el.setAttribute("class", "dg-n s-" + status + " u-" + ucls + (n.managed ? "" : " ext") + (hold ? " hold" : "") + (dg.sel === id ? " sel" : ""));
    V.stt.textContent = status === "EXT" ? "EXT" : status;
    let sw = 40; try { sw = V.stt.getComputedTextLength() || 40; } catch (e) {}
    dgSetText(V.nm, n.name, DG_W - 27 - sw - 16);                       // name never runs into the UP / DOWN label
    dgSetText(V.ipt, n.managed ? (n.host + (n.platform || (h && h.platform) ? " \u00b7 " + (h && h.platform || n.platform) : ""))
                               : "not onboarded" + (n.platform ? " \u00b7 " + n.platform : ""), DG_W - 14 - 10);
    dgSetText(V.stripT, u.text || (hold ? "HOLD: uplink of " + hold[0] : ""), DG_W - 16 - 14);
    V.title.textContent = [n.name, n.host || "", n.platform || (h && h.platform) || "", n.managed ? "Status: " + status : "Not onboarded in this site",
      h && h.version ? "Version: " + h.version : "", hold ? "HOLD: uplink of " + hold.join(", ") + " (upgrade in progress)" : "", u.text, h && h.upgrade && h.upgrade.message && u.cls !== "none" ? h.upgrade.message : ""].filter(Boolean).join("\n");
  }
  for (const E of v.edges) {
    const A = info.get(E.e.a), B = info.get(E.e.b);
    const down = A.status === "DOWN" || B.status === "DOWN";
    const upg = A.u.cls === "running" || B.u.cls === "running";
    const hold = (A.u.cls === "running" && holds.has(E.e.b)) || (B.u.cls === "running" && holds.has(E.e.a));
    E.path.setAttribute("class", "dg-e" + (down ? " down" : "") + (upg ? " upg" : "") + (hold ? " hold" : ""));
    E.path.style.strokeWidth = (2 + Math.min(E.e.links.length - 1, 3) + (upg ? 1 : 0)) + "px";
  }
  if (active.length) {
    const names = a => a.map(id => v.meta.get(id).name).join(", ");
    banner.textContent = "Upgrade in progress on " + names(active) + "." + (holds.size ? " Do not start an upgrade on its uplink (" + names([...holds.keys()]) + ") until it finishes." : "");
    banner.hidden = false;
  } else banner.hidden = true;
  dgRenderDetail(info);
}

// ---- details panel for the clicked switch ----
function dgSelect(id) { dg.sel = dg.sel === id ? null : id; dg.detailKey = ""; dgPaintLive(); }
function dgRenderDetail(info) {
  const box = $("dg-detail"), v = dg.view;
  if (!dg.sel || !v || !v.nodes.has(dg.sel)) { box.hidden = true; dg.detailKey = ""; return; }
  const live = dgLive(), n = v.nodes.get(dg.sel).n, h = n.managed ? live.get(n.host) : null, u = dgUpg(h);
  const status = !n.managed ? "" : (h && h.status) || "PENDING";
  const rows = [["Address", n.host || "-"], ["Platform", n.platform || (h && h.platform) || "-"]];
  if (n.managed) {
    rows.push(["Status", status === "UP" ? "Up" : status === "DOWN" ? "Down" : "Not checked yet"]);
    rows.push(["Version", (h && h.version) || "-"]);
    rows.push(["Credentials", (h && h.profile) || "-"]);
    rows.push(["Upgrade", u.text || "No upgrade"]);
    if (h && h.upgrade && h.upgrade.message && u.cls !== "none" && u.cls !== "queued") rows.push(["Detail", dgClip(h.upgrade.message, 120)]);
    rows.push(["Last scan", n.scan_time ? n.scan_time + (n.scan_ok === false ? " (latest scan failed: " + n.scan_error + ")" : "") : (n.scan_error || "not scanned")]);
  } else rows.push(["Note", "Seen by CDP but not onboarded in this site"]);
  const peers = [];
  for (const e of dg.data.graph.edges) {
    if (e.a !== n.id && e.b !== n.id) continue;
    const other = v.meta.get(e.a === n.id ? e.b : e.a); if (!other) continue;
    const pl = live.get(other.host), ps = other.managed ? ((pl && pl.status) || "PENDING") : "";
    peers.push('<li><i class="dot ' + (ps === "UP" ? "ok" : ps === "DOWN" ? "bad" : "") + '"></i><b>' + esc(other.name) + "</b>" +
      e.links.map(l => '<div class="pt">' + esc(ifShort(e.a === n.id ? l.a_if : l.b_if)) + " \u2194 " + esc(ifShort(e.a === n.id ? l.b_if : l.a_if)) + "</div>").join("") + "</li>");
  }
  const html = '<h5><span>' + esc(n.name) + '</span><button type="button" data-dg="close" title="Close">&times;</button></h5><dl>' +
    rows.map(r => "<dt>" + esc(r[0]) + "</dt><dd>" + esc(r[1]) + "</dd>").join("") + "</dl>" +
    '<div class="note" style="margin:0 0 4px">Connected to (' + peers.length + ")</div><ul class=\"nbs\">" + (peers.join("") || "<li>No links found</li>") + "</ul>";
  if (html === dg.detailKey) return;
  dg.detailKey = html; box.innerHTML = html; box.hidden = false;
}
$("dg-detail").addEventListener("click", async e => {
  const b = e.target.closest("[data-dg]"); if (!b) return;
  if (b.dataset.dg === "close") { dg.sel = null; dgPaintLive(); return; }
});

// ---- toolbar ----
$("dg-ports").addEventListener("change", e => { dg.ports = e.target.checked; dgPlaceEdges(); });
$("dg-other").addEventListener("change", e => { dg.showOther = e.target.checked; dg.graphKey = ""; dgLoad(true); });
$("dg-refresh").addEventListener("click", () => dgLoad(true));
$("dg-reset").addEventListener("click", () => { dg.pos = {}; dgSavePos(); dg.fitted = false; dgDraw(); });
$("dg-fit").addEventListener("click", () => { dgFitView(); });

// Export the diagram as a PNG: the SVG is copied with its styles written into every element (a picture cannot use the
// page's stylesheet), drawn onto a canvas at 2-3x for a sharp result, and saved. The whole diagram is exported, whatever
// the current pan or zoom, with the switches where they were dragged to.
function dgInlineStyles(src, dst) {
  const cs = getComputedStyle(src);
  const props = ["fill", "fill-opacity", "stroke", "stroke-width", "stroke-opacity", "stroke-dasharray", "stroke-linecap", "stroke-linejoin",
                 "opacity", "font-family", "font-size", "font-weight", "font-style", "text-anchor", "letter-spacing", "visibility", "display",
                 "paint-order", "stroke-miterlimit", "stroke-dashoffset", "dominant-baseline", "text-transform"];
  dst.setAttribute("style", props.map(k => k + ":" + cs.getPropertyValue(k)).join(";"));
  const a = src.children, b = dst.children;
  for (let i = 0; i < a.length && i < b.length; i++) dgInlineStyles(a[i], b[i]);
}
function dgExportPng() {
  if (!dg.root) { dgSay("err", "Nothing to export yet - tick some switches and click Discover."); return; }
  let bb; try { bb = dg.root.getBBox(); } catch (e) { bb = null; }
  if (!bb || !bb.width) { dgSay("err", "Nothing to export yet."); return; }
  const pad = 36, head = 44, x = bb.x - pad, y = bb.y - pad - head, w = bb.width + 2 * pad, h = bb.height + 2 * pad + head;
  const scale = Math.max(1, Math.min(3, 8000 / Math.max(w, h)));
  const clone = dg.root.cloneNode(true);
  dgInlineStyles(dg.root, clone);
  clone.removeAttribute("transform"); clone.removeAttribute("id");
  clone.querySelectorAll("title").forEach(t => t.remove());
  const NS = "http://www.w3.org/2000/svg", mk = (n, a) => { const e = document.createElementNS(NS, n); for (const k in a) e.setAttribute(k, a[k]); return e; };
  const bg = getComputedStyle($("dg-wrap")).backgroundColor || "#ffffff", fg = getComputedStyle(document.body).color || "#111111";
  const font = getComputedStyle(document.body).fontFamily;
  const when = new Date(), pad2 = n => String(n).padStart(2, "0");
  const stamp = when.getFullYear() + "-" + pad2(when.getMonth() + 1) + "-" + pad2(when.getDate()) + " " + pad2(when.getHours()) + ":" + pad2(when.getMinutes());
  const svg = mk("svg", {xmlns: NS, width: Math.round(w * scale), height: Math.round(h * scale), viewBox: [x, y, w, h].join(" ")});
  svg.appendChild(mk("rect", {x: x, y: y, width: w, height: h, fill: bg}));
  const t1 = mk("text", {x: x + pad, y: y + 28, "font-family": font, "font-size": 18, "font-weight": 700, fill: fg});
  t1.textContent = "Topology" + (siteName ? " - " + siteName : "");
  const t2 = mk("text", {x: x + pad, y: y + 46, "font-family": font, "font-size": 11, fill: fg, "fill-opacity": 0.65});
  t2.textContent = "Exported " + stamp;
  svg.appendChild(t1); svg.appendChild(t2); svg.appendChild(clone);
  const img = new Image();
  img.onload = () => {
    const c = document.createElement("canvas"); c.width = Math.round(w * scale); c.height = Math.round(h * scale);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    c.toBlob(blob => {
      if (!blob) { dgSay("err", "Could not create the PNG in this browser."); return; }
      const url = URL.createObjectURL(blob), a = document.createElement("a");
      a.href = url; a.download = "topology_" + (siteName || "site").replace(/[^A-Za-z0-9._-]+/g, "_") + "_" +
        when.getFullYear() + pad2(when.getMonth() + 1) + pad2(when.getDate()) + "_" + pad2(when.getHours()) + pad2(when.getMinutes()) + ".png";
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 5000);
      dgSay("ok", "Saved " + a.download);
    }, "image/png");
  };
  img.onerror = () => dgSay("err", "Could not create the PNG in this browser.");
  img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(svg));
}
$("dg-png").addEventListener("click", dgExportPng);
const dgMid = () => { const w = $("dg-wrap"); return [w.clientWidth / 2, w.clientHeight / 2]; };
$("dg-zin").addEventListener("click", () => { const c = dgMid(); dgZoomAt(1.25, c[0], c[1]); });
$("dg-zout").addEventListener("click", () => { const c = dgMid(); dgZoomAt(0.8, c[0], c[1]); });

// Mouse wheel: zoom around the pointer.
$("dg-svg").addEventListener("wheel", ev => {
  if (!dg.root) return;
  ev.preventDefault();
  const r = $("dg-wrap").getBoundingClientRect();
  const dy = ev.deltaMode === 1 ? ev.deltaY * 16 : ev.deltaY;
  dgZoomAt(Math.exp(-dy * 0.0016), ev.clientX - r.left, ev.clientY - r.top);
}, {passive: false});

// Drag empty space (left or middle button): move the whole diagram freely, in any direction.
$("dg-svg").addEventListener("pointerdown", ev => {
  if (!dg.root || (ev.button !== 0 && ev.button !== 1) || ev.target.closest(".dg-n")) return;
  ev.preventDefault();
  const svg = $("dg-svg"), wrap = $("dg-wrap");
  const sx = ev.clientX, sy = ev.clientY, ox = dg.px, oy = dg.py; let moved = false;
  try { svg.setPointerCapture(ev.pointerId); } catch (e) {}
  wrap.classList.add("panning");
  const mv = e2 => {
    const dx = e2.clientX - sx, dy = e2.clientY - sy;
    if (!moved && Math.abs(dx) + Math.abs(dy) > 3) moved = true;
    if (!moved) return;
    dg.px = ox + dx; dg.py = oy + dy; dgApplyView();
  };
  const up = e2 => {
    svg.removeEventListener("pointermove", mv); svg.removeEventListener("pointerup", up); svg.removeEventListener("pointercancel", up);
    try { svg.releasePointerCapture(e2.pointerId); } catch (e) {}
    wrap.classList.remove("panning");
    if (!moved && dg.sel) { dg.sel = null; dgPaintLive(); }      // a plain click on the background closes the details panel
  };
  svg.addEventListener("pointermove", mv); svg.addEventListener("pointerup", up); svg.addEventListener("pointercancel", up);
});

// Full screen: the whole Topology card (toolbar, canvas, legend, details) fills the screen.
// Uses the browser's Fullscreen API and falls back to a fixed full-window layer where that is not allowed.
function dgIsFull() { const c = $("dg-card"); return document.fullscreenElement === c || c.classList.contains("dg-full"); }
function dgFullUI() {
  const on = dgIsFull(), b = $("dg-full");
  b.innerHTML = IC[on ? "shrink" : "expand"]; b.title = on ? "Exit full screen (Esc)" : "Full screen (Esc to exit)";
  b.setAttribute("aria-label", on ? "Exit full screen" : "Full screen"); b.setAttribute("aria-pressed", on ? "true" : "false");
  if (dg.view) setTimeout(() => { if (dgFitView()) dg.fitted = true; }, 60);   // the canvas changed size: fit the diagram to it
}
function dgExitFull() {
  const c = $("dg-card"); if (!c) return;
  if (document.fullscreenElement === c && document.exitFullscreen) document.exitFullscreen().catch(() => {});
  if (c.classList.contains("dg-full")) { c.classList.remove("dg-full"); dgFullUI(); }
}
$("dg-full").addEventListener("click", async () => {
  const c = $("dg-card");
  if (dgIsFull()) { dgExitFull(); return; }
  try { if (!c.requestFullscreen) throw new Error("no fullscreen api"); await c.requestFullscreen(); }
  catch (e) { c.classList.add("dg-full"); dgFullUI(); }
});
document.addEventListener("fullscreenchange", dgFullUI);
document.addEventListener("keydown", e => { if (e.key === "Escape" && $("dg-card").classList.contains("dg-full")) dgExitFull(); });

if (location.hash === "#storage") showView("storage", false);
else if (location.hash === "#diagram") showView("diagram", false);
else if (location.hash === "#settings") showView("settings", false);

// ============================================================================================
// Device drawer: click a switch name in the device list -> a panel on the right with everything
// known about that switch. The live parts (versions, check results, upgrade progress) come from
// the same /api/status poll as the table; the rest (facts read from backups, checklist, saved
// files, command log, events, history) comes from /api/device, refreshed every 3 s while open.
// ============================================================================================
const DR = {host: null, tab: "overview", data: null, err: "", html: "", tabsHtml: "", timer: null, busy: false, src: "all", viewing: false};
const DR_TABS = [["overview", "Overview"], ["checks", "Checks"], ["backups", "Backups"], ["commands", "Commands"], ["events", "Events"], ["console", "Console"]];
const CAN_VIEW = ME.role !== "viewer";          // viewing saved files is for Administrators / Operators (same as Storage)
const DASH = '<span class="muted">\u2014</span>';
const PHASE_PCT = {waiting_backup: 3, starting: 5, connecting: 8, connected: 12, prechecking: 16, boot: 18, cleaning: 22,
                   space_check: 28, uploading: 40, md5: 52, installing: 65, reloading: 75, waiting_down: 78,
                   waiting_up: 85, waiting_ssh: 91, committing: 94, postcheck: 96, skipped: 100};
const READY_PILL = {checking: ["CHECKING", "RUN"], ready: ["READY", "OK"], warning: ["WARNING", "CHECK"], not_ready: ["NOT READY", "DOWN"]};

const drDash = v => (v == null || v === "") ? DASH : esc(v);
const drSize = n => n < 1024 ? n + " B" : n < 1048576 ? (n / 1024).toFixed(1) + " KB" : (n / 1048576).toFixed(1) + " MB";
let DR_FOLD = {};
try { DR_FOLD = JSON.parse(localStorage.getItem("drFold") || "{}") || {}; } catch (e) {}
const drFoldOpen = (k, dflt) => DR_FOLD[k] === undefined ? dflt : !!DR_FOLD[k];
// An expandable section: a header with a down arrow; the choice is remembered in this browser.
function drFold(key, title, right, inner, dflt) {
  const open = drFoldOpen(key, dflt);
  return '<div class="dr-sec dr-fold' + (open ? " open" : "") + '"><h5><button type="button" class="fold-btn" data-fold="' + key +
    '" aria-expanded="' + open + '">' + esc(title) + '<span class="chev">' + IC.dn + "</span></button>" + (right || "") + "</h5>" +
    (open ? '<div class="fold-body">' + inner + "</div>" : "") + "</div>";
}
const drHost = () => last && last.hosts.find(x => x.host === DR.host);

// rows: [label, value, valueIsHtml?]. showEmpty: keep rows without a value (shown as a dash).
function drKV(rows, showEmpty) {
  const out = [];
  for (const r of rows) {
    const empty = r[1] == null || r[1] === "";
    if (empty && !showEmpty) continue;
    out.push("<tr><td>" + esc(r[0]) + "</td><td>" + (empty ? DASH : (r[2] ? r[1] : esc(r[1]))) + "</td></tr>");
  }
  return out.length ? '<table class="dr-kv">' + out.join("") + "</table>" : "";
}
function drSec(title, right, inner) {
  return '<div class="dr-sec"><h5>' + esc(title) + '<span class="sp"></span>' + (right || "") + "</h5>" + inner + "</div>";
}
function drCard(k, v, s, bad) {
  return '<div class="dr-card"><div class="k">' + esc(k) + '</div><div class="v">' + v + "</div>" +
         (s ? '<div class="s' + (bad ? " bad" : "") + '">' + s + "</div>" : "") + "</div>";
}
function drViewBtn(kind, name) {
  return CAN_VIEW ? '<button type="button" class="btn sm" data-view="' + esc(kind) + '" data-name="' + esc(name) + '">View</button>' : "";
}
function drFileRow(title, sub, size, kind, name) {
  return '<div class="fl-row"><div class="nm">' + esc(title) + (sub ? "<small>" + esc(sub) + "</small>" : "") + "</div>" +
         '<span class="sz">' + drSize(size) + "</span>" + drViewBtn(kind, name) + "</div>";
}

// ---- upgrade progress (steps + percentage) ----
function drProgress(h) {
  const u = h.upgrade;
  if (u.state === "completed") return 100;
  if (u.state === "idle") return 0;
  return PHASE_PCT[u.phase] || 5;
}
// Image stage steps: Distribute = image is on the switch's flash; Activate = installed and the switch reloaded onto it.
function distStep(h) {
  const u = h.upgrade, st = h.stage || {}, act = u.action;
  if (ACTIVE_STATES.includes(u.state)) {
    if (act === "distribute") return "active";
    if (act === "full") return DIST_PHASES.includes(u.phase) ? "active" : "done";
    return "done";                                                  // activate-only: it was distributed earlier
  }
  if (u.state === "completed" && act !== "distribute" && !/^Already running/.test(u.message || "")) return "done";
  if (u.state === "completed" && act === "distribute" && /^DISTRIBUTED/.test(u.message || "")) return "done";
  if (u.state === "failed" && act !== "activate" && DIST_PHASES.includes(u.phase)) return "failed";
  if (st.staged || (st.upgraded && !(u.state === "failed" || u.state === "attention"))) return "done";
  return "pending";
}
function actStep(h) {
  const u = h.upgrade, st = h.stage || {}, act = u.action;
  if (ACTIVE_STATES.includes(u.state)) return (act === "activate" || (act === "full" && !DIST_PHASES.includes(u.phase))) ? "active" : "pending";
  if (u.state === "completed" && act !== "distribute" && !/^Already running/.test(u.message || "")) return "done";
  if (u.state === "attention" && act !== "distribute") return "warn";
  if (u.state === "failed" && act !== "distribute" && !DIST_PHASES.includes(u.phase)) return "failed";
  if (st.upgraded && !st.staged) return "done";
  return "pending";
}
function drSteps(h, d) {
  const u = h.upgrade, hasBackup = !!(d && d.backups.runs.length);
  const steps = [
    ["Credentials", {ok: "done", testing: "active", failed: "failed"}[h.creds.state] || "pending"],
    ["Readiness", {ready: "done", warning: "warn", not_ready: "failed", checking: "active"}[h.readiness.state] || "pending"],
    ["Backup", {completed: "done", running: "active", queued: "active", failed: "failed"}[h.backup.state] || (hasBackup ? "done" : "pending")],
    ["Distribute", distStep(h)],
    ["Activate", actStep(h)],
    ["Post-check", {passed: "done", warning: "warn", failed: "failed", checking: "active"}[h.postcheck.state] || "pending"]
  ];
  const mark = {done: "\u2713", warn: "!", failed: "\u2715", active: "\u25CF"};
  return '<div class="steps">' + steps.map((s, i) =>
    '<div class="step ' + s[1] + '"><span class="c">' + (mark[s[1]] || (i + 1)) + "</span>" + esc(s[0]) + "</div>").join("") + "</div>";
}

function drStageLine(h) {
  const s = stageInfo(h);
  if (!s) return '<div class="dr-note">Image stage: no image distributed or upgraded by this tool yet.</div>';
  return '<div class="dr-note">Image stage: ' + pill(s[0], s[1]) + (s[2] ? " <b>" + esc(s[2]) + "</b>" : "") + (s[3] ? " \u00B7 " + esc(s[3]) : "") + "</div>";
}

// ---- Overview ----
function drOverview(h, d) {
  const f = d ? d.facts : {}, u = h.upgrade, bk = h.backup, r = h.readiness;
  const newest = d && d.backups.runs[0];
  const img = $("image-select").value || "";
  const tgt = d && d.target ? d.target
            : img ? {image: img, version: (img.match(/(\d+\.\d+\.\d+[a-z]?)/) || [])[1] || null, problems: []} : null;

  let bv, bs;
  if (bk.state === "queued" || bk.state === "running") { bv = pill("RUNNING", "RUN"); bs = esc(bk.detail); }
  else if (bk.state === "failed") { bv = pill("FAILED", "DOWN"); bs = esc(bk.detail); }
  else if (bk.state === "completed") { bv = pill("SUCCESS", "OK"); bs = esc(newest ? newest.time : (bk.time || "")); }
  else if (newest) { bv = pill("SAVED", "OK"); bs = "Last backup " + esc(newest.time); }
  else { bv = pill("NOT RUN", "PENDING"); bs = "No backup yet"; }
  const rm = READY_PILL[h.readiness.state];

  let out = '<div class="dr-grid">' +
    drCard("Current Version", h.version ? esc(h.version) : DASH, h.version_note ? esc(h.version_note) : "") +
    drCard("Target Version", tgt ? (tgt.version ? esc(tgt.version) : "unknown") : DASH,
           tgt ? esc(tgt.image) + (tgt.problems.length ? "<br>" + esc(tgt.problems.join("; ")) : "") : (ME.role === "viewer" ? "No target image chosen yet" : "Pick an image under Firmware upgrade"),
           !!(tgt && tgt.problems.length)) +
    drCard("Configuration Backup", bv, bs) +
    drCard("Preflight Check (readiness)", rm ? pill(rm[0], rm[1]) : pill("NOT RUN", "PENDING"),
           esc(r.time || (rm ? "" : "Run the readiness check"))) +
    "</div>";

  // device information
  const m = d ? d.monitor : {};
  const rows = [
    ["Hostname", h.label || f.hostname], ["Management IP", h.host], ["Model", h.platform],
    ["Serial Number", f.serial], ["Hardware Version", f.hardware_version], ["Software Version", h.version],
    ["Boot mode", h.boot && h.boot.mode ? h.boot.mode + (h.boot.detail ? " \u2014 " + h.boot.detail : "") : null],
    ["Credential profile", h.profile], ["Uptime (at last backup)", f.uptime], ["Last reload reason", f.last_reload],
    ["Running image", f.image_file],
    ["Availability", h.availability != null ? h.availability.toFixed(1) + "% (" + h.successes + " of " + h.checks + " pings)" : null],
    ["Last checked", m.last_checked], ["Last status change", h.last_change]];
  for (const s of (f.members || []))
    rows.push(["Stack member " + s.switch + (s.active ? " (active)" : ""), s.model + " \u00B7 " + s.version + " \u00B7 " + s.mode]);
  out += drFold("devinfo", "Device Information", "", drKV(rows, true) +
    (d && !f.serial ? '<div class="dr-note">Serial number, hardware version and uptime are read from the newest saved ' +
      "<i>show version</i> / <i>show inventory</i> backup. Run a backup to fill them in.</div>" :
     d && f.read_from_backup ? '<div class="dr-note">Read from the backup of ' + esc(f.read_from_backup) + ".</div>" : ""), false);

  // upgrade progress
  const pct = drProgress(h);
  const text = h.scheduled && ["idle", "completed", "failed", "attention"].includes(u.state) ? (h.scheduled.queued ? "Queued - waiting for a free upgrade slot" : "Scheduled for " + h.scheduled.run_at_str)
             : u.state === "idle" ? (d && d.missing === 0 ? "Ready to upgrade" : "Not started")
             : u.message;
  out += drSec("Upgrade Progress", u.job_id ? '<a href="#" class="clink" data-console="1">Open live console</a>' : "",
    drStageLine(h) + drSteps(h, d) + '<div class="dr-barrow"><div class="dr-bar"><i style="width:' + pct + '%"></i></div><b>' + pct + "%</b></div>" +
    '<div class="dr-note">' + esc(text) + "</div>");

  // recent events
  const ev = d ? d.events.slice(0, 5) : [];
  out += drSec("Recent Events", d && d.events.length > 5 ? '<button type="button" class="btn sm" data-goto="events">View All Events</button>' : "",
    ev.length ? ev.map(drEventRow).join("") : '<div class="dr-empty">' + (d ? "No events for this switch yet." : "Loading ...") + "</div>");
  return out;
}
function drEventRow(e) {
  return '<div class="dr-ev"><span class="dot ' + esc(e.level) + '"></span><span class="x">' + esc(e.text) + '</span><span class="t">' + esc(e.time) + "</span></div>";
}

// ---- Checks: what was read from the switch by the credential test, readiness check, pre- and post-check ----
function drChecks(h, d) {
  const c = h.creds, r = h.readiness, ck = r.checks || {}, pc = h.postcheck, pk = pc.checks || {};
  const empty = '<div class="dr-empty">Not run yet.</div>';
  let out = "";

  const cm = {testing: ["TESTING", "RUN"], ok: ["OK", "OK"], failed: ["FAILED", "DOWN"]}[c.state];
  out += drSec("Credentials", cm ? pill(cm[0], cm[1]) : pill("NOT TESTED", "PENDING"),
    drKV([["Profile", h.profile || "No profile attached"], ["Result", c.detail], ["Tested at", c.time]]) || empty);

  const rm = READY_PILL[r.state];
  const boot = ck.boot ? (ck.boot.mode || "?") + (ck.boot.detail ? " \u2014 " + ck.boot.detail : "") : ck.boot_error;
  out += drSec("Readiness check", rm ? pill(rm[0], rm[1]) : pill("NOT RUN", "PENDING"),
    drKV([["Result", r.detail], ["Checked at", r.time], ["Version", ck.version], ["Model", ck.model], ["Boot", boot],
          ["Flash space", ck.space_detail], ["CPU", ck.cpu_pct != null ? ck.cpu_pct + "% (1-minute average)" : null],
          ["Memory", ck.mem_free_pct != null ? ck.mem_free_pct + "% free" : null], ["Stack", ck.stack],
          ["Interfaces / VLANs", ck.snapshot_text]]) || empty);

  out += drSec("Pre-upgrade snapshot", "", d && d.pre_snapshot
    ? drKV([["Taken at", d.pre_snapshot.time], ["Network state", d.pre_snapshot.summary]]) +
      '<div class="dr-note">Taken just before the last upgrade; the post-check compares against it.</div>'
    : '<div class="dr-empty">' + (d ? "None yet \u2014 it is taken automatically when an upgrade starts." : "Loading ...") + "</div>");

  const pm = {checking: ["CHECKING", "RUN"], passed: ["PASSED", "OK"], warning: ["DIFFERENCES", "CHECK"], failed: ["FAILED", "DOWN"]}[pc.state];
  const issues = (pk.issues || []).map(esc).join("<br>");
  out += drSec("Post-upgrade check", pm ? pill(pm[0], pm[1]) : pill("NOT RUN", "PENDING"),
    drKV([["Result", pc.detail], ["Checked at", pc.time], ["Version", pk.version],
          ["CPU", pk.cpu_pct != null ? pk.cpu_pct + "% (1-minute average)" : null],
          ["Memory", pk.mem_free_pct != null ? pk.mem_free_pct + "% free" : null], ["Stack", pk.stack],
          ["Interfaces / VLANs", pk.summary],
          ["Compared with pre-upgrade snapshot", pk.compared == null ? null : pk.compared ? "Yes" : "No snapshot to compare"],
          ["Differences", issues, true]]) || '<div class="dr-empty">Runs automatically after an upgrade.</div>');

  const files = d ? d.check_files : [];
  out += drSec("Saved check results and logs", "", files.length
    ? files.map(x => drFileRow(x.label, x.time + " \u00B7 " + x.name, x.size, "run", x.name)).join("")
    : '<div class="dr-empty">' + (d ? "Nothing saved yet." : "Loading ...") + "</div>");
  return out;
}

// ---- Backups ----
function drBackups(h, d) {
  const bk = h.backup;
  const m = {queued: ["QUEUED", "PENDING"], running: ["RUNNING", "RUN"], completed: ["SUCCESS", "OK"], failed: ["FAILED", "DOWN"]}[bk.state];
  let out = drSec("Latest backup", m ? pill(m[0], m[1]) : pill("NOT RUN THIS SESSION", "PENDING"),
    drKV([["Result", bk.detail], ["Finished at", bk.time]]) || '<div class="dr-empty">No backup has been run since the program started.</div>');
  if (!d) return out + '<div class="dr-empty">Loading ...</div>';
  out += drSec("Saved backups", d.backups.runs.length ? '<span class="muted" style="font-weight:400">newest first</span>' : "",
    d.backups.runs.length
      ? d.backups.runs.map(run => '<div class="dr-run"><div class="rh">' + esc(run.time) + '<span class="m">' + run.files.length + " file" +
          (run.files.length === 1 ? "" : "s") + "</span></div>" +
          run.files.map(f => drFileRow(f.command, f.name, f.size, "backup", f.name)).join("") + "</div>").join("")
      : '<div class="dr-empty">No backup files for this switch yet.</div>');
  out += '<div class="dr-note">Folder: Switch Backups / ' + esc(d.backups.folder) + ". Every backup adds new time-stamped files here" +
         (CAN_VIEW ? "; use the Storage tab to download them." : ". Viewing files needs an Operator or Administrator account.") + "</div>";
  return out;
}

// ---- Commands ----
function drCommands(h, d) {
  if (!d) return '<div class="dr-empty">Loading ...</div>';
  const all = d.commands;
  if (!all.length) return '<div class="dr-empty">No commands recorded for this switch yet. From now on, every command the program sends ' +
    "to it (backup, credential test, readiness check, version check, upgrade) is listed here.</div>";
  const srcs = [...new Set(all.map(c => c.src))];
  const cur = srcs.includes(DR.src) ? DR.src : "all";
  const count = s => s === "all" ? all.length : all.filter(c => c.src === s).length;
  let out = '<div class="chips">' + ["all"].concat(srcs).map(s => '<button type="button" data-src="' + esc(s) + '" class="' + (cur === s ? "on" : "") + '">' +
    (s === "all" ? "All" : esc(s)) + " " + count(s) + "</button>").join("") + "</div>";
  out += '<table class="cm">' + all.filter(c => cur === "all" || c.src === cur).map(c =>
    '<tr><td class="w">' + esc(c.t.slice(5)) + '</td><td><span class="src">' + esc(c.src) + "</span><br><code>" + esc(c.cmd) + "</code>" +
    '<div class="w">' + (c.err ? esc(c.err) : c.lines != null ? c.lines + " line" + (c.lines === 1 ? "" : "s") + " returned" : "") +
    (c.sec != null ? " \u00B7 " + c.sec + " s" : "") + "</div></td><td>" +
    (c.ok ? '<span class="ok">\u2713</span>' : '<span class="no">\u2715</span>') + "</td></tr>").join("") + "</table>";
  out += '<div class="dr-note">Only the command is recorded \u2014 never what the switch answered. Newest first.</div>';
  return out;
}

// ---- Events + upgrade history ----
function drEvents(h, d) {
  if (!d) return '<div class="dr-empty">Loading ...</div>';
  let out = drSec("Events", "", d.events.length ? d.events.map(drEventRow).join("")
    : '<div class="dr-empty">No events for this switch yet.</div>');
  const st = {completed: ["OK", "OK"], failed: ["FAILED", "DOWN"], attention: ["CHECK", "WARN"]};
  out += drSec("Upgrade history", "", d.history.length ? d.history.map(x => {
    const s = st[x.status] || [String(x.status || "").toUpperCase(), "PENDING"];
    const ver = x.previous_image || x.current_image ? " \u00B7 " + esc(x.previous_image || "?") + " \u2192 " + esc(x.current_image || "?") : "";
    return '<div class="hrow"><div class="a"><b>' + esc(x.action) + "</b>" + pill(s[0], s[1]) + '<span class="muted">' + esc(x.finished) + "</span></div>" +
      '<div class="b">' + (x.image_file ? esc(x.image_file) : "") + ver + (x.upgrade_duration ? " \u00B7 " + esc(x.upgrade_duration) : "") +
      (x.postcheck_result ? " \u00B7 post-check: " + esc(x.postcheck_result) : "") + (x.started_by ? " \u00B7 by " + esc(x.started_by) : "") +
      (x.message && x.status !== "completed" ? "<br>" + esc(x.message) : "") + "</div></div>";
  }).join("") : '<div class="dr-empty">No finished upgrades for this switch yet.</div>');
  return out;
}

// ---- drawer plumbing ----
function drTabsRender(d) {
  const n = {backups: d ? d.backups.runs.length : "", commands: d ? d.commands.length : "", events: d ? d.events.length : ""};
  const html = DR_TABS.filter(t => t[0] !== "console" || ME.role !== "viewer" || upHas()).map(t => '<button type="button" role="tab" data-tab="' + t[0] + '" class="' + (DR.tab === t[0] ? "on" : "") +
    '" aria-selected="' + (DR.tab === t[0]) + '">' + t[1] + (n[t[0]] !== undefined && n[t[0]] !== "" ? '<span class="n">' + n[t[0]] + "</span>" : "") +
    (t[0] === "console" && DR.host && TERMS.has(DR.host) && TERMS.get(DR.host).state === "on" ? ' <span style="color:var(--gr)">\u25CF</span>' : "") +
    (t[0] === "console" && upActive() ? ' <span style="color:var(--am)">\u25CF</span>' : "") + "</button>").join("");
  if (html !== DR.tabsHtml) { $("dr-tabs").innerHTML = html; DR.tabsHtml = html; }
}
function drRender() {
  if (!DR.host) return;
  const h = drHost();
  if (!h) { closeDrawer(); return; }                       // switch deleted or site changed
  $("dr-name").innerHTML = esc(h.label || h.host) + " " + pill(h.status, h.status);
  $("dr-sub").textContent = h.host + (h.platform ? " \u00B7 " + h.platform : "");
  if (DR.tab === "console" && ME.role === "viewer" && !(h.upgrade && h.upgrade.job_id)) DR.tab = "overview";   // nothing for a viewer to see
  drTabsRender(DR.data);
  const con = DR.tab === "console";
  $("dr-body").hidden = con; $("dr-term").hidden = !con;
  if (con) { if (DR.termShown !== DR.host) { DR.termShown = DR.host; termShow(); } else termPaint(); return; }
  DR.termShown = null;
  const view = {overview: drOverview, checks: drChecks, backups: drBackups, commands: drCommands, events: drEvents}[DR.tab];
  let html = (DR.err ? '<div class="dr-banner bad">' + esc(DR.err) + "</div>" : "") + view(h, DR.data);
  if (html !== DR.html) {
    const body = $("dr-body"), top = body.scrollTop;
    body.innerHTML = html; body.scrollTop = top; DR.html = html;
  }
}
async function drFetch(force) {
  if (!DR.host || DR.busy || (document.hidden && !force)) return;
  DR.busy = true;
  const host = DR.host;
  try {
    const r = await fetch("/api/device?host=" + encodeURIComponent(host) + "&image=" + encodeURIComponent($("image-select").value || ""), {cache: "no-store"});
    let d = {}; try { d = await r.json(); } catch (e) {}
    if (host !== DR.host) return;
    if (r.ok) { DR.data = d; DR.err = ""; } else { DR.err = d.error || "Could not load the details (HTTP " + r.status + ")"; }
  } catch (e) {
    if (host === DR.host) DR.err = "Lost connection to the script";
  } finally { DR.busy = false; }
  drRender();
}
function drMarkRow() {
  $("rows").querySelectorAll("tr").forEach(tr => {
    const a = tr.querySelector("a[data-detail]");
    tr.classList.toggle("dopen", !!a && a.dataset.detail === DR.host);
  });
}
function openDrawer(host) {
  if (DR.host !== host) {
    DR.host = host; DR.data = null; DR.err = ""; DR.html = ""; DR.src = "all"; DR.tab = "overview"; closeViewer();
    $("dr-body").scrollTop = 0;
  }
  $("drawer").hidden = false;
  document.body.classList.add("drawer-open");
  drMarkRow(); drRender(); drFetch(true);
  clearInterval(DR.timer); DR.timer = setInterval(() => drFetch(false), 3000);
}
function closeDrawer() {
  termMax(false);
  clearInterval(DR.timer); DR.timer = null;
  DR.host = null; DR.data = null; DR.html = "";
  closeViewer();
  $("drawer").hidden = true;
  document.body.classList.remove("drawer-open");
  drMarkRow();
}
function drSetTab(t) { if (DR.tab === "console" && t !== "console") termMax(false); DR.tab = t; DR.html = ""; closeViewer(); $("dr-body").scrollTop = 0; drRender(); }
function drawerSync() {
  if (!DR.host) return;
  if (!last.hosts.some(x => x.host === DR.host)) { closeDrawer(); return; }
  drRender();
}
function closeViewer() { DR.viewing = false; $("dr-view").hidden = true; $("dr-view-text").textContent = ""; }
const DRV = {kind: "", name: "", tail: false, truncated: false};
async function drView(kind, name, tail) {
  const host = DR.host;
  DRV.kind = kind; DRV.name = name;
  DR.viewing = true;
  $("dr-view").hidden = false;
  $("dr-view-name").textContent = name;
  $("dr-view-text").textContent = "Loading ...";
  $("dr-view-note").hidden = true;
  try {
    const r = await fetch("/api/device/file?host=" + encodeURIComponent(host) + "&kind=" + encodeURIComponent(kind) + "&name=" + encodeURIComponent(name) + (tail ? "&tail=1" : ""), {cache: "no-store"});
    let d = {}; try { d = await r.json(); } catch (e) {}
    if (!DR.viewing || host !== DR.host) return;
    if (!r.ok) throw new Error(d.error || "HTTP " + r.status);
    const pre = $("dr-view-text");
    DRV.tail = !!d.tail; DRV.truncated = !!d.truncated;
    pre.textContent = d.text || "(empty file)";
    pre.scrollTop = tail || d.tail ? pre.scrollHeight : 0;
    if (d.truncated) {
      $("dr-view-note").textContent = "Large file (" + drSize(d.size) + ") \u2014 showing the " + (d.tail ? "last" : "first") + " " + drSize(d.limit || 0) + ". Use " + (d.tail ? "Top" : "End") + " for the other part.";
      $("dr-view-note").hidden = false;
    }
  } catch (e) {
    if (DR.viewing) $("dr-view-text").textContent = "Could not open the file: " + e.message;
  }
}
$("dr-view-top").addEventListener("click", () => { if (DRV.tail && !/\.(log|console\.txt)$/.test(DRV.name)) drView(DRV.kind, DRV.name, false); else $("dr-view-text").scrollTop = 0; });
$("dr-view-end").addEventListener("click", () => { if (DRV.truncated && !DRV.tail) drView(DRV.kind, DRV.name, true); else $("dr-view-text").scrollTop = $("dr-view-text").scrollHeight; });
function drConsole() {
  const h = drHost();
  if (h && h.upgrade.job_id) { UP.pref[h.host] = "upgrade"; drSetTab("console"); }
}

$("dr-ic").innerHTML = IC[1];
$("dr-close").innerHTML = I('<path d="M6 6l12 12M18 6 6 18"/>');
$("dr-close").addEventListener("click", closeDrawer);
$("dr-view-back").addEventListener("click", closeViewer);
$("rows").addEventListener("click", e => {
  const a = e.target.closest("a[data-detail]");
  if (!a) return;
  e.preventDefault();
  openDrawer(a.dataset.detail);
});
$("drawer").addEventListener("click", e => {
  const a = e.target.closest("[data-tab],[data-goto],[data-src],[data-view],[data-console],[data-fold]");
  if (!a) return;
  if (a.tagName === "A") e.preventDefault();
  if (a.dataset.fold) {
    DR_FOLD[a.dataset.fold] = a.getAttribute("aria-expanded") !== "true";
    try { localStorage.setItem("drFold", JSON.stringify(DR_FOLD)); } catch (x) {}
    DR.html = ""; drRender();
  } else if (a.dataset.tab) drSetTab(a.dataset.tab);
  else if (a.dataset.goto) drSetTab(a.dataset.goto);
  else if (a.dataset.src !== undefined) { DR.src = a.dataset.src; drRender(); }
  else if (a.dataset.view) drView(a.dataset.view, a.dataset.name);
  else if (a.dataset.console !== undefined) drConsole();
});

// ---- Console tab: an interactive SSH session to the switch. The SSH connection is held by this program
// (the login never reaches the browser); the page sends lines and reads new output every half second. ----
const TERMS = new Map();           // host -> {sid, offset, text, state: off|connecting|on|closed, hist, hi, polling}
const TUI = {max: false};
const termGet = h => { if (!TERMS.has(h)) TERMS.set(h, {sid: null, offset: 0, text: "", state: "off", hist: [], hi: -1, polling: false}); return TERMS.get(h); };
const termVisible = h => DR.host === h && DR.tab === "console";
const termPrompt = t => /(?:password|secret|passphrase)[^\n]*:\s*$/i.test(t.text.slice(t.text.lastIndexOf("\n") + 1));

function termPaint() {
  const h = DR.host; if (!h) return;
  const hh = drHost();
  if (termUpgradeView(hh)) { upPaint(hh); return; }
  termSshChrome(hh);
  const t = termGet(h), out = $("dr-term-out"), st = $("dr-term-state");
  const live = t.state === "on" || t.state === "connecting";
  const atEnd = out.scrollHeight - out.scrollTop - out.clientHeight < 40;
  if (out.textContent !== t.text) { out.textContent = t.text || ("Not connected. Click Connect to open " + (hostProto(h) === "Telnet" ? "a Telnet" : "an SSH") + " session to this switch."); if (atEnd || !t.text) out.scrollTop = out.scrollHeight; }
  st.textContent = {off: "Not connected", connecting: "Connecting ...", on: "Connected", closed: "Session closed"}[t.state];
  st.className = "tst" + (t.state === "on" ? " on" : t.state === "connecting" ? " busy" : "");
  $("dr-term-title").textContent = (drHost() && (drHost().label || h)) || h;
  $("dr-term-connect").textContent = live ? "Disconnect" : "Connect";
  $("dr-term-connect").className = "tb" + (live ? "" : " go");
  $("dr-term-ctrlc").disabled = !live;
  const inp = $("dr-term-in");
  const wasOff = inp.disabled;
  inp.disabled = !live;
  if (wasOff && live) inp.focus();          // just connected: type straight away
  const pw = live && termPrompt(t);
  inp.type = pw ? "password" : "text";
  inp.placeholder = pw ? "Enter the password (not shown, not recorded)" : "Type a command and press Enter";
  DR.tabsHtml = ""; drTabsRender(DR.data);
}
function termShow() {
  const h = DR.host, t = termGet(h);
  termPaint();
  if (t.sid && !t.polling) termPoll(h);
  if (!$("dr-term-in").disabled) $("dr-term-in").focus();
}
function termMax(on) {
  TUI.max = !!on;
  $("dr-term").classList.toggle("max", TUI.max);
  $("dr-term-bd").hidden = !TUI.max;
  $("dr-term-max").textContent = TUI.max ? "Restore" : "Maximise";
  const out = $("dr-term-out"); out.scrollTop = out.scrollHeight;
  if (!$("dr-term-in").disabled) $("dr-term-in").focus();
}
async function termConnect(host) {
  const t = termGet(host);
  t.text = ""; t.offset = 0; t.state = "connecting"; termPaint();
  try {
    const d = await post("/api/term/open", {host: host});
    t.sid = d.sid; t.offset = 0;
    termPoll(host);
  } catch (e) {
    t.state = "off"; t.text = "[" + e.message + "]\n"; t.sid = null;
  }
  if (DR.host === host) termPaint();
}
async function termDisconnect(host) {
  const t = termGet(host), sid = t.sid;
  t.sid = null; t.state = "off"; t.text += "\n[Disconnected]\n";
  if (DR.host === host) termPaint();
  if (sid) { try { await post("/api/term/close", {host: host, sid: sid}); } catch (e) {} }
}
async function termPoll(host) {
  const t = termGet(host);
  if (t.polling) return;
  t.polling = true;
  try {
    while (t.sid && termVisible(host)) {
      const sid = t.sid;
      let d = null;
      try {
        const r = await fetch("/api/term/read?host=" + encodeURIComponent(host) + "&sid=" + encodeURIComponent(sid) + "&offset=" + t.offset, {cache: "no-store"});
        d = await r.json();
        if (!r.ok) { if (r.status === 404) { t.sid = null; t.state = "closed"; t.text += "\n[The session no longer exists on the server]\n"; break; } throw new Error(d.error || r.status); }
      } catch (e) { await new Promise(res => setTimeout(res, 1500)); continue; }
      if (t.sid !== sid) break;                       // disconnected meanwhile
      if (d.text) { t.text += d.text; if (t.text.length > 300000) t.text = t.text.slice(-200000); }
      t.offset = d.offset;
      t.state = d.closed ? "closed" : (d.ready ? "on" : "connecting");
      if (termVisible(host)) termPaint();
      if (d.closed) { t.sid = null; break; }
      await new Promise(res => setTimeout(res, d.text ? 120 : 450));
    }
  } finally { t.polling = false; }
  if (termVisible(host)) termPaint();
}
async function termSend(payload) {
  const host = DR.host, t = termGet(host);
  if (!t.sid) return;
  try { await post("/api/term/send", Object.assign({host: host, sid: t.sid}, payload)); }
  catch (e) { t.text += "\n[" + e.message + "]\n"; termPaint(); }
}
$("dr-term-connect").addEventListener("click", () => {
  const h = DR.host, t = termGet(h);
  (t.state === "on" || t.state === "connecting") ? termDisconnect(h) : termConnect(h);
});
$("dr-term-ctrlc").addEventListener("click", () => { termSend({raw: "ctrl_c"}); $("dr-term-in").focus(); });
$("dr-term-clear").addEventListener("click", () => { const t = termGet(DR.host); t.text = ""; termPaint(); $("dr-term-in").focus(); });
$("dr-term-max").addEventListener("click", () => termMax(!TUI.max));
$("dr-term-bd").addEventListener("click", () => termMax(false));
$("dr-term-out").addEventListener("click", () => { if (!String(getSelection())) $("dr-term-in").focus(); });
$("dr-term-in").addEventListener("keydown", e => {
  const t = termGet(DR.host), inp = e.target;
  if (e.key === "Enter") {
    e.preventDefault();
    const line = inp.value, secret = inp.type === "password";
    inp.value = "";
    if (!secret && line.trim()) { t.hist.push(line); t.hist = t.hist.slice(-100); }
    t.hi = -1;
    termSend({line: line, secret: secret});
  } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
    if (inp.type === "password" || !t.hist.length) return;
    e.preventDefault();
    t.hi = e.key === "ArrowUp" ? Math.min(t.hist.length - 1, t.hi + 1) : Math.max(-1, t.hi - 1);
    inp.value = t.hi < 0 ? "" : t.hist[t.hist.length - 1 - t.hi];
  } else if (e.ctrlKey && (e.key === "c" || e.key === "C") && inp.selectionStart === inp.selectionEnd) {
    e.preventDefault(); inp.value = ""; termSend({raw: "ctrl_c"});
  } else if (e.ctrlKey && (e.key === "z" || e.key === "Z")) {
    e.preventDefault(); termSend({raw: "ctrl_z"});
  }
});
$("image-select").addEventListener("change", () => { if (DR.host) { DR.html = ""; drRender(); drFetch(true); } });
document.addEventListener("keydown", e => {
  if (e.key !== "Escape" || !DR.host) return;
  if (TUI.max) { termMax(false); return; }
  if (document.querySelector(".modal:not([hidden])")) return;     // those handle Escape themselves
  if (DR.viewing) closeViewer(); else closeDrawer();
});

loadSites();
applySite(null);

refresh();
setInterval(refresh, 2000);
</script>
</body>
</html>
"""


def make_handler(manager, allowed_hosts, any_host):
    class Handler(BaseHTTPRequestHandler):
        def _send(self, code, body, ctype, extra=None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _me(self):
            """The signed-in user and role, for the dashboard. Without the access guard (local
            command-line use) there is no sign-in, so the person is a local administrator."""
            role = getattr(self, "guard_role", None) or ROLE_ADMIN
            return {"user": self._who(), "role": role, "role_label": ROLE_LABELS.get(role, "Viewer"),
                    "sign_out": hasattr(self, "guard_user")}

        def _who(self):
            """Signed-in user (set by the access guard); 'local' when no guard is in use."""
            return getattr(self, "guard_user", None) or "local"

        def _json(self, code, obj):
            self._send(code, json.dumps(obj).encode("utf-8"), "application/json")

        def _send_stream(self, fobj, size, filename, ctype="application/octet-stream"):
            """Send an open file as a download, in chunks (backups can be large, and a
            folder ZIP is built on disk rather than in memory)."""
            ascii_name = re.sub(r"[^A-Za-z0-9._-]", "_", filename) or "download"
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(size))
            self.send_header("Content-Disposition",
                             f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}")
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            self.end_headers()
            remaining = size
            try:
                while remaining > 0:
                    chunk = fobj.read(min(STORAGE_CHUNK, remaining))
                    if not chunk:
                        break
                    self.wfile.write(chunk)
                    remaining -= len(chunk)
            except OSError:            # browser cancelled / connection dropped
                pass

        def _storage_download(self, site, rel):
            name, root = storage_root(site)
            path, clean = storage_resolve(root, rel)
            if path.is_dir():                       # a folder (or the whole site) -> ZIP
                fobj, size, filename = storage_zip_folder(site, clean)
                try:
                    self._send_stream(fobj, size, filename, "application/zip")
                finally:
                    fobj.close()
            elif path.is_file():
                try:
                    fobj = open(path, "rb")
                except OSError as exc:
                    raise ApiError(f"Could not read that file ({type(exc).__name__}).", 500)
                try:
                    self._send_stream(fobj, os.fstat(fobj.fileno()).st_size, path.name)
                finally:
                    fobj.close()
            else:
                raise ApiError("File not found.", 404)

        UPLOAD_PATHS = ("/api/images/upload", "/api/images/checksum-upload", "/api/devices/upload")

        def _check_request(self, post, path=None):
            """Reject DNS-rebinding and cross-site requests (this GUI can start upgrades)."""
            host_header = self.headers.get("Host", "")
            hostname = (urlsplit("//" + host_header).hostname or "").lower()
            if not any_host and hostname not in allowed_hosts:
                raise ApiError("Host not allowed.", 403)
            if post:
                ctype = self.headers.get("Content-Type", "").split(";")[0].strip().lower()
                expected = "multipart/form-data" if path in self.UPLOAD_PATHS else "application/json"
                if ctype != expected:
                    raise ApiError(f"Content-Type must be {expected}.", 415)
                origin = self.headers.get("Origin")
                if origin is not None and urlsplit(origin).netloc != host_header:
                    raise ApiError("Cross-origin requests are not allowed.", 403)

        def do_GET(self):  # noqa: N802
            try:
                self._check_request(False)
                parts = urlparse(self.path)
                path, query = parts.path, parse_qs(parts.query)
                qv = lambda k: (query.get(k) or [""])[0]  # noqa: E731
                if path in ("/", "/index.html"):
                    self._send(200, DASHBOARD_HTML.replace("__APP_VERSION__", APP_VERSION)
                                              .replace("__APP_AUTHOR__", APP_AUTHOR)
                                              .replace("__APP_ICON__", APP_ICON_DATA_URI)
                                              .replace("__APP_REPO_URL__", APP_REPO_URL)
                                              .replace("__ROLE__", self._me()["role"])
                                              .replace("__ME__", json.dumps(self._me()).replace("<", "\\u003c"))
                                              .encode("utf-8"),
                               "text/html; charset=utf-8")
                elif path == "/api/status":
                    snap = manager.snapshot()
                    snap["me"] = {"role": self._me()["role"]}     # lets the page notice a role change
                    self._json(200, snap)
                elif path == "/api/me":
                    self._json(200, self._me())
                elif path == "/api/images":
                    self._json(200, {"images": list_images()})
                elif path == "/api/images/details":
                    self._json(200, {"images": list_image_details()})
                elif path == "/api/settings":
                    self._json(200, manager.settings_view())
                elif path == "/api/timeout-profiles":
                    self._json(200, manager.timeout_profiles_view())
                elif path == "/api/device":
                    self._json(200, manager.device_detail(qv("host"), qv("image"), shared_target=self._me()["role"] == ROLE_VIEWER))
                elif path == "/api/device/file":
                    self._json(200, manager.device_file(qv("host"), qv("kind"), qv("name"), qv("tail") == "1"))
                elif path == "/api/term/read":
                    self._json(200, manager.term_read(qv("host"), qv("sid"), qv("offset")))
                elif path == "/api/console":
                    q = lambda k: (query.get(k) or [""])[0]  # noqa: E731
                    self._json(200, manager.console_read(q("host"), q("job"), q("offset")))
                elif path == "/api/sites":
                    self._json(200, {"sites": list_site_profiles(), "current": manager.site_name})
                elif path == "/api/topology":
                    self._json(200, manager.topology_view())
                elif path == "/api/storage/list":
                    self._json(200, storage_list(qv("site"), qv("path")))
                elif path == "/api/storage/search":
                    self._json(200, storage_search(qv("site"), qv("q")))
                elif path == "/api/storage/backups":
                    self._json(200, storage_backup_runs(qv("site"), qv("path")))
                elif path == "/api/storage/compare":
                    self._json(200, storage_compare(qv("site"), qv("path"), qv("a"), qv("b"), qv("ignore") != "0", qv("cmd") or None))
                elif path == "/api/storage/read":
                    self._json(200, storage_read(qv("site"), qv("path"), qv("tail") == "1"))
                elif path == "/api/storage/download":
                    self._storage_download(qv("site"), qv("path"))
                elif path == "/api/devices/template.xlsx":
                    self._send(200, device_list_template_xlsx(),
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               {"Content-Disposition": 'attachment; filename="device_list_template.xlsx"'})
                elif path == "/api/notifications/email":
                    self._json(200, manager.email_settings_view())
                elif path == "/api/upgrade-report.csv":
                    self._send(200, manager.upgrade_report_csv().encode("utf-8"), "text/csv; charset=utf-8",
                               {"Content-Disposition": f'attachment; filename="upgrade_report_v{APP_VERSION}.csv"'})
                elif path == "/api/upgrade-report.xlsx":
                    self._send(200, manager.upgrade_report_xlsx(self._who()),
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               {"Content-Disposition": f'attachment; filename="upgrade_report_v{APP_VERSION}.xlsx"'})
                elif path == "/api/upgrade-report/print":
                    self._send(200, manager.upgrade_report_html(self._who()).encode("utf-8"),
                               "text/html; charset=utf-8")
                else:
                    self._send(404, b"Not found", "text/plain")
            except ApiError as exc:
                self._json(exc.status, {"error": str(exc)})

        def do_POST(self):  # noqa: N802
            # A large image upload, or any error partway through one, must not leave the
            # connection in a state where leftover bytes get read as the next request.
            self.close_connection = True
            path = self.path.split("?", 1)[0]
            try:
                self._check_request(True, path)
                if path in self.UPLOAD_PATHS:
                    self._handle_upload(path)
                    return
                length = int(self.headers.get("Content-Length") or 0)
                if length > 65536:
                    raise ApiError("Request too large.", 413)
                try:
                    body = json.loads(self.rfile.read(length) or b"{}")
                    if not isinstance(body, dict):
                        raise ValueError
                except ValueError:
                    raise ApiError("Invalid JSON.")

                if path == "/api/site/select":
                    result = manager.switch_site(body.get("name"), mode="select")
                elif path == "/api/site/create":
                    result = manager.switch_site(body.get("name"), mode="create")
                elif path == "/api/monitor/start":
                    result = manager.start_monitoring()
                elif path == "/api/monitor/stop":
                    result = manager.stop_monitoring()
                elif path == "/api/devices/add":
                    result = manager.add_device_manual(body.get("hostname"), body.get("ip"), body.get("platform"))
                elif path == "/api/devices/delete":
                    result = manager.delete_devices(body.get("hosts"))
                elif path == "/api/backup":
                    result = manager.start_backup(body.get("hosts"), body.get("custom_commands"))
                elif path == "/api/version":
                    result = manager.start_version_check(body.get("hosts"))
                elif path == "/api/test-credentials":
                    result = manager.test_credentials(body.get("hosts"), body.get("profile"))
                elif path == "/api/profile/test":
                    result = manager.test_credentials(body.get("hosts"), None, {
                        "username": body.get("username"), "password": body.get("password"),
                        "secret": body.get("secret"), "protocol": body.get("protocol"),
                        "port": body.get("port"), "original": body.get("original"),
                        "clear_secret": bool(body.get("clear_secret"))})
                elif path == "/api/profile/save":
                    result = manager.save_profile(body.get("name"), body.get("username"), body.get("password"),
                                                  body.get("secret"), bool(body.get("clear_secret")),
                                                  body.get("original"), body.get("protocol") or "ssh", body.get("port"))
                elif path == "/api/notifications/clear":
                    result = manager.clear_notifications()
                elif path == "/api/profile/delete":
                    result = manager.delete_profile(body.get("name"))
                elif path == "/api/topology/discover":
                    result = manager.start_topology_discovery(body.get("hosts"))
                elif path == "/api/topology/root":
                    result = manager.topology_set_root(body.get("host"), bool(body.get("on")))
                elif path == "/api/profile/assign":
                    result = manager.assign_profile(body.get("hosts"), body.get("profile"))
                elif path == "/api/term/open":
                    result = manager.term_open(body.get("host"), self._who())
                elif path == "/api/term/send":
                    result = manager.term_send(body.get("host"), body.get("sid"), body.get("line"), body.get("raw"),
                                               bool(body.get("secret")), self._who())
                elif path == "/api/term/close":
                    result = manager.term_close(body.get("host"), body.get("sid"), self._who())
                elif path == "/api/readiness-check":
                    result = manager.start_readiness_check(body.get("hosts"), body.get("image_path"))
                elif path == "/api/upgrade":
                    result = manager.start_upgrade(body.get("hosts"), body.get("image_path"),
                                                   bool(body.get("backup_first")), body.get("action") or "full",
                                                   self._who(), bool(body.get("skip_image_check")),
                                                   bool(body.get("ignore_space")))
                elif path == "/api/schedule":
                    result = manager.schedule_upgrade(body.get("hosts"), body.get("image_path"), body.get("run_at"),
                                                      bool(body.get("backup_first")),
                                                      body.get("action") or "full",
                                                      self._who(), bool(body.get("skip_image_check")),
                                                      bool(body.get("ignore_space")))
                elif path == "/api/schedule/cancel":
                    result = manager.cancel_schedule(body.get("hosts"))
                elif path == "/api/settings":
                    result = manager.update_settings(body.get("values"), self._who())
                elif path == "/api/notifications/email":
                    result = manager.update_email_settings(body.get("values"), self._who())
                elif path == "/api/notifications/email/test":
                    result = manager.test_email_settings(body.get("values"), self._who())
                elif path == "/api/timeout-profiles/save":
                    result = manager.timeout_profile_save(body.get("name"), body.get("values"), body.get("note"),
                                                          body.get("source"), body.get("site"), body.get("switch"),
                                                          bool(body.get("overwrite")), self._who())
                elif path == "/api/timeout-profiles/delete":
                    result = manager.timeout_profile_delete(body.get("name"), self._who())
                elif path == "/api/timeout-profiles/apply":
                    result = manager.timeout_profile_apply(body.get("name"), self._who())
                elif path == "/api/command-test/start":
                    result = manager.command_test_start(body.get("host"), body.get("id"), self._who(), body.get("image"))
                elif path == "/api/command-test/stop":
                    result = manager.command_test_stop(body.get("job"))
                elif path == "/api/command-test/poll":
                    result = manager.command_test_poll(body.get("job"), body.get("offset"))
                elif path == "/api/images/verify":
                    name = safe_upload_filename(str(body.get("name") or ""))
                    image = images_dir() / name
                    if not name.lower().endswith(".bin") or not image.is_file():
                        raise ApiError("That image is not in the IOS Images folder.", 404)
                    try:
                        verify_image_checksum_sidecar(image)
                    except ValueError as exc:
                        raise ApiError(f"{name}: {exc}", 409)
                    result = {"message": f"{name}: the file matches its saved {image_checksum_kind(image)} checksum."}
                else:
                    raise ApiError("Not found.", 404)
                self._json(200, result)
            except ApiError as exc:
                self._json(exc.status, {"error": str(exc)})
            except Exception as exc:  # noqa: BLE001 - never echo details (could contain input)
                print(f"Internal error: {type(exc).__name__}", flush=True)
                self._json(500, {"error": f"Internal error ({type(exc).__name__}). See the console."})

        def _handle_upload(self, path):
            length = int(self.headers.get("Content-Length") or 0)
            ctype = self.headers.get("Content-Type", "")
            try:
                if path == "/api/images/upload":
                    expected_md5 = (self.headers.get("X-Image-MD5") or "").strip().lower()
                    chosen = {}

                    def dest_for_image(filename):
                        name = safe_upload_filename(filename)
                        if not name.lower().endswith(".bin"):
                            raise ValueError("Only .bin image files can be uploaded.")
                        final_path = images_dir() / name
                        temp_path = images_dir() / ("." + name + "." + secrets.token_hex(8) + ".upload")
                        chosen.update(name=name, final_path=final_path, temp_path=temp_path)
                        return str(temp_path)

                    _, _, size, temp_path = read_multipart_file(
                        self.rfile, length, ctype, MAX_IMAGE_UPLOAD_BYTES, dest_for_image)
                    temp_path = Path(temp_path)
                    checksum_temp = None
                    backups = []
                    installed = []
                    try:
                        if not re.fullmatch(r"[0-9a-f]{32}", expected_md5):
                            raise ValueError("Paste the 32-character MD5 checksum shown on Cisco's download page.")
                        verify_image_md5(temp_path, expected_md5)
                        sidecar_path = Path(str(chosen["final_path"]) + ".md5")
                        older_sidecars = [Path(str(chosen["final_path"]) + ext) for ext in (".sha512", ".sha256")]
                        checksum_temp = Path(str(sidecar_path) + "." + secrets.token_hex(8) + ".upload")
                        checksum_temp.write_text(expected_md5 + "\n", encoding="ascii")

                        # Back up existing files so a failed replacement can be rolled back (a replaced image
                        # must not keep an older SHA sidecar that belongs to the previous file).
                        for old_path in (chosen["final_path"], sidecar_path, *older_sidecars):
                            if old_path.exists():
                                backup = Path(str(old_path) + "." + secrets.token_hex(8) + ".backup")
                                os.replace(old_path, backup)
                                backups.append((old_path, backup))
                        os.replace(temp_path, chosen["final_path"])
                        installed.append(chosen["final_path"])
                        os.replace(checksum_temp, sidecar_path)
                        installed.append(sidecar_path)
                    except (OSError, ValueError) as exc:
                        for installed_path in installed:
                            try:
                                installed_path.unlink()
                            except OSError:
                                pass
                        for old_path, backup_path in reversed(backups):
                            try:
                                os.replace(backup_path, old_path)
                            except OSError:
                                pass
                        for leftover in (temp_path, checksum_temp):
                            if leftover is not None:
                                try:
                                    leftover.unlink()
                                except OSError:
                                    pass
                        raise ApiError("Image rejected: " + str(exc).replace(temp_path.name, chosen["name"])) from exc
                    for _, backup_path in backups:
                        try:
                            backup_path.unlink()
                        except OSError:
                            pass
                    self._json(200, {
                        "message": f"Uploaded {chosen['name']} and verified its MD5 checksum "
                                   f"({size / 1_048_576:.1f} MB).",
                        "filename": chosen["name"], "images": list_images(),
                    })

                elif path == "/api/images/checksum-upload":
                    chosen = {}

                    def dest_for_checksum(filename):
                        name = safe_upload_filename(filename)
                        if not name.lower().endswith(".bin.sha256"):
                            raise ValueError("Checksum file must be named <image>.bin.sha256.")
                        image_name = name[:-len(".sha256")]
                        image_path = images_dir() / image_name
                        if not image_path.is_file() or image_path.suffix.lower() != ".bin":
                            raise ValueError("Upload the matching .bin image before its checksum.")
                        chosen.update(name=name, image_path=image_path,
                                      final_path=Path(str(image_path) + ".sha256"))
                        return str(images_dir() / ("." + name + "." + secrets.token_hex(8) + ".upload"))

                    _, _, size, temp_path = read_multipart_file(
                        self.rfile, length, ctype, MAX_CHECKSUM_UPLOAD_BYTES, dest_for_checksum)
                    try:
                        verify_image_sha256_sidecar(chosen["image_path"], sidecar_path=temp_path)
                        os.replace(temp_path, chosen["final_path"])
                    except (OSError, ValueError) as exc:
                        try:
                            os.remove(temp_path)
                        except OSError:
                            pass
                        raise ApiError(f"Checksum manifest rejected: {exc}") from exc
                    self._json(200, {"message": f"Verified Cisco SHA-256 manifest for {chosen['name']}."})

                elif path == "/api/devices/upload":
                    filename, data, size, _ = read_multipart_file(
                        self.rfile, length, ctype, MAX_DEVICE_LIST_UPLOAD_BYTES, dest_path_fn=None)
                    try:
                        hosts = load_hosts_from_bytes(data)
                    except ValueError as exc:
                        raise ApiError(str(exc))
                    if not hosts:
                        raise ApiError("No devices found in that file.")
                    added, skipped = manager.add_devices_bulk(hosts, filename)
                    msg = f"Added {added} device(s) from {filename}."
                    if skipped:
                        msg += f" {skipped} were already in the list."
                    self._json(200, {"message": msg, "added": added, "skipped": skipped})

                else:
                    raise ApiError("Not found.", 404)
            except ApiError as exc:
                self._json(exc.status, {"error": str(exc)})
            except ValueError as exc:  # malformed upload body
                self._json(400, {"error": str(exc)})
            except Exception as exc:  # noqa: BLE001
                print(f"Internal error during upload: {type(exc).__name__}: {exc}", flush=True)
                self._json(500, {"error": f"Internal error ({type(exc).__name__}). See the console."})

        def log_message(self, *args):  # keep the console clean
            pass

    return Handler


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def make_tls_server(address, handler, certfile, keyfile):
    """HTTPS version of ThreadingHTTPServer. The TLS handshake is done in the
    request's own thread, so a slow or broken client cannot stall the server."""
    import ssl
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.minimum_version = ssl.TLSVersion.TLSv1_2
    ctx.load_cert_chain(certfile, keyfile)

    class TLSServer(ThreadingHTTPServer):
        def get_request(self):
            sock, addr = super().get_request()
            return ctx.wrap_socket(sock, server_side=True, do_handshake_on_connect=False), addr

        def process_request_thread(self, request, client_address):
            try:
                request.settimeout(10)
                request.do_handshake()
                request.settimeout(None)
            except Exception:
                self.shutdown_request(request)       # e.g. a browser refusing the certificate
                return
            super().process_request_thread(request, client_address)

        def handle_error(self, request, client_address):
            import traceback
            exc = sys.exc_info()[1]
            if isinstance(exc, (ssl.SSLError, ConnectionError, TimeoutError)):
                return                                # normal noise: client went away
            traceback.print_exc()

    return TLSServer(address, handler)


def apply_access_control(base_handler, access):
    """Wrap the dashboard's request handler with login + live access rules.

    `access` is a dict owned by the launcher window and read on EVERY request,
    so changes made in the launcher apply instantly, without a restart:
        access["public"]     True: other PCs may connect; False: this PC only
        access["allowed"]    allowed client IPs/subnets for OTHER PCs (empty = any)
        access["users"]      {username: {"salt": hex, "hash": hex, "admin": bool}}
                             EVERYONE must log in, including the PC owner.
                             No users at all = nobody can get in (fail closed).
        access["epoch"]      the launcher raises it when public access is blocked,
                             which signs every REMOTE user out (this PC is unaffected)
        access["https"]      set by gui_main(); adds the Secure flag to the cookie
        access["audit_path"] file that receives the audit log
    This function adds access["active_clients"], a function the launcher calls
    to list the users who are logged in right now.

    The dashboard is shown inside a thin wrapper page with a top bar (signed-in
    user name, Sign out). The wrapper also checks every 2 seconds that the
    session is still valid, and reports the buttons the user clicks."""
    import hashlib
    import hmac
    import html as _html
    import ipaddress
    import json
    import os
    import re
    import secrets
    import traceback
    import urllib.parse
    import http.server

    if not (isinstance(base_handler, type)
            and issubclass(base_handler, http.server.BaseHTTPRequestHandler)):
        sys.exit("Access control could not be applied to this dashboard handler.")

    COOKIE = "iosxe_session"
    MAX_FAILS, WINDOW = 5, 300            # 5 wrong logins -> blocked for 5 minutes
    SESSION_IDLE = 12 * 3600              # sign out after 12 hours without use
    ACTIVE_SECS = 90                      # "logged in" = browser checked in within 90 s
    AUDIT_MAX = 5_000_000                 # rotate the audit log at 5 MB
    DUMMY_SALT = bytes(16)
    lock = threading.Lock()
    alock = threading.Lock()
    sessions = {}                         # token -> {user, hash, epoch, seen, login, ip, local}
    failures = {}                         # client ip -> [timestamps of wrong logins]
    ui_rate = {}                          # user -> [window start, clicks logged]
    net_cache = {"key": None, "nets": []}

    def clean(text, limit=200):
        return re.sub(r"[\x00-\x1f\x7f]+", " ", str(text))[:limit]

    def audit(user, ip, text):
        notify = access.get("notify")
        if notify and text in ("login ok", "logout"):            # shown under the bell icon
            try:
                notify("login" if text == "login ok" else "logout",
                       "User signed in" if text == "login ok" else "User signed out",
                       f"{user} from {'this PC' if ip in ('127.0.0.1', '::1') else ip}")
            except Exception:  # noqa: BLE001
                pass
        path = access.get("audit_path")
        if not path:
            return
        line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {clean(user or '-', 40)} | {ip} | {clean(text)}\n"
        try:
            with alock:
                try:
                    if os.path.getsize(path) > AUDIT_MAX:
                        os.replace(path, path + ".1")
                except OSError:
                    pass
                with open(path, "a", encoding="utf-8") as f:
                    f.write(line)
        except Exception:
            traceback.print_exc()

    CSS = ("*{box-sizing:border-box}body{margin:0;min-height:100vh;display:flex;align-items:center;"
           "justify-content:center;font-family:system-ui,Segoe UI,sans-serif;background:#f3f5f8;color:#1b2430}"
           "main{background:#fff;padding:28px 30px;border-radius:12px;box-shadow:0 2px 14px #0002;"
           "width:min(380px,92vw)}h1{font-size:20px;margin:0 0 6px}p{margin:6px 0 14px;color:#51606f}"
           "label{display:block;font-size:13px;margin:10px 0 0}input{width:100%;padding:9px;margin-top:4px;"
           "border:1px solid #b9c3cf;border-radius:6px;font-size:15px}button{width:100%;margin-top:16px;"
           "padding:10px;border:0;border-radius:6px;background:#1f6feb;color:#fff;font-size:15px;cursor:pointer}"
           ".err{background:#fde8e8;color:#9b1c1c;padding:8px 10px;border-radius:6px;font-size:14px}"
           "@media(prefers-color-scheme:dark){body{background:#12171d;color:#e6ebf1}main{background:#1b232c}"
           "p{color:#9fb0c1}input{background:#12171d;color:#e6ebf1;border-color:#3a4654}}"
           ".brand{display:flex;align-items:center;gap:14px;margin:0 0 18px;padding-bottom:16px;"
           "border-bottom:1px solid #d5dce4}.brand img{width:52px;height:52px;object-fit:contain;flex:none}"
           ".brand b{display:block;font-size:24px;font-weight:600;line-height:1.1}"
           ".brand small{display:block;margin-top:3px;font-size:12px;color:#6b7a8c}"
           ".credit{margin:18px 0 0;padding-top:12px;border-top:1px solid #d5dce4;font-size:12px;"
           "text-align:center;color:#6b7a8c}.credit a{color:#1f6feb;text-decoration:none;font-weight:600}")

    def page(title, inner, script=""):
        return ("<!doctype html><html><head><meta charset='utf-8'>"
                "<meta name='viewport' content='width=device-width,initial-scale=1'>"
                f"<title>{title} - {APP_NAME}</title><link rel='icon' href='{APP_ICON_DATA_URI}'>"
                f"<style>{CSS}</style></head><body><main>"
                f"<div class='brand'><img src='{APP_ICON_DATA_URI}' alt=''>"
                f"<div><b>{APP_NAME}</b><small>Version {APP_VERSION}</small></div></div>{inner}"
                f"<p class='credit'>Developed by <b>{_html.escape(APP_AUTHOR)}</b> | "
                f"<a href='{_html.escape(APP_REPO_URL, quote=True)}' target='_blank' "
                f"rel='noopener noreferrer'>GitHub</a></p></main>"
                f"<script>{script}</script></body></html>")

    KICK = "if(top!==self){top.location.replace('/__guard/login');}"

    def login_page(next_url, error=""):
        err = f"<div class='err'>{_html.escape(error)}</div>" if error else ""
        return page("Sign in", (
            "<h1>Sign in</h1><p>Enter your username and password to continue.</p>" + err +
            "<form method='post' action='/__guard/login'>"
            f"<input type='hidden' name='next' value='{_html.escape(next_url, quote=True)}'>"
            "<label>Username<input name='username' autocomplete='username' autofocus required></label>"
            "<label>Password<input name='password' type='password' autocomplete='current-password' required></label>"
            "<button type='submit'>Sign in</button></form>"), KICK)

    def blocked_page():
        return page("Access turned off", (
            "<h1>Access turned off</h1><p>The owner of this tool has turned off public access. "
            "This page reloads by itself when access is allowed again.</p>"),
            "setInterval(async function(){try{var r=await fetch('/__guard/status',{cache:'no-store'});"
            "var j=await r.json();if(j.state!=='blocked'){location.href='/';}}catch(e){}},3000);")

    def setup_page():
        return page("Setup needed", (
            "<h1>Setup needed</h1><p>The administrator account has not been created yet. "
            "Open the NetUpgrade window on the PC that runs this tool and create it first.</p>"))

    SHELL_JS = r"""
var f=document.getElementById('f'), last='', lastT=0;
var ICON='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"/></svg>';
var CSS='#iosxe-user{position:absolute;top:6px;right:16px;z-index:2147483000;font-size:13px;line-height:1;color:inherit}'
 +'#iosxe-user button{all:unset;box-sizing:border-box;display:flex;align-items:center;gap:6px;height:26px;padding:0 10px 0 7px;border-radius:13px;cursor:pointer;border:1px solid rgba(128,128,128,.45);background:rgba(128,128,128,.10);color:inherit;font-size:13px}'
 +'#iosxe-user button:hover{background:rgba(128,128,128,.22)}'
 +'#iosxe-user button:focus-visible{outline:2px solid #1f6feb;outline-offset:1px}'
 +'#iosxe-user .m{position:absolute;top:31px;right:0;min-width:190px;padding:6px;border-radius:8px;border:1px solid rgba(128,128,128,.5);background:Canvas;color:CanvasText;box-shadow:0 4px 14px rgba(0,0,0,.28)}'
 +'#iosxe-user .w{padding:6px 8px 8px;font-size:12px;opacity:.8;border-bottom:1px solid rgba(128,128,128,.35);margin-bottom:4px}'
 +'#iosxe-user .m a{display:block;padding:7px 8px;border-radius:5px;color:inherit;text-decoration:none;font-size:13px}'
 +'#iosxe-user .m a:hover{background:rgba(128,128,128,.22)}'
 +'@media print{#iosxe-user{display:none}}';
function send(t){
  var n=Date.now(); if(t===last && n-lastT<400){return;} last=t; lastT=n;
  try{fetch('/__guard/audit',{method:'POST',headers:{'Content-Type':'text/plain'},body:t,keepalive:true,credentials:'same-origin'});}catch(e){}
}
function label(el){
  var t=(el.innerText||el.value||el.getAttribute('aria-label')||el.title||el.id||el.name||'').replace(/\s+/g,' ').trim();
  if(!t&&el.tagName==='A'){t=el.getAttribute('href')||'';}
  return t.slice(0,60);
}
function addChip(d){
  try{
    if(!d.body||d.getElementById('iosxe-user')){return;}
    if(!d.__iosxeCss){
      d.__iosxeCss=true;
      try{var sh=new d.defaultView.CSSStyleSheet();sh.replaceSync(CSS);d.adoptedStyleSheets=d.adoptedStyleSheets.concat([sh]);}
      catch(e1){var st=d.createElement('style');st.textContent=CSS;(d.head||d.documentElement).appendChild(st);}
    }
    var box=d.createElement('div');box.id='iosxe-user';
    box.innerHTML='<button type="button" aria-haspopup="true" aria-expanded="false">'+ICON+'<span></span></button>'
      +'<div class="m" hidden><div class="w"></div><a href="/__guard/logout" target="_top">Sign out</a></div>';
    box.querySelector('span').textContent=WHO.name;
    box.querySelector('.w').textContent='Signed in as '+WHO.name+' ('+WHO.role+')';
    var btn=box.querySelector('button'), menu=box.querySelector('.m');
    btn.addEventListener('click',function(e){e.stopPropagation();menu.hidden=!menu.hidden;btn.setAttribute('aria-expanded',String(!menu.hidden));});
    d.addEventListener('click',function(){menu.hidden=true;btn.setAttribute('aria-expanded','false');});
    d.addEventListener('keydown',function(e){if(e.key==='Escape'){menu.hidden=true;}});
    d.body.appendChild(box);
  }catch(x){}
}
function hook(){
  try{
    var d=f.contentDocument;
    if(!d){return;}
    // addChip(d);  user icon turned off for now
    if(d.__iosxeHooked){return;}
    d.__iosxeHooked=true;
    d.addEventListener('click',function(e){
      var el=e.target&&e.target.closest?e.target.closest('button,a,input[type=button],input[type=submit],[role=button],summary'):null;
      if(!el||el.closest('#iosxe-user')){return;}
      send('clicked '+el.tagName.toLowerCase()+' "'+label(el)+'"');
    },true);
  }catch(x){}
}
f.addEventListener('load',function(){try{document.title=f.contentDocument.title||document.title;}catch(e){} hook();});
setInterval(async function(){
  try{var r=await fetch('/__guard/status',{cache:'no-store'});var j=await r.json();
      if(j.state!=='ok'){location.replace('/__guard/login');}}catch(e){}
},2000);
"""

    def shell_page(path, user, role):
        """The dashboard in a full-window frame. A small user chip (icon + name + Sign out menu)
        is added inside the dashboard page by SHELL_JS, so there is no extra bar."""
        src = _html.escape(path, quote=True)
        who = json.dumps({"name": user, "role": role}).replace("<", "\\u003c")
        return ("<!doctype html><html><head><meta charset='utf-8'>"
                "<meta name='viewport' content='width=device-width,initial-scale=1'>"
                f"<title>{APP_NAME}</title><link rel='icon' href='{APP_ICON_DATA_URI}'><style>"
                "html,body{margin:0;height:100%;overflow:hidden}iframe{display:block;border:0;width:100%;height:100%}"
                f"</style></head><body><iframe id='f' src='{src}'></iframe><script>var WHO={who};{SHELL_JS}</script></body></html>")

    def cookie_hdr(value=None):
        flags = "Path=/; HttpOnly; SameSite=Lax" + ("; Secure" if access.get("https") else "")
        return f"{COOKIE}=; {flags}; Max-Age=0" if value is None else f"{COOKIE}={value}; {flags}"

    def networks():
        items = tuple(access.get("allowed", []))
        if net_cache["key"] != items:
            nets = []
            for it in items:
                try:
                    nets.append(ipaddress.ip_network(it, strict=False))
                except ValueError:
                    pass
            net_cache["nets"] = nets
            net_cache["key"] = items
        return net_cache["nets"]

    def ip_allowed(ip):
        nets = networks()
        if not nets:
            return True
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            return False
        return any(addr in n for n in nets)

    def check_password(user, pw):
        rec = (access.get("users") or {}).get(user)
        salt = bytes.fromhex(rec["salt"]) if rec else DUMMY_SALT      # same work either way
        got = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, 200_000).hex()
        return bool(rec) and hmac.compare_digest(got, rec["hash"])

    def session_valid(s, rec, now, local):
        # remote sessions die when Block is pressed (epoch changes); this PC's sessions survive it
        return not (rec is None or rec["hash"] != s["hash"] or now - s["seen"] > SESSION_IDLE
                    or (not local and s["epoch"] != access.get("epoch", 0)))

    def session_user(token, ip, local):
        if not token:
            return None
        now = time.time()
        with lock:
            s = sessions.get(token)
            if not s:
                return None
            rec = (access.get("users") or {}).get(s["user"])
            if not session_valid(s, rec, now, local):
                sessions.pop(token, None)          # user removed, new password, Block pressed, or idle
                return None
            s["seen"] = now
            s["ip"] = ip
            return s["user"]

    def new_session(user, ip, local):
        token = secrets.token_hex(32)
        now = time.time()
        with lock:
            if len(sessions) > 500:
                for t in [t for t, s in sessions.items() if now - s["seen"] > SESSION_IDLE]:
                    sessions.pop(t, None)
            sessions[token] = {"user": user, "hash": access["users"][user]["hash"], "epoch": access.get("epoch", 0),
                               "seen": now, "login": now, "ip": ip, "local": local}
        return token

    def active_clients():
        """Users logged in right now (for the launcher window)."""
        now = time.time()
        out, dead = [], []
        with lock:
            users = access.get("users") or {}
            for token, s in sessions.items():
                rec = users.get(s["user"])
                if not session_valid(s, rec, now, s["local"]):
                    dead.append(token)
                    continue
                if now - s["seen"] <= ACTIVE_SECS:
                    out.append({"user": s["user"], "ip": s["ip"], "local": s["local"], "admin": bool(rec.get("admin")), "role": role_of(rec),
                                "since": s["login"], "last": s["seen"]})
            for token in dead:
                sessions.pop(token, None)
        out.sort(key=lambda c: c["since"])
        return out

    access["active_clients"] = active_clients

    def safe_next(url):
        if url.startswith("/") and not url.startswith("//") and "\\" not in url \
                and not url.startswith("/__guard"):
            return url
        return "/"

    class Guarded(base_handler):
        # ---- small helpers -------------------------------------------------
        def _gsend(self, code, body, ctype="text/html; charset=utf-8", extra=()):
            data = body.encode("utf-8") if isinstance(body, str) else body
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("Connection", "close")
            for k, v in extra:
                self.send_header(k, v)
            self.end_headers()
            if self.command != "HEAD":
                self.wfile.write(data)
            self.close_connection = True

        def _redirect(self, location, extra=()):
            self.send_response(303)
            self.send_header("Location", location)
            self.send_header("Content-Length", "0")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Connection", "close")
            for k, v in extra:
                self.send_header(k, v)
            self.end_headers()
            self.close_connection = True

        def _token(self):
            for part in self.headers.get("Cookie", "").split(";"):
                k, _, v = part.strip().partition("=")
                if k == COOKIE:
                    return v
            return None

        def _user(self, ip, local):
            return session_user(self._token(), ip, local)

        def _body(self, limit):
            try:
                length = min(int(self.headers.get("Content-Length", "0") or 0), limit)
            except ValueError:
                length = 0
            return self.rfile.read(length).decode("utf-8", "replace") if length else ""

        def send_header(self, keyword, value):
            # The dashboard is shown inside a same-origin wrapper frame.
            k = keyword.lower()
            if k == "x-frame-options":
                value = "SAMEORIGIN"
            elif k == "content-security-policy":
                value = re.sub(r"frame-ancestors[^;]*", "frame-ancestors 'self'", value)
            super().send_header(keyword, value)

        # ---- the guard -----------------------------------------------------
        def parse_request(self):
            if not super().parse_request():
                return False
            try:
                return self._guard()
            except Exception:
                traceback.print_exc()
                try:
                    self._gsend(500, "Access control error", "text/plain")
                except Exception:
                    pass
                return False          # fail closed

        def _guard(self):
            ip = self.client_address[0]
            local = ip in ("127.0.0.1", "::1") or ip == self.connection.getsockname()[0]
            path = self.path.split("?", 1)[0]
            dest = self.headers.get("Sec-Fetch-Dest", "").lower()
            wants_page = dest in ("document", "iframe") or (
                not dest and "text/html" in self.headers.get("Accept", ""))
            public = bool(access.get("public", False))
            users = access.get("users") or {}

            if path.startswith("/__guard/"):
                self._guard_route(ip, local, path, public, users)
                return False

            if not users:                                  # no accounts yet: nobody gets in
                if wants_page:
                    self._gsend(503, setup_page())
                else:
                    self._gsend(503, "Setup needed: create the admin account first.", "text/plain")
                return False
            if not local:                                  # this PC is not subject to Block / allowed IPs
                if not public:
                    if wants_page:
                        self._gsend(403, blocked_page().replace("</script>", KICK + "</script>"))
                    else:
                        self._gsend(403, "Public access is turned off.", "text/plain")
                    return False
                if not ip_allowed(ip):
                    self._gsend(403, "Your address is not allowed to use this tool.", "text/plain")
                    return False
            user = self._user(ip, local)
            if not user:
                if dest == "iframe":
                    self._gsend(401, f"<script>{KICK}</script>")
                elif wants_page:
                    self._redirect("/__guard/login?next=" + urllib.parse.quote(self.path))
                else:
                    self._gsend(401, "Login required.", "text/plain")
                return False
            self.guard_user = user                         # lets the dashboard record who did what
            role = role_of(users[user])
            self.guard_role = role
            allowed, why = role_allows(role, self.command, self.path)
            if not allowed:                                # enforced here, whatever the page shows
                audit(user, ip, f"DENIED ({role}) {self.command} {path}")
                if path.startswith("/api/"):
                    self._gsend(403, json.dumps({"error": why}), "application/json")
                else:
                    self._gsend(403, why, "text/plain")
                return False
            if self.command not in ("GET", "HEAD"):        # record every action that changes something
                audit(user, ip, f"{self.command} {path}")
            if (dest == "document" and self.command == "GET"
                    and (path in ("/", "/index.html") or path.endswith(".html"))):
                self._gsend(200, shell_page(self.path, user, role))
                return False
            return True

        def _guard_route(self, ip, local, path, public, users):
            remote_blocked = (not local) and (not public or not ip_allowed(ip))
            if path == "/__guard/status":
                if not users:
                    state = "setup"
                elif remote_blocked:
                    state = "blocked"
                elif not self._user(ip, local):
                    state = "login"
                else:
                    state = "ok"
                self._gsend(200, json.dumps({"state": state}), "application/json")
                return
            if path == "/__guard/logout":
                token = self._token()
                with lock:
                    s = sessions.pop(token, None)
                if s:
                    audit(s["user"], ip, "logout")
                self._redirect("/__guard/login", [("Set-Cookie", cookie_hdr(None))])
                return
            if path == "/__guard/audit":                   # button clicks reported by the wrapper page
                user = self._user(ip, local) if self.command == "POST" else None
                if not user:
                    self._gsend(401, "Login required.", "text/plain")
                    return
                text = self._body(300)
                now = time.time()
                with lock:
                    win = ui_rate.get(user)
                    if not win or now - win[0] > 60:
                        win = ui_rate[user] = [now, 0]
                    win[1] += 1
                    allowed_now = win[1] <= 120
                if allowed_now and text:
                    audit(user, ip, "ui " + text)
                self._gsend(204, "", "text/plain")
                return
            if path != "/__guard/login":
                self._gsend(404, "Not found.", "text/plain")
                return
            if not users:
                self._gsend(503, setup_page())
                return
            if remote_blocked:
                self._gsend(403, blocked_page())
                return
            query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
            if self.command != "POST":
                nxt = safe_next(query.get("next", ["/"])[0])
                if self._user(ip, local):
                    self._redirect(nxt)
                else:
                    self._gsend(200, login_page(nxt))
                return
            # ---- POST: check the login form
            form = urllib.parse.parse_qs(self._body(4096))
            user = form.get("username", [""])[0].strip().lower()
            pw = form.get("password", [""])[0]
            nxt = safe_next(form.get("next", ["/"])[0])
            now = time.time()
            with lock:
                recent = [t for t in failures.get(ip, []) if now - t < WINDOW]
                failures[ip] = recent
                locked = len(recent) >= MAX_FAILS
            if locked:
                self._gsend(429, login_page(nxt, "Too many wrong logins. Try again in a few minutes."))
                return
            if check_password(user, pw):
                token = new_session(user, ip, local)
                with lock:
                    failures.pop(ip, None)
                audit(user, ip, "login ok")
                self._redirect(nxt, [("Set-Cookie", cookie_hdr(token))])
            else:
                with lock:
                    failures.setdefault(ip, []).append(now)
                audit(user, ip, "login FAILED")
                self._gsend(401, login_page(nxt, "Wrong username or password."))

    return Guarded


apply_access_control.version = 5      # the launcher checks this to detect an out-of-date script


def gui_main(embedded=False, argv=None, access=None, tls=None):
    if getattr(sys, "frozen", False):
        sys.excepthook = _frozen_excepthook
    p = argparse.ArgumentParser(description="NetUpgrade - monitor, back up and upgrade "
                                            "Cisco C9200 switches.")
    p.add_argument("--version", action="version", version=f"{APP_NAME} v{APP_VERSION}")
    p.add_argument("excel", nargs="?", default=None,
                   help="Optional: a .xlsx file to seed a site profile's device list with on "
                        "first activation. Only used together with --site; otherwise devices "
                        "are uploaded or added by hand from the dashboard.")
    p.add_argument("--site", default=None,
                   help="Activate this site profile automatically on startup (creating it if "
                        "new). Without it, select or create a site from the dashboard.")
    p.add_argument("--sheet", help="Sheet name (default: first sheet)")
    p.add_argument("--column", help="Header of the IP/hostname column (default: auto-detect)")
    p.add_argument("--mode", choices=["icmp", "tcp"], default="icmp",
                   help="Availability check: icmp = ping (default); tcp = TCP connect to --port")
    p.add_argument("--port", type=int, default=22, help="TCP port for --mode tcp (default: 22)")
    p.add_argument("--interval", type=float, default=5.0, help="Seconds between check cycles (default: 5)")
    p.add_argument("--timeout", type=float, default=2.0, help="Per-check timeout in seconds (default: 2)")
    p.add_argument("--duration", type=float, default=0,
                   help=f"Stop monitoring after this many minutes (min {MIN_DURATION_MIN}); "
                        "0 = keep monitoring until stopped (default). Monitoring can be restarted in the GUI.")
    p.add_argument("--fail-threshold", type=int, default=2,
                   help="Consecutive failed checks before a switch is marked DOWN (default: 2)")
    p.add_argument("--host", default="127.0.0.1", help="Dashboard bind address (default: 127.0.0.1)")
    p.add_argument("--allow-remote", action="store_true",
                   help="Request remote binding; requires the desktop launcher's login and HTTPS controls")
    p.add_argument("--web-port", type=int, default=8080, help="Dashboard port (default: 8080)")
    p.add_argument("--open", action="store_true", help="Open the dashboard in your browser")
    args = p.parse_args(argv)

    if args.duration != 0 and args.duration < MIN_DURATION_MIN:
        p.error(f"--duration must be at least {MIN_DURATION_MIN} minutes (or 0 for no limit)")
    if args.interval <= 0 or args.timeout <= 0 or args.fail_threshold < 1:
        p.error("--interval and --timeout must be > 0 and --fail-threshold must be >= 1")
    loopback = args.host in ("127.0.0.1", "localhost", "::1")
    exposure_error = dashboard_exposure_error(args.host, args.allow_remote, access, tls)
    if exposure_error:
        p.error(exposure_error)
    if args.mode == "icmp" and shutil.which("ping") is None:
        sys.exit("The 'ping' command was not found. Install it or use: --mode tcp --port 22")

    print(f"{APP_NAME} v{APP_VERSION}  |  Developed by {APP_AUTHOR}  |  {APP_REPO_URL}")
    events = EventLog()
    monitor = Monitor([], events, args.mode, args.port, args.interval, args.timeout,
                      args.duration, args.fail_threshold)
    manager = Manager(monitor, events)

    if args.site:
        try:
            manager.switch_site(args.site)
        except ApiError as exc:
            sys.exit(f"Could not activate site '{args.site}': {exc}")
        print(f"Active site: {manager.site_name} ({len(manager.order)} device(s) loaded)")
        if args.excel:
            try:
                seed_hosts = load_hosts(args.excel, args.sheet, args.column)
                added, skipped = manager.add_devices_bulk(seed_hosts, Path(args.excel).name)
                print(f"Seeded from {args.excel}: {added} device(s) added"
                     + (f", {skipped} already present" if skipped else ""))
            except SystemExit:
                raise
    else:
        if args.excel:
            print("NOTE: a switch list file was given but no --site was specified, so it is "
                 "being ignored. Pass --site NAME to activate a site and seed it from this "
                 "file, or choose/create a site from the dashboard and upload it there.")
        print("No site profile selected yet - choose or create one from the dashboard.")

    EMAILER.cfg_provider = manager.email_live_config
    legacy = (access or {}).get("email_legacy")          # e-mail settings the launcher's old Notifications tab saved
    if isinstance(legacy, dict) and legacy.get("host") and not email_settings_path().exists():
        try:
            merged = dict(EMAIL_DEFAULTS)
            merged.update({k: legacy[k] for k in EMAIL_DEFAULTS if k in legacy})
            save_email_settings(merged)
            manager.email_cfg = load_email_settings()
            print("E-mail notification settings moved from the launcher to Settings > Notifications.", flush=True)
        except Exception as exc:  # noqa: BLE001
            print(f"Could not move the old e-mail settings: {type(exc).__name__}", flush=True)
    allowed = {"localhost", "127.0.0.1", "::1"}
    handler = make_handler(manager, allowed, any_host=not loopback)
    if access is not None:
        access["notify"] = lambda kind, title, text: manager._notify("-", kind, title, text)
        access["https"] = bool(tls)
        handler = apply_access_control(handler, access)
    try:
        if tls:
            server = make_tls_server((args.host, args.web_port), handler, tls[0], tls[1])
        else:
            server = ThreadingHTTPServer((args.host, args.web_port), handler)
    except OSError as exc:
        sys.exit(f"Could not start the dashboard on {args.host}:{args.web_port}: {exc}\n"
                 "Try a different port with --web-port.")
    threading.Thread(target=server.serve_forever, daemon=True).start()
    # Monitoring no longer starts on its own - the person clicks "Start Monitoring" in the GUI.

    shown = "127.0.0.1" if args.host in ("0.0.0.0", "") else args.host
    url = f"{'https' if tls else 'http'}://{shown}:{args.web_port}"

    # When started from the launcher window, hand control back instead of
    # blocking here; the launcher opens the browser and handles Exit.
    if embedded:
        def shutdown():
            monitor.stop()
            server.shutdown()
            server.server_close()
            manager.close()
        return url, shutdown

    print(f"Dashboard: {url}")
    if not loopback:
        print("WARNING: the dashboard is reachable from the network and has no login.")
    print("Press Ctrl+C to stop.\n")
    if args.open or getattr(sys, "frozen", False):
        webbrowser.open(url)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down ...")

    running = manager.active_upgrades()
    if running:
        print(f"NOTE: {len(running)} upgrade(s) are still running in the background and will "
              f"continue (logs are in the site's Runs folder): {', '.join(running)}")
    monitor.stop()
    server.shutdown()
    manager.close()


if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == '--job':
        sys.exit(job_main(sys.argv[2]))            # per-switch worker (launched by the GUI)
    elif '--cli' in sys.argv:
        sys.argv.remove('--cli')
        interactive_main()                         # original no-GUI batch upgrade flow
    else:
        gui_main()                                 # web GUI (default)
