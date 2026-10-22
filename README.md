# NetUpgrade

NetUpgrade is a Windows desktop application for monitoring Cisco Catalyst switches, managing device inventories, collecting configuration backups, and running or scheduling IOS-XE upgrades. The desktop launcher runs in the Windows notification area and opens a local web dashboard.

The current application version is **5.6.1**. The upgrade workflow is designed for Cisco Catalyst 9200/9300 switches and uses IOS-XE INSTALL mode.

## Screenshots

### Dashboard and live upgrade console

The dashboard shows switch availability, backup and upgrade state, events, and live worker output.

![Dashboard with live upgrade console](Dashboardview.png)

## Features

- Monitor switch availability using ICMP or a TCP connection.
- Organize devices into site profiles and import inventories from Excel workbooks.
- Test and manage device credentials.
- Back up running configuration, version, boot, inventory, and optional additional commands.
- Run readiness checks before upgrading, distribute IOS-XE images, activate upgrades, and verify the version after reload.
- Check each IOS-XE image against Cisco's MD5 checksum during upload; images with missing or mismatched checksums are rejected.
- Configure the maximum concurrent reloads per site (1–10); additional switches wait in a queue, with an administrator-selectable stop-on-failure policy.
- Start upgrades immediately or schedule them for a later date and time; view each worker's live progress in the dashboard.
- Review device events, backups, and per-device command history. Command logs record commands and results metadata, not switch command output.
- Use role-based dashboard accounts, access controls, and optional HTTPS from the desktop launcher.
- Run the interactive command-line upgrade flow with `--cli`.

## Requirements

- Windows 10 or later for the desktop launcher.
- Python 3.10 or later to run from source or build the executable.
- Network reachability to the switches and SSH credentials with the privileges required for backups and upgrades.
- An Excel workbook (`.xlsx`) when importing a switch inventory.

## Run from source

Install the application dependencies in PowerShell:

```powershell
py -m pip install -r requirements.txt
```

Start the desktop launcher:

```powershell
py launcher.py
```

The launcher runs in the notification area. On first start, create the administrator account, then open the dashboard from the launcher. The dashboard uses HTTPS by default; its configured address and port are shown in the launcher.

To start the manager directly without the desktop launcher:

```powershell
py ios_xe_upgrade_manager.py
```

Open `http://127.0.0.1:8080` in a browser when using the direct manager defaults. To seed a site profile from a workbook:

```powershell
py ios_xe_upgrade_manager.py switches.xlsx --site SiteName
```

The workbook's first sheet must have a header row with an `ip` column. Optional columns include `hostname`, `platform`, and `type`. A workbook supplied without `--site` is not imported automatically; inventories can also be uploaded or entered through the dashboard.

### Interactive CLI upgrade

Run the interactive whole-inventory upgrade flow:

```powershell
py launcher.py --cli
```

The CLI prompts for the inventory, credentials, and image. It upgrades one switch at a time by default;
pass `--parallel N` to process up to N switches in waves. By default, a failed wave stops later waves;
add `--continue-on-failure` to continue. Dashboard rollout concurrency is configured per site by an
Administrator, from 1 to 10. Queued devices start as slots become available, and the rollout can stop
starting queued devices after a failure.

### Direct manager options

```powershell
py ios_xe_upgrade_manager.py --help
```

For example, use TCP checks when ICMP is blocked, set a custom dashboard port, or open the dashboard automatically:

```powershell
py ios_xe_upgrade_manager.py --mode tcp --port 22 --web-port 9000 --open
```

Monitoring starts when requested from the dashboard. The direct manager listens on `127.0.0.1` by default and cannot be exposed remotely. Remote access must be enabled through the desktop launcher, with configured users and HTTPS.

## Build the Windows executable

Install PyInstaller and the application dependencies, then build the launcher:

```powershell
py -m pip install -r requirements.txt pyinstaller==6.22.3
py -m PyInstaller --noconsole --onefile --name NetUpgrade --icon app.ico --add-data "app.ico;." --hidden-import ios_xe_upgrade_manager --collect-all netmiko --collect-all cryptography launcher.py
```

The executable is created at `dist\NetUpgrade.exe`. Keep `ios_xe_upgrade_manager.py` available to the build process; the launcher packages it as a hidden import and invokes its worker mode for device upgrades.

## Security and local data

The desktop launcher supports administrator, operator, and viewer roles, an administrator-managed user list, allowed-client controls, and optional HTTPS. Protect administrator credentials and restrict access to trusted networks. The manager can issue privileged commands that change switch software.

Application settings, audit logs, application logs, and TLS files are stored under:

```text
%LOCALAPPDATA%\NetUpgrade
```

The direct manager defaults to loopback. Enabling remote access exposes switch-management capabilities to reachable clients, so configure authentication and network restrictions appropriately when using the launcher.

This application can run privileged commands that change network software. Use it in a controlled pilot until it has been reviewed and validated for your environment. See [SECURITY.md](SECURITY.md) for a deployment baseline and [DEPLOYMENT_READINESS.md](DEPLOYMENT_READINESS.md) for the engineering and operational gates that remain before broad deployment.

To upload an image, choose its `.bin` file, paste the 32-character MD5 checksum shown on Cisco's software download page, then select **Upload and verify**. NetUpgrade checks the image before saving it and rejects missing or mismatched checksums. It saves the verified MD5 checksum beside the image and checks it again before an upgrade. Copy the checksum from Cisco's download page; do not calculate the expected value from the local image.

## Verification

Install the pinned development tools and run the automated checks from PowerShell:

```powershell
py -m pip install -r requirements-dev.txt
py -m unittest discover -s tests -p "test_*.py" -v
py tests/smoke_test.py
pip-audit -r requirements.txt
```

CI runs the unit and smoke suites on supported Python versions, performs static checks and dependency auditing, and builds the Windows executable. The tests do not replace hardware-in-the-loop validation on every supported switch and IOS-XE release.

## Project files

- `launcher.py` — Windows desktop launcher, tray controls, dashboard access management, and packaged-app entry point.
- `ios_xe_upgrade_manager.py` — dashboard server, device monitoring, backups, upgrade workers, scheduling, and CLI implementation.
- `Dashboardview.png` — application dashboard screenshot shown above.
- `app.ico` — application icon used by the launcher and executable build.
- `.github/workflows/ci.yml` — continuous integration workflow.
