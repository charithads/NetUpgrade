# NetUpgrade

NetUpgrade is a Windows desktop application for monitoring Cisco Catalyst switches, managing device inventories, collecting configuration backups, and running or scheduling IOS-XE upgrades. The desktop launcher runs in the Windows notification area and opens a local web dashboard.

The current application version is **5.12.1**. The upgrade workflow is designed for Cisco Catalyst 9200/9300 switches and uses IOS-XE INSTALL mode.

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
- Set the dashboard-wide upgrade limit (1–10 switches at a time). Extra switches wait in a queue, and an Administrator can choose whether a failure stops queued upgrades.
- Start upgrades now or schedule them for later; view each switch's live progress and console output in the dashboard.
- Adjust command timeouts, save named timeout profiles, and run command-time checks on a switch before a rollout. Review device events, backups, and per-device command history.
- Get optional email alerts for upgrade starts, completion, failures, attention items, post-check results, missed schedules, and sign-ins. Configure SMTP and test delivery in **Settings > Notifications**.
- Use Administrator, Operator, and Viewer dashboard roles, access controls, and optional HTTPS from the desktop launcher.
- Run the interactive command-line upgrade flow with `--cli`.
- Check for releases from **About**. The installed Windows app verifies the installer, asks before installing, then closes and restarts. Updates are blocked while a switch upgrade is running.

## Requirements

- Windows 10 or later for the desktop launcher.
- Python 3.11 or later to run from source or build the executable.
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
add `--continue-on-failure` to continue. In the dashboard, an Administrator sets a global rollout limit from 1 to 10 and chooses whether a failure stops queued upgrades. Queued devices start as slots become available.

### Direct manager options

```powershell
py ios_xe_upgrade_manager.py --help
```

For example, use TCP checks when ICMP is blocked, set a custom dashboard port, or open the dashboard automatically:

```powershell
py ios_xe_upgrade_manager.py --mode tcp --port 22 --web-port 9000 --open
```

Monitoring starts when requested from the dashboard. The direct manager listens on `127.0.0.1` by default and cannot be exposed remotely. Remote access must be enabled through the desktop launcher, with configured users and HTTPS.

## Install the Windows app

Download `NetUpgrade-Setup-v<version>.exe` from the [GitHub Releases page](https://github.com/charithads/NetUpgrade/releases) and run it. The installer installs NetUpgrade for the current Windows user and can add a desktop shortcut. The first installation must be done with the setup file; after that, use **About > Check for updates** in the app. NetUpgrade checks GitHub, asks before an update, verifies the installer checksum published by GitHub, installs it, and starts the app again. It will not install an update while a switch upgrade is running.

The installer is not code-signed yet, so Windows may show a SmartScreen warning. Releases are built by the Windows GitHub Actions workflow after tests, dependency audit, and executable packaging succeed. Publishing a GitHub release automatically attaches its version-matched Windows installer.

## Build from source

Install PyInstaller and the application dependencies, then build from PowerShell:

```powershell
py -m pip install -r requirements.txt pyinstaller==6.22.3
py -m PyInstaller --noconsole --onefile --name NetUpgrade --icon app.ico --add-data "app.ico;." --hidden-import ios_xe_upgrade_manager --collect-all netmiko --collect-all cryptography launcher.py
py -m PyInstaller --noconsole --onefile --name NetUpgradeUpdater --hidden-import psutil updater_helper.py
```

The Windows installer is compiled from `NetUpgrade.iss` with Inno Setup 6 after both executables exist in `dist`. CI performs the complete build and release upload.

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
