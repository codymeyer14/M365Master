# M365 Read-Only Reporting Dashboard

This repository contains a Python application that signs in with your Microsoft 365 admin account and builds a **read-only dashboard** of:

- Conditional Access policies and their key settings.
- Intune managed devices.
- Intune configuration profiles and their group assignments.
- Intune applications and their group assignments.

The tool uses **Microsoft Graph GET requests only** and does not create, update, or delete anything.

## Features

- Device-code sign-in with your global admin credentials (MFA-friendly).
- Automatic Graph pagination handling.
- Group assignment resolution (group IDs -> display names).
- JSON exports for raw reporting data.
- Single-file HTML dashboard for easy review/share.

## Prerequisites

1. Python 3.10+
2. An Entra app registration for this reporting tool.
3. Microsoft Graph **Application/Delegated read permissions** (delegated is used by this tool):
   - `Policy.Read.All`
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementApps.Read.All`
   - `Group.Read.All`
4. Admin consent granted where required.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Copy configuration template:

```bash
cp config.example.yaml config.yaml
```

Edit `config.yaml` with your tenant/client values from your app registration.

## Usage

```bash
python m365_reporter.py --config config.yaml --output-dir output
```

What happens:
1. You are prompted to complete device-code sign-in in a browser.
2. The tool calls Microsoft Graph read endpoints.
3. Files are generated in the output folder:
   - `report.json`
   - `dashboard.html`

Open `dashboard.html` in your browser for the dashboard report.

## Security notes

- The tool does **not** store your password.
- Access tokens remain in memory for the runtime only.
- The script only uses `GET` requests and is intentionally read-only.

## Disclaimer

Microsoft Graph permissions determine exactly what can be read. If permissions are missing, sections may be incomplete.
