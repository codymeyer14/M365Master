#!/usr/bin/env python3
"""Read-only Microsoft 365 reporting dashboard generator."""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
from pathlib import Path
from typing import Any

import msal
import requests
import yaml

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = [
    "Policy.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementApps.Read.All",
    "Group.Read.All",
]


class GraphClient:
    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
            }
        )

    def get(self, url: str) -> dict[str, Any]:
        response = self.session.get(url, timeout=60)
        if response.status_code >= 400:
            raise RuntimeError(f"Graph GET failed {response.status_code}: {response.text}")
        return response.json()

    def get_all(self, relative_url: str) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        next_url = f"{GRAPH_ROOT}{relative_url}"
        while next_url:
            payload = self.get(next_url)
            items.extend(payload.get("value", []))
            next_url = payload.get("@odata.nextLink")
        return items


def load_config(config_path: Path) -> dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}

    missing = [k for k in ("tenant_id", "client_id") if not cfg.get(k)]
    if missing:
        raise ValueError(f"Missing required config keys: {', '.join(missing)}")

    cfg.setdefault("scopes", DEFAULT_SCOPES)
    return cfg


def acquire_token(tenant_id: str, client_id: str, scopes: list[str]) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id=client_id, authority=authority)

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not create device flow: {flow}")

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description', result)}")

    return result["access_token"]


def group_name_lookup(client: GraphClient, group_ids: set[str]) -> dict[str, str]:
    names: dict[str, str] = {}
    for group_id in sorted(group_ids):
        try:
            g = client.get(f"{GRAPH_ROOT}/groups/{group_id}?$select=id,displayName")
            names[group_id] = g.get("displayName", group_id)
        except RuntimeError:
            names[group_id] = f"Unknown group ({group_id})"
    return names


def read_conditional_access(client: GraphClient) -> list[dict[str, Any]]:
    policies = client.get_all("/identity/conditionalAccess/policies?$top=200")
    output: list[dict[str, Any]] = []
    for p in policies:
        output.append(
            {
                "id": p.get("id"),
                "displayName": p.get("displayName"),
                "state": p.get("state"),
                "conditions": p.get("conditions", {}),
                "grantControls": p.get("grantControls", {}),
                "sessionControls": p.get("sessionControls", {}),
            }
        )
    return output


def extract_group_ids_from_assignments(assignments: list[dict[str, Any]]) -> set[str]:
    group_ids: set[str] = set()
    for assignment in assignments:
        target = assignment.get("target", {})
        group_id = target.get("groupId")
        if group_id:
            group_ids.add(group_id)
    return group_ids


def read_device_configurations(client: GraphClient) -> list[dict[str, Any]]:
    profiles = client.get_all(
        "/deviceManagement/deviceConfigurations?$top=200"
        "&$select=id,displayName,description,lastModifiedDateTime,@odata.type"
    )
    enriched: list[dict[str, Any]] = []

    for profile in profiles:
        pid = profile["id"]
        assignments = client.get_all(f"/deviceManagement/deviceConfigurations/{pid}/assignments?$top=200")
        enriched.append(
            {
                **profile,
                "assignments": assignments,
            }
        )
    return enriched


def read_mobile_apps(client: GraphClient) -> list[dict[str, Any]]:
    apps = client.get_all(
        "/deviceAppManagement/mobileApps?$top=200"
        "&$select=id,displayName,publisher,isAssigned,@odata.type"
    )
    enriched: list[dict[str, Any]] = []

    for app in apps:
        app_id = app["id"]
        assignments = client.get_all(f"/deviceAppManagement/mobileApps/{app_id}/assignments?$top=200")
        enriched.append(
            {
                **app,
                "assignments": assignments,
            }
        )
    return enriched


def read_managed_devices(client: GraphClient) -> list[dict[str, Any]]:
    return client.get_all(
        "/deviceManagement/managedDevices?$top=200"
        "&$select=id,deviceName,operatingSystem,complianceState,managementState,"
        "ownerType,userDisplayName,userPrincipalName,lastSyncDateTime,enrolledDateTime"
    )


def friendly_assignment_targets(assignments: list[dict[str, Any]], groups: dict[str, str]) -> str:
    if not assignments:
        return "None"

    labels: list[str] = []
    for assignment in assignments:
        target = assignment.get("target", {})
        odata_type = target.get("@odata.type", "")
        group_id = target.get("groupId")

        if "allLicensedUsers" in odata_type:
            labels.append("All Users")
        elif "allDevices" in odata_type:
            labels.append("All Devices")
        elif group_id:
            labels.append(groups.get(group_id, group_id))
        else:
            labels.append("Unknown Target")

    return ", ".join(sorted(set(labels)))


def build_html(report: dict[str, Any]) -> str:
    generated = html.escape(report["generatedUtc"])

    ca_rows = "\n".join(
        (
            f"<tr><td>{html.escape(p.get('displayName', ''))}</td>"
            f"<td>{html.escape(p.get('state', ''))}</td>"
            f"<td><pre>{html.escape(json.dumps(p.get('conditions', {}), indent=2))}</pre></td>"
            f"<td><pre>{html.escape(json.dumps(p.get('grantControls', {}), indent=2))}</pre></td>"
            f"<td><pre>{html.escape(json.dumps(p.get('sessionControls', {}), indent=2))}</pre></td></tr>"
        )
        for p in report["conditionalAccessPolicies"]
    )

    device_rows = "\n".join(
        (
            f"<tr><td>{html.escape(d.get('deviceName', ''))}</td>"
            f"<td>{html.escape(d.get('operatingSystem', ''))}</td>"
            f"<td>{html.escape(d.get('complianceState', ''))}</td>"
            f"<td>{html.escape(d.get('managementState', ''))}</td>"
            f"<td>{html.escape(d.get('userPrincipalName', ''))}</td>"
            f"<td>{html.escape(str(d.get('lastSyncDateTime', '')))}</td></tr>"
        )
        for d in report["managedDevices"]
    )

    config_rows = "\n".join(
        (
            f"<tr><td>{html.escape(c.get('displayName', ''))}</td>"
            f"<td>{html.escape(c.get('@odata.type', ''))}</td>"
            f"<td>{html.escape(c.get('assignmentTargets', 'None'))}</td>"
            f"<td>{html.escape(str(c.get('lastModifiedDateTime', '')))}</td></tr>"
        )
        for c in report["deviceConfigurations"]
    )

    app_rows = "\n".join(
        (
            f"<tr><td>{html.escape(a.get('displayName', ''))}</td>"
            f"<td>{html.escape(a.get('publisher', ''))}</td>"
            f"<td>{html.escape(a.get('@odata.type', ''))}</td>"
            f"<td>{html.escape(a.get('assignmentTargets', 'None'))}</td></tr>"
        )
        for a in report["mobileApps"]
    )

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\">
  <title>M365 Reporting Dashboard</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 16px; }}
    h1, h2 {{ margin-bottom: 0.2rem; }}
    p.meta {{ color: #666; margin-top: 0; }}
    table {{ border-collapse: collapse; width: 100%; margin-bottom: 24px; table-layout: fixed; }}
    th, td {{ border: 1px solid #ccc; padding: 8px; vertical-align: top; word-wrap: break-word; }}
    th {{ background: #f5f5f5; }}
    pre {{ white-space: pre-wrap; margin: 0; }}
  </style>
</head>
<body>
  <h1>M365 Read-Only Dashboard</h1>
  <p class=\"meta\">Generated (UTC): {generated}</p>

  <h2>Conditional Access Policies ({len(report['conditionalAccessPolicies'])})</h2>
  <table>
    <tr><th>Name</th><th>State</th><th>Conditions</th><th>Grant Controls</th><th>Session Controls</th></tr>
    {ca_rows}
  </table>

  <h2>Intune Managed Devices ({len(report['managedDevices'])})</h2>
  <table>
    <tr><th>Device</th><th>OS</th><th>Compliance</th><th>Management</th><th>User UPN</th><th>Last Sync</th></tr>
    {device_rows}
  </table>

  <h2>Intune Configuration Profiles ({len(report['deviceConfigurations'])})</h2>
  <table>
    <tr><th>Name</th><th>Type</th><th>Assignments</th><th>Last Modified</th></tr>
    {config_rows}
  </table>

  <h2>Intune Applications ({len(report['mobileApps'])})</h2>
  <table>
    <tr><th>Name</th><th>Publisher</th><th>Type</th><th>Assignments</th></tr>
    {app_rows}
  </table>
</body>
</html>"""


def build_report(client: GraphClient) -> dict[str, Any]:
    ca_policies = read_conditional_access(client)
    devices = read_managed_devices(client)
    configs = read_device_configurations(client)
    apps = read_mobile_apps(client)

    group_ids: set[str] = set()
    for c in configs:
        group_ids.update(extract_group_ids_from_assignments(c.get("assignments", [])))
    for a in apps:
        group_ids.update(extract_group_ids_from_assignments(a.get("assignments", [])))

    group_map = group_name_lookup(client, group_ids)

    for c in configs:
        c["assignmentTargets"] = friendly_assignment_targets(c.get("assignments", []), group_map)
    for a in apps:
        a["assignmentTargets"] = friendly_assignment_targets(a.get("assignments", []), group_map)

    return {
        "generatedUtc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "conditionalAccessPolicies": ca_policies,
        "managedDevices": devices,
        "deviceConfigurations": configs,
        "mobileApps": apps,
        "groupLookup": group_map,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a read-only M365/Intune dashboard report")
    parser.add_argument("--config", required=True, help="Path to YAML config")
    parser.add_argument("--output-dir", default="output", help="Directory for report artifacts")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    cfg = load_config(Path(args.config))

    token = acquire_token(cfg["tenant_id"], cfg["client_id"], cfg.get("scopes", DEFAULT_SCOPES))
    client = GraphClient(token)

    report = build_report(client)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    json_path = out_dir / "report.json"
    html_path = out_dir / "dashboard.html"

    json_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    html_path.write_text(build_html(report), encoding="utf-8")

    print(f"Wrote: {json_path}")
    print(f"Wrote: {html_path}")


if __name__ == "__main__":
    main()
