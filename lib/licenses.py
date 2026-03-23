#!/usr/bin/env python3
# List Microsoft 365 licenses per user with start/end dates.
# Auth: client credentials (AZURE_CLIENT_SECRET) when set, else MSAL device code flow.
# Token cache: ~/.office365-mcp-tokens.json (device code only)

import asyncio
import os
import time
from pathlib import Path

import msal
import httpx
from azure.core.credentials import AccessToken

TOKEN_CACHE_PATH = Path.home() / ".office365-mcp-tokens.json"
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"] if CLIENT_SECRET else [
    "https://graph.microsoft.com/User.Read.All",
    "https://graph.microsoft.com/Directory.Read.All",
]

# Teams service plan IDs
# TEAMS1 (standard Teams), Teams Free, Teams Exploratory
TEAMS_SERVICE_PLAN_IDS = {
    "57ff2da0-773e-42df-b2af-ffb7a2317929",  # TEAMS1 (included in most M365/O365 plans)
    "bd8f45e8-5e4e-4c4e-9e0d-e7b27a3c0f4c",  # Microsoft Teams (Free)
    "a9e39479-e56e-4b3c-8b60-c8d93ab80b3d",  # Teams Exploratory
}


def _get_token() -> str:
    if CLIENT_SECRET:
        app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    else:
        cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_PATH.exists():
            cache.deserialize(TOKEN_CACHE_PATH.read_text())
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result or "access_token" not in result:
            flow = app.initiate_device_flow(scopes=SCOPES)
            print(f"Visit: {flow['verification_uri']}")
            print(f"Code:  {flow['user_code']}")
            result = app.acquire_token_by_device_flow(flow)
        if cache.has_state_changed:
            TOKEN_CACHE_PATH.write_text(cache.serialize())

    if not result or "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result.get('error_description')}")

    return result["access_token"]


class _TokenCredential:
    def get_token(self, *scopes, **kwargs):
        return AccessToken(_get_token(), int(time.time()) + 3600)


def _sku_has_teams(sku: dict) -> bool:
    """Return True if the SKU includes any Teams service plan."""
    for plan in sku.get("servicePlans", []):
        if plan.get("servicePlanId") in TEAMS_SERVICE_PLAN_IDS:
            return True
    return False


def _format_date(dt_str: str | None) -> str:
    if not dt_str:
        return "—"
    # ISO 8601 → YYYY-MM-DD
    return dt_str[:10]


async def list_teams_licenses():
    token = _get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    async with httpx.AsyncClient(timeout=30) as http:
        # 1. Get tenant subscriptions (beta) for start/end dates keyed by skuId
        sub_resp = await http.get(
            "https://graph.microsoft.com/beta/directory/subscriptions",
            headers=headers,
        )
        sub_resp.raise_for_status()
        subscriptions: dict[str, dict] = {}
        for sub in sub_resp.json().get("value", []):
            sku_id = sub.get("skuId", "")
            subscriptions[sku_id] = sub

        # 2. Get tenant-level subscribed SKUs (v1.0) – filter for Teams
        sku_resp = await http.get(
            "https://graph.microsoft.com/v1.0/subscribedSkus",
            headers=headers,
        )
        sku_resp.raise_for_status()
        teams_skus: dict[str, dict] = {}
        for sku in sku_resp.json().get("value", []):
            if _sku_has_teams(sku):
                teams_skus[sku["skuId"]] = sku

        if not teams_skus:
            print("No Teams-related SKUs found in this tenant.")
            return

        # 3. Get all users with their assigned licenses
        users_url = (
            "https://graph.microsoft.com/v1.0/users"
            "?$select=id,displayName,userPrincipalName,assignedLicenses"
            "&$top=999"
        )
        rows: list[tuple] = []
        while users_url:
            u_resp = await http.get(users_url, headers=headers)
            u_resp.raise_for_status()
            u_data = u_resp.json()
            for user in u_data.get("value", []):
                display_name = user.get("displayName") or ""
                upn = user.get("userPrincipalName") or ""
                for assignment in user.get("assignedLicenses", []):
                    sku_id = assignment.get("skuId", "")
                    if sku_id not in teams_skus:
                        continue
                    sku = teams_skus[sku_id]
                    sku_name = sku.get("skuPartNumber", sku_id)
                    capability = sku.get("capabilityStatus", "")
                    is_trial = "Trial" in capability or "TRIAL" in sku_name or "EXPLORATORY" in sku_name
                    disabled_plans = set(assignment.get("disabledPlans", []))
                    teams_disabled = bool(disabled_plans & TEAMS_SERVICE_PLAN_IDS)
                    if teams_disabled:
                        license_type = "Disabled"
                    elif is_trial:
                        license_type = "Trial"
                    else:
                        license_type = "License"

                    sub = subscriptions.get(sku_id, {})
                    start_date = _format_date(sub.get("createdDateTime"))
                    end_date = _format_date(sub.get("nextLifecycleDateTime"))

                    rows.append((display_name, upn, sku_name, license_type, start_date, end_date))

            users_url = u_data.get("@odata.nextLink")

    if not rows:
        print("No users found with Teams licenses.")
        return

    rows.sort(key=lambda r: (r[0].lower(), r[2]))

    # Print table
    c = (32, 36, 34, 8, 12, 12)
    header = (
        f"{'User':<{c[0]}} {'UPN':<{c[1]}} {'SKU':<{c[2]}} {'Type':<{c[3]}} "
        f"{'Start':<{c[4]}} {'End':<{c[5]}}"
    )
    sep = "-" * (sum(c) + len(c) - 1)
    print(header)
    print(sep)
    for display_name, upn, sku_name, license_type, start_date, end_date in rows:
        print(
            f"{display_name:<{c[0]}} {upn:<{c[1]}} {sku_name:<{c[2]}} "
            f"{license_type:<{c[3]}} {start_date:<{c[4]}} {end_date:<{c[5]}}"
        )


if __name__ == "__main__":
    asyncio.run(list_teams_licenses())
