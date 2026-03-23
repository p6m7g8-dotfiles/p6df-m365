#!/usr/bin/env python3
# List Microsoft 365 users via Microsoft Graph SDK.
# Auth: client credentials (AZURE_CLIENT_SECRET) when set, else MSAL device code flow.
# Token cache: ~/.office365-mcp-tokens.json (device code only)

import asyncio
import os
import time
from pathlib import Path

import msal
from azure.core.credentials import AccessToken
from msgraph import GraphServiceClient

TOKEN_CACHE_PATH = Path.home() / ".office365-mcp-tokens.json"
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES        = ["https://graph.microsoft.com/.default"] if CLIENT_SECRET else ["https://graph.microsoft.com/User.Read.All"]


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


async def list_users():
    client = GraphServiceClient(credentials=_TokenCredential(), scopes=SCOPES)
    result = await client.users.get()
    if not result or not result.value:
        print("No users found.")
        return
    print(f"{'Display Name':<40} {'UPN':<50} {'ID'}")
    print("-" * 110)
    for user in result.value:
        print(f"{(user.display_name or ''):<40} {(user.user_principal_name or ''):<50} {user.id}")


if __name__ == "__main__":
    asyncio.run(list_users())
