"""Autenticação Microsoft 365 via MSAL Client Credentials."""

import os

import msal
from dotenv import load_dotenv

load_dotenv()

AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
SCOPE = ["https://graph.microsoft.com/.default"]


def get_token() -> str:
    """Obtém access_token via Client Credentials flow."""
    app = msal.ConfidentialClientApplication(
        client_id=AZURE_CLIENT_ID,
        client_credential=AZURE_CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
    )

    result = app.acquire_token_for_client(scopes=SCOPE)

    if "access_token" in result:
        return result["access_token"]

    raise RuntimeError(
        f"Falha na autenticação MSAL: {result.get('error_description', result)}"
    )


if __name__ == "__main__":
    token = get_token()
    print("Token OK:", token[:50], "...")
