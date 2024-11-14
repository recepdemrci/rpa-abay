import os
import msal
import logging
from dotenv import load_dotenv

# Load environment variables from .env file
# Azure Active Directory (Azure AD) secrets
load_dotenv()
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")


def get_access_token():
    # Create a PublicClientApplication object
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    # Try to acquire a token silently from cache
    scopes = [
        "User.Read",
        "Files.ReadWrite.All",
        "Sites.ReadWrite.All",
        "Sites.Manage.All",
        "Mail.Send",
    ]
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=scopes, account=accounts[0])
        if result:
            return result["access_token"]

    # If a token cannot be acquired silently, prompt the user to sign in
    result = app.acquire_token_interactive(scopes=scopes)
    if "access_token" in result:
        logging.info("Authentication successful.")
        return result["access_token"]
    else:
        raise Exception("Authentication failed.")
