import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/calendar"
]


def get_calendar():

    info = json.loads(
        os.environ["GOOGLE_SERVICE_ACCOUNT"]
    )

    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=SCOPES,
    )

    return build(
        "calendar",
        "v3",
        credentials=creds,
    )
