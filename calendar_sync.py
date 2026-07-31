import json
import os
from datetime import timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/calendar"
]

CALENDAR_ID = "logoali231@gmail.com"


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


def criar_evento(
    service,
    titulo,
    descricao,
    inicio,
    link,
):

    body = {

        "summary": titulo,

        "description": f"{descricao}\n\n{link}",

        "start": {
            "dateTime": inicio.isoformat(),
            "timeZone": "America/Sao_Paulo",
        },

        "end": {
            "dateTime": (
                inicio + timedelta(hours=1)
            ).isoformat(),
            "timeZone": "America/Sao_Paulo",
        },
    }

    service.events().insert(
        calendarId=CALENDAR_ID,
        body=body,
    ).execute()
