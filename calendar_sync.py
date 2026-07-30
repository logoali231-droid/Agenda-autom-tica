from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import timedelta
import os
import pickle

def criar_evento(
    service,
    titulo,
    inicio,
    fim,
    descricao,
    url,
):
    evento = {

        "summary": titulo,

        "description":
            descricao + "\n\n" + url,

        "start": {
            "dateTime": inicio.isoformat(),
            "timeZone": "America/Sao_Paulo",
        },

        "end": {
            "dateTime": fim.isoformat(),
            "timeZone": "America/Sao_Paulo",
        },
    }

    service.events().insert(
        calendarId="primary",
        body=evento,
    ).execute()


SCOPES = [
    "https://www.googleapis.com/auth/calendar"
]

SERVICE = None


def get_calendar():

    creds = None

    if os.path.exists("token.pickle"):

        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:

            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json",
                SCOPES,
            )

            creds = flow.run_local_server(port=0)

        with open("token.pickle", "wb") as f:
            pickle.dump(creds, f)

    return build(
        "calendar",
        "v3",
        credentials=creds,
    )
