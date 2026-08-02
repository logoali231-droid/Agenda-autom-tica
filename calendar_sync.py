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


def procurar_evento(service, canvas_id):

    eventos = service.events().list(
        calendarId=CALENDAR_ID,
        privateExtendedProperty=f"canvas_id={canvas_id}",
    ).execute()

    itens = eventos.get("items", [])

    if itens:
        return itens[0]

    return None


def criar_ou_atualizar_evento(
    service,
    canvas_id,
    titulo,
    descricao,
    inicio,
    link,
):

    fim = inicio + timedelta(hours=1)

    body = {
        "summary": titulo,
        "description": f"{descricao}\n\n{link}",
        "start": {
            "dateTime": inicio.isoformat(),
            "timeZone": "America/Sao_Paulo",
        },
        "end": {
            "dateTime": fim.isoformat(),
            "timeZone": "America/Sao_Paulo",
        },
        "extendedProperties": {
            "private": {
                "canvas_id": canvas_id
            }
        }
    }

    evento = procurar_evento(service, canvas_id)

    if evento:

        print(f"ATUALIZANDO EVENTO: {titulo}")

        service.events().update(
            calendarId=CALENDAR_ID,
            eventId=evento["id"],
            body=body,
        ).execute()

    else:

        print(f"CRIANDO EVENTO: {titulo}")

        service.events().insert(
            calendarId=CALENDAR_ID,
            body=body,
        ).execute()

    print("OK")
