#!/usr/bin/env python3
"""
Script de configuración — ejecutar UNA SOLA VEZ antes del torneo.

Rellena el Google Sheet de resultados del torneo JPQ con todos los
partidos, listo para cargar resultados durante el fin de semana.

Pasos previos (solo la primera vez):
  1. Crear un Google Sheet vacío en https://sheets.google.com
  2. Compartirlo con Editor con: padelstats@padelstats-492015.iam.gserviceaccount.com
  3. Copiar el ID del Sheet de la URL (el tramo largo entre /d/ y /edit)
  4. Ejecutar: python crear_resultados_sheet.py <SHEET_ID>

Al finalizar imprime el ID del Sheet que hay que configurar como
JPQ_RESULTS_SHEET_ID en Streamlit Cloud y en .env local.
"""
from __future__ import annotations

import importlib
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd

SERVICE_ACCOUNT_FILE = os.getenv("JPQ_GOOGLE_SERVICE_ACCOUNT_FILE", "")
SERVICE_ACCOUNT_JSON = os.getenv("JPQ_GOOGLE_SERVICE_ACCOUNT_JSON", "")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]


def _build_credentials():
    oauth2_module = importlib.import_module("google.oauth2.service_account")
    Credentials = getattr(oauth2_module, "Credentials")

    if SERVICE_ACCOUNT_FILE:
        cred_path = Path(SERVICE_ACCOUNT_FILE)
        if not cred_path.exists():
            raise RuntimeError(f"Archivo de credenciales no encontrado: {cred_path}")
        return Credentials.from_service_account_file(str(cred_path), scopes=SCOPES)

    if SERVICE_ACCOUNT_JSON:
        try:
            info = json.loads(SERVICE_ACCOUNT_JSON)
        except json.JSONDecodeError as exc:
            raise RuntimeError("JPQ_GOOGLE_SERVICE_ACCOUNT_JSON no es JSON válido") from exc
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    raise RuntimeError(
        "No se encontraron credenciales. "
        "Configurar JPQ_GOOGLE_SERVICE_ACCOUNT_FILE o JPQ_GOOGLE_SERVICE_ACCOUNT_JSON."
    )


def _parse_pairs(pareja: str) -> tuple[str, str]:
    """Divide el campo 'pareja' en (pareja_1, pareja_2)."""
    pareja = pareja.strip()

    # "TEAM1 (espera ganador del partido N)"
    m = re.match(r"^(.+?)\s*\(espera ganador del partido\s+(\d+)\)", pareja, re.IGNORECASE)
    if m:
        return m.group(1).strip(), f"Ganador partido {m.group(2)}"

    # "TEAM1 vs TEAM2" or "#57 vs #58"
    if " vs " in pareja:
        parts = pareja.split(" vs ", 1)
        return parts[0].strip(), parts[1].strip()

    return pareja, ""


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        print("ERROR: Falta el ID del Sheet como argumento.")
        print("Uso: python crear_resultados_sheet.py <SHEET_ID>")
        sys.exit(1)

    sheet_id = sys.argv[1].strip()

    csv_path = Path(__file__).parent / "partidos_jpq.csv"
    if not csv_path.exists():
        raise RuntimeError(f"No se encontró {csv_path}")

    df = pd.read_csv(csv_path, sep=";")
    df.columns = df.columns.str.strip()

    headers = [
        "categoria",
        "numero_partido",
        "dia_horario",
        "complejo",
        "pareja_1",
        "pareja_2",
        "ganador",
        "marcador",
    ]
    sheet_rows: list[list[str]] = [headers]

    for _, row in df.iterrows():
        cat = str(row.get("categoria", "")).strip()
        num = str(row.get("numero de partido", "")).strip()
        try:
            num = str(int(float(num)))
        except (ValueError, TypeError):
            pass
        dia = str(row.get("dia y horario", "")).strip()
        complejo = str(row.get("complejo", "")).strip()
        pareja = str(row.get("pareja", "")).strip()
        p1, p2 = _parse_pairs(pareja)
        sheet_rows.append([cat, num, dia, complejo, p1, p2, "", ""])

    print(f"Total de partidos a cargar: {len(sheet_rows) - 1}")

    creds = _build_credentials()
    transport_module = importlib.import_module("google.auth.transport.requests")
    AuthorizedSession = getattr(transport_module, "AuthorizedSession")
    session = AuthorizedSession(creds)

    # Limpiar contenido previo y escribir datos
    clear_resp = session.delete(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/A1:H5000:clear",
        timeout=20,
    )
    # Ignorar error de clear (puede ser método incorrecto)
    clear_resp2 = session.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/A1:H5000:clear",
        timeout=20,
    )

    update_resp = session.put(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/A1",
        json={"values": sheet_rows},
        params={"valueInputOption": "RAW"},
        timeout=30,
    )
    if update_resp.status_code != 200:
        raise RuntimeError(
            f"Error al escribir datos ({update_resp.status_code}): {update_resp.text[:400]}"
        )

    # Negrita + fondo gris en encabezado
    session.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}:batchUpdate",
        json={
            "requests": [
                {
                    "repeatCell": {
                        "range": {"startRowIndex": 0, "endRowIndex": 1},
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.85, "green": 0.85, "blue": 0.85},
                            }
                        },
                        "fields": "userEnteredFormat(textFormat,backgroundColor)",
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": 8,
                        }
                    }
                },
            ]
        },
        timeout=20,
    )

    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    print("\n" + "=" * 60)
    print("  DATOS CARGADOS EXITOSAMENTE")
    print("=" * 60)
    print(f"  URL : {sheet_url}")
    print(f"  ID  : {sheet_id}")
    print()
    print("  Próximos pasos:")
    print("  1. Configurar en Streamlit Cloud (Settings → Secrets):")
    print(f'       JPQ_RESULTS_SHEET_ID = "{sheet_id}"')
    print("  2. Configurar localmente:")
    print(f'       export JPQ_RESULTS_SHEET_ID="{sheet_id}"')
    print("=" * 60)


if __name__ == "__main__":
    main()
