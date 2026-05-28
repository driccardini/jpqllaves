#!/usr/bin/env python3
"""
Script de configuración — ejecutar UNA SOLA VEZ antes del torneo.

Crea el Google Sheet "JPQ Resultados 1er Torneo" con todas las filas
de partidos listas para cargar resultados durante el fin de semana.

Uso:
    python crear_resultados_sheet.py

Requiere la variable de entorno JPQ_GOOGLE_SERVICE_ACCOUNT_FILE (o
JPQ_GOOGLE_SERVICE_ACCOUNT_JSON) configurada con credenciales que tengan
permisos de escritura en Google Sheets (scope: spreadsheets).

Al finalizar imprime el ID del Sheet que hay que configurar como
JPQ_RESULTS_SHEET_ID en el entorno de Streamlit y en .env local.
"""
from __future__ import annotations

import importlib
import json
import os
import re
from pathlib import Path

import pandas as pd

SERVICE_ACCOUNT_FILE = os.getenv("JPQ_GOOGLE_SERVICE_ACCOUNT_FILE", "")
SERVICE_ACCOUNT_JSON = os.getenv("JPQ_GOOGLE_SERVICE_ACCOUNT_JSON", "")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
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
        # Normalize float-like numbers ("50.0" → "50")
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

    # Crear el spreadsheet
    create_body = {
        "properties": {"title": "JPQ Resultados 1er Torneo"},
        "sheets": [{"properties": {"title": "resultados", "index": 0}}],
    }
    resp = session.post(
        "https://sheets.googleapis.com/v4/spreadsheets",
        json=create_body,
        timeout=20,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Error al crear el Sheet ({resp.status_code}): {resp.text[:400]}")

    sheet_data = resp.json()
    sheet_id: str = sheet_data["spreadsheetId"]
    sheet_url: str = sheet_data["spreadsheetUrl"]
    print(f"Sheet creado: {sheet_url}")

    # Escribir los datos
    update_resp = session.put(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/resultados!A1",
        json={"values": sheet_rows},
        params={"valueInputOption": "RAW"},
        timeout=30,
    )
    if update_resp.status_code != 200:
        raise RuntimeError(
            f"Error al escribir datos ({update_resp.status_code}): {update_resp.text[:400]}"
        )

    # Negrita en la fila de encabezado
    fmt_resp = session.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}:batchUpdate",
        json={
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                            }
                        },
                        "fields": "userEnteredFormat(textFormat,backgroundColor)",
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 0,
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
    if fmt_resp.status_code != 200:
        print(f"Advertencia: no se pudo formatear el encabezado ({fmt_resp.status_code})")

    print("\n" + "=" * 60)
    print("  SHEET CREADO EXITOSAMENTE")
    print("=" * 60)
    print(f"  URL : {sheet_url}")
    print(f"  ID  : {sheet_id}")
    print()
    print("  Próximos pasos:")
    print(f"  1. Compartí el Sheet con tu cuenta de servicio (si no")
    print(f"     tiene acceso automático).")
    print(f"  2. Configurar en Streamlit Cloud (Settings → Secrets):")
    print(f"       JPQ_RESULTS_SHEET_ID = \"{sheet_id}\"")
    print(f"  3. Configurar localmente en tu .env o shell:")
    print(f"       export JPQ_RESULTS_SHEET_ID=\"{sheet_id}\"")
    print("=" * 60)


if __name__ == "__main__":
    main()
