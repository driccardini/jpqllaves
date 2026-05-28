#!/usr/bin/env python3
"""
Propaga ganadores al Sheet de resultados.

Cuando el partido 50 tiene ganador, escribe ese nombre en la columna
pareja_2 de la fila del partido 57 (y así con todos los cruces del torneo).

Uso:
    python propagar_resultados.py

Requiere JPQ_GOOGLE_SERVICE_ACCOUNT_FILE y JPQ_RESULTS_SHEET_ID configurados.
Puede ejecutarse varias veces sin problema — solo actualiza celdas vacías.
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
RESULTS_SHEET_ID = os.getenv("JPQ_RESULTS_SHEET_ID", "")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _build_credentials():
    oauth2 = importlib.import_module("google.oauth2.service_account")
    Creds = getattr(oauth2, "Credentials")
    if SERVICE_ACCOUNT_FILE:
        return Creds.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    if SERVICE_ACCOUNT_JSON:
        return Creds.from_service_account_info(json.loads(SERVICE_ACCOUNT_JSON), scopes=SCOPES)
    raise RuntimeError(
        "Configurar JPQ_GOOGLE_SERVICE_ACCOUNT_FILE o JPQ_GOOGLE_SERVICE_ACCOUNT_JSON"
    )


def _build_topology(csv_path: Path) -> dict[tuple[str, str], list[tuple[str, int]]]:
    """Lee el CSV y construye un mapa: (categoria, src_match) → [(tgt_match, position), ...]

    position: 1 = pareja_1, 2 = pareja_2 del partido destino.
    """
    df = pd.read_csv(csv_path, sep=";")
    df.columns = df.columns.str.strip()

    topology: dict[tuple[str, str], list[tuple[str, int]]] = {}

    for _, row in df.iterrows():
        cat = str(row.get("categoria", "")).strip().lower()
        tgt_num = str(row.get("numero de partido", "")).strip()
        try:
            tgt_num = str(int(float(tgt_num)))
        except (ValueError, TypeError):
            continue
        pareja = str(row.get("pareja", "")).strip()

        # "TEAM (espera ganador del partido N)" → position 2
        m = re.match(r"^(.+?)\s*\(espera ganador del partido\s+(\d+)\)", pareja, re.IGNORECASE)
        if m:
            src_num = m.group(2)
            topology.setdefault((cat, src_num), []).append((tgt_num, 2))
            continue

        # "#N vs #M" → each feeds as position 1 and 2
        m2 = re.match(r"^#(\d+)\s+vs\s+#(\d+)$", pareja, re.IGNORECASE)
        if m2:
            src1, src2 = m2.group(1), m2.group(2)
            topology.setdefault((cat, src1), []).append((tgt_num, 1))
            topology.setdefault((cat, src2), []).append((tgt_num, 2))

    return topology


def main() -> None:
    if not RESULTS_SHEET_ID:
        raise RuntimeError("Configurar JPQ_RESULTS_SHEET_ID")

    csv_path = Path(__file__).parent / "partidos_jpq.csv"
    topology = _build_topology(csv_path)

    creds = _build_credentials()
    transport = importlib.import_module("google.auth.transport.requests")
    Session = getattr(transport, "AuthorizedSession")
    session = Session(creds)

    sid = RESULTS_SHEET_ID

    # Read current sheet data
    resp = session.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values/A1:H5000",
        timeout=20,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Error leyendo sheet ({resp.status_code}): {resp.text[:300]}")

    values = resp.json().get("values", [])
    if not values:
        print("Sheet vacío.")
        return

    header = [h.strip().lower() for h in values[0]]
    col = {name: idx for idx, name in enumerate(header)}

    # Build index: (cat, num) → row_index (0-based in values, so row 2 in sheet = values[1])
    row_index: dict[tuple[str, str], int] = {}
    results: dict[tuple[str, str], str] = {}

    for i, row in enumerate(values[1:], start=1):
        padded = list(row) + [""] * max(0, len(header) - len(row))
        cat = padded[col["categoria"]].strip().lower()
        num = padded[col["numero_partido"]].strip()
        ganador = padded[col["ganador"]].strip() if col.get("ganador") is not None and len(padded) > col["ganador"] else ""
        row_index[(cat, num)] = i
        if ganador:
            results[(cat, num)] = ganador

    # Compute propagations
    updates: list[dict] = []
    col_ganador = col.get("ganador", 6)
    col_p1 = col.get("pareja_1", 4)
    col_p2 = col.get("pareja_2", 5)

    for (cat, src_num), winner in results.items():
        for tgt_num, position in topology.get((cat, src_num), []):
            tgt_row_i = row_index.get((cat, tgt_num))
            if tgt_row_i is None:
                continue

            tgt_row = list(values[tgt_row_i])
            padded_tgt = tgt_row + [""] * max(0, len(header) - len(tgt_row))

            pareja_col = col_p1 if position == 1 else col_p2
            current_val = padded_tgt[pareja_col].strip() if len(padded_tgt) > pareja_col else ""

            # Only fill if currently empty or still placeholder
            is_placeholder = re.match(r"^ganador partido \d+$", current_val, re.IGNORECASE)
            if current_val == "" or is_placeholder:
                # Sheet row number = tgt_row_i + 1 (1-based) + 1 for header
                sheet_row = tgt_row_i + 1  # already 1-based for the data, +1 for header row = tgt_row_i + 1 in values = sheet row tgt_row_i + 1
                col_letter = chr(ord("A") + pareja_col)
                range_notation = f"{col_letter}{sheet_row + 1}"
                updates.append({
                    "range": range_notation,
                    "values": [[winner]],
                })
                print(f"  [{cat.upper()}] Partido {src_num} → Partido {tgt_num} pareja_{position}: {winner}")

    if not updates:
        print("No hay propagaciones pendientes.")
        return

    batch_body = {
        "valueInputOption": "RAW",
        "data": updates,
    }
    update_resp = session.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values:batchUpdate",
        json=batch_body,
        timeout=20,
    )
    if update_resp.status_code != 200:
        raise RuntimeError(f"Error actualizando sheet ({update_resp.status_code}): {update_resp.text[:300]}")

    print(f"\n✅ {len(updates)} celda(s) propagada(s) al Sheet.")


if __name__ == "__main__":
    main()
