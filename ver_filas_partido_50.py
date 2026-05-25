import pandas as pd
from io import BytesIO
from urllib.request import urlopen

SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"

def ver_filas_cercanas_partido_50():
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    df = pd.read_excel(BytesIO(workbook_bytes), sheet_name="c5", header=None)
    # Buscar la fila donde está el partido 50 en la columna E (índice 4)
    partido_50_row = None
    for idx in range(len(df)):
        val = str(df.iat[idx, 4]) if pd.notna(df.iat[idx, 4]) else ""
        val_clean = val.strip().replace(".0", "")
        if val_clean == "50":
            partido_50_row = idx
            break
    if partido_50_row is not None:
        print(f"Partido 50 encontrado en fila {partido_50_row} (índice DataFrame)")
        print("Filas cercanas en columna E (índice 4):")
        for i in range(max(0, partido_50_row-10), min(len(df), partido_50_row+11)):
            print(f"Fila {i}: {df.iat[i, 4] if df.shape[1] > 4 else ''}")
    else:
        print("No se encontró el partido 50 en la columna E. Listando todos los valores de la columna E:")
        for idx in range(len(df)):
            print(f"Fila {idx}: {df.iat[idx, 4] if df.shape[1] > 4 else ''}")

if __name__ == "__main__":
    ver_filas_cercanas_partido_50()
