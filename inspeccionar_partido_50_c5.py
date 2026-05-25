import pandas as pd
from io import BytesIO
from urllib.request import urlopen

SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"

def inspeccionar_partido_50():
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    df = pd.read_excel(BytesIO(workbook_bytes), sheet_name="c5", header=None)
    # Buscar la fila donde está el partido 50 en la columna 2 (índice 2)
    for idx in range(len(df)):
        val = str(df.iat[idx, 2]) if pd.notna(df.iat[idx, 2]) else ""
        if val.strip() == "50":
            print(f"Partido 50 encontrado en fila {idx}, columna 2 (C{idx+1})")
            # Mostrar 10 filas arriba y 10 abajo para inspección
            for i in range(max(0, idx-10), min(len(df), idx+11)):
                fila = [str(df.iat[i, j]) if pd.notna(df.iat[i, j]) else "" for j in range(df.shape[1])]
                print(f"Fila {i}: {fila}")
            break

if __name__ == "__main__":
    inspeccionar_partido_50()