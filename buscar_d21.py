import pandas as pd
from io import BytesIO
from urllib.request import urlopen

SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"

def buscar_d21():
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    df = pd.read_excel(BytesIO(workbook_bytes), sheet_name="c5", header=None)
    # D21 es columna 3 (índice 3), fila 20 (índice 20)
    valor = df.iat[20, 3] if df.shape[0] > 20 and df.shape[1] > 3 else None
    print(f"D21 (fila 21, columna D) en DataFrame: índice fila 20, columna 3")
    print(f"Valor en esa celda: {valor}")

if __name__ == "__main__":
    buscar_d21()
