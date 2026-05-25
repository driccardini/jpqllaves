import pandas as pd
from io import BytesIO
from urllib.request import urlopen

SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"

def extraer_cruces_c5():
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    df = pd.read_excel(BytesIO(workbook_bytes), sheet_name="c5", header=None)
    partidos = {}
    # Buscar todos los números de partido entre 34 y 49 en la columna 2 (índice 2)
    for idx in range(len(df)):
        val = str(df.iat[idx, 2]) if pd.notna(df.iat[idx, 2]) else ""
        if val.isdigit() and 34 <= int(val) <= 49:
            nro_partido = int(val)
            # Buscar hacia arriba la primera pareja (dos filas arriba)
            pareja1 = []
            for arriba in range(idx-1, idx-10, -1):
                if arriba < 0:
                    break
                nombre = str(df.iat[arriba, 2]) if pd.notna(df.iat[arriba, 2]) else ""
                if nombre.strip() != "" and not nombre.strip().isdigit():
                    pareja1.insert(0, nombre.strip())
                    if len(pareja1) == 2:
                        break
            # Buscar hacia abajo la segunda pareja (dos filas abajo)
            pareja2 = []
            for abajo in range(idx+1, idx+10):
                if abajo >= len(df):
                    break
                nombre = str(df.iat[abajo, 2]) if pd.notna(df.iat[abajo, 2]) else ""
                if nombre.strip() != "" and not nombre.strip().isdigit():
                    pareja2.append(nombre.strip())
                    if len(pareja2) == 2:
                        break
            partidos[nro_partido] = {"pareja_1": pareja1, "pareja_2": pareja2}
    return partidos

if __name__ == "__main__":
    partidos = extraer_cruces_c5()
    for nro in sorted(partidos):
        p = partidos[nro]
        print(f"Partido {nro}:")
        print(f"  Pareja 1: {', '.join(p['pareja_1'])}")
        print(f"  Pareja 2: {', '.join(p['pareja_2'])}")