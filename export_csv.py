import pandas as pd
from main import _filter_visible_categories, _extract_match_schedule_labels, _normalize_cell_text, _cell_class
import pandas as pd
import os

def exportar_partidos_csv(path_csv="partidos_jpq.csv"):
    from io import BytesIO
    from urllib.request import urlopen
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"
    brackets = {}
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    excel_file = pd.ExcelFile(BytesIO(workbook_bytes))
    for sheet_name in excel_file.sheet_names:
        if sheet_name.lower().startswith("base"):
            continue
        bracket_df = pd.read_excel(BytesIO(workbook_bytes), sheet_name=sheet_name, header=None)
        not_empty = bracket_df.notna() & (bracket_df.astype(str).applymap(lambda x: x.strip() != ""))
        rows_with_data = not_empty.any(axis=1)
        cols_with_data = not_empty.any(axis=0)
        if not rows_with_data.any() or not cols_with_data.any():
            continue
        first_used_row = rows_with_data.idxmax()
        last_used_row = len(rows_with_data) - 1 - rows_with_data[::-1].idxmax()
        first_used_col = cols_with_data.idxmax()
        last_used_col = len(cols_with_data) - 1 - cols_with_data[::-1].idxmax()
        cropped = bracket_df.iloc[
            first_used_row : last_used_row + 1,
            first_used_col : last_used_col + 1,
        ].reset_index(drop=True)
        cropped.columns = range(cropped.shape[1])
        category = sheet_name.strip()
        brackets[category] = {"sheet": sheet_name, "grid": cropped}
    brackets = _filter_visible_categories(brackets)
    rows = []
    # DEBUG: Volcar primeras filas y columnas de la hoja c5
    # Volcar nombres de hojas detectadas
    from urllib.request import urlopen
    from io import BytesIO
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"
    workbook_bytes = urlopen(SHEET_URL, timeout=20).read()
    excel_file = pd.ExcelFile(BytesIO(workbook_bytes))
    sheet_names = excel_file.sheet_names
    with open("debug_c5.txt", "w", encoding="utf-8") as f:
        f.write("Hojas detectadas en el archivo:\n")
        for name in sheet_names:
            f.write(f"- {name}\n")

    if "c5" not in excel_file.sheet_names:
        with open("debug_c5.txt", "a", encoding="utf-8") as f:
            f.write("ERROR: La hoja 'c5' no fue encontrada en el archivo.\n")
        print("ERROR: La hoja 'c5' no fue encontrada en el archivo.")
    else:
        bracket_df = pd.read_excel(BytesIO(workbook_bytes), sheet_name="c5", header=None)
        with open("debug_c5.txt", "a", encoding="utf-8") as f:
            f.write("--- VOLCADO COMPLETO HOJA c5 ---\n")
            f.write(bracket_df.to_string(index=True, header=True))
            f.write("\n--- FIN VOLCADO ---\n")

    for categoria, payload in brackets.items():
        grid = payload["grid"]
        n_rows, n_cols = grid.shape
        for row_idx in range(n_rows):
            for col_idx in range(n_cols):
                cell = str(grid.iat[row_idx, col_idx]) if pd.notna(grid.iat[row_idx, col_idx]) else ""
                cell_norm = _normalize_cell_text(cell)
                # Buscar "1°" (puede ser "1° A", "1° B", etc)
                if cell_norm.startswith("1°"):
                    # pareja 1: columna a la derecha, dos filas
                    p1 = str(grid.iat[row_idx, col_idx+1]) if col_idx+1 < n_cols and pd.notna(grid.iat[row_idx, col_idx+1]) else ""
                    p2 = str(grid.iat[row_idx+1, col_idx+1]) if row_idx+1 < n_rows and col_idx+1 < n_cols and pd.notna(grid.iat[row_idx+1, col_idx+1]) else ""
                    pareja_1 = f"{_normalize_cell_text(p1)} {_normalize_cell_text(p2)}".strip()
                    # Buscar "2°" hacia abajo (máx 20 filas)
                    for k in range(row_idx+1, min(row_idx+20, n_rows)):
                        cell2 = str(grid.iat[k, col_idx]) if pd.notna(grid.iat[k, col_idx]) else ""
                        cell2_norm = _normalize_cell_text(cell2)
                        if cell2_norm.startswith("2°"):
                            # pareja 2: columna a la derecha, dos filas
                            p3 = str(grid.iat[k, col_idx+1]) if col_idx+1 < n_cols and pd.notna(grid.iat[k, col_idx+1]) else ""
                            p4 = str(grid.iat[k+1, col_idx+1]) if k+1 < n_rows and col_idx+1 < n_cols and pd.notna(grid.iat[k+1, col_idx+1]) else ""
                            pareja_2 = f"{_normalize_cell_text(p3)} {_normalize_cell_text(p4)}".strip()
                            # Buscar número de partido entre ambas parejas, en la columna dos a la derecha
                            match_number = ""
                            for m in range(row_idx, k+2):
                                if col_idx+2 < n_cols:
                                    val = str(grid.iat[m, col_idx+2]) if pd.notna(grid.iat[m, col_idx+2]) else ""
                                    if val.strip().isdigit():
                                        match_number = val.strip()
                                        break
                            if pareja_1 and pareja_2 and match_number:
                                rows.append({
                                    "categoria": categoria,
                                    "pareja_1": pareja_1,
                                    "pareja_2": pareja_2,
                                    "partido": match_number
                                })
                            break
    df = pd.DataFrame(rows)
    df.to_csv(path_csv, index=False)

if __name__ == "__main__":
    exportar_partidos_csv()
