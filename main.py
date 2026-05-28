from pathlib import Path
from html import escape
import os
import base64
import json
import re
import importlib
from io import BytesIO
import textwrap
from typing import Any, Dict, List, Optional, Tuple
from urllib.request import urlopen
from urllib.parse import quote

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


def _get_config_value(name: str, default: str = "") -> str:
    env_value = os.getenv(name)
    if env_value is not None and str(env_value).strip() != "":
        return str(env_value).strip()

    try:
        secret_value = st.secrets.get(name, "")
    except Exception:
        secret_value = ""

    if secret_value is None:
        return default

    secret_value = str(secret_value).strip()
    return secret_value if secret_value != "" else default


def _get_config_bool(name: str, default: bool) -> bool:
    default_text = "1" if default else "0"
    value = _get_config_value(name, default_text).lower()
    return value in {"1", "true", "yes"}


DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Wm5ieVgYn1NCNNtzyjs-bOF6kKzCjkk76Bdq0KaUNzE/export?format=xlsx"
SHEET_URL = _get_config_value("JPQ_SHEET_URL", DEFAULT_SHEET_URL)
SHEET_ID = _get_config_value("JPQ_SHEET_ID", "")
SERVICE_ACCOUNT_FILE = _get_config_value("JPQ_GOOGLE_SERVICE_ACCOUNT_FILE", "")
SERVICE_ACCOUNT_JSON = _get_config_value("JPQ_GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_EXPORT_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GOOGLE_SERVICE_ACCOUNT_SCOPES = (
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
)
DEFAULT_VISIBLE_CATEGORIES = (
    "c5 40",
    "c6 35",
    "c6 45",
    "c7 45",
    "d2",
    "d4",
    "d5",
    "d6",
    "d7",
)
VISIBLE_CATEGORIES_ENV = os.getenv("JPQ_VISIBLE_CATEGORIES", "all").strip()
SHEET_TIMEOUT_SECONDS = 20
RESULTS_SHEET_ID = _get_config_value("JPQ_RESULTS_SHEET_ID", "")
LOGO_GLOB = "Logo JPQ*"
CELL_WIDTH = 132
CELL_HEIGHT = 26
LOCAL_FALLBACK_DIR = Path(__file__).parent / "LLAVES 1er JPQ"
LAST_LOAD_ERROR: Optional[str] = None


APP_CSS = """
<style>
    .section-title {
        margin: 10px 0 6px 0;
        font-weight: 700;
    }
    .header-wrap {
        display: flex;
        align-items: center;
        gap: 14px;
    }
    .header-logo {
        width: 66px;
        height: 66px;
        border-radius: 12px;
        object-fit: cover;
        opacity: 0.72;
        mix-blend-mode: screen;
    }
    .header-title {
        margin: 0;
    }
    div[data-baseweb="select"] > div {
        border-radius: 10px;
        min-height: 44px;
    }
    @media (max-width: 768px) {
        .header-logo {
            width: 50px;
            height: 50px;
        }
        .header-title {
            font-size: 1.45rem;
        }
    }
</style>
"""


@st.cache_data(show_spinner=False)
def load_logo_data_uri() -> Optional[str]:
    logo_matches = sorted(Path(__file__).parent.glob(LOGO_GLOB))
    if not logo_matches:
        return None

    logo_path = logo_matches[0]
    suffix = logo_path.suffix.lower()
    mime_type = "image/jpeg" if suffix in {".jpg", ".jpeg"} else "image/png"
    encoded = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"


@st.cache_data(show_spinner=False, ttl=3600)
def load_brackets() -> Dict[str, Dict[str, object]]:
    global LAST_LOAD_ERROR
    use_local_fallback = _get_config_bool("JPQ_USE_LOCAL_FALLBACK", True)

    try:
        workbook_bytes = _download_workbook_bytes()
        LAST_LOAD_ERROR = None
        return _load_brackets_from_workbook_bytes(workbook_bytes)
    except Exception as e:
        LAST_LOAD_ERROR = str(e)
        print("ERROR AL DESCARGAR PLANILLA:", e)
        if not use_local_fallback:
            raise

    local_brackets = _load_local_brackets()
    if local_brackets:
        LAST_LOAD_ERROR = None
        return local_brackets

    return {}


@st.cache_data(show_spinner=False, ttl=60)
def load_results() -> Dict[Tuple[str, str], Dict[str, str]]:
    """Lee el Sheet de resultados del torneo.

    Retorna un dict keyed por (categoria_lower, numero_partido) con los
    campos 'ganador' y 'marcador' de cada partido jugado.
    Retorna un dict vacío si no hay resultados o no está configurado el Sheet.
    """
    if not RESULTS_SHEET_ID:
        return {}

    try:
        creds = _build_service_account_credentials()
        if creds is None:
            return {}

        transport_module = importlib.import_module("google.auth.transport.requests")
        AuthorizedSession = getattr(transport_module, "AuthorizedSession")
        session = AuthorizedSession(creds)

        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{RESULTS_SHEET_ID}"
            f"/values/A1:H5000"
        )
        resp = session.get(url, timeout=SHEET_TIMEOUT_SECONDS)
        if resp.status_code != 200:
            print(f"RESULTADOS: HTTP {resp.status_code} al leer el Sheet de resultados")
            return {}

        values = resp.json().get("values", [])
        if not values:
            return {}

        header = [h.strip().lower() for h in values[0]]
        results: Dict[Tuple[str, str], Dict[str, str]] = {}

        for row in values[1:]:
            # Pad row to avoid index errors
            padded = list(row) + [""] * max(0, len(header) - len(row))
            row_dict = dict(zip(header, padded))
            cat = str(row_dict.get("categoria", "")).strip().lower()
            num = str(row_dict.get("numero_partido", "")).strip()
            ganador = str(row_dict.get("ganador", "")).strip()
            marcador = str(row_dict.get("marcador", "")).strip()

            if not cat or not num:
                continue
            if ganador:
                results[(cat, num)] = {"ganador": ganador, "marcador": marcador}

        return results

    except Exception as exc:
        print(f"RESULTADOS: error cargando resultados: {exc}")
        return {}


def _download_workbook_bytes() -> bytes:
    """Download the workbook, preferring public URL and falling back to authenticated export."""
    errors: List[str] = []

    try:
        return urlopen(SHEET_URL, timeout=SHEET_TIMEOUT_SECONDS).read()
    except Exception as e:
        errors.append(f"descarga publica: {e}")

    sheet_id = _resolve_sheet_id()
    if not sheet_id:
        detail = " | ".join(errors)
        raise RuntimeError(f"No se pudo descargar la planilla y no hay SHEET_ID utilizable. Detalle: {detail}")

    creds = _build_service_account_credentials()
    if creds is None:
        detail = " | ".join(errors)
        raise RuntimeError(
            "No se pudo descargar de forma publica y faltan credenciales de service account "
            f"(JPQ_GOOGLE_SERVICE_ACCOUNT_FILE o JPQ_GOOGLE_SERVICE_ACCOUNT_JSON). Detalle: {detail}"
        )

    export_url = (
        f"https://www.googleapis.com/drive/v3/files/{quote(sheet_id)}/export"
        f"?mimeType={quote(GOOGLE_EXPORT_MIME)}"
    )
    transport_requests_module = importlib.import_module("google.auth.transport.requests")
    AuthorizedSession = getattr(transport_requests_module, "AuthorizedSession")

    session = AuthorizedSession(creds)
    response = session.get(export_url, timeout=SHEET_TIMEOUT_SECONDS)
    if response.status_code != 200:
        body = response.text.strip()
        raise RuntimeError(
            "Export autenticado fallido "
            f"(HTTP {response.status_code}): {body[:300]}"
        )

    return response.content


def _resolve_sheet_id() -> Optional[str]:
    if SHEET_ID:
        return SHEET_ID
    return _extract_sheet_id_from_url(SHEET_URL)


def _extract_sheet_id_from_url(url: str) -> Optional[str]:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    return None


def _build_service_account_credentials() -> Optional[Any]:
    oauth2_module = importlib.import_module("google.oauth2.service_account")
    credentials_factory = getattr(oauth2_module, "Credentials")

    if SERVICE_ACCOUNT_FILE:
        credentials_path = Path(SERVICE_ACCOUNT_FILE)
        if not credentials_path.exists():
            raise RuntimeError(f"No existe el archivo de credenciales: {credentials_path}")
        return credentials_factory.from_service_account_file(
            str(credentials_path),
            scopes=GOOGLE_SERVICE_ACCOUNT_SCOPES,
        )

    if SERVICE_ACCOUNT_JSON:
        try:
            account_info = json.loads(SERVICE_ACCOUNT_JSON)
        except json.JSONDecodeError as e:
            raise RuntimeError("JPQ_GOOGLE_SERVICE_ACCOUNT_JSON no es JSON valido") from e
        return credentials_factory.from_service_account_info(
            account_info,
            scopes=GOOGLE_SERVICE_ACCOUNT_SCOPES,
        )

    # Streamlit Cloud friendly format:
    # [gcp_service_account]
    # type = "service_account"
    # project_id = "..."
    # ...
    try:
        gcp_secret = st.secrets.get("gcp_service_account", None)
    except Exception:
        gcp_secret = None

    if gcp_secret:
        if isinstance(gcp_secret, dict):
            account_info = dict(gcp_secret)
        else:
            # Some Streamlit secret providers may return a mapping-like object.
            account_info = {k: gcp_secret[k] for k in gcp_secret}
        return credentials_factory.from_service_account_info(
            account_info,
            scopes=GOOGLE_SERVICE_ACCOUNT_SCOPES,
        )

    return None


def _load_brackets_from_workbook_bytes(workbook_bytes: bytes) -> Dict[str, Dict[str, object]]:
    brackets: Dict[str, Dict[str, object]] = {}
    excel_file = pd.ExcelFile(BytesIO(workbook_bytes))

    for sheet_name in excel_file.sheet_names:
        if sheet_name.lower().startswith("base"):
            continue

        bracket_df = pd.read_excel(BytesIO(workbook_bytes), sheet_name=sheet_name, header=None)
        cropped = _crop_bracket_df(bracket_df)
        if cropped is None:
            continue

        category = sheet_name.strip()
        brackets[category] = {"sheet": sheet_name, "grid": cropped}

    return brackets


def _load_local_brackets() -> Dict[str, Dict[str, object]]:
    brackets: Dict[str, Dict[str, object]] = {}
    if not LOCAL_FALLBACK_DIR.exists():
        return brackets

    for workbook_path in sorted(LOCAL_FALLBACK_DIR.glob("*.xlsx")):
        try:
            excel_file = pd.ExcelFile(workbook_path)
        except Exception as e:
            print(f"ERROR AL LEER ARCHIVO LOCAL {workbook_path.name}:", e)
            continue

        for sheet_name in excel_file.sheet_names:
            if sheet_name.lower().startswith("base"):
                continue

            bracket_df = pd.read_excel(workbook_path, sheet_name=sheet_name, header=None)
            cropped = _crop_bracket_df(bracket_df)
            if cropped is None:
                continue

            category = workbook_path.stem.strip()
            brackets[category] = {"sheet": sheet_name, "grid": cropped}

    return brackets


def _crop_bracket_df(bracket_df: pd.DataFrame) -> Optional[pd.DataFrame]:
    first_used_row: Optional[int] = None
    first_used_col: Optional[int] = None
    last_used_row = 0
    last_used_col = 0

    for row_idx in range(bracket_df.shape[0]):
        for col_idx in range(bracket_df.shape[1]):
            value = bracket_df.iat[row_idx, col_idx]
            if pd.notna(value) and str(value).strip() != "":
                if first_used_row is None:
                    first_used_row = row_idx
                else:
                    first_used_row = min(first_used_row, row_idx)
                if first_used_col is None:
                    first_used_col = col_idx
                else:
                    first_used_col = min(first_used_col, col_idx)
                last_used_row = max(last_used_row, row_idx)
                last_used_col = max(last_used_col, col_idx)

    if first_used_row is None or first_used_col is None:
        return None

    cropped = bracket_df.iloc[
        first_used_row : last_used_row + 1,
        first_used_col : last_used_col + 1,
    ].reset_index(drop=True)
    cropped.columns = range(cropped.shape[1])
    return cropped


def _cell_class(value: str) -> str:
    text = _normalize_cell_text(value)
    if text.lower().startswith("categoría"):
        return "cat-title"
    if text.upper() in {"Nº PARTIDO", "DIA - HORA", "COMPLEJO"}:
        return "legend"
    digits = text.replace(".", "", 1)
    if digits.isdigit():
        try:
            number = int(float(text))
            if 1 <= number <= 999:
                return "match-id"
        except ValueError:
            pass
    if "°" in text:
        return "seed"
    return "team"


def _extract_match_schedule_labels(grid: pd.DataFrame) -> Dict[str, str]:
    labels: Dict[str, str] = {}

    for row_idx in range(grid.shape[0]):
        row_texts = {
            col_idx: _normalize_cell_text(str(grid.iat[row_idx, col_idx]))
            if pd.notna(grid.iat[row_idx, col_idx])
            else ""
            for col_idx in range(grid.shape[1])
        }

        partido_col = next(
            (col for col, text in row_texts.items() if text.upper() == "Nº PARTIDO"),
            None,
        )
        dia_hora_col = next(
            (col for col, text in row_texts.items() if text.upper() == "DIA - HORA"),
            None,
        )
        complejo_col = next(
            (col for col, text in row_texts.items() if text.upper() == "COMPLEJO"),
            None,
        )

        if partido_col is None or dia_hora_col is None or complejo_col is None:
            continue

        empty_streak = 0
        for data_row in range(row_idx + 1, grid.shape[0]):
            partido_raw = _normalize_cell_text(
                str(grid.iat[data_row, partido_col]) if pd.notna(grid.iat[data_row, partido_col]) else ""
            )
            dia_hora = _normalize_cell_text(
                str(grid.iat[data_row, dia_hora_col]) if pd.notna(grid.iat[data_row, dia_hora_col]) else ""
            )
            complejo = _normalize_cell_text(
                str(grid.iat[data_row, complejo_col]) if pd.notna(grid.iat[data_row, complejo_col]) else ""
            )

            if partido_raw == "" and dia_hora == "" and complejo == "":
                empty_streak += 1
                if empty_streak >= 2:
                    break
                continue
            empty_streak = 0

            digits = partido_raw.replace(".", "", 1)
            if not digits.isdigit():
                continue

            match_number = str(int(float(partido_raw)))
            if dia_hora and complejo:
                labels[match_number] = f"{match_number} | {dia_hora} | {complejo}"
            elif dia_hora:
                labels[match_number] = f"{match_number} | {dia_hora}"
            elif complejo:
                labels[match_number] = f"{match_number} | {complejo}"

        break

    return labels


def _normalize_cell_text(value: str) -> str:
    text = value.replace("\xa0", " ").strip()
    if len(text) >= 2 and text[:-1].isdigit() and text[-1].isalpha() and text[-1].isupper() and "°" not in text:
        return f"{int(text[:-1])}° {text[-1]}"
    return text


def _normalize_category_name(value: str) -> str:
    return " ".join(value.strip().lower().split())


def _filter_visible_categories(brackets: Dict[str, Dict[str, object]]) -> Dict[str, Dict[str, object]]:
    env_value = VISIBLE_CATEGORIES_ENV.strip()
    if not env_value or env_value.lower() == "all":
        return brackets

    allowed = {
        _normalize_category_name(item)
        for item in env_value.split(",")
        if item.strip()
    }
    return {
        category: payload
        for category, payload in brackets.items()
        if _normalize_category_name(category) in allowed
    }


def _match_stage(value: str) -> Optional[int]:
    text = value.strip()
    digits = text.replace(".", "", 1)
    if not digits.isdigit():
        return None

    match_number = int(float(text))
    if 34 <= match_number <= 49:
        return 0  # 16vos
    if 50 <= match_number <= 57:
        return 1  # 8vos
    if 58 <= match_number <= 61:
        return 2  # 4tos
    if 62 <= match_number <= 63:
        return 3  # Semi
    if match_number == 64:
        return 4  # Final
    return None


def _build_connectors(
    nodes: List[Dict[str, object]],
    board_rows: int,
    category: Optional[str] = None,
) -> str:
    connector_pairs = _compute_connector_pairs(nodes=nodes, category=category)

    def route(x1: int, y1: int, x2: int, y2: int) -> str:
        if x2 >= x1:
            mid_x = x1 + max(14, (x2 - x1) // 2)
        else:
            mid_x = x2 - 14
        return f"M{x1},{y1} H{mid_x} V{y2} H{x2}"

    def target_x(right_node: Dict[str, object]) -> int:
        default_x = int(right_node["x"]) + 88
        category_key = (category or "").lower()

        if category_key in {"c2", "d2", "d4", "d5", "d6", "d7"}:
            return int(right_node["x"]) - 12

        if category_key != "c3":
            return default_x

        right_text = str(right_node["text"]).strip()
        digits = right_text.replace(".", "", 1)
        if digits.isdigit():
            match_number = int(float(right_text))
            if match_number >= 57:
                return int(right_node["x"]) - 12
        return default_x

    def source_x(left_node: Dict[str, object]) -> int:
        default_x = int(left_node["x"]) + 78
        if (category or "").lower() != "c3":
            return default_x

        if left_node.get("class") != "match-id":
            return default_x

        left_text = str(left_node["text"]).strip()
        digits = left_text.replace(".", "", 1)
        if digits.isdigit():
            match_number = int(float(left_text))
            if match_number >= 57:
                return int(left_node["x"]) + 106
        return default_x

    connector_paths: List[str] = []
    c3_shared_trunk_x_by_target: Dict[str, int] = {}
    for left_node, right_node in connector_pairs:
        x1 = source_x(left_node)
        y1 = int(left_node["y"]) + 12
        x2 = target_x(right_node)
        y2 = int(right_node["y"]) + 12

        if (category or "").lower() == "c3":
            left_text = str(left_node["text"]).strip()
            right_text = str(right_node["text"]).strip()
            if (left_text, right_text) in {("50", "57"), ("54", "59"), ("55", "60")}:
                trunk_x = x1 + max(14, (x2 - x1) // 2)
                c3_shared_trunk_x_by_target[right_text] = trunk_x
                connector_paths.append(f"M{x1},{y1} H{trunk_x} V{y2} H{x2}")
                continue

        connector_paths.append(route(x1, y1, x2, y2))

    if (category or "").lower() == "c2":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]) == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            target_left_x = int(target_node["x"]) - 12
            x2 = target_left_x
            y2 = int(target_node["y"]) + 12

            if (
                seed_label.strip() == "1° A"
                and target_number == "57"
                and "57" in c3_shared_trunk_x_by_target
            ):
                x2 = target_left_x
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            if (
                seed_label.strip() == "1° C"
                and target_number == "59"
                and "59" in c3_shared_trunk_x_by_target
            ):
                x2 = target_left_x
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            if (
                seed_label.strip() == "1° B"
                and target_number == "60"
                and "60" in c3_shared_trunk_x_by_target
            ):
                x2 = target_left_x
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1°", "57")
        append_direct_connector("9°", "50")
        append_direct_connector("8°", "50")
        append_direct_connector("5°", "51")
        append_direct_connector("12°", "51")
        append_direct_connector("4°", "58")
        append_direct_connector("3°", "59")
        append_direct_connector("11°", "54")
        append_direct_connector("6°", "54")
        append_direct_connector("7°", "55")
        append_direct_connector("10°", "55")
        append_direct_connector("2°", "60")

    if (category or "").lower() == "c3":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]) == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = target_x(target_node)
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "57")
        append_direct_connector("1° C", "59")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "d5":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "57")
        append_direct_connector("1° E", "58")
        append_direct_connector("1° D", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("1° B", "60")
        append_direct_connector("3° A", "54")
        append_direct_connector("2° E", "54")
        append_direct_connector("2° D", "55")
        append_direct_connector("2° A", "55")

    if (category or "").lower() == "d2":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("3° A", "58")
        append_direct_connector("2° B", "58")
        append_direct_connector("2° A", "59")
        append_direct_connector("3° B", "59")
        append_direct_connector("1° A", "61")
        append_direct_connector("1° B", "62")

    if (category or "").lower() == "d4":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("3° A", "50")
        append_direct_connector("2° B", "50")
        append_direct_connector("1° A", "57")
        append_direct_connector("2° C", "58")
        append_direct_connector("1° D", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("2° D", "59")
        append_direct_connector("2° A", "60")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "d6":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "57")

    if (category or "").lower() == "d7":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("2° B", "50")
        append_direct_connector("2° C", "50")
        append_direct_connector("2° D", "55")
        append_direct_connector("2° A", "55")
        append_direct_connector("1° A", "57")
        append_direct_connector("1° E", "58")
        append_direct_connector("1° D", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("2° E", "59")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "c6 35":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("2° C", "34")
        append_direct_connector("2° F", "34")
        append_direct_connector("2° G", "35")
        append_direct_connector("2° B", "35")
        append_direct_connector("2° A", "36")
        append_direct_connector("2° H", "36")
        append_direct_connector("2° E", "37")
        append_direct_connector("2° D", "37")
        append_direct_connector("1° A", "49")
        append_direct_connector("1° I", "50")
        append_direct_connector("1° H", "50")
        append_direct_connector("1° E", "51")
        append_direct_connector("2° J", "51")
        append_direct_connector("1° D", "52")
        append_direct_connector("1° C", "53")
        append_direct_connector("2° I", "54")
        append_direct_connector("1° F", "54")
        append_direct_connector("1° G", "55")
        append_direct_connector("1° J", "55")
        append_direct_connector("1° B", "56")

    if (category or "").lower() == "c6 45":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("2° B", "50")
        append_direct_connector("2° C", "50")
        append_direct_connector("1° E", "51")
        append_direct_connector("3° B", "51")
        append_direct_connector("3° A", "54")
        append_direct_connector("2° E", "54")
        append_direct_connector("2° D", "55")
        append_direct_connector("2° A", "55")
        append_direct_connector("1° A", "57")
        append_direct_connector("1° D", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "c7 45":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("2° C", "50")
        append_direct_connector("2° F", "50")
        append_direct_connector("1° E", "51")
        append_direct_connector("2° B", "51")
        append_direct_connector("2° A", "54")
        append_direct_connector("1° F", "54")
        append_direct_connector("2° E", "55")
        append_direct_connector("2° D", "55")
        append_direct_connector("1° A", "57")
        append_direct_connector("1° D", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "c6 40":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "57")
        append_direct_connector("2° B", "57")
        append_direct_connector("1° D", "58")
        append_direct_connector("2° C", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("2° D", "59")
        append_direct_connector("2° A", "60")
        append_direct_connector("1° B", "60")

    if (category or "").lower() == "c7 40":
        def find_direct_anchor(seed_label: str) -> Optional[tuple[int, int]]:
            seed_text = seed_label.strip()
            seed_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and str(node["text"]).strip() == seed_text
                ),
                None,
            )
            if seed_node is None:
                return None

            seed_col = int(seed_node["col"])
            seed_row = int(seed_node["row"])
            paired_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and int(node["row"]) in {seed_row, seed_row + 1}
                ],
                key=lambda node: int(node["row"]),
            )
            if not paired_teams:
                return None

            x = int(paired_teams[0]["x"])
            if len(paired_teams) >= 2:
                y = (int(paired_teams[0]["y"]) + int(paired_teams[1]["y"])) // 2 + 12
            else:
                y = int(paired_teams[0]["y"]) + 12
            return x, y

        def append_direct_connector(seed_label: str, target_number: str) -> None:
            anchor = find_direct_anchor(seed_label)
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id" and str(node["text"]).strip() == target_number
                ),
                None,
            )
            if anchor is None or target_node is None:
                return

            x1 = anchor[0] + 78
            y1 = anchor[1]
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "57")
        append_direct_connector("2° B", "57")
        append_direct_connector("1° D", "58")
        append_direct_connector("2° C", "58")
        append_direct_connector("1° C", "59")
        append_direct_connector("2° D", "59")
        append_direct_connector("2° A", "60")
        append_direct_connector("1° B", "60")

    return "".join(
        f'<path d="{path}" class="connector"></path>' for path in connector_paths
    )


def _compute_connector_pairs(
    nodes: List[Dict[str, object]],
    category: Optional[str] = None,
) -> List[tuple[Dict[str, object], Dict[str, object]]]:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    match_nodes = [
        node
        for node in nodes
        if node["class"] == "match-id"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]
    team_nodes = [
        node
        for node in nodes
        if node["class"] == "team"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]

    if len(match_nodes) < 2:
        return []

    stage_map: Dict[int, List[Dict[str, object]]] = {}
    for node in match_nodes:
        stage = _match_stage(str(node["text"]))
        if stage is None:
            continue
        stage_map.setdefault(stage, []).append(node)

    for stage in stage_map:
        stage_map[stage].sort(key=lambda item: int(item["row"]))

    sorted_stages = sorted(stage_map.keys())
    connector_pairs: List[tuple[Dict[str, object], Dict[str, object]]] = []

    def connect(left_node: Dict[str, object], right_node: Dict[str, object]) -> None:
        connector_pairs.append((left_node, right_node))

    for idx in range(len(sorted_stages) - 1):
        left_stage = sorted_stages[idx]
        right_stage = sorted_stages[idx + 1]
        if right_stage != left_stage + 1:
            continue

        left_nodes = stage_map[left_stage]
        right_nodes = stage_map[right_stage]
        if not right_nodes:
            continue

        capacities: List[int] = []
        for right_node in right_nodes:
            right_col = int(right_node["col"])
            right_row = int(right_node["row"])
            nearby_direct_team_rows = [
                team
                for team in team_nodes
                if int(team["col"]) == right_col and abs(int(team["row"]) - right_row) <= 8
            ]
            direct_pairs = len(nearby_direct_team_rows) // 2
            incoming_needed = max(0, 2 - direct_pairs)
            capacities.append(incoming_needed)

        total_incoming_needed = sum(capacities)

        if total_incoming_needed == len(left_nodes):
            left_idx = 0
            for right_node, incoming_needed in zip(right_nodes, capacities):
                for _ in range(incoming_needed):
                    if left_idx >= len(left_nodes):
                        break
                    connect(left_nodes[left_idx], right_node)
                    left_idx += 1
        else:
            left_count = len(left_nodes)
            right_count = len(right_nodes)
            for right_idx, right_node in enumerate(right_nodes):
                start_idx = round(right_idx * left_count / right_count)
                end_idx = round((right_idx + 1) * left_count / right_count)
                for left_node in left_nodes[start_idx:end_idx]:
                    connect(left_node, right_node)

    if (category or "").lower() == "c2":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        explicit_c2_pairs = [
            ("50", "57"),
            ("51", "58"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_c2_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower().startswith(("c5", "c6", "c7")) and (category or "").lower() not in {"c5 40", "c6 35", "c6 40", "c7 40"}:
        def as_match_number(node: Dict[str, object]) -> Optional[int]:
            text = str(node["text"]).strip()
            digits = text.replace(".", "", 1)
            if not digits.isdigit():
                return None
            return int(float(text))

        nodes_by_number: Dict[int, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_number = as_match_number(node)
            if match_number is None:
                continue
            nodes_by_number.setdefault(match_number, []).append(node)

        for number in nodes_by_number:
            nodes_by_number[number].sort(key=lambda item: int(item["row"]))

        connector_pairs = [
            (left_node, right_node)
            for left_node, right_node in connector_pairs
            if not (
                _match_stage(str(left_node["text"])) == 0
                and _match_stage(str(right_node["text"])) == 1
            )
        ]

        for idx, target in enumerate(range(50, 58)):
            left_a = 34 + idx * 2
            left_b = left_a + 1
            target_nodes = nodes_by_number.get(target, [])
            left_a_nodes = nodes_by_number.get(left_a, [])
            left_b_nodes = nodes_by_number.get(left_b, [])
            if not target_nodes or not left_a_nodes or not left_b_nodes:
                continue
            target_node = target_nodes[0]
            connector_pairs.append((left_a_nodes[0], target_node))
            connector_pairs.append((left_b_nodes[0], target_node))

    if (category or "").lower() == "c5 40":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []

        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        explicit_c5_40_pairs = [
            ("49", "57"),
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("53", "59"),
            ("54", "59"),
            ("55", "60"),
            ("56", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_c5_40_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "d5":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        explicit_d5_pairs = [
            ("50", "57"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_d5_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "d2":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_d2_pairs = [
            ("58", "61"),
            ("59", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_d2_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "d4":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        explicit_d4_pairs = [
            ("50", "57"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_d4_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "d6":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_d6_pairs = [
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("53", "59"),
            ("54", "59"),
            ("55", "60"),
            ("56", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_d6_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "d7":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_d7_pairs = [
            ("50", "57"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_d7_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c6 35":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return
            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1
            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_c6_35_pairs = [
            ("34", "49"),
            ("35", "52"),
            ("36", "53"),
            ("37", "56"),
            ("49", "57"),
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("53", "59"),
            ("54", "59"),
            ("55", "60"),
            ("56", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_c6_35_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c6 40":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_c6_40_pairs = [
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_c6_40_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c6 45":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return
            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1
            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_c6_45_pairs = [
            ("50", "57"),
            ("51", "58"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_c6_45_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c7 45":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return
            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1
            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_c7_45_pairs = [
            ("50", "57"),
            ("51", "58"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_c7_45_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c7 40":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        final_target = "63" if "63" in match_nodes_by_number else "64"
        explicit_c7_40_pairs = [
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", final_target),
            ("62", final_target),
        ]

        for left_number, right_number in explicit_c7_40_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c3":
        match_nodes_by_number: Dict[str, List[Dict[str, object]]] = {}
        for node in match_nodes:
            match_nodes_by_number.setdefault(str(node["text"]).strip(), []).append(node)

        for key in match_nodes_by_number:
            match_nodes_by_number[key].sort(key=lambda item: int(item["row"]))

        connector_pairs = []
        used_left_by_number: Dict[str, int] = {}

        def connect_by_number(left_number: str, right_number: str) -> None:
            left_candidates = match_nodes_by_number.get(left_number, [])
            right_candidates = match_nodes_by_number.get(right_number, [])
            if not left_candidates or not right_candidates:
                return

            left_index = used_left_by_number.get(left_number, 0)
            if left_index >= len(left_candidates):
                left_index = len(left_candidates) - 1

            left_node = left_candidates[left_index]
            right_node = right_candidates[0]
            used_left_by_number[left_number] = left_index + 1
            connector_pairs.append((left_node, right_node))

        explicit_c3_pairs = [
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_c3_pairs:
            connect_by_number(left_number, right_number)

    return connector_pairs


def _find_direct_team_center(
    nodes: List[Dict[str, object]],
    target_node: Dict[str, object],
    legend_start_row: Optional[int],
) -> Optional[int]:
    target_col = int(target_node["col"])
    target_row = int(target_node["row"])
    nearby_teams = sorted(
        [
            node
            for node in nodes
            if node["class"] == "team"
            and int(node["col"]) == target_col
            and abs(int(node["row"]) - target_row) <= 8
            and (legend_start_row is None or int(node["row"]) < legend_start_row)
        ],
        key=lambda node: int(node["row"]),
    )
    if len(nearby_teams) < 2:
        return None

    best_center: Optional[int] = None
    best_distance: Optional[float] = None
    for idx in range(len(nearby_teams) - 1):
        top_team = nearby_teams[idx]
        bottom_team = nearby_teams[idx + 1]
        row_gap = abs(int(bottom_team["row"]) - int(top_team["row"]))
        if row_gap > 2:
            continue

        center_y = round((int(top_team["y"]) + int(bottom_team["y"])) / 2) + 12
        center_row = (int(top_team["row"]) + int(bottom_team["row"])) / 2
        distance = abs(center_row - target_row)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_center = center_y

    return best_center


def _align_match_nodes(
    nodes: List[Dict[str, object]],
    category: Optional[str] = None,
) -> None:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    connector_pairs = _compute_connector_pairs(nodes=nodes, category=category)
    incoming_by_target: Dict[int, List[Dict[str, object]]] = {}
    for left_node, right_node in connector_pairs:
        incoming_by_target.setdefault(id(right_node), []).append(left_node)

    staged_targets: Dict[int, List[Dict[str, object]]] = {}
    for node in nodes:
        if node["class"] != "match-id":
            continue
        if legend_start_row is not None and int(node["row"]) >= legend_start_row:
            continue
        stage = _match_stage(str(node["text"]))
        if stage is None or stage == 0:
            continue
        staged_targets.setdefault(stage, []).append(node)

    for stage in sorted(staged_targets):
        for target_node in sorted(staged_targets[stage], key=lambda item: int(item["row"])):
            candidate_centers = [
                int(source_node["y"]) + 12
                for source_node in incoming_by_target.get(id(target_node), [])
            ]

            direct_center = _find_direct_team_center(
                nodes=nodes,
                target_node=target_node,
                legend_start_row=legend_start_row,
            )
            if direct_center is not None:
                candidate_centers.append(direct_center)

            if not candidate_centers:
                continue

            aligned_center = round(sum(candidate_centers) / len(candidate_centers))
            target_node["y"] = aligned_center - 12

    if (category or "").lower() == "c2":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        c2_seed_pairs = {
            "50": ("9°", "8°"),
            "51": ("5°", "12°"),
            "54": ("11°", "6°"),
            "55": ("7°", "10°"),
        }

        for match_number, (seed_a, seed_b) in c2_seed_pairs.items():
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id"
                    and str(node["text"]).strip() == match_number
                    and (legend_start_row is None or int(node["row"]) < legend_start_row)
                ),
                None,
            )
            if target_node is None:
                continue

            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if center_a is None or center_b is None:
                continue

            target_node["y"] = round((center_a + center_b) / 2) - 12

        # Fine-tune 58/59 so they sit midway between prior-match winner and fixed opponent.
        c2_midpoint_pairs = {
            "58": "51",
            "59": "54",
        }

        for target_match, source_match in c2_midpoint_pairs.items():
            target_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id"
                    and str(node["text"]).strip() == target_match
                    and (legend_start_row is None or int(node["row"]) < legend_start_row)
                ),
                None,
            )
            source_node = next(
                (
                    node
                    for node in nodes
                    if node["class"] == "match-id"
                    and str(node["text"]).strip() == source_match
                    and (legend_start_row is None or int(node["row"]) < legend_start_row)
                ),
                None,
            )
            if target_node is None or source_node is None:
                continue

            direct_center = _find_direct_team_center(
                nodes=nodes,
                target_node=target_node,
                legend_start_row=legend_start_row,
            )
            if direct_center is None:
                continue

            source_center = int(source_node["y"]) + 12
            target_node["y"] = round((source_center + direct_center) / 2) - 12

    if (category or "").lower() == "c3":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        c3_seed_pairs = {
            "50": ("2° C", "2° F"),
            "51": ("1° E", "2° B"),
            "52": ("3° A", "1° D"),
            "54": ("2° A", "1° F"),
            "55": ("2° E", "2° D"),
        }

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        for match_number, (seed_a, seed_b) in c3_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_51 = match_nodes_by_number.get("51")
        match_52 = match_nodes_by_number.get("52")
        match_58 = match_nodes_by_number.get("58")
        if match_58 is not None and match_51 is not None and match_52 is not None:
            center_51 = int(match_51["y"]) + 12
            center_52 = int(match_52["y"]) + 12
            match_58["y"] = round((center_51 + center_52) / 2) - 12

        match_50 = match_nodes_by_number.get("50")

        match_57 = match_nodes_by_number.get("57")
        center_1a = seed_centers.get("1° A")
        if match_57 is not None and match_50 is not None and center_1a is not None:
            center_50 = int(match_50["y"]) + 12
            match_57["y"] = round((center_1a + center_50) / 2) - 12

        match_58 = match_nodes_by_number.get("58")
        match_61 = match_nodes_by_number.get("61")
        if match_61 is not None and match_57 is not None and match_58 is not None:
            center_57 = int(match_57["y"]) + 12
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_57 + center_58) / 2) - 12

        match_59 = match_nodes_by_number.get("59")
        match_60 = match_nodes_by_number.get("60")
        match_62 = match_nodes_by_number.get("62")
        if match_62 is not None and match_59 is not None and match_60 is not None:
            center_59 = int(match_59["y"]) + 12
            center_60 = int(match_60["y"]) + 12
            match_62["y"] = round((center_59 + center_60) / 2) - 12

    if (category or "").lower() == "d5":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        d5_pairs = {
            "50": ("2° B", "2° C"),
            "54": ("3° A", "2° E"),
            "55": ("2° D", "2° A"),
            "57": ("1° A", "50"),
            "58": ("1° E", "1° D"),
            "59": ("1° C", "54"),
            "60": ("55", "1° B"),
            "61": ("57", "58"),
            "62": ("59", "60"),
            "64": ("61", "62"),
        }

        for match_number, (left_ref, right_ref) in d5_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            if target_node is None:
                continue
            left_center = seed_centers.get(left_ref)
            if left_center is None and left_ref in match_nodes_by_number:
                left_center = int(match_nodes_by_number[left_ref]["y"]) + 12
            right_center = seed_centers.get(right_ref)
            if right_center is None and right_ref in match_nodes_by_number:
                right_center = int(match_nodes_by_number[right_ref]["y"]) + 12
            if left_center is None or right_center is None:
                continue
            target_node["y"] = round((left_center + right_center) / 2) - 12

    if (category or "").lower() == "d2":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        match_58 = match_nodes_by_number.get("58")
        center_3a = seed_centers.get("3° A")
        center_2b = seed_centers.get("2° B")
        if match_58 is not None and center_3a is not None and center_2b is not None:
            match_58["y"] = round((center_3a + center_2b) / 2) - 12

        match_59 = match_nodes_by_number.get("59")
        center_2a = seed_centers.get("2° A")
        center_3b = seed_centers.get("3° B")
        if match_59 is not None and center_2a is not None and center_3b is not None:
            match_59["y"] = round((center_2a + center_3b) / 2) - 12

        match_61 = match_nodes_by_number.get("61")
        center_1a = seed_centers.get("1° A")
        if match_61 is not None and match_58 is not None and center_1a is not None:
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_1a + center_58) / 2) - 12

        match_62 = match_nodes_by_number.get("62")
        center_1b = seed_centers.get("1° B")
        if match_62 is not None and match_59 is not None and center_1b is not None:
            center_59 = int(match_59["y"]) + 12
            match_62["y"] = round((center_59 + center_1b) / 2) - 12

        match_64 = match_nodes_by_number.get("64")
        if match_64 is not None and match_61 is not None and match_62 is not None:
            center_61 = int(match_61["y"]) + 12
            center_62 = int(match_62["y"]) + 12
            match_64["y"] = round((center_61 + center_62) / 2) - 12

        # Add horizontal breathing room so 58/59 are not visually stuck to 61/62.
        for match_number, x_delta in {"61": 40, "62": 40, "64": 76}.items():
            node = match_nodes_by_number.get(match_number)
            if node is not None:
                node["x"] = int(node["x"]) + x_delta

    if (category or "").lower() == "d4":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        def find_d4_playing_pair_center(target_node: Dict[str, object]) -> Optional[int]:
            target_row = int(target_node["row"])
            target_x = int(target_node["x"])
            nearby_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and (legend_start_row is None or int(node["row"]) < legend_start_row)
                    and abs(int(node["row"]) - target_row) <= 10
                    and abs(int(node["x"]) - target_x) <= 180
                ],
                key=lambda node: int(node["row"]),
            )
            if len(nearby_teams) < 2:
                return None

            best_center: Optional[int] = None
            best_score: Optional[float] = None
            for idx in range(len(nearby_teams) - 1):
                top_team = nearby_teams[idx]
                bottom_team = nearby_teams[idx + 1]
                row_gap = abs(int(bottom_team["row"]) - int(top_team["row"]))
                if row_gap > 2:
                    continue

                center_y = round((int(top_team["y"]) + int(bottom_team["y"])) / 2) + 12
                center_row = (int(top_team["row"]) + int(bottom_team["row"])) / 2
                avg_x = (int(top_team["x"]) + int(bottom_team["x"])) / 2
                score = abs(center_row - target_row) + (abs(avg_x - target_x) / 200)
                if best_score is None or score < best_score:
                    best_score = score
                    best_center = center_y

            return best_center

        match_50 = match_nodes_by_number.get("50")
        center_2b = seed_centers.get("2° B")
        center_3a = seed_centers.get("3° A")
        if match_50 is not None and center_2b is not None and center_3a is not None:
            match_50["y"] = round((center_2b + center_3a) / 2) - 12

        match_57 = match_nodes_by_number.get("57")
        center_1a = seed_centers.get("1° A")
        if match_57 is not None and match_50 is not None and center_1a is not None:
            center_50 = int(match_50["y"]) + 12
            match_57["y"] = round((center_1a + center_50) / 2) - 12

        match_58 = match_nodes_by_number.get("58")
        center_1d = seed_centers.get("1° D")
        center_2c = seed_centers.get("2° C")
        if match_58 is not None and center_2c is not None and center_1d is not None:
            match_58["y"] = round((center_2c + center_1d) / 2) - 12

        match_59 = match_nodes_by_number.get("59")
        center_1c = seed_centers.get("1° C")
        center_2d = seed_centers.get("2° D")
        if match_59 is not None and center_1c is not None and center_2d is not None:
            match_59["y"] = round((center_1c + center_2d) / 2) - 12

        match_60 = match_nodes_by_number.get("60")
        center_1b = seed_centers.get("1° B")
        center_2a = seed_centers.get("2° A")
        if match_60 is not None and center_2a is not None and center_1b is not None:
            match_60["y"] = round((center_2a + center_1b) / 2) - 12

    if (category or "").lower().startswith(("c5", "c6", "c7")) and (category or "").lower() not in {"c5 40", "c6 35", "c6 40", "c7 40"}:
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        for target in range(50, 56):
            feeder_a = str(34 + (target - 50) * 2)
            feeder_b = str(35 + (target - 50) * 2)
            target_node = match_nodes_by_number.get(str(target))
            feeder_a_node = match_nodes_by_number.get(feeder_a)
            feeder_b_node = match_nodes_by_number.get(feeder_b)
            if target_node is None or feeder_a_node is None or feeder_b_node is None:
                continue

            center_a = int(feeder_a_node["y"]) + 12
            center_b = int(feeder_b_node["y"]) + 12
            target_node["y"] = round((center_a + center_b) / 2) - 12

    if (category or "").lower() == "c5 40":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_50 = next(
            (
                node
                for node in nodes
                if node["class"] == "match-id"
                and str(node["text"]).strip() == "50"
                and (legend_start_row is None or int(node["row"]) < legend_start_row)
            ),
            None,
        )
        center_2f = seed_centers.get("2° F")
        center_2g = seed_centers.get("2° G")
        if match_50 is not None and center_2f is not None and center_2g is not None:
            match_50["y"] = round((center_2f + center_2g) / 2) - 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c5_40_seed_pairs = {
            "49": ("1° A", "2° B"),
            "50": ("2° G", "1° H"),
            "51": ("1° E", "2° F"),
            "52": ("2° C", "1° D"),
            "53": ("1° C", "2° D"),
            "54": ("2° E", "1° F"),
            "55": ("1° G", "2° H"),
            "56": ("2° A", "1° B"),
        }

        for match_number, (seed_a, seed_b) in c5_40_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_57 = match_nodes_by_number.get("57")
        match_49 = match_nodes_by_number.get("49")
        match_50 = match_nodes_by_number.get("50")
        if match_57 is not None and match_49 is not None and match_50 is not None:
            center_49 = int(match_49["y"]) + 12
            center_50 = int(match_50["y"]) + 12
            match_57["y"] = round((center_49 + center_50) / 2) - 12

        match_58 = match_nodes_by_number.get("58")
        match_51 = match_nodes_by_number.get("51")
        match_52 = match_nodes_by_number.get("52")
        if match_58 is not None and match_51 is not None and match_52 is not None:
            center_51 = int(match_51["y"]) + 12
            center_52 = int(match_52["y"]) + 12
            match_58["y"] = round((center_51 + center_52) / 2) - 12

        match_59 = match_nodes_by_number.get("59")
        match_53 = match_nodes_by_number.get("53")
        match_54 = match_nodes_by_number.get("54")
        if match_59 is not None and match_53 is not None and match_54 is not None:
            center_53 = int(match_53["y"]) + 12
            center_54 = int(match_54["y"]) + 12
            match_59["y"] = round((center_53 + center_54) / 2) - 12

        match_60 = match_nodes_by_number.get("60")
        match_55 = match_nodes_by_number.get("55")
        match_56 = match_nodes_by_number.get("56")
        if match_60 is not None and match_55 is not None and match_56 is not None:
            center_55 = int(match_55["y"]) + 12
            center_56 = int(match_56["y"]) + 12
            match_60["y"] = round((center_55 + center_56) / 2) - 12

        match_61 = match_nodes_by_number.get("61")
        if match_61 is not None and match_57 is not None and match_58 is not None:
            center_57 = int(match_57["y"]) + 12
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_57 + center_58) / 2) - 12

        match_62 = match_nodes_by_number.get("62")
        if match_62 is not None and match_59 is not None and match_60 is not None:
            center_59 = int(match_59["y"]) + 12
            center_60 = int(match_60["y"]) + 12
            match_62["y"] = round((center_59 + center_60) / 2) - 12

        match_64 = match_nodes_by_number.get("64")
        if match_64 is not None and match_61 is not None and match_62 is not None:
            center_61 = int(match_61["y"]) + 12
            center_62 = int(match_62["y"]) + 12
            match_64["y"] = round((center_61 + center_62) / 2) - 12

    if (category or "").lower() == "c6 40":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c6_40_seed_pairs = {
            "57": ("1° A", "2° B"),
            "58": ("1° D", "2° C"),
            "59": ("1° C", "2° D"),
            "60": ("2° A", "1° B"),
        }

        for match_number, (seed_a, seed_b) in c6_40_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_57 = match_nodes_by_number.get("57")
        match_58 = match_nodes_by_number.get("58")
        match_59 = match_nodes_by_number.get("59")
        match_60 = match_nodes_by_number.get("60")

        match_61 = match_nodes_by_number.get("61")
        if match_61 is not None and match_57 is not None and match_58 is not None:
            center_57 = int(match_57["y"]) + 12
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_57 + center_58) / 2) - 12

        match_62 = match_nodes_by_number.get("62")
        if match_62 is not None and match_59 is not None and match_60 is not None:
            center_59 = int(match_59["y"]) + 12
            center_60 = int(match_60["y"]) + 12
            match_62["y"] = round((center_59 + center_60) / 2) - 12

        final_match_c6_40 = match_nodes_by_number.get("63") or match_nodes_by_number.get("64")
        if final_match_c6_40 is not None and match_61 is not None and match_62 is not None:
            center_61 = int(match_61["y"]) + 12
            center_62 = int(match_62["y"]) + 12
            final_match_c6_40["y"] = round((center_61 + center_62) / 2) - 12

    if (category or "").lower() == "c7 40":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c7_40_seed_pairs = {
            "57": ("1° A", "2° B"),
            "58": ("1° D", "2° C"),
            "59": ("1° C", "2° D"),
            "60": ("2° A", "1° B"),
        }

        for match_number, (seed_a, seed_b) in c7_40_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_57 = match_nodes_by_number.get("57")
        match_58 = match_nodes_by_number.get("58")
        match_59 = match_nodes_by_number.get("59")
        match_60 = match_nodes_by_number.get("60")

        match_61 = match_nodes_by_number.get("61")
        if match_61 is not None and match_57 is not None and match_58 is not None:
            center_57 = int(match_57["y"]) + 12
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_57 + center_58) / 2) - 12

        match_62 = match_nodes_by_number.get("62")
        if match_62 is not None and match_59 is not None and match_60 is not None:
            center_59 = int(match_59["y"]) + 12
            center_60 = int(match_60["y"]) + 12
            match_62["y"] = round((center_59 + center_60) / 2) - 12

        final_match_c7_40 = match_nodes_by_number.get("63") or match_nodes_by_number.get("64")
        if final_match_c7_40 is not None and match_61 is not None and match_62 is not None:
            center_61 = int(match_61["y"]) + 12
            center_62 = int(match_62["y"]) + 12
            final_match_c7_40["y"] = round((center_61 + center_62) / 2) - 12

    if (category or "").lower() == "d6":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        d6_seed_pairs = {
            "50": ("2° F", "2° G"),
            "51": ("1° E", "2° C"),
            "52": ("2° B", "1° D"),
            "53": ("1° C", "2° A"),
            "54": ("2° D", "1° F"),
            "55": ("1° G", "2° E"),
            "56": ("3° A", "1° B"),
        }

        for match_number, (seed_a, seed_b) in d6_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            center_b = seed_centers.get(seed_b)
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_57 = match_nodes_by_number.get("57")
        match_50 = match_nodes_by_number.get("50")
        center_1a = seed_centers.get("1° A")
        if match_57 is not None and match_50 is not None and center_1a is not None:
            center_50 = int(match_50["y"]) + 12
            match_57["y"] = round((center_1a + center_50) / 2) - 12

        match_58 = match_nodes_by_number.get("58")
        match_51 = match_nodes_by_number.get("51")
        match_52 = match_nodes_by_number.get("52")
        if match_58 is not None and match_51 is not None and match_52 is not None:
            center_51 = int(match_51["y"]) + 12
            center_52 = int(match_52["y"]) + 12
            match_58["y"] = round((center_51 + center_52) / 2) - 12

        match_59 = match_nodes_by_number.get("59")
        match_53 = match_nodes_by_number.get("53")
        match_54 = match_nodes_by_number.get("54")
        if match_59 is not None and match_53 is not None and match_54 is not None:
            center_53 = int(match_53["y"]) + 12
            center_54 = int(match_54["y"]) + 12
            match_59["y"] = round((center_53 + center_54) / 2) - 12

        match_60 = match_nodes_by_number.get("60")
        match_55 = match_nodes_by_number.get("55")
        match_56 = match_nodes_by_number.get("56")
        if match_60 is not None and match_55 is not None and match_56 is not None:
            center_55 = int(match_55["y"]) + 12
            center_56 = int(match_56["y"]) + 12
            match_60["y"] = round((center_55 + center_56) / 2) - 12

        match_61 = match_nodes_by_number.get("61")
        if match_61 is not None and match_57 is not None and match_58 is not None:
            center_57 = int(match_57["y"]) + 12
            center_58 = int(match_58["y"]) + 12
            match_61["y"] = round((center_57 + center_58) / 2) - 12

        match_62 = match_nodes_by_number.get("62")
        if match_62 is not None and match_59 is not None and match_60 is not None:
            center_59 = int(match_59["y"]) + 12
            center_60 = int(match_60["y"]) + 12
            match_62["y"] = round((center_59 + center_60) / 2) - 12

        final_match = match_nodes_by_number.get("63") or match_nodes_by_number.get("64")
        if final_match is not None and match_61 is not None and match_62 is not None:
            center_61 = int(match_61["y"]) + 12
            center_62 = int(match_62["y"]) + 12
            final_match["y"] = round((center_61 + center_62) / 2) - 12

    if (category or "").lower() == "d7":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        d7_pairs = {
            "50": ("2° B", "2° C"),
            "55": ("2° D", "2° A"),
            "57": ("1° A", "50"),
            "58": ("1° E", "1° D"),
            "59": ("1° C", "2° E"),
            "60": ("55", "1° B"),
            "61": ("57", "58"),
            "62": ("59", "60"),
            "64": ("61", "62"),
        }

        for match_number, (left_ref, right_ref) in d7_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            if target_node is None:
                continue
            left_center = seed_centers.get(left_ref)
            if left_center is None and left_ref in match_nodes_by_number:
                left_center = int(match_nodes_by_number[left_ref]["y"]) + 12
            right_center = seed_centers.get(right_ref)
            if right_center is None and right_ref in match_nodes_by_number:
                right_center = int(match_nodes_by_number[right_ref]["y"]) + 12
            if left_center is None or right_center is None:
                continue
            target_node["y"] = round((left_center + right_center) / 2) - 12

    if (category or "").lower() == "c6 35":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c6_35_seed_pairs = {
            "34": ("2° C", "2° F"),
            "35": ("2° G", "2° B"),
            "36": ("2° A", "2° H"),
            "37": ("2° E", "2° D"),
            "49": ("1° A", "34"),
            "50": ("1° I", "1° H"),
            "51": ("1° E", "2° J"),
            "52": ("35", "1° D"),
            "53": ("1° C", "36"),
            "54": ("2° I", "1° F"),
            "55": ("1° G", "1° J"),
            "56": ("37", "1° B"),
        }

        for match_number, (seed_a, seed_b) in c6_35_seed_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            center_a = seed_centers.get(seed_a)
            if center_a is None and seed_a in match_nodes_by_number:
                center_a = int(match_nodes_by_number[seed_a]["y"]) + 12
            center_b = seed_centers.get(seed_b)
            if center_b is None and seed_b in match_nodes_by_number:
                center_b = int(match_nodes_by_number[seed_b]["y"]) + 12
            if target_node is None or center_a is None or center_b is None:
                continue
            target_node["y"] = round((center_a + center_b) / 2) - 12

        match_49 = match_nodes_by_number.get("49")
        match_57 = match_nodes_by_number.get("57")
        match_50 = match_nodes_by_number.get("50")
        if match_57 is not None and match_49 is not None and match_50 is not None:
            center_49 = int(match_49["y"]) + 12
            center_50 = int(match_50["y"]) + 12
            match_57["y"] = round((center_49 + center_50) / 2) - 12


    if (category or "").lower() == "c6 45":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c6_45_pairs = {
            "50": ("2° B", "2° C"),
            "51": ("1° E", "3° B"),
            "54": ("3° A", "2° E"),
            "55": ("2° D", "2° A"),
            "57": ("1° A", "50"),
            "58": ("51", "1° D"),
            "59": ("1° C", "54"),
            "60": ("55", "1° B"),
            "61": ("57", "58"),
            "62": ("59", "60"),
            "64": ("61", "62"),
        }

        for match_number, (left_ref, right_ref) in c6_45_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            if target_node is None:
                continue
            left_center = seed_centers.get(left_ref)
            if left_center is None and left_ref in match_nodes_by_number:
                left_center = int(match_nodes_by_number[left_ref]["y"]) + 12
            right_center = seed_centers.get(right_ref)
            if right_center is None and right_ref in match_nodes_by_number:
                right_center = int(match_nodes_by_number[right_ref]["y"]) + 12
            if left_center is None or right_center is None:
                continue
            target_node["y"] = round((left_center + right_center) / 2) - 12

    if (category or "").lower() == "c7 45":
        seed_centers: Dict[str, int] = {}
        for node in nodes:
            if node["class"] not in {"seed", "seed-between"}:
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            seed_centers[str(node["text"]).strip()] = int(node["y"]) + 12

        match_nodes_by_number: Dict[str, Dict[str, object]] = {}
        for node in nodes:
            if node["class"] != "match-id":
                continue
            if legend_start_row is not None and int(node["row"]) >= legend_start_row:
                continue
            match_nodes_by_number[str(node["text"]).strip()] = node

        c7_45_pairs = {
            "50": ("2° C", "2° F"),
            "51": ("1° E", "2° B"),
            "54": ("2° A", "1° F"),
            "55": ("2° E", "2° D"),
            "57": ("1° A", "50"),
            "58": ("51", "1° D"),
            "59": ("1° C", "54"),
            "60": ("55", "1° B"),
            "61": ("57", "58"),
            "62": ("59", "60"),
            "64": ("61", "62"),
        }

        for match_number, (left_ref, right_ref) in c7_45_pairs.items():
            target_node = match_nodes_by_number.get(match_number)
            if target_node is None:
                continue
            left_center = seed_centers.get(left_ref)
            if left_center is None and left_ref in match_nodes_by_number:
                left_center = int(match_nodes_by_number[left_ref]["y"]) + 12
            right_center = seed_centers.get(right_ref)
            if right_center is None and right_ref in match_nodes_by_number:
                right_center = int(match_nodes_by_number[right_ref]["y"]) + 12
            if left_center is None or right_center is None:
                continue
            target_node["y"] = round((left_center + right_center) / 2) - 12


def _build_round_labels(nodes: List[Dict[str, object]], category: Optional[str] = None) -> str:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    match_cols = sorted(
        {
            int(node["col"])
            for node in nodes
            if node["class"] == "match-id"
            and (legend_start_row is None or int(node["row"]) < legend_start_row)
        }
    )
    if not match_cols:
        return ""

    stage_titles = {
        0: "16avos",
        1: "8vos",
        2: "4tos",
        3: "Semi",
        4: "Final",
    }

    stage_counts_by_col: Dict[int, Dict[int, int]] = {}
    for node in nodes:
        if node["class"] != "match-id":
            continue
        if legend_start_row is not None and int(node["row"]) >= legend_start_row:
            continue
        stage = _match_stage(str(node["text"]))
        if stage is None:
            continue
        col = int(node["col"])
        stage_counts_by_col.setdefault(col, {})
        stage_counts_by_col[col][stage] = stage_counts_by_col[col].get(stage, 0) + 1

    stage_order_from_final = ["Final", "Semi", "4tos", "8vos", "16avos", "32avos", "64avos"]
    total_rounds = len(match_cols)
    round_names: List[str] = []
    for idx in range(total_rounds):
        distance_from_final = total_rounds - 1 - idx
        if distance_from_final < len(stage_order_from_final):
            round_names.append(stage_order_from_final[distance_from_final])
        else:
            round_names.append(f"Ronda {idx + 1}")

    labels: List[str] = []
    for idx, col in enumerate(match_cols):
        stage_counts = stage_counts_by_col.get(col)
        if stage_counts:
            selected_stage = sorted(stage_counts.items(), key=lambda item: (-item[1], item[0]))[0][0]
            label = stage_titles.get(selected_stage, f"Ronda {idx + 1}")
        else:
            label = round_names[idx] if idx < len(round_names) else f"Ronda {idx + 1}"
        x = col * CELL_WIDTH
        labels.append(
            f'<div class="round-label" style="left:{x}px;">{escape(label)}</div>'
        )

    return "".join(labels)


def _build_matchups_html(nodes: List[Dict[str, object]], category: Optional[str] = None) -> str:
    connector_pairs = _compute_connector_pairs(nodes=nodes, category=category)
    if not connector_pairs:
        return ""

    stage_titles = {1: "8vos", 2: "4tos", 3: "Semi", 4: "Final"}
    grouped: Dict[str, Dict[str, object]] = {}

    for left_node, right_node in connector_pairs:
        right_text = str(right_node["text"])
        stage = _match_stage(right_text)
        if stage is None or stage == 0:
            continue
        key = f"{stage}:{right_text}:{int(right_node['row'])}"
        entry = grouped.setdefault(
            key,
            {
                "stage": stage,
                "target": right_text,
                "target_row": int(right_node["row"]),
                "sources": [],
            },
        )
        source_text = str(left_node["text"])
        if source_text not in entry["sources"]:
            entry["sources"].append(source_text)

    direct_sources_by_category: Dict[str, Dict[str, List[str]]] = {
        "d5": {
            "57": ["1° A"],
            "58": ["1° D"],
            "59": ["1° C"],
            "60": ["1° B"],
        },
        "d6": {
            "57": ["1° A"],
        },
        "d7": {
            "57": ["1° A"],
            "60": ["1° B"],
        }
    }

    for target_text, direct_sources in direct_sources_by_category.get((category or "").lower(), {}).items():
        target_node = next(
            (
                node
                for node in nodes
                if node["class"] == "match-id" and str(node["text"]).strip() == target_text
            ),
            None,
        )
        if target_node is None:
            continue

        stage = _match_stage(target_text)
        if stage is None or stage == 0:
            continue

        key = f"{stage}:{target_text}:{int(target_node['row'])}"
        entry = grouped.setdefault(
            key,
            {
                "stage": stage,
                "target": target_text,
                "target_row": int(target_node["row"]),
                "sources": [],
            },
        )
        for source_text in direct_sources:
            if source_text not in entry["sources"]:
                entry["sources"].append(source_text)

    if not grouped:
        return ""

    def format_source_label(source: str) -> str:
        text = escape(source)
        return text if "°" in source else f"#{text}"

    sections: Dict[int, List[str]] = {}
    for entry in sorted(grouped.values(), key=lambda item: (int(item["stage"]), int(item["target_row"]))):
        stage = int(entry["stage"])
        target = escape(str(entry["target"]))
        sources = [format_source_label(src) for src in entry["sources"]]
        if len(sources) >= 2:
            pair_text = f"{sources[0]} <span>vs</span> {sources[1]}"
        elif len(sources) == 1:
            pair_text = f"{sources[0]} <span>vs</span> ingreso directo"
        else:
            pair_text = "por definir"

        item_html = (
            f'<div class="stage-matchup-item">'
            f'<div class="stage-matchup-head">Partido #{target}</div>'
            f'<div class="stage-matchup-pair">{pair_text}</div>'
            f'</div>'
        )
        sections.setdefault(stage, []).append(item_html)

    html_sections: List[str] = []
    for stage in sorted(sections):
        title = stage_titles.get(stage, f"Ronda {stage}")
        html_sections.append(
            f'<div class="stage-matchups-group"><div class="stage-matchups-title">{escape(title)}</div><div class="stage-matchups-grid">{"".join(sections[stage])}</div></div>'
        )

    return f'<div class="matchups-box"><div class="matchups-title">Cruces claros</div>{"".join(html_sections)}</div>'


def _build_matchup_guides_svg(
    nodes: List[Dict[str, object]],
    category: Optional[str] = None,
) -> str:
    if (category or "").lower() in {"c2", "c6 40", "c7 40"}:
        return ""

    matches = [
        node
        for node in nodes
        if node["class"] == "match-id"
    ]
    if not matches:
        return ""

    first_round_col = min(int(node["col"]) for node in matches)
    first_round_matches = sorted(
        [node for node in matches if int(node["col"]) == first_round_col],
        key=lambda node: int(node["row"]),
    )
    teams = [
        node
        for node in nodes
        if node["class"] == "team"
        and int(node["col"]) == first_round_col
    ]

    guides: List[str] = []
    for match in first_round_matches:
        match_row = int(match["row"])
        candidate_teams = sorted(
            [team for team in teams],
            key=lambda team: int(team["row"]),
        )
        upper_teams = [team for team in candidate_teams if int(team["row"]) <= match_row][-2:]
        lower_teams = [team for team in candidate_teams if int(team["row"]) > match_row][:2]

        if (category or "").lower() == "d2" and (not upper_teams or not lower_teams):
            continue

        nearby_participants = sorted(upper_teams + lower_teams, key=lambda team: int(team["row"]))

        if len(nearby_participants) < 2:
            candidate_seeds = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] in {"seed", "seed-between"}
                    and abs(int(node["row"]) - match_row) <= 8
                ],
                key=lambda node: int(node["row"]),
            )
            upper_seeds = [node for node in candidate_seeds if int(node["row"]) <= match_row][-2:]
            lower_seeds = [node for node in candidate_seeds if int(node["row"]) > match_row][:2]
            nearby_participants = sorted(upper_seeds + lower_seeds, key=lambda node: int(node["row"]))

        if len(nearby_participants) < 2:
            continue

        # Find seeds on the same rows as teams
        team_rows = {int(t["row"]) for t in nearby_participants}
        nearby_seeds = [
            n
            for n in nodes
            if n["class"] in {"seed", "seed-between"}
            and (
                int(n["row"]) in team_rows
                or int(n["row"]) in {row - 1 for row in team_rows}
                or int(n["row"]) in {row + 1 for row in team_rows}
            )
        ]
        
        # Bounding box from seeds + participants + match
        all_nodes = nearby_seeds + nearby_participants + [match]
        min_x = min(int(n["x"]) for n in all_nodes)
        min_y = min(int(n["y"]) for n in all_nodes)
        max_y = max(int(n["y"]) + CELL_HEIGHT for n in all_nodes)

        # Extend box to the left to fully cover seeds
        box_x = min_x - 50
        box_y = min_y - 6
        box_w = 280   # Wide enough from left seed edge to match
        box_h = max_y - min_y + 12

        guides.append(
            f'<rect x="{box_x}" y="{box_y}" width="{box_w}" height="{box_h}" rx="8" class="matchup-guide"></rect>'
        )

    return "".join(guides)

LABEL_ROW_HEIGHT = 28


def _apply_results_to_nodes(
    nodes: List[Dict[str, object]],
    results: Dict[Tuple[str, str], Dict[str, str]],
    category: str,
) -> None:
    """Sustituye referencias a partidos (#N / 'espera ganador del partido N')
    con el nombre real del ganador cuando el resultado ya está cargado."""
    if not results:
        return

    cat_key = (category or "").lower()

    for node in nodes:
        if node.get("class") not in {"team"}:
            continue

        text: str = str(node.get("display_text", node.get("text", "")))

        def _replace_hash(m: re.Match) -> str:  # type: ignore[type-arg]
            num = m.group(1)
            r = results.get((cat_key, num))
            return r["ganador"] if r and r.get("ganador") else m.group(0)

        new_text = re.sub(r"#(\d+)", _replace_hash, text)

        def _replace_espera(m: re.Match) -> str:  # type: ignore[type-arg]
            team = m.group(1).strip()
            num = m.group(2)
            r = results.get((cat_key, num))
            if r and r.get("ganador"):
                return f"{team} vs {r['ganador']}"
            return m.group(0)

        new_text = re.sub(
            r"(.+?)\s*\(espera ganador del partido\s+(\d+)\)",
            _replace_espera,
            new_text,
            flags=re.IGNORECASE,
        )

        node["display_text"] = new_text


def render_bracket(
    grid: pd.DataFrame,
    category: str,
    sheet_name: str,
    results: Optional[Dict[Tuple[str, str], Dict[str, str]]] = None,
) -> None:
    rows, cols = grid.shape
    width = max(900, cols * CELL_WIDTH + 40)
    height = max(380, rows * CELL_HEIGHT + LABEL_ROW_HEIGHT + 20)
    board_height = height + 26

    legend_start_row: Optional[int] = None
    first_round_col: Optional[int] = None
    first_round_rows: set[int] = set()
    for row_idx in range(rows):
        for col_idx in range(cols):
            value = grid.iat[row_idx, col_idx]
            if pd.isna(value):
                continue
            text = _normalize_cell_text(str(value))
            if not text:
                continue
            cls = _cell_class(text)
            if cls == "legend":
                legend_start_row = row_idx if legend_start_row is None else min(legend_start_row, row_idx)

    for row_idx in range(rows):
        for col_idx in range(cols):
            value = grid.iat[row_idx, col_idx]
            if pd.isna(value):
                continue
            text = _normalize_cell_text(str(value))
            if not text:
                continue
            cls = _cell_class(text)
            if cls != "match-id":
                continue
            if legend_start_row is not None and row_idx >= legend_start_row:
                continue
            first_round_col = col_idx if first_round_col is None else min(first_round_col, col_idx)

    if first_round_col is not None:
        for row_idx in range(rows):
            value = grid.iat[row_idx, first_round_col]
            if pd.isna(value):
                continue
            text = _normalize_cell_text(str(value))
            if not text:
                continue
            if _cell_class(text) == "match-id":
                first_round_rows.add(row_idx)

    c3_top_seed_pos: Optional[tuple[int, int]] = None
    if (category or "").lower() == "c3":
        for row_idx in range(rows):
            for col_idx in range(cols):
                value = grid.iat[row_idx, col_idx]
                if pd.isna(value):
                    continue
                text = _normalize_cell_text(str(value))
                if text == "1° A":
                    c3_top_seed_pos = (row_idx, col_idx)
                    break
            if c3_top_seed_pos is not None:
                break

    match_schedule_labels: Dict[str, str] = _extract_match_schedule_labels(grid)

    cell_html: List[str] = []
    node_data: List[Dict[str, object]] = []
    for row_idx in range(rows):
        for col_idx in range(cols):
            value = grid.iat[row_idx, col_idx]
            if pd.isna(value):
                continue
            text = _normalize_cell_text(str(value))
            if text == "":
                continue

            cls = _cell_class(text)
            if (category or "").lower() == "c2" and cls == "match-id" and col_idx <= 2:
                if text.isdigit() and 1 <= int(text) <= 16:
                    cls = "seed"
                    text = f"{int(text)}°"
            display_text = text
            if cls == "match-id":
                try:
                    text = str(int(float(text)))
                except ValueError:
                    pass
                if legend_start_row is None or row_idx < legend_start_row:
                    display_text = match_schedule_labels.get(text, text)
            extra_class = ""
            top = row_idx * CELL_HEIGHT + LABEL_ROW_HEIGHT
            left = col_idx * CELL_WIDTH

            if first_round_col is not None:
                if cls == "team" and col_idx < first_round_col:
                    columns_away = first_round_col - col_idx
                    left = first_round_col * CELL_WIDTH - 108 - max(0, columns_away - 1) * 28
                elif cls == "match-id" and col_idx == first_round_col:
                    left = first_round_col * CELL_WIDTH

            if cls == "seed" and col_idx + 1 < cols and row_idx + 1 < rows:
                right_value = grid.iat[row_idx, col_idx + 1]
                below_right_value = grid.iat[row_idx + 1, col_idx + 1]
                right_text = _normalize_cell_text(str(right_value)) if pd.notna(right_value) else ""
                below_right_text = (
                    _normalize_cell_text(str(below_right_value)) if pd.notna(below_right_value) else ""
                )
                if right_text and below_right_text:
                    if _cell_class(right_text) == "team" and _cell_class(below_right_text) == "team":
                        top = row_idx * CELL_HEIGHT + CELL_HEIGHT // 2 + LABEL_ROW_HEIGHT
                        left = (col_idx + 1) * CELL_WIDTH - 44
                        cls = "seed-between"

            if (category or "").lower() == "c3" and c3_top_seed_pos is not None:
                top_seed_row, top_seed_col = c3_top_seed_pos
                if cls in {"seed", "seed-between"} and text == "1° A":
                    left -= 56
                if (
                    cls == "team"
                    and col_idx == top_seed_col + 1
                    and row_idx in {top_seed_row, top_seed_row + 1}
                ):
                    left -= 56

            safe_text = escape(display_text)
            node_data.append(
                {
                    "row": row_idx,
                    "col": col_idx,
                    "class": cls,
                    "x": left,
                    "y": top,
                    "text": text,
                    "display_text": display_text,
                }
            )
            cell_html.append(
                f'<div class="node {cls}{extra_class}" style="top:{top}px;left:{left}px;">{safe_text}</div>'
            )

    _align_match_nodes(nodes=node_data, category=category)
    _apply_results_to_nodes(nodes=node_data, results=results or {}, category=category)

    cat_key = (category or "").lower()
    cell_html = []
    for node in node_data:
        safe_text = escape(str(node.get("display_text", node["text"])))
        cell_html.append(
            f'<div class="node {node["class"]}" style="top:{int(node["y"])}px;left:{int(node["x"])}px;">{safe_text}</div>'
        )
        # If this match-id has a result, show winner badge below the node
        if node.get("class") == "match-id" and results:
            match_num = node.get("text", "")
            result = (results or {}).get((cat_key, match_num))
            if result and result.get("ganador"):
                ganador = escape(result["ganador"])
                marcador = escape(result.get("marcador", ""))
                badge_top = int(node["y"]) + CELL_HEIGHT
                badge_left = int(node["x"]) - 30
                marcador_html = f' <span class="result-score">{marcador}</span>' if marcador else ""
                cell_html.append(
                    f'<div class="node result-badge" style="top:{badge_top}px;left:{badge_left}px;">'
                    f'&#10003; {ganador}{marcador_html}</div>'
                )

    connectors_svg = _build_connectors(nodes=node_data, board_rows=rows, category=category)
    matchup_guides_svg = _build_matchup_guides_svg(nodes=node_data, category=category)
    logo_watermark_data_uri = load_logo_data_uri()
    watermark_html = ""
    if logo_watermark_data_uri:
        watermark_html = f'<div class="bracket-watermark"><img src="{logo_watermark_data_uri}" alt="Logo JPQ" /></div>'

    category_key = (category or "").lower()
    if category_key.startswith(("c4", "c5", "c6", "c7")):
        watermark_width_css = "100%"
        watermark_height_css = "100%"
        watermark_fit_css = "cover"
        watermark_opacity_css = "0.16"
    else:
        watermark_width_css = "100%"
        watermark_height_css = "100%"
        watermark_fit_css = "cover"
        watermark_opacity_css = "0.16"


    html = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes, minimum-scale=0.5, maximum-scale=4.0">
</head>
<body style="margin:0;padding:0;background:#0f1520;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#e8edf9;">
    <style>
      .bracket-wrap {{
                border: 1px solid rgba(255, 255, 255, 0.13);
                border-radius: 16px;
        overflow: hidden;
                padding: 14px;
            background: linear-gradient(140deg, rgba(28, 36, 54, 0.95) 0%, rgba(21, 28, 42, 0.98) 100%);
                box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.04);
        margin-bottom: 1rem;
      }}
      .bracket-board {{
        position: relative;
        width: {width}px;
                height: {board_height}px;
        transform-origin: top left;
      }}
            .bracket-watermark {{
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: -1;
                pointer-events: none;
            }}
            .bracket-watermark img {{
                width: {watermark_width_css};
                height: {watermark_height_css};
                object-fit: {watermark_fit_css};
                opacity: {watermark_opacity_css};
            }}
            .bracket-viewport {{
                position: relative;
                overflow: auto;
                width: 100%;
                -webkit-overflow-scrolling: touch;
                touch-action: pan-x pan-y pinch-zoom;
            }}
            .zoom-controls {{
                position: absolute;
                top: 8px;
                right: 8px;
                z-index: 20;
                display: flex;
                gap: 6px;
            }}
            .zoom-btn {{
                border: 1px solid rgba(255, 255, 255, 0.3);
                background: rgba(18, 26, 42, 0.82);
                color: #e8edf9;
                border-radius: 8px;
                padding: 4px 8px;
                font-size: 0.78rem;
                font-weight: 700;
                line-height: 1;
                cursor: pointer;
            }}
            .bracket-lines {{
                position: absolute;
                left: 0;
                top: 0;
                z-index: 0;
                pointer-events: none;
            }}
            .connector {{
                fill: none;
                stroke: rgba(92, 147, 255, 0.62);
                stroke-width: 2.4;
                stroke-linecap: round;
                stroke-linejoin: round;
            }}
            .matchup-guide {{
                fill: rgba(255, 255, 255, 0.06);
                stroke: rgba(255, 255, 255, 0.72);
                stroke-width: 2.4;
                filter: drop-shadow(0 0 5px rgba(255, 255, 255, 0.28));
            }}
            .round-label {{
                position: absolute;
                top: 0;
                padding: 3px 11px;
                border-radius: 999px;
                border: 1px solid rgba(214, 183, 99, 0.44);
                color: #f3d47a;
                background: rgba(214, 183, 99, 0.11);
                font-size: 0.75rem;
                font-weight: 800;
                letter-spacing: 0.1px;
            }}
            .round-labels-container {{
                position: sticky;
                top: 0;
                z-index: 10;
                height: {LABEL_ROW_HEIGHT}px;
            }}
      .node {{
        position: absolute;
                z-index: 1;
        min-width: 90px;
        max-width: 240px;
            padding: 3px 8px;
        line-height: 1.18;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
                border-radius: 9px;
        font-size: 0.84rem;
      }}
      .node.team {{
                border: 1px solid rgba(255, 255, 255, 0.16);
                background: rgba(21, 28, 42, 0.94);
                color: #f2f6ff;
      }}
      .node.seed {{
        font-weight: 600;
      }}
            .node.seed-between {{
                                font-weight: 700;
                                border: 1px solid rgba(214, 183, 99, 0.52);
                                color: #f8df95;
                                                                background: rgba(64, 52, 20, 0.92);
                        min-width: unset;
                        max-width: unset;
                        padding: 2px 7px;
            }}
      .node.match-id {{
        font-weight: 700;
                                border: 1px solid rgba(92, 147, 255, 0.62);
                                color: #8db7ff;
                                                                background: rgba(16, 40, 82, 0.92);
                .node.result-badge {{
                    font-size: 0.68rem;
                    font-weight: 600;
                    color: #4ade80;
                    background: rgba(20, 50, 30, 0.92);
                    border: 1px solid rgba(74, 222, 128, 0.45);
                    border-radius: 5px;
                    padding: 1px 5px;
                    white-space: nowrap;
                    max-width: 200px;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    z-index: 10;
                }}
                .result-score {{
                    color: #a3e4b8;
                    font-size: 0.64rem;
                    margin-left: 3px;
                }}
      }}
      .node.cat-title {{
                font-weight: 800;
                color: #d8e1f3;
                font-size: 0.88rem;
      }}
      .node.legend {{
                font-weight: 700;
                color: #adbbd7;
      }}
            .matchups-box {{
                margin-top: 8px;
                border: 1px solid rgba(255, 255, 255, 0.16);
                border-radius: 12px;
                padding: 10px 12px;
                background: rgba(14, 20, 34, 0.72);
            }}
            .matchups-title {{
                font-weight: 700;
                margin-bottom: 6px;
                color: #d8e1f3;
            }}
            .matchup-item {{
                display: flex;
                gap: 8px;
                align-items: baseline;
                font-size: 0.85rem;
                margin-bottom: 4px;
            }}
            .matchup-num {{
                min-width: 40px;
                font-weight: 700;
                color: #9fc0ff;
            }}
            .matchup-pairs {{
                color: #e8edf9;
            }}
            .stage-matchups-group {{
                margin-top: 10px;
            }}
            .stage-matchups-title {{
                font-size: 0.82rem;
                font-weight: 800;
                color: #f3d47a;
                margin-bottom: 8px;
                letter-spacing: 0.2px;
            }}
            .stage-matchups-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
                gap: 8px;
            }}
            .stage-matchup-item {{
                border: 1px solid rgba(255, 255, 255, 0.12);
                background: rgba(255, 255, 255, 0.04);
                border-radius: 10px;
                padding: 8px 10px;
            }}
            .stage-matchup-head {{
                color: #9fc0ff;
                font-size: 0.78rem;
                font-weight: 800;
                margin-bottom: 4px;
            }}
            .stage-matchup-pair {{
                color: #e8edf9;
                font-size: 0.83rem;
                line-height: 1.25;
            }}
            .stage-matchup-pair span {{
                color: #f3d47a;
                font-weight: 700;
            }}
            @media (max-width: 768px) {{
                .bracket-wrap {{
                    padding: 6px;
                    border-radius: 12px;
                }}
                .stage-matchups-grid {{
                    grid-template-columns: 1fr;
                }}
            }}
    </style>
    <script>
    (function() {{
      var BOARD_W = {width};
      var BOARD_H = {board_height};
            var userScale = 1;
            var fitScale = 1;
            var currentScale = 1;

            function clamp(value, min, max) {{
                return Math.max(min, Math.min(max, value));
            }}

      function applyScale() {{
        var wrap = document.querySelector('.bracket-wrap');
        var viewport = document.querySelector('.bracket-viewport');
        var board = document.querySelector('.bracket-board');
        if (!wrap || !board) return;
        var padH = parseInt(getComputedStyle(wrap).paddingLeft) +
                   parseInt(getComputedStyle(wrap).paddingRight);
        var available = wrap.clientWidth - padH;
                fitScale = Math.min(1, available / BOARD_W);
                currentScale = clamp(fitScale * userScale, fitScale * 0.7, fitScale * 4);
                board.style.transform = 'scale(' + currentScale + ')';
                viewport.style.height = Math.ceil(BOARD_H * currentScale) + 'px';
            }}

            window.zoomIn = function() {{
                userScale = clamp(userScale * 1.2, 0.7, 4);
                applyScale();
            }};

            window.zoomOut = function() {{
                userScale = clamp(userScale / 1.2, 0.7, 4);
                applyScale();
            }};

            window.zoomReset = function() {{
                userScale = 1;
                applyScale();
      }}

      document.addEventListener('DOMContentLoaded', applyScale);
      window.addEventListener('resize', applyScale);
    }})();
    </script>
    <div class="bracket-wrap">
                        <div class="zoom-controls">
                                <button class="zoom-btn" type="button" onclick="window.zoomOut()">−</button>
                                <button class="zoom-btn" type="button" onclick="window.zoomReset()">100%</button>
                                <button class="zoom-btn" type="button" onclick="window.zoomIn()">+</button>
                        </div>
            <div class="bracket-viewport">
                <div class="bracket-board">
                                                                        {watermark_html}
                                    <svg class="bracket-lines" width="{width}" height="{height}">
                                          {matchup_guides_svg}
                                            {connectors_svg}
                                    </svg>
                    {''.join(cell_html)}
                </div>
      </div>
    </div>
</body>
</html>
    """

    st.subheader(f"Categoría {category}")
    component_height = min(1800, max(520, board_height + 260))
    components.html(textwrap.dedent(html), height=component_height, scrolling=True)



def main() -> None:
    st.set_page_config(page_title="JPQ Llaves", layout="wide", page_icon="🏆")
    st.markdown(textwrap.dedent(APP_CSS), unsafe_allow_html=True)
    logo_data_uri = load_logo_data_uri()
    if logo_data_uri:
        st.markdown(
            textwrap.dedent(
                f"""
                <div class="header-wrap">
                  <img src="{logo_data_uri}" class="header-logo" alt="Logo JPQ" />
                  <h1 class="header-title">JPQ Llaves</h1>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
    else:
        st.title("JPQ Llaves")

    brackets = load_brackets()
    if not brackets:
        st.error(
            "No se pudieron cargar datos desde Google Sheets. "
            "Verificá si usás URL pública de exportación (.xlsx) o Service Account "
            "(JPQ_SHEET_ID + JPQ_GOOGLE_SERVICE_ACCOUNT_FILE/JPQ_GOOGLE_SERVICE_ACCOUNT_JSON)."
        )
        if LAST_LOAD_ERROR:
            st.caption(f"Detalle técnico: {LAST_LOAD_ERROR}")
        st.stop()

    brackets = _filter_visible_categories(brackets)
    if not brackets:
        st.error("No hay categorías publicadas con la configuración actual.")
        st.stop()

    categories = sorted(brackets.keys())
    category_labels = {category.upper(): category for category in categories}
    ordered_labels = sorted(category_labels.keys())
    st.markdown("<div class='section-title'>Seleccionar categoría</div>", unsafe_allow_html=True)
    selected_category_label = st.selectbox(
        "Categoría",
        options=ordered_labels,
        label_visibility="collapsed",
    )
    selected_category = category_labels[selected_category_label]
    payload = brackets[selected_category]
    results = load_results()
    render_bracket(
        grid=payload["grid"],
        category=selected_category,
        sheet_name=str(payload["sheet"]),
        results=results,
    )


if __name__ == "__main__":
    main()
