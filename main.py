from pathlib import Path
from html import escape
import os
import base64
from io import BytesIO
import textwrap
from typing import Dict, List, Optional
from urllib.request import urlopen

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


DATA_DIR = Path(__file__).parent / "LLAVES 1er JPQ"
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1kiw5oIs3dw5yj2ME26tHoJ7Y5_TEvqYBnxHoDVeoi_4/export?format=xlsx"
SHEET_URL = os.getenv("JPQ_SHEET_URL", DEFAULT_SHEET_URL).strip()
USE_LOCAL_FALLBACK = os.getenv("JPQ_USE_LOCAL_FALLBACK", "1").strip().lower() not in {
    "0",
    "false",
    "no",
}
SHEET_TIMEOUT_SECONDS = 20
LOGO_GLOB = "Logo JPQ*"
CELL_WIDTH = 132
CELL_HEIGHT = 26


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


def _find_data_sheet(sheet_names: List[str]) -> Optional[str]:
    for sheet_name in sheet_names:
        if not sheet_name.lower().startswith("base"):
            return sheet_name
    return None


@st.cache_data(show_spinner=False)
def load_brackets() -> Dict[str, Dict[str, object]]:
    brackets: Dict[str, Dict[str, object]] = {}

    try:
        workbook_bytes = urlopen(SHEET_URL, timeout=SHEET_TIMEOUT_SECONDS).read()
        excel_file = pd.ExcelFile(BytesIO(workbook_bytes))

        for sheet_name in excel_file.sheet_names:
            if sheet_name.lower().startswith("base"):
                continue

            bracket_df = pd.read_excel(BytesIO(workbook_bytes), sheet_name=sheet_name, header=None)

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
                continue

            cropped = bracket_df.iloc[
                first_used_row : last_used_row + 1,
                first_used_col : last_used_col + 1,
            ].reset_index(drop=True)
            cropped.columns = range(cropped.shape[1])
            category = sheet_name.strip()
            brackets[category] = {"sheet": sheet_name, "grid": cropped}

        if brackets:
            return brackets
    except Exception:
        pass

    if not USE_LOCAL_FALLBACK:
        return {}

    for file_path in sorted(DATA_DIR.glob("*.xlsx")):
        category = file_path.stem
        excel_file = pd.ExcelFile(file_path)
        data_sheet = _find_data_sheet(excel_file.sheet_names)

        if data_sheet is None:
            continue

        bracket_df = pd.read_excel(file_path, sheet_name=data_sheet, header=None)

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
            continue

        cropped = bracket_df.iloc[
            first_used_row : last_used_row + 1,
            first_used_col : last_used_col + 1,
        ].reset_index(drop=True)
        cropped.columns = range(cropped.shape[1])
        brackets[category] = {"sheet": data_sheet, "grid": cropped}

    return brackets


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


def _normalize_cell_text(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text[:-1].isdigit() and text[-1].isalpha() and text[-1].isupper() and "°" not in text:
        return f"{int(text[:-1])}° {text[-1]}"
    return text


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
        mid_x = x1 + max(14, (x2 - x1) // 2)
        return f"M{x1},{y1} H{mid_x} V{y2} H{x2}"

    def target_x(right_node: Dict[str, object]) -> int:
        default_x = int(right_node["x"]) + 88
        if (category or "").lower() != "c3":
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
            y2 = int(target_node["y"]) + 12

            if (
                seed_label.strip() == "1° A"
                and target_number == "57"
                and "57" in c3_shared_trunk_x_by_target
            ):
                x2 = int(target_node["x"]) - 12
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            if (
                seed_label.strip() == "1° C"
                and target_number == "59"
                and "59" in c3_shared_trunk_x_by_target
            ):
                x2 = int(target_node["x"]) - 12
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            if (
                seed_label.strip() == "1° B"
                and target_number == "60"
                and "60" in c3_shared_trunk_x_by_target
            ):
                x2 = int(target_node["x"]) - 12
                x1_left = int(anchor[0]) - 16
                connector_paths.append(f"M{x1_left},{y1} H{x2} V{y2}")
                return

            x2 = int(target_node["x"]) - 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("9°", "50")
        append_direct_connector("8°", "50")
        append_direct_connector("5°", "51")
        append_direct_connector("12°", "51")
        append_direct_connector("13°", "52")
        append_direct_connector("4°", "52")
        append_direct_connector("3°", "53")
        append_direct_connector("14°", "53")
        append_direct_connector("11°", "54")
        append_direct_connector("6°", "54")
        append_direct_connector("7°", "55")
        append_direct_connector("10°", "55")
        append_direct_connector("15°", "56")
        append_direct_connector("2°", "56")

    if (category or "").lower() == "c4":
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
            x2 = int(target_node["x"]) - 4
            y2 = int(target_node["y"]) + 12
            connector_paths.append(route(x1, y1, x2, y2))

        append_direct_connector("1° A", "49")
        append_direct_connector("1° E", "51")
        append_direct_connector("1° D", "52")
        append_direct_connector("1° C", "53")
        append_direct_connector("1° B", "56")

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

    if (category or "").lower() == "c5 40":
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
            adjacent_teams = sorted(
                [
                    node
                    for node in nodes
                    if node["class"] == "team"
                    and int(node["col"]) == seed_col + 1
                    and abs(int(node["row"]) - seed_row) <= 8
                ],
                key=lambda node: int(node["row"]),
            )
            if len(adjacent_teams) < 2:
                return None

            candidate_pairs: List[tuple[Dict[str, object], Dict[str, object]]] = []
            for idx in range(len(adjacent_teams) - 1):
                top_team = adjacent_teams[idx]
                bottom_team = adjacent_teams[idx + 1]
                if abs(int(bottom_team["row"]) - int(top_team["row"])) > 2:
                    continue
                candidate_pairs.append((top_team, bottom_team))

            if not candidate_pairs:
                return None

            below_pairs = [
                (top_team, bottom_team)
                for top_team, bottom_team in candidate_pairs
                if int(top_team["row"]) >= seed_row
            ]

            if below_pairs:
                selected_top, selected_bottom = min(
                    below_pairs,
                    key=lambda pair: int(pair[0]["row"]) - seed_row,
                )
            else:
                selected_top, selected_bottom = min(
                    candidate_pairs,
                    key=lambda pair: abs(
                        ((int(pair[0]["row"]) + int(pair[1]["row"])) / 2) - seed_row
                    ),
                )

            x = int(selected_top["x"])
            y = (int(selected_top["y"]) + int(selected_bottom["y"])) // 2 + 12
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
            ("34", "49"),
            ("49", "57"),
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("57", "61"),
            ("58", "61"),
            ("53", "59"),
            ("54", "59"),
            ("55", "60"),
            ("56", "60"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_c2_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower() == "c4":
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

        explicit_c4_pairs = [
            ("34", "49"),
            ("35", "50"),
            ("36", "50"),
            ("37", "51"),
            ("38", "52"),
            ("39", "53"),
            ("40", "54"),
            ("41", "54"),
            ("42", "55"),
            ("43", "55"),
            ("44", "56"),
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

        for left_number, right_number in explicit_c4_pairs:
            connect_by_number(left_number, right_number)

    if (category or "").lower().startswith(("c5", "c6", "c7")):
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
            ("50", "57"),
            ("51", "58"),
            ("52", "58"),
            ("53", "59"),
            ("54", "59"),
            ("55", "60"),
            ("57", "61"),
            ("58", "61"),
            ("59", "62"),
            ("60", "62"),
            ("61", "64"),
            ("62", "64"),
        ]

        for left_number, right_number in explicit_c5_40_pairs:
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
            and abs(int(node["row"]) - target_row) <= 6
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
            "52": ("13°", "4°"),
            "53": ("3°", "14°"),
            "54": ("11°", "6°"),
            "55": ("7°", "10°"),
            "56": ("15°", "2°"),
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

    if (category or "").lower().startswith(("c5", "c6", "c7")):
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
            "50": ("2° F", "2° G"),
            "51": ("1° E", "2° C"),
            "52": ("2° B", "1° D"),
            "53": ("1° C", "2° A"),
            "54": ("2° D", "1° F"),
            "55": ("1° G", "2° E"),
        }

        for match_number, (seed_a, seed_b) in c5_40_seed_pairs.items():
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
        center_1b = seed_centers.get("1° B")
        if match_60 is not None and match_55 is not None and center_1b is not None:
            center_55 = int(match_55["y"]) + 12
            match_60["y"] = round((center_55 + center_1b) / 2) - 12

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
        "c5 40": {
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
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

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
            [team for team in teams if abs(int(team["row"]) - match_row) <= 8],
            key=lambda team: int(team["row"]),
        )
        upper_teams = [team for team in candidate_teams if int(team["row"]) <= match_row][-2:]
        lower_teams = [team for team in candidate_teams if int(team["row"]) > match_row][:2]
        nearby_teams = sorted(upper_teams + lower_teams, key=lambda team: int(team["row"]))
        if len(nearby_teams) < 2:
            continue

        # Find seeds on the same rows as teams
        team_rows = {int(t["row"]) for t in nearby_teams}
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
        
        # Bounding box from seeds + teams + match
        all_nodes = nearby_seeds + nearby_teams + [match]
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


def render_bracket(grid: pd.DataFrame, category: str, sheet_name: str) -> None:
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
            if cls == "match-id":
                try:
                    text = str(int(float(text)))
                except ValueError:
                    pass
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

            safe_text = escape(text)
            node_data.append(
                {
                    "row": row_idx,
                    "col": col_idx,
                    "class": cls,
                    "x": left,
                    "y": top,
                    "text": text,
                }
            )
            cell_html.append(
                f'<div class="node {cls}{extra_class}" style="top:{top}px;left:{left}px;">{safe_text}</div>'
            )

    _align_match_nodes(nodes=node_data, category=category)

    cell_html = []
    for node in node_data:
        safe_text = escape(str(node["text"]))
        cell_html.append(
            f'<div class="node {node["class"]}" style="top:{int(node["y"])}px;left:{int(node["x"])}px;">{safe_text}</div>'
        )

    connectors_svg = _build_connectors(nodes=node_data, board_rows=rows, category=category)
    matchup_guides_svg = _build_matchup_guides_svg(nodes=node_data, category=category)

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
                background: rgba(255, 255, 255, 0.06);
                color: #f2f6ff;
      }}
      .node.seed {{
        font-weight: 600;
      }}
            .node.seed-between {{
                                font-weight: 700;
                                border: 1px solid rgba(214, 183, 99, 0.52);
                                color: #f8df95;
                                background: rgba(214, 183, 99, 0.13);
                        min-width: unset;
                        max-width: unset;
                        padding: 2px 7px;
            }}
      .node.match-id {{
        font-weight: 700;
                                border: 1px solid rgba(92, 147, 255, 0.62);
                                color: #8db7ff;
                                background: rgba(92, 147, 255, 0.14);
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
        st.warning("No se pudieron cargar datos desde Google Sheets ni desde la carpeta local.")
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
    render_bracket(
        grid=payload["grid"],
        category=selected_category,
        sheet_name=str(payload["sheet"]),
    )


if __name__ == "__main__":
    main()
