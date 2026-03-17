from pathlib import Path
from html import escape
import base64
import textwrap
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


DATA_DIR = Path(__file__).parent / "LLAVES 1er JPQ"
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
def load_brackets(data_dir: Path) -> Dict[str, Dict[str, object]]:
    brackets: Dict[str, Dict[str, object]] = {}

    for file_path in sorted(data_dir.glob("*.xlsx")):
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
    text = value.strip()
    if text.lower().startswith("categoría"):
        return "cat-title"
    if text.upper() in {"Nº PARTIDO", "DIA - HORA", "COMPLEJO"}:
        return "legend"
    digits = text.replace(".", "", 1)
    if digits.isdigit() and len(text) <= 3:
        return "match-id"
    if "°" in text:
        return "seed"
    return "team"


def _build_connectors(
    nodes: List[Dict[str, object]],
    board_rows: int,
) -> str:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    match_nodes = [
        node
        for node in nodes
        if node["class"] == "match-id"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]

    if len(match_nodes) < 2:
        return ""

    col_map: Dict[int, List[Dict[str, object]]] = {}
    for node in match_nodes:
        col = int(node["col"])
        col_map.setdefault(col, []).append(node)

    for col in col_map:
        col_map[col].sort(key=lambda item: int(item["row"]))

    sorted_cols = sorted(col_map.keys())
    connector_paths: List[str] = []

    def route(x1: int, y1: int, x2: int, y2: int) -> str:
        mid_x = x1 + max(14, (x2 - x1) // 2)
        return f"M{x1},{y1} H{mid_x} V{y2} H{x2}"

    def connect(left_node: Dict[str, object], right_node: Dict[str, object]) -> None:
        x1 = int(left_node["x"]) + 78
        y1 = int(left_node["y"]) + 12
        x2 = int(right_node["x"]) - 4
        y2 = int(right_node["y"]) + 12
        connector_paths.append(route(x1, y1, x2, y2))

    for idx in range(len(sorted_cols) - 1):
        left_nodes = col_map[sorted_cols[idx]]
        right_nodes = col_map[sorted_cols[idx + 1]]
        if not right_nodes:
            continue

        left_count = len(left_nodes)
        right_count = len(right_nodes)

        for right_idx, right_node in enumerate(right_nodes):
            start_idx = round(right_idx * left_count / right_count)
            end_idx = round((right_idx + 1) * left_count / right_count)
            for left_node in left_nodes[start_idx:end_idx]:
                connect(left_node, right_node)

    return "".join(
        f'<path d="{path}" class="connector"></path>' for path in connector_paths
    )


def _build_round_labels(nodes: List[Dict[str, object]]) -> str:
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
        label = round_names[idx] if idx < len(round_names) else f"Ronda {idx + 1}"
        x = col * CELL_WIDTH
        labels.append(
            f'<div class="round-label" style="left:{x}px;">{escape(label)}</div>'
        )

    return "".join(labels)


def _build_matchups_html(nodes: List[Dict[str, object]]) -> str:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    teams = [
        node
        for node in nodes
        if node["class"] == "team"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]
    matches = [
        node
        for node in nodes
        if node["class"] == "match-id"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]

    if not matches or not teams:
        return ""

    first_round_col = min(int(node["col"]) for node in matches)
    matches = sorted(
        [node for node in matches if int(node["col"]) == first_round_col],
        key=lambda node: int(node["row"]),
    )
    items: List[str] = []

    for match in matches:
        match_row = int(match["row"])
        match_number = str(match["text"])

        nearby = [
            team
            for team in teams
            if int(team["col"]) == first_round_col
            and abs(int(team["row"]) - match_row) <= 9
        ]
        nearby = sorted(
            nearby,
            key=lambda team: (abs(int(team["row"]) - match_row), int(team["row"])),
        )

        selected = nearby[:4]
        if not selected:
            continue

        selected = sorted(selected, key=lambda team: int(team["row"]))
        names = [escape(str(team["text"])) for team in selected]

        if len(names) >= 4:
            pair_a = f"{names[0]} / {names[1]}"
            pair_b = f"{names[2]} / {names[3]}"
        elif len(names) == 3:
            pair_a = f"{names[0]} / {names[1]}"
            pair_b = names[2]
        elif len(names) == 2:
            pair_a = names[0]
            pair_b = names[1]
        else:
            pair_a = names[0]
            pair_b = "-"

        items.append(
            f'<div class="matchup-item"><span class="matchup-num">#{escape(match_number)}</span><span class="matchup-pairs">{pair_a} <b>vs</b> {pair_b}</span></div>'
        )

    if not items:
        return ""

    return f'<div class="matchups-box"><div class="matchups-title">Cruces de la etapa</div>{"".join(items)}</div>'


def _build_matchup_guides_svg(nodes: List[Dict[str, object]]) -> str:
    legend_rows = [int(node["row"]) for node in nodes if node["class"] == "legend"]
    legend_start_row = min(legend_rows) if legend_rows else None

    matches = [
        node
        for node in nodes
        if node["class"] == "match-id"
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
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
        and (legend_start_row is None or int(node["row"]) < legend_start_row)
    ]

    guides: List[str] = []
    for match in first_round_matches:
        match_row = int(match["row"])
        nearby_teams = sorted(
            [team for team in teams if abs(int(team["row"]) - match_row) <= 5],
            key=lambda team: (abs(int(team["row"]) - match_row), int(team["row"])),
        )[:4]
        if len(nearby_teams) < 2:
            continue

        # Find seeds on the same rows as teams
        team_rows = {int(t["row"]) for t in nearby_teams}
        nearby_seeds = [n for n in nodes if n["class"] == "seed" and int(n["row"]) in team_rows and (legend_start_row is None or int(n["row"]) < legend_start_row)]
        
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
            text = str(value).strip()
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
            text = str(value).strip()
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
            text = str(value).strip()
            if not text:
                continue
            if _cell_class(text) == "match-id":
                first_round_rows.add(row_idx)

    cell_html: List[str] = []
    node_data: List[Dict[str, object]] = []
    for row_idx in range(rows):
        for col_idx in range(cols):
            value = grid.iat[row_idx, col_idx]
            if pd.isna(value):
                continue
            text = str(value).strip()
            if text == "":
                continue

            cls = _cell_class(text)
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
                right_text = str(right_value).strip() if pd.notna(right_value) else ""
                below_right_text = (
                    str(below_right_value).strip() if pd.notna(below_right_value) else ""
                )
                if right_text and below_right_text:
                    if _cell_class(right_text) == "team" and _cell_class(below_right_text) == "team":
                        top = row_idx * CELL_HEIGHT + CELL_HEIGHT // 2 + LABEL_ROW_HEIGHT
                        left = (col_idx + 1) * CELL_WIDTH - 44
                        cls = "seed-between"

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

    connectors_svg = _build_connectors(nodes=node_data, board_rows=rows)
    matchup_guides_svg = _build_matchup_guides_svg(nodes=node_data)
    round_labels_html = _build_round_labels(nodes=node_data)
    matchups_html = _build_matchups_html(nodes=node_data)

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
                overflow: hidden;
                width: 100%;
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
            @media (max-width: 768px) {{
                .bracket-wrap {{
                    padding: 6px;
                    border-radius: 12px;
                }}
            }}
    </style>
    <script>
    (function() {{
      var BOARD_W = {width};
      var BOARD_H = {board_height};

      function applyScale() {{
        var wrap = document.querySelector('.bracket-wrap');
        var viewport = document.querySelector('.bracket-viewport');
        var board = document.querySelector('.bracket-board');
        if (!wrap || !board) return;
        var padH = parseInt(getComputedStyle(wrap).paddingLeft) +
                   parseInt(getComputedStyle(wrap).paddingRight);
        var available = wrap.clientWidth - padH;
        var scale = Math.min(1, available / BOARD_W);
        board.style.transform = 'scale(' + scale + ')';
        viewport.style.height = Math.ceil(BOARD_H * scale) + 'px';
      }}

      document.addEventListener('DOMContentLoaded', applyScale);
      window.addEventListener('resize', applyScale);
    }})();
    </script>
    <div class="bracket-wrap">
            <div class="bracket-viewport">
                <div class="bracket-board">
                                    <div class="round-labels-container">
                                        {round_labels_html}
                                    </div>
                                    <svg class="bracket-lines" width="{width}" height="{height}">
                                          {matchup_guides_svg}
                                            {connectors_svg}
                                    </svg>
                    {''.join(cell_html)}
                </div>
      </div>
            {matchups_html}
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

    if not DATA_DIR.exists():
        st.error(f"No se encontró la carpeta de datos: {DATA_DIR}")
        st.stop()

    brackets = load_brackets(DATA_DIR)
    if not brackets:
        st.warning("No se encontraron datos para mostrar.")
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
