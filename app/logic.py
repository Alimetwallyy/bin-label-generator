# app/logic.py
from typing import List, Dict, Any
import pandas as pd
from app.utils import parse_bay_id, normalize_bay_id
from functools import lru_cache
import streamlit as st
import plotly.graph_objs as go
import plotly.express as px

def check_duplicate_bay_ids(groups: List[List[str]]) -> Dict[str, Any]:
    """
    Returns dict with 'duplicates' key listing normalized duplicates across groups.
    """
    seen = {}
    duplicates = set()
    for gi, group in enumerate(groups):
        for bay in group:
            nb = normalize_bay_id(bay)
            if nb in seen:
                duplicates.add(nb)
            else:
                seen[nb] = (gi, bay)
    return {"duplicates": sorted(list(duplicates)), "count": len(duplicates)}


@st.cache_data(ttl=600)
def generate_bin_labels_table_cached(groups: List[List[str]], shelves: List[str], bins_per_shelf: int) -> pd.DataFrame:
    return generate_bin_labels_table(groups, shelves, bins_per_shelf)


def generate_bin_labels_table(groups: List[List[str]], shelves: List[str], bins_per_shelf: int) -> pd.DataFrame:
    rows = []
    for gi, group in enumerate(groups, start=1):
        group_name = f"Group {gi}"
        for bay in group:
            parsed = parse_bay_id(bay)
            if parsed is None:
                # keep original raw but mark invalid
                base_number = None
                base_prefix = normalize_bay_id(bay)
            else:
                base_number = parsed.get("number") or ""
                # use a prefix composed of Bay normalized without trailing number
                raw = normalize_bay_id(parsed["raw"])
                # strip trailing number if present
                if base_number and raw.endswith(base_number):
                    base_prefix = raw[: -len(base_number)]
                else:
                    base_prefix = raw + "-" if not raw.endswith("-") else raw

            for shelf in shelves:
                for i in range(bins_per_shelf):
                    if base_number and base_number.isdigit():
                        # increment numeric suffix safely
                        try:
                            num = int(base_number) + i
                            label = f"{base_prefix}{shelf}{num:03d}"
                        except Exception:
                            label = f"{base_prefix}{shelf}{i+1}"
                    else:
                        # fallback if we couldn't parse base number
                        label = f"{base_prefix}{shelf}{i+1}"
                    rows.append(
                        {
                            "group": group_name,
                            "bay_input": bay,
                            "normalized_bay": normalize_bay_id(bay),
                            "shelf": shelf,
                            "bin_label": label,
                        }
                    )
    df = pd.DataFrame(rows)
    return df


def plot_bin_diagram(group_bays: List[str], shelves: List[str], bins_per_shelf: int):
    """
    Simple responsive plotly diagram: shows bays on x-axis and stacked shelf rows.
    This intentionally avoids huge widths — width is responsive.
    """
    # Build a small layout dataset
    data = []
    for bi, bay in enumerate(group_bays):
        for si, shelf in enumerate(shelves):
            for b in range(bins_per_shelf):
                label = f"{normalize_bay_id(bay)}-{shelf}-{b+1}"
                data.append({"bay": normalize_bay_id(bay), "shelf": shelf, "bin": b + 1, "label": label, "x": bi, "y": si})

    if not data:
        fig = go.Figure()
        fig.update_layout(title="No data to plot")
        return fig

    df = pd.DataFrame(data)

    # Use Plotly categorical colors
    palette = px.colors.qualitative.Plotly
    shelves_unique = list(df["shelf"].unique())

    fig = go.Figure()
    for idx, shelf in enumerate(shelves_unique):
        df_s = df[df["shelf"] == shelf]
        fig.add_trace(
            go.Bar(
                x=df_s["bay"],
                y=[1] * len(df_s),
                name=shelf,
                text=df_s["label"],
                hoverinfo="text",
                marker=dict(color=palette[idx % len(palette)]),
            )
        )

    fig.update_layout(barmode="stack", title="Bin diagram (stacked by shelf)", xaxis_title="Bay", yaxis_title="Shelf stack (visual only)", legend_title="Shelf")
    return fig
