# app/ui.py
from typing import List, Dict
import streamlit as st
import pandas as pd

from app.logic import (
    generate_bin_labels_table_cached,
    check_duplicate_bay_ids,
    plot_bin_diagram,
)
from app.excel import build_excel_bytes

def run_app():
    st.title("Bin Label Generator — Refactored")

    st.sidebar.header("About")
    st.sidebar.markdown(
        "Refactored Streamlit app: robust bay/bin parsing, cached generation, excel export, and tests."
    )

    # Input panel
    with st.expander("Input / Groups", expanded=True):
        col1, col2 = st.columns([2, 1])
        with col1:
            st.markdown("**Enter groups of bays (one group per block)**")
            st.markdown(
                "Format: paste bay IDs separated by commas or new lines. Example bay IDs: `BAY-001-001`, `BAY-002-010`"
            )
            raw_groups = st.text_area(
                "Groups (separate groups with an empty line):",
                value="BAY-001-001, BAY-001-002\n\nBAY-002-001, BAY-002-002",
                height=200,
            )
        with col2:
            st.markdown("**Global options**")
            shelves_input = st.text_input(
                "Shelves (comma-separated labels)", value="S1,S2"
            )
            bins_per_shelf = st.number_input(
                "Bins per shelf (default for each shelf)", min_value=1, max_value=999, value=3
            )

    # parse groups
    def parse_groups_text(raw: str):
        groups = []
        for block in raw.split("\n\n"):
            block = block.strip()
            if not block:
                continue
            # allow commas or newlines inside block
            items = [x.strip() for part in block.splitlines() for x in part.split(",") if x.strip()]
            groups.append(items)
        return groups

    groups = parse_groups_text(raw_groups)
    shelves = [s.strip() for s in shelves_input.split(",") if s.strip()]

    # Show parsed preview
    st.markdown("**Preview:**")
    st.write({"num_groups": len(groups), "shelves": shelves, "bins_per_shelf": int(bins_per_shelf)})

    # Duplicate check
    dup_result = check_duplicate_bay_ids(groups)
    if dup_result["duplicates"]:
        st.warning(f"Duplicate bay IDs detected across groups: {dup_result['duplicates']}")

    # Use form to avoid reruns on every input change
    with st.form("generate_form"):
        submitted = st.form_submit_button("Generate labels & diagram")

    if submitted:
        try:
            # heavy computation (cached)
            df = generate_bin_labels_table_cached(groups=groups, shelves=shelves, bins_per_shelf=int(bins_per_shelf))
            st.success(f"Generated {len(df)} label rows.")

            st.dataframe(df.head(200), use_container_width=True)

            # Plot diagram for first group as example
            if len(groups) >= 1 and groups[0]:
                fig = plot_bin_diagram(groups[0], shelves=shelves, bins_per_shelf=int(bins_per_shelf))
                st.plotly_chart(fig, use_container_width=True)

            # Excel download
            excel_bytes = build_excel_bytes(df)
            st.download_button("Download Excel", data=excel_bytes, file_name="bin_labels.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.exception(e)

    st.sidebar.markdown("---")
    st.sidebar.markdown("Created by Alimomet — refactor edition")
