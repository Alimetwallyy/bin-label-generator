import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
import string

# --- 🔧 Generate bin labels ---
def generate_bin_labels_table(bay_groups):
    all_dataframes = []
    for group in bay_groups:
        bay_ids = group['bays']
        shelves = group['shelves']
        bins_per_shelf = group['bins_per_shelf']
        group_name = group['group_name']
        data = []

        for bay in bay_ids:
            base_label = bay.replace("BAY-", "")
            aisle = base_label[9:12] if len(base_label) >= 12 else ""
            base_number = int(base_label[-3:])
            max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

            for i in range(max_bins):
                row = {
                    'BAY TYPE': group_name,
                    'AISLE': aisle,
                    'BAY ID': bay
                }
                for shelf in shelves:
                    shelf_bin_count = bins_per_shelf.get(shelf, 0)
                    if i < shelf_bin_count:
                        bin_label = base_label[:-4] + shelf + f"{base_number + i:03d}"
                        row[shelf] = bin_label
                    else:
                        row[shelf] = None
                data.append(row)

        df = pd.DataFrame(data)
        all_dataframes.append((group_name, df, shelves))
    return all_dataframes

# --- 📊 Draw bin diagram ---
def plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number):
    fig, ax = plt.subplots(figsize=(len(shelves) * 2, max(bins_per_shelf.values()) * 0.6))
    ax.set_title(f"Bin Layout for {bay_id}", fontsize=14)
    ax.axis('off')
    colors = ['lightblue', 'lightgreen', 'salmon', 'khaki', 'plum', 'coral', 'lightpink', 'wheat']
    shelf_colors = {shelf: colors[i % len(colors)] for i, shelf in enumerate(shelves)}
    for col_idx, shelf in enumerate(shelves):
        shelf_bins = bins_per_shelf.get(shelf, 0)
        for i in range(shelf_bins):
            bin_label = bay_id.replace("BAY-", "")[:-4] + shelf + f"{base_number + i:03d}"
            ax.text(col_idx, -i, bin_label, va='center', ha='center', fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor=shelf_colors[shelf]))
    return fig

# --- 🖥️ Streamlit UI ---
st.title("📦 Bin Label Generator")

if st.button("Reset Form"):
    st.session_state.clear()
    st.experimental_rerun()

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

used_bays = set()
errors = []

for group_idx in range(num_groups):
    st.header(f"🧱 Bay Group {group_idx + 1}")
    group_name = st.text_input(f"Group Name", value=f"Bay Group {group_idx + 1}", key=f"group_name_{group_idx}")
    bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")
    num_shelves = st.number_input(f"Number of shelves (auto A-Z)", min_value=1, max_value=26, value=3, key=f"num_shelves_{group_idx}")
    shelves = list(string.ascii_uppercase[:num_shelves])
    bins_per_shelf = {}
    for shelf in shelves:
        count = st.number_input(f"Number of bins in shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
        bins_per_shelf[shelf] = count

    if bays_input:
        bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
        group_duplicates = set([b for b in bay_list if bay_list.count(b) > 1])
        cross_duplicates = used_bays.intersection(set(bay_list))
        if group_duplicates:
            errors.append(f"❌ Group '{group_name}' contains duplicate bays: {', '.join(group_duplicates)}")
        if cross_duplicates:
            errors.append(f"❌ Group '{group_name}' has bays already used in previous groups: {', '.join(cross_duplicates)}")
        used_bays.update(bay_list)
        bay_groups.append({
            "group_name": group_name,
            "bays": bay_list,
            "shelves": shelves,
            "bins_per_shelf": bins_per_shelf
        })

if errors:
    for e in errors:
        st.error(e)

if st.button("Generate Bin Labels") and not errors:
    all_dfs = generate_bin_labels_table(bay_groups)
    st.success("✅ Bin labels generated successfully!")

    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    hex_colors = ["339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017", "FF00FF", "996600", "00FF00", "FF6565", "9999FE"]

    for group_name, df, shelves in all_dfs:
        ws = wb.create_sheet(title=group_name)

        # Merge A1 to C1 and set header
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        ws.cell(row=1, column=1, value="HEX COLOR CODES ->")
        ws.cell(row=1, column=1).fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

        for idx, color in enumerate(hex_colors):
            cell = ws.cell(row=1, column=4 + idx, value=color)
            cell.fill = PatternFill(start_color=color, fill_type="solid")

        # Second row: headers
        ws.cell(row=2, column=1, value="BAY TYPE")
        ws.cell(row=2, column=2, value="AISLE")
        ws.cell(row=2, column=3, value="BAY ID")

        for idx, shelf in enumerate(shelves):
            col = 4 + idx
            cell = ws.cell(row=2, column=col, value=shelf)
            if idx < len(hex_colors):
                cell.fill = PatternFill(start_color=hex_colors[idx], fill_type="solid")

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=3):
            for c_idx, val in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=val)

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Download Formatted Excel File",
        data=output,
        file_name="bin_labels_formatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("🖼️ Bin Layout Diagrams")
    for group in bay_groups:
        for bay_id in group['bays']:
            shelves = group['shelves']
            bins_per_shelf = group['bins_per_shelf']
            base_label = bay_id.replace("BAY-", "")
            base_number = int(base_label[-3:])
            fig = plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number)
            st.pyplot(fig)
