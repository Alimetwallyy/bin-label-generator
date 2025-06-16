import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Bin Label Generator", layout="wide")

# --- 🧠 Helper Functions ---
def generate_shelf_labels(count):
    return [chr(i) for i in range(ord('A'), ord('A') + count)]

def extract_aisle_number(bay_id):
    parts = bay_id.split('-')
    for part in parts:
        if part.isdigit() and len(part) == 3:
            return part
    return ""

def get_duplicates_within(lst):
    return set([x for x in lst if lst.count(x) > 1])

def get_duplicates_across(groups):
    seen = {}
    duplicates = {}
    for group in groups:
        for bay in group["bays"]:
            if bay in seen:
                duplicates.setdefault(bay, set()).update([group["group_name"], seen[bay]])
            else:
                seen[bay] = group["group_name"]
    return {k: list(v) for k, v in duplicates.items()}

def generate_bin_labels_table(group_name, bay_ids, shelves, bins_per_shelf):
    data = []
    for bay in bay_ids:
        base_label = bay.replace("BAY-", "")
        base_number = int(base_label[-3:])
        max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

        for i in range(max_bins):
            row = {
                'BAY TYPE': group_name,
                'AISLE': extract_aisle_number(bay),
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
    return pd.DataFrame(data)

def export_to_excel(grouped_dfs):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    hex_colors = ["339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017", "FF00FF", "996600", "00FF00", "FF6565", "9999FE"]
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    bold_font = Font(bold=True)

    for group_name, df in grouped_dfs.items():
        ws = wb.create_sheet(title=group_name)

        shelves = [col for col in df.columns if col not in ['BAY TYPE', 'AISLE', 'BAY ID']]
        has_shelves = len(shelves) > 0

        if has_shelves:
            # ABC1 cell
            ws.merge_cells('A1:C1')
            cell = ws['A1']
            cell.value = "HEX COLOR CODES ->"
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.font = bold_font

            for i, color in enumerate(hex_colors):
                col_letter = get_column_letter(4 + i)
                c = ws[f"{col_letter}1"]
                c.value = color
                c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                c.font = bold_font
                c.border = border

                if i < len(shelves):
                    c2 = ws[f"{col_letter}2"]
                    c2.value = shelves[i]
                    c2.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    c2.font = bold_font
                    c2.border = border

        # Write headers
        for col_idx, col in enumerate(df.columns, 1):
            cell = ws.cell(row=2, column=col_idx, value=col)
            cell.font = bold_font
            cell.border = border

        # Write data
        for row_idx, row in enumerate(df.values, start=3):
            for col_idx, val in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = bold_font
                cell.border = border

    wb.save(output)
    output.seek(0)
    return output

# --- 🖥️ Streamlit UI ---
st.title("📦 Bin Label Generator")

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

for i in range(num_groups):
    st.header(f"Bay Group {i+1}")
    group_name = st.text_input(f"Enter a name for Bay Group {i+1}", key=f"group_name_{i}")
    bay_input = st.text_area("Enter Bay IDs (one per line)", key=f"bay_input_{i}")
    shelf_count = st.number_input("Number of shelves (A-Z)", min_value=1, max_value=26, value=3, key=f"shelf_count_{i}")

    shelves = generate_shelf_labels(shelf_count)
    bins_per_shelf = {}
    for shelf in shelves:
        bins = st.number_input(f"Bins in Shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{i}_{shelf}")
        bins_per_shelf[shelf] = bins

    bay_ids = [b.strip() for b in bay_input.splitlines() if b.strip()]
    bay_groups.append({
        "group_name": group_name if group_name else f"Bay Group {i+1}",
        "bays": bay_ids,
        "shelves": shelves,
        "bins_per_shelf": bins_per_shelf
    })

# Validation
within_duplicates = {}
for group in bay_groups:
    dups = get_duplicates_within(group["bays"])
    if dups:
        within_duplicates[group["group_name"]] = dups

across_duplicates = get_duplicates_across(bay_groups)

if st.button("Generate Bin Labels"):
    if within_duplicates:
        for group, dups in within_duplicates.items():
            st.error(f"❌ Duplicated bays found within {group}: {', '.join(dups)}")
    elif across_duplicates:
        for bay, groups in across_duplicates.items():
            st.error(f"❌ Bay ID '{bay}' appears in multiple groups: {', '.join(groups)}")
    else:
        st.success("✅ No duplicate bays found. Generating...")

        result_dfs = {}
        for group in bay_groups:
            df = generate_bin_labels_table(group["group_name"], group["bays"], group["shelves"], group["bins_per_shelf"])
            result_dfs[group["group_name"]] = df
            st.subheader(f"📊 {group['group_name']}")
            st.dataframe(df)

        excel_data = export_to_excel(result_dfs)

        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name="bin_labels.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
