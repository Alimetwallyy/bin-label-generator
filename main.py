import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Define HEX colors
HEX_COLORS = ["339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017", "FF00FF", "996600", "00FF00", "FF6565", "9999FE"]

# Generate shelves based on number
def generate_shelves(n):
    return [chr(65 + i) for i in range(n)]

# Extract aisle number
def extract_aisle(bay_id):
    try:
        digits = ''.join(filter(str.isdigit, bay_id[9:]))
        return digits[:3] if digits else None
    except:
        return None

# Generate bin labels
def generate_bin_labels_table(bay_group_name, bay_ids, shelves, bins_per_shelf):
    data = []

    for bay in bay_ids:
        base_label = bay.replace("BAY-", "")
        base_number = int(base_label[-3:]) if base_label[-3:].isdigit() else 0
        aisle_number = extract_aisle(bay)

        max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

        for i in range(max_bins):
            row = {
                'BAY TYPE': bay_group_name,
                'AISLE': aisle_number,
                'BAY ID': bay
            }
            shelf_has_data = False
            for shelf in shelves:
                count = bins_per_shelf.get(shelf, 0)
                if i < count:
                    row[shelf] = base_label[:-4] + shelf + f"{base_number + i:03d}"
                    shelf_has_data = True
                else:
                    row[shelf] = None
            if shelf_has_data:
                data.append(row)
    return pd.DataFrame(data)

# Plot bin diagram
def plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number):
    fig, ax = plt.subplots(figsize=(len(shelves) * 2, max(bins_per_shelf.values()) * 0.6))
    ax.set_title(f"Bin Layout for {bay_id}", fontsize=14)
    ax.axis('off')

    colors = ['lightblue', 'lightgreen', 'salmon', 'khaki', 'plum', 'coral', 'lightpink', 'wheat']
    shelf_colors = {shelf: colors[i % len(colors)] for i, shelf in enumerate(shelves)}

    for col_idx, shelf in enumerate(shelves):
        for i in range(bins_per_shelf.get(shelf, 0)):
            bin_label = bay_id.replace("BAY-", "")[:-4] + shelf + f"{base_number + i:03d}"
            ax.text(col_idx, -i, bin_label, ha='center', va='center', fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor=shelf_colors[shelf]))
    return fig

# Style Excel Sheet
def style_excel_sheet(ws, shelves):
    bold_font = Font(bold=True)
    center = Alignment(horizontal='center', vertical='center')
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Header: merged and colored
    if shelves:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        ws['A1'] = "HEX COLOR CODES ->"
        ws['A1'].font = bold_font
        ws['A1'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws['A1'].alignment = center

        for col_idx, color in enumerate(HEX_COLORS, start=4):
            cell = ws.cell(row=1, column=col_idx, value=color)
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            cell.font = bold_font
            cell.alignment = center

        for col_idx, shelf in enumerate(shelves, start=4):
            cell = ws.cell(row=2, column=col_idx, value=shelf)
            cell.fill = PatternFill(start_color=HEX_COLORS[col_idx - 4], end_color=HEX_COLORS[col_idx - 4], fill_type="solid")
            cell.font = bold_font
            cell.alignment = center

    headers = ["BAY TYPE", "AISLE", "BAY ID"] + shelves
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = bold_font
        cell.alignment = center

    # Apply borders and bold to all cells with data
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                cell.border = border
                cell.font = bold_font
                cell.alignment = center

# Streamlit UI
st.title("📦 Bin Label Generator")

if 'bay_groups' not in st.session_state:
    st.session_state.bay_groups = []

if st.button("➕ Add New Bay Group"):
    st.session_state.bay_groups.append({
        "name": f"Bay Group {len(st.session_state.bay_groups) + 1}",
        "bays": "",
        "shelf_count": 3,
        "bins_per_shelf": {}
    })

# Show groups
valid = True
duplicate_errors = []
all_bays = []

for i, group in enumerate(st.session_state.bay_groups):
    st.subheader(f"🧱 {group['name']}")
    st.session_state.bay_groups[i]['name'] = st.text_input("Group Name", value=group['name'], key=f"name_{i}")
    bays_text = st.text_area("Bay IDs (one per line)", value=group['bays'], key=f"bays_{i}")
    bay_list = [b.strip() for b in bays_text.splitlines() if b.strip()]
    st.session_state.bay_groups[i]['bays'] = bays_text

    shelf_count = st.number_input("Number of Shelves", min_value=1, max_value=26, value=group['shelf_count'], key=f"shelf_count_{i}")
    shelves = generate_shelves(shelf_count)
    st.session_state.bay_groups[i]['shelf_count'] = shelf_count

    bins_per_shelf = {}
    for shelf in shelves:
        bins = st.number_input(f"Bins in Shelf {shelf}", min_value=1, max_value=100, value=group['bins_per_shelf'].get(shelf, 5), key=f"bins_{i}_{shelf}")
        bins_per_shelf[shelf] = bins
    st.session_state.bay_groups[i]['bins_per_shelf'] = bins_per_shelf

    # Check for duplicates within the same group
    if len(set(bay_list)) != len(bay_list):
        valid = False
        st.error("❌ Duplicate bay IDs found **within this group**.")

    # Check for duplicates across groups
    for bay in bay_list:
        if bay in all_bays:
            valid = False
            duplicate_errors.append(f"{bay} (duplicate in another group)")
        else:
            all_bays.append(bay)

if duplicate_errors:
    st.error("❌ Duplicate bay IDs across groups:\n" + "\n".join(duplicate_errors))

if st.button("🧹 Reset All"):
    st.session_state.bay_groups = []

if st.button("🚀 Generate Bin Labels") and valid:
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    for group in st.session_state.bay_groups:
        bays = [b.strip() for b in group['bays'].splitlines() if b.strip()]
        shelves = generate_shelves(group['shelf_count'])
        df = generate_bin_labels_table(group['name'], bays, shelves, group['bins_per_shelf'])

        if not df.empty:
            ws = wb.create_sheet(title=group['name'][:31])
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
            style_excel_sheet(ws, shelves)

    wb.save(output)
    output.seek(0)

    st.success("✅ Bin labels generated and formatted successfully!")
    st.download_button("📥 Download Excel File", output, "bin_labels.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
