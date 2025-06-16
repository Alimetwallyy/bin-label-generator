import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import string
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

def generate_bin_labels_table(group_name, bay_ids, shelves, bins_per_shelf):
    data = []
    for bay in bay_ids:
        base_label = bay.replace("BAY-", "")
        base_number = int(base_label[-3:])
        aisle = base_label[9:12]  # extract aisle number from bay ID

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
    return pd.DataFrame(data)

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
            x = col_idx
            y = -i
            ax.text(x, y, bin_label, va='center', ha='center', fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor=shelf_colors[shelf]))

    ax.set_xlim(-0.5, len(shelves) - 0.5)
    ax.set_ylim(-max(bins_per_shelf.values()), 1)
    ax.set_xticks(range(len(shelves)))
    ax.set_xticklabels(shelves)

    return fig

def style_excel(writer, sheet_name, df, shelves):
    ws = writer.sheets[sheet_name]
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    hex_colors = [
        "339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017",
        "FF00FF", "996600", "00FF00", "FF6565", "9999FE"
    ]

    if shelves:
        ws.merge_cells('A1:C1')
        ws['A1'] = "HEX COLOR CODES ->"
        ws['A1'].fill = yellow_fill
        ws['A1'].font = bold_font
        ws['A1'].alignment = center_align
        ws['A1'].border = border

        for i, hex_color in enumerate(hex_colors[:len(shelves)]):
            col_letter = get_column_letter(4 + i)  # D is column 4
            ws[f"{col_letter}1"] = hex_color
            ws[f"{col_letter}1"].fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            ws[f"{col_letter}1"].font = bold_font
            ws[f"{col_letter}1"].alignment = center_align
            ws[f"{col_letter}1"].border = border

            ws[f"{col_letter}2"] = shelves[i]
            ws[f"{col_letter}2"].fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            ws[f"{col_letter}2"].font = bold_font
            ws[f"{col_letter}2"].alignment = center_align
            ws[f"{col_letter}2"].border = border

    # Style header (row 2 if hex row exists, else row 1)
    header_row = 2 if shelves else 1
    for col in range(1, df.shape[1] + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = border

    # Style all data rows
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                cell.font = bold_font
                cell.alignment = center_align
                cell.border = border

def check_duplicate_bay_ids(bay_groups):
    """Check for duplicate bay IDs within and across bay groups."""
    all_bay_ids = set()
    errors = []

    for group_idx, group in enumerate(bay_groups):
        group_name = group["name"]
        bay_ids = group["bays"]
        
        # Check for duplicates within the group
        seen_in_group = set()
        duplicates_in_group = set()
        for bay_id in bay_ids:
            if bay_id in seen_in_group:
                duplicates_in_group.add(bay_id)
            seen_in_group.add(bay_id)
        if duplicates_in_group:
            errors.append(f"⚠️ Duplicate bay IDs found in {group_name}: {', '.join(duplicates_in_group)}")

        # Check for duplicates across groups
        for bay_id in bay_ids:
            if bay_id in all_bay_ids:
                errors.append(f"⚠️ Bay ID {bay_id} is duplicated across multiple groups.")
            all_bay_ids.add(bay_id)

    return errors

# --- Streamlit App ---
st.title("📦 Bin Label Generator")
st.markdown("Define bay groups, shelves, and bins per shelf to generate structured bin labels.")

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

for group_idx in range(num_groups):
    st.header(f"🧱 Bay Group {group_idx + 1}")
    group_name = st.text_input(f"Group Name", value=f"Bay Group {group_idx + 1}", key=f"group_name_{group_idx}")
    bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")
    shelf_count = st.number_input("How many shelves?", min_value=1, max_value=26, value=3, key=f"shelf_count_{group_idx}")
    shelves = list(string.ascii_uppercase[:shelf_count])

    bins_per_shelf = {}
    for shelf in shelves:
        count = st.number_input(f"Number of bins in shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
        bins_per_shelf[shelf] = count

    if bays_input:
        bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
        if bay_list:
            bay_groups.append({
                "name": group_name,
                "bays": bay_list,
                "shelves": shelves,
                "bins_per_shelf": bins_per_shelf
            })

# Check for duplicates and display errors
if bay_groups:
    duplicate_errors = check_duplicate_bay_ids(bay_groups)
    for error in duplicate_errors:
        st.error(error)

if st.button("Generate Bin Labels", disabled=bool(duplicate_errors)):
    if bay_groups:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for group in bay_groups:
                df = generate_bin_labels_table(group["name"], group["bays"], group["shelves"], group["bins_per_shelf"])
                df.to_excel(writer, index=False, startrow=1, sheet_name=group["name"])
                style_excel(writer, group["name"], df, group["shelves"])
        output.seek(0)

        st.success("✅ Bin labels generated successfully!")
        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name="bin_labels.xlsx",
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
    else:
        st.warning("⚠️ Please define at least one bay group.")
