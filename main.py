import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import string
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- 🔧 Generate bin labels ---
def generate_bin_labels_table(bay_groups):
    data = []

    for group in bay_groups:
        group_name = group['name']
        bay_ids = group['bays']
        shelves = group['shelves']
        bins_per_shelf = group['bins_per_shelf']

        for bay in bay_ids:
            base_label = bay.replace("BAY-", "")
            base_number = int(base_label[-3:])
            aisle = ''.join(filter(str.isdigit, base_label[9:12]))

            max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

            for i in range(max_bins):
                row = {
                    'Bay_Type': group_name,
                    'Aisle': aisle,
                    'Bay_ID': bay
                }
                for shelf in shelves:
                    shelf_bin_count = bins_per_shelf.get(shelf, 0)
                    if i < shelf_bin_count:
                        bin_label = base_label[:-4] + shelf + f"{base_number + i:03d}"
                        row[f"{shelf}"] = bin_label
                    else:
                        row[f"{shelf}"] = None
                data.append(row)

    return pd.DataFrame(data)

# --- 📊 Draw bin diagram ---
def plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number):
    fig, ax = plt.subplots(figsize=(len(shelves) * 2, max(bins_per_shelf.values()) * 0.6))
    ax.set_title(f"Bin Layout for {bay_id}", fontsize=14)
    ax.axis('off')
    colors = ['#339900', '#9B30FF', '#FFFF00', '#00FFFF', '#CC0000', '#F88017', '#FF00FF', '#996600', '#00FF00', '#FF6565', '#9999FE']
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

# --- 📁 Format Excel ---
def style_excel(writer, df, group_name, shelves):
    workbook = writer.book
    worksheet = writer.sheets[group_name]

    # Bold font and border styles
    bold_font = Font(bold=True)
    border = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )

    # Prepare shelf list (remove shelves with 0 bins)
    valid_shelves = [shelf for shelf in shelves if any(col for col in df.columns if col == shelf)]

    # Hex colors
    hex_colors = ['339900', '9B30FF', 'FFFF00', '00FFFF', 'CC0000', 'F88017', 'FF00FF', '996600', '00FF00', 'FF6565', '9999FE']

    if valid_shelves:
        # Row 1: "HEX COLOR CODES ->"
        worksheet.merge_cells(start_row=1, start_column=3, end_row=1, end_column=3)
        cell = worksheet.cell(row=1, column=3)
        cell.value = "HEX COLOR CODES ->"
        cell.fill = PatternFill(start_color='FFFF00', fill_type='solid')
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')

        for i, color in enumerate(hex_colors[:len(valid_shelves)]):
            col = 4 + i
            worksheet.cell(row=1, column=col, value=color).fill = PatternFill(start_color=color, fill_type='solid')
            worksheet.cell(row=2, column=col, value=valid_shelves[i]).fill = PatternFill(start_color=color, fill_type='solid')
            worksheet.cell(row=1, column=col).font = bold_font
            worksheet.cell(row=2, column=col).font = bold_font

    # Headers in row 2
    for col_num, column_title in enumerate(df.columns, 1):
        cell = worksheet.cell(row=2, column=col_num)
        cell.value = column_title
        cell.font = bold_font
        cell.border = border

    # Write data starting from row 3
    for row_num, row in enumerate(df.values, 3):
        for col_num, cell_val in enumerate(row, 1):
            cell = worksheet.cell(row=row_num, column=col_num, value=cell_val)
            cell.border = border

# --- 🖥️ Streamlit UI ---
st.title("📦 Bin Label Generator")
st.markdown("Define bay groups, shelves, and bins per shelf to generate structured bin labels.")

if "group_names" not in st.session_state:
    st.session_state.group_names = []

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

duplicate_error = False
all_bays = set()

for group_idx in range(num_groups):
    st.header(f"🧱 Bay Group {group_idx + 1}")

    group_name = st.text_input(f"Enter a name for Bay Group {group_idx + 1}", key=f"group_name_{group_idx}")
    bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")
    num_shelves = st.number_input(f"Number of shelves (A-Z)", min_value=1, max_value=26, value=3, key=f"num_shelves_{group_idx}")
    shelves = list(string.ascii_uppercase[:num_shelves])
    bins_per_shelf = {}
    for shelf in shelves:
        count = st.number_input(f"Number of bins in shelf {shelf}", min_value=0, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
        bins_per_shelf[shelf] = count

    if bays_input:
        bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
        local_duplicates = set([b for b in bay_list if bay_list.count(b) > 1])
        cross_duplicates = set(bay_list) & all_bays

        if local_duplicates:
            st.error(f"❌ Duplicate bay IDs within this group: {', '.join(local_duplicates)}")
            duplicate_error = True
        elif cross_duplicates:
            st.error(f"❌ Duplicate bay IDs used in other groups: {', '.join(cross_duplicates)}")
            duplicate_error = True

        all_bays.update(bay_list)
        st.session_state.group_names.append(group_name)
        bay_groups.append({
            "name": group_name,
            "bays": bay_list,
            "shelves": shelves,
            "bins_per_shelf": bins_per_shelf
        })

if st.button("Generate Bin Labels") and not duplicate_error:
    if bay_groups:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for group in bay_groups:
                df = generate_bin_labels_table([group])
                df.to_excel(writer, sheet_name=group['name'], startrow=2, index=False)
                style_excel(writer, df, group['name'], group['shelves'])
        output.seek(0)

        st.success("✅ Excel file generated successfully!")
        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name="bin_labels.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("🖼️ Bin Layout Diagrams")
        for group in bay_groups:
            for bay_id in group['bays']:
                shelves = [s for s in group['shelves'] if group['bins_per_shelf'].get(s, 0) > 0]
                bins_per_shelf = {k: v for k, v in group['bins_per_shelf'].items() if v > 0}
                base_label = bay_id.replace("BAY-", "")
                base_number = int(base_label[-3:])
                fig = plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number)
                st.pyplot(fig)
    else:
        st.warning("⚠️ Please define at least one bay group.")
