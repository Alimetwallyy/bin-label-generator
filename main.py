import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
import string

# --- 🔧 Generate bin labels ---
def generate_bin_labels_table(bay_groups):
    data = []
    for group in bay_groups:
        bay_ids = group['bays']
        shelves = group['shelves']
        bins_per_shelf = group['bins_per_shelf']
        group_name = group['group_name']

        for bay in bay_ids:
            base_label = bay.replace("BAY-", "")
            base_number = int(base_label[-3:]) if base_label[-3:].isdigit() else 0
            max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

            for i in range(max_bins):
                row = {'Bay_Group': group_name, 'Bay_ID': bay}
                for shelf in shelves:
                    shelf_bin_count = bins_per_shelf.get(shelf, 0)
                    if i < shelf_bin_count:
                        bin_label = base_label[:-4] + shelf + f"{base_number + i:03d}"
                        row[f"Shelf_{shelf}"] = bin_label
                    else:
                        row[f"Shelf_{shelf}"] = None
                data.append(row)
    return pd.DataFrame(data)

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
            x = col_idx
            y = -i
            ax.text(x, y, bin_label, va='center', ha='center', fontsize=8,
                    bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor=shelf_colors[shelf]))

    ax.set_xlim(-0.5, len(shelves) - 0.5)
    ax.set_ylim(-max(bins_per_shelf.values()), 1)
    ax.set_xticks(range(len(shelves)))
    ax.set_xticklabels(shelves)

    return fig

# --- 🖥️ Streamlit UI ---
st.title("📦 Bin Label Generator")
st.markdown("Define bay groups, number of shelves, and bins per shelf to generate structured bin labels.")

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

all_bay_ids = {}
duplicates_to_highlight = set()

for group_idx in range(num_groups):
    st.header(f"🧱 Bay Group {group_idx + 1}")
    group_name = st.text_input("Group name", value=f"Group {group_idx + 1}", key=f"group_name_{group_idx}")

    bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")

    num_shelves = st.number_input(f"How many shelves? (max 26)", min_value=1, max_value=26, value=3, key=f"num_shelves_{group_idx}")
    shelves = list(string.ascii_uppercase[:num_shelves])

    bins_per_shelf = {}
    for shelf in shelves:
        count = st.number_input(f"Number of bins in shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
        bins_per_shelf[shelf] = count

    bay_list_raw = [b.strip() for b in bays_input.splitlines() if b.strip()]
    seen_in_group = set()
    skipped_duplicates_within = []
    skipped_duplicates_across = []

    for bay in bay_list_raw:
        if bay in seen_in_group:
            skipped_duplicates_within.append(bay)
        elif bay in all_bay_ids:
            skipped_duplicates_across.append((bay, all_bay_ids[bay]))
        seen_in_group.add(bay)

    for bay in bay_list_raw:
        all_bay_ids[bay] = group_name

    if skipped_duplicates_within:
        st.warning(f"⚠️ Duplicate bay IDs *within this group*: {', '.join(set(skipped_duplicates_within))}")
    if skipped_duplicates_across:
        for bay, other_group in skipped_duplicates_across:
            st.warning(f"⚠️ Bay ID `{bay}` already exists in **{other_group}** – consider removing or renaming.")

    duplicates_to_highlight.update(skipped_duplicates_within + [b for b, _ in skipped_duplicates_across])

    if bay_list_raw:
        bay_groups.append({
            "group_name": group_name,
            "bays": bay_list_raw,
            "shelves": shelves,
            "bins_per_shelf": bins_per_shelf
        })

# --- 🚀 Generate Bin Labels ---
if st.button("✅ Generate Bin Labels"):
    if bay_groups:
        df = generate_bin_labels_table(bay_groups)
        st.success("✅ Bin labels generated successfully!")
        st.dataframe(df)

        # --- Excel export with highlights ---
        wb = Workbook()
        ws = wb.active
        ws.title = "Bin Labels"
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):
            for cell in row:
                if cell.value in duplicates_to_highlight:
                    cell.fill = red_fill

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download Excel File (with highlights)",
            data=output,
            file_name="bin_labels_highlighted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("🖼️ Bin Layout Diagrams")
        for group in bay_groups:
            for bay_id in group['bays']:
                shelves = group['shelves']
                bins_per_shelf = group['bins_per_shelf']
                base_label = bay_id.replace("BAY-", "")
                base_number = int(base_label[-3:]) if base_label[-3:].isdigit() else 0
                fig = plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number)
                st.pyplot(fig)
    else:
        st.warning("⚠️ Please define at least one valid bay group.")
