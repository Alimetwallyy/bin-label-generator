import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
import seaborn as sns
import string
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# Add "Created By Alimomet" in top left
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@700&display=swap');
    .created-by {
        position: absolute;
        top: 10px;
        left: 20px;
        font-family: 'Roboto', Arial, Helvetica, sans-serif;
        font-size: 14px;
        font-weight: bold;
        color: #333333;
        z-index: 1000;
    }
    </style>
    <div class="created-by">Created By Alimomet</div>
""", unsafe_allow_html=True)

def generate_bin_labels_table(group_name, bay_ids, shelves, bins_per_shelf):
    data = []
    for bay in bay_ids:
        try:
            base_label = bay.replace("BAY-", "")
            base_number = int(base_label[-3:])
            aisle_match = re.search(r'\d{3}', base_label)
            aisle = aisle_match.group(0) if aisle_match else ""

            max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves) if shelves else 1

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
        except Exception as e:
            st.error(f"Error processing bay ID '{bay}': {str(e)}")
    return pd.DataFrame(data)

def plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number):
    try:
        fig = go.Figure()
        colors = sns.color_palette("colorblind", len(shelves) if shelves else 1).as_hex()
        shelf_colors = {shelf: colors[i % len(colors)] for i, shelf in enumerate(shelves)} if shelves else {}

        for col_idx, shelf in enumerate(shelves):
            shelf_bins = bins_per_shelf.get(shelf, 0)
            for i in range(shelf_bins):
                bin_label = bay_id.replace("BAY-", "")[:-4] + shelf + f"{base_number + i:03d}"
                x0, x1 = col_idx - 0.4, col_idx + 0.4
                y0, y1 = -i - 0.4, -i + 0.4
                fig.add_shape(
                    type="rect",
                    x0=x0,
                    x1=x1,
                    y0=y0,
                    y1=y1,
                    fillcolor=shelf_colors.get(shelf, "lightblue"),
                    line=dict(color="black"),
                    label=dict(text=bin_label, textposition="middle center", font=dict(size=10)),
                )
                fig.add_trace(
                    go.Scatter(
                        x=[(x0 + x1) / 2],
                        y=[(y0 + y1) / 2],
                        text=[bin_label],
                        mode="text",
                        hoverinfo="text",
                        showlegend=False,
                    )
                )

        fig.update_layout(
            title=f"Bin Layout for {bay_id}",
            xaxis=dict(
                tickmode="array",
                tickvals=list(range(len(shelves))) if shelves else [0],
                ticktext=shelves if shelves else ["No Shelves"],
                showgrid=False,
                zeroline=False,
            ),
            yaxis=dict(
                showgrid=False,
                zeroline=False,
                autorange="reversed",
            ),
            showlegend=bool(shelves),
            legend_title_text="Shelves",
            width=200 * (len(shelves) if shelves else 1),
            height=100 * (max(bins_per_shelf.values(), default=1) if bins_per_shelf else 1),
            margin=dict(l=20, r=20, t=50, b=20),
        )

        for shelf in shelves:
            fig.add_trace(
                go.Scatter(
                    x=[None],
                    y=[None],
                    mode="markers",
                    name=shelf,
                    marker=dict(size=10, color=shelf_colors.get(shelf, "lightblue")),
                )
            )

        return fig
    except Exception as e:
        st.error(f"Error generating diagram for '{bay_id}': {str(e)}")
        return None

def style_excel(writer, sheet_name, df, shelves):
    try:
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
                col_letter = get_column_letter(4 + i)
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

        header_row = 2 if shelves else 1
        for col in range(1, df.shape[1] + 1):
            cell = ws.cell(row=header_row, column=col)
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = border

        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.font = bold_font
                    cell.alignment = center_align
                    cell.border = border
    except Exception as e:
        st.error(f"Error styling Excel sheet '{sheet_name}': {str(e)}")

def check_duplicate_bay_ids(bay_groups):
    errors = []
    all_bay_ids = {}

    for group_idx, group in enumerate(bay_groups):
        group_name = group["name"]
        bay_ids = [bay_id.strip().upper() for bay_id in group["bays"] if bay_id.strip()]

        seen_in_group = set()
        for bay_id in bay_ids:
            if bay_id in seen_in_group:
                errors.append(f"⚠️ Duplicate bay ID '{bay_id}' found in {group_name}.")
            seen_in_group.add(bay_id)

            if bay_id not in all_bay_ids:
                all_bay_ids[bay_id] = [group_name]
            elif group_name not in all_bay_ids[bay_id]:
                all_bay_ids[bay_id].append(group_name)

    for bay_id, groups in all_bay_ids.items():
        if len(groups) > 1:
            errors.append(f"⚠️ Bay ID '{bay_id}' is duplicated across groups: {', '.join(groups)}.")

    return errors

def check_duplicate_bin_ids(bay_groups):
    errors = []
    all_bin_ids = {}

    for group_idx, group in enumerate(bay_groups):
        group_name = group["name"]
        bin_ids = [bin_id.strip().upper() for bin_id in group["bin_ids"] if bin_id.strip()]

        seen_in_group = set()
        for bin_id in bin_ids:
            if bin_id in seen_in_group:
                errors.append(f"⚠️ Duplicate bin ID '{bin_id}' found in {group_name}.")
            seen_in_group.add(bin_id)

            if bin_id not in all_bin_ids:
                all_bin_ids[bin_id] = [group_name]
            elif group_name not in all_bin_ids[bin_id]:
                all_bin_ids[bin_id].append(group_name)

    for bin_id, groups in all_bin_ids.items():
        if len(groups) > 1:
            errors.append(f"⚠️ Bin ID '{bin_id}' is duplicated across groups: {', '.join(groups)}.")

    return errors

def parse_bay_definition(bay_definition):
    try:
        # Extract dimensions using regex: e.g., "250x253x120cm"
        match = re.match(r'^(\d+)x(\d+)x(\d+)cm', bay_definition)
        if not match:
            raise ValueError("Invalid dimension format. Expected 'HxWxDcm' (e.g., '250x253x120cm').")

        height_cm = float(match.group(1))
        width_cm = float(match.group(2))
        depth_cm = float(match.group(3))

        # Convert cm to inches (1 cm = 0.393701 inches)
        cm_to_inch = 0.393701
        height_inch = round(height_cm * cm_to_inch, 2)
        width_inch = round(width_cm * cm_to_inch, 2)
        depth_inch = round(depth_cm * cm_to_inch, 2)

        # Format depth for bin_size
        bin_size = f"{int(depth_cm)}Deep"

        return {
            "height_inch": height_inch,
            "width_inch": width_inch,
            "depth_inch": depth_inch,
            "bin_size": bin_size
        }
    except Exception as e:
        return {"error": str(e)}

# --- Streamlit App ---
st.title("Space Launch Quick Tools")
st.markdown("A collection of tools for space launch operations.")

# Create tabs
tab1, tab2 = st.tabs(["Bin Label Generator", "Bin Bay Mapping"])

with tab1:
    st.header("Bin Label Generator")
    st.markdown("Define bay groups, shelves, and bins per shelf to generate structured bin labels. Bay IDs must be unique (e.g., BAY-001-001-001).")

    bay_groups = []
    duplicate_errors = []
    num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1, key="num_groups_bin_label")

    for group_idx in range(num_groups):
        # Initialize session state for group name
        if f"group_name_{group_idx}" not in st.session_state:
            st.session_state[f"group_name_{group_idx}"] = f"Bay Group {group_idx + 1}"

        # Callback to force rerender on name change
        def update_group_name(group_idx=group_idx):
            st.session_state[f"group_name_{group_idx}"] = st.session_state[f"group_name_input_{group_idx}"]

        # Use session state for header
        header = st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}"

        with st.expander(header, expanded=True):
            # Text input with on_change callback
            st.text_input(
                "Group Name",
                value=st.session_state[f"group_name_{group_idx}"],
                key=f"group_name_input_{group_idx}",
                on_change=update_group_name
            )

            bays_input = st.text_area(f"Enter bay IDs (one per line, e.g., BAY-001-001-001)", key=f"bays_{group_idx}")
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
                        "name": st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}",
                        "bays": bay_list,
                        "shelves": shelves,
                        "bins_per_shelf": bins_per_shelf
                    })
                    temp_errors = check_duplicate_bay_ids(bay_groups)
                    if temp_errors:
                        with st.container():
                            st.markdown("**Errors in this group:**")
                            for error in temp_errors:
                                st.warning(error)

    if bay_groups:
        duplicate_errors = check_duplicate_bay_ids(bay_groups)
        with st.expander("Duplicate Errors", expanded=bool(duplicate_errors)):
            if duplicate_errors:
                for error in duplicate_errors:
                    st.warning(error)
            else:
                st.info("No duplicate bay IDs detected.")
    else:
        st.warning("⚠️ Please define at least one bay group with valid bay IDs.")

    if st.button("Generate Bin Labels", disabled=bool(duplicate_errors or not bay_groups), key="generate_bin_labels"):
        with st.spinner("Generating bin labels and diagrams..."):
            output = io.BytesIO()
            try:
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for group in bay_groups:
                        df = generate_bin_labels_table(group["name"], group["bays"], group["shelves"], group["bins_per_shelf"])
                        if not df.empty:
                            df.to_excel(writer, index=False, startrow=1, sheet_name=group["name"])
                            style_excel(writer, group["name"], df, group["shelves"])
                output.seek(0)

                st.success("✅ Bin labels generated successfully!")
                st.download_button(
                    label="📥 Download Excel File",
                    data=output,
                    file_name="bin_labels.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )

                st.subheader("🖼️ Interactive Bin Layout Diagrams")
                for group in bay_groups:
                    for bay_id in group['bays']:
                        shelves = group['shelves']
                        bins_per_shelf = group['bins_per_shelf']
                        try:
                            base_label = bay_id.replace("BAY-", "")
                            base_number = int(base_label[-3:])
                            fig = plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number)
                            if fig:
                                st.plotly_chart(fig, use_container_width=True)
                        except Exception as e:
                            st.error(f"Error processing bay ID '{bay_id}': {str(e)}")
            except Exception as e:
                st.error(f"Error generating output: {str(e)}")

with tab2:
    st.header("Bin Bay Mapping")
    st.markdown("Define bay definition groups and map bin IDs to bay types.")

    # List of unique bay types for dropdown
    bay_types = [
        "Bulk Stock", "Case Flow", "Drawer", "Flat Apparel", "Hanger Rod", "Hangers",
        "Jewelry", "Library", "Library Deep", "Pallet", "Shoes", "Random Other Bin",
        "PassThrough"
    ]

    # List of bay usage options for dropdown
    bay_usage_options = [
        "45F Produce", "Aerosol", "Ambient", "Apparel", "BATTERIES", "BWS",
        "BWS_HIGH_FLAMMABLE", "BWS_LOW_FLAMMABLE", "BWS_MEDIUM_FLAMMABLE", "Book",
        "Chilled", "Chilled-FMP", "Corrosive", "Damage", "Damage Human Food",
        "Damage Pet Food", "Damage_HRV", "Damaged Aerosol", "Damaged Corrosive",
        "Damaged Flammable", "Damaged Flammable Aerosols", "Damaged Misc Health Hazard",
        "Damaged Non Flammable Aerosols", "Damaged Oxidizer", "Damaged Restricted Hazmat",
        "Damaged Toxic", "Dry Produce", "FMP", "Flammable", "Flammable Aerosols",
        "Flammables_HRV", "Frozen", "HRV", "Hazmat", "Hazmat_HRV", "Meat-Beef",
        "Meat-Deli", "Meat-Pork", "Meat-Poultry", "Meat-Seafood", "Misc Health Hazard",
        "Non Flammable Aerosols", "Non Inventory Storage-Facilities",
        "Non Inventory Storage-Other", "Non Inventory Storage-Stores",
        "Non Inventory-Black Totes", "Non Sort-Team Lift", "Non-Storage",
        "Non-TC Food", "Oxidizer", "Pet Food", "Produce", "Produce Backstock",
        "Produce Wetracks", "Reserve-Ambient", "Restricted Hazmat", "Semi-Chilled",
        "Shoes", "TC-Food", "Toxic", "Tropical"
    ]

    num_groups = st.number_input("How many bay definition groups do you want to define?", min_value=1, max_value=10, value=1, key="num_groups_bin_mapping")

    bay_groups = []
    for group_idx in range(num_groups):
        # Initialize session state for group name
        if f"bin_group_name_{group_idx}" not in st.session_state:
            st.session_state[f"bin_group_name_{group_idx}"] = f"Bay Definition Group {group_idx + 1}"

        # Callback to force rerender on name change
        def update_bin_group_name(group_idx=group_idx):
            st.session_state[f"bin_group_name_{group_idx}"] = st.session_state[f"bin_group_name_input_{group_idx}"]

        # Use session state for header
        header = st.session_state[f"bin_group_name_{group_idx}"].strip() or f"Bay Definition Group {group_idx + 1}"

        with st.expander(header, expanded=True):
            # Text input for group name
            st.text_input(
                "Group Name",
                value=st.session_state[f"bin_group_name_{group_idx}"],
                key=f"bin_group_name_input_{group_idx}",
                on_change=update_bin_group_name
            )

            # Text area for bin IDs
            bin_ids_input = st.text_area(f"Enter bin IDs (one per line, e.g., BIN-001-001-001)", key=f"bin_ids_{group_idx}")

            # Text input for bay definition (max 48 chars)
            bay_definition = st.text_input(
                "Enter Bay Definition",
                max_chars=48,
                key=f"bay_definition_{group_idx}"
            )

            # Dropdown for bay usage
            bay_usage = st.selectbox(
                "Select Bay Usage",
                options=bay_usage_options,
                index=0,  # Default to 45F Produce
                key=f"bay_usage_{group_idx}"
            )

            # Dropdown for bay type
            bay_type = st.selectbox(
                "Select Bay Type",
                options=bay_types,
                index=0,  # Default to Bulk Stock
                key=f"bay_type_{group_idx}"
            )

            # Description and text input for zone
            st.markdown("Enter Zone bins are inside followed by depth of bays. ex: Library (30D)")
            zone = st.text_input(
                "Zone",
                max_chars=25,
                key=f"zone_{group_idx}"
            )

            if bin_ids_input:
                bin_list = [b.strip() for b in bin_ids_input.splitlines() if b.strip()]
                if bin_list:
                    bay_groups.append({
                        "name": st.session_state[f"bin_group_name_{group_idx}"].strip() or f"Bay Definition Group {group_idx + 1}",
                        "bin_ids": bin_list,
                        "bay_definition": bay_definition,
                        "bay_usage": bay_usage,
                        "bay_type": bay_type,
                        "zone": zone
                    })
                    temp_errors = check_duplicate_bin_ids(bay_groups)
                    if temp_errors:
                        with st.container():
                            st.markdown("**Errors in this group:**")
                            for error in temp_errors:
                                st.warning(error)

    if bay_groups:
        duplicate_errors = check_duplicate_bin_ids(bay_groups)
        with st.expander("Duplicate Errors", expanded=bool(duplicate_errors)):
            if duplicate_errors:
                for error in duplicate_errors:
                    st.warning(error)
            else:
                st.info("No duplicate bin IDs detected.")
    else:
        st.warning("⚠️ Please define at least one bay definition group with valid bin IDs.")

    if st.button("Generate Excel", disabled=bool(duplicate_errors or not bay_groups), key="generate_bin_mapping_excel"):
        with st.spinner("Generating Excel file..."):
            output = io.BytesIO()
            try:
                data = []
                for group in bay_groups:
                    bay_def = group["bay_definition"]
                    parsed = parse_bay_definition(bay_def)
                    if "error" in parsed:
                        st.error(f"Invalid bay definition in {group['name']}: {parsed['error']}")
                        break

                    for bin_id in group["bin_ids"]:
                        data.append({
                            "ScannableId": bin_id,
                            "Distance Index": None,
                            "Depth(inch)": parsed["depth_inch"],
                            "Width(inch)": parsed["width_inch"],
                            "Height(inch)": parsed["height_inch"],
                            "Zone": group["zone"],
                            "Bay Definition": bay_def,
                            "bin_size": parsed["bin_size"],
                            "Bay Type": group["bay_type"],
                            "Bay Usage": group["bay_usage"]
                        })
                else:  # Execute if no break (i.e., no errors)
                    df = pd.DataFrame(data)
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Bin Bay Mapping")
                    output.seek(0)

                    st.success("✅ Excel file generated successfully!")
                    st.download_button(
                        label="📥 Download Excel File",
                        data=output,
                        file_name="bin_bay_mapping.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_bin_mapping_excel"
                    )
            except Exception as e:
                st.error(f"Error generating Excel: {str(e)}")
