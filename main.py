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

        # --- MODIFICATION START ---
        # Original color list for shelves starting from Shelf B
        hex_colors = [
            "339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017",
            "FF00FF", "996600", "00FF00", "FF6565", "9999FE"
        ]
        
        # New color list for styling, with White ("FFFFFF") added at the beginning for Shelf A
        styling_colors = ["FFFFFF"] + hex_colors
        # --- MODIFICATION END ---


        if shelves:
            ws.merge_cells('A1:C1')
            ws['A1'] = "HEX COLOR CODES ->"
            ws['A1'].fill = yellow_fill
            ws['A1'].font = bold_font
            ws['A1'].alignment = center_align
            ws['A1'].border = border
            
            # The loop now uses the 'styling_colors' list to apply colors to shelves
            for i, hex_color in enumerate(styling_colors[:len(shelves)]):
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
        if not bay_definition:
            raise ValueError("Bay Definition cannot be empty.")
        return {"bay_definition": bay_definition}
    except Exception as e:
        return {"error": str(e)}

def check_duplicate_aisles(mod_groups):
    errors = []
    all_aisles = {}
    for group_idx, group in enumerate(mod_groups):
        mod = group["mod"]
        aisles = list(range(group["aisle_start"], group["aisle_end"] + 1))
        for aisle in aisles:
            aisle_key = f"{mod}-{aisle}"
            if aisle_key in all_aisles:
                errors.append(f"⚠️ Aisle {aisle} in module {mod} is duplicated in module {all_aisles[aisle_key]}.")
            else:
                all_aisles[aisle_key] = mod
    return errors

# --- Streamlit App ---
st.title("Space Launch Quick Tools")
st.markdown("A collection of tools for space launch operations.")

# Create tabs
tab1, tab2, tab3 = st.tabs(["Bin Label Generator", "Bin Bay Mapping", "EOA Generator"])

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
        "*",  # Default option
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
            bin_ids_input = st.text_area(
                f"Enter bin IDs (e.g., P-1-B217A262)",
                key=f"bin_ids_{group_idx}",
                help="Paste Bin IDs from Excel (tab-separated, space-separated, or one per line)."
            )

            # Text input for bay definition (max 48 chars)
            bay_definition = st.text_input(
                "Enter Bay Definition",
                max_chars=48,
                key=f"bay_definition_{group_idx}"
            )

            # Text inputs for dimensions in CM
            height_cm = st.number_input(
                "Height (CM)",
                min_value=0.0,
                value=0.0,
                key=f"height_cm_{group_idx}"
            )
            width_cm = st.number_input(
                "Width (CM)",
                min_value=0.0,
                value=0.0,
                key=f"width_cm_{group_idx}"
            )
            depth_cm = st.number_input(
                "Depth (CM)",
                min_value=0.0,
                value=0.0,
                key=f"depth_cm_{group_idx}"
            )

            # Dropdown for bay usage
            bay_usage = st.selectbox(
                "Select Bay Usage",
                options=bay_usage_options,
                index=0,  # Default to *
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
                bin_list = [b.strip() for line in bin_ids_input.splitlines() for b in re.split(r'[\t\s]+', line) if b.strip()]
                if bin_list:
                    bay_groups.append({
                        "name": st.session_state[f"bin_group_name_{group_idx}"].strip() or f"Bay Definition Group {group_idx + 1}",
                        "bin_ids": bin_list,
                        "bay_definition": bay_definition,
                        "height_cm": height_cm,
                        "width_cm": width_cm,
                        "depth_cm": depth_cm,
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
                cm_to_inch = 0.393701
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
                            "Depth(inch)": round(group["depth_cm"] * cm_to_inch, 2) if group["depth_cm"] else None,
                            "Width(inch)": round(group["width_cm"] * cm_to_inch, 2) if group["width_cm"] else None,
                            "Height(inch)": round(group["height_cm"] * cm_to_inch, 2) if group["height_cm"] else None,
                            "Zone": group["zone"],
                            "Bay Definition": bay_def,
                            "bin_size": f"{int(group['depth_cm'])}Deep" if group["depth_cm"] else "",
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

with tab3:
    st.header("EOA Generator")
    st.markdown("Generate End of Aisle signage based on FC design drawing. Define modules, aisle ranges, slot ranges, and facing aisles.")

    # Input for modules
    modules_input = st.text_area(
        "Enter modules (comma-separated, e.g., P-1-A, P-1-B, P-2-A)",
        key="modules_input",
        help="Enter module IDs including floor (e.g., P-1-A for floor 1, module A)."
    )

    # Input for facing aisles
    facing_aisles_input = st.text_area(
        "Enter facing aisle pairs (comma-separated, e.g., P-1-A-200/P-1-A-201, P-1-B-200/P-1-B-202)",
        key="facing_aisles_input",
        help="Specify which aisles face each other (e.g., P-1-A-200/P-1-A-201 for aisles 200 and 201 in P-1-A)."
    )

    mod_groups = []
    duplicate_errors = []

    if modules_input:
        modules = [mod.strip() for mod in modules_input.split(",") if mod.strip()]
        if modules:
            for mod_idx, mod in enumerate(modules):
                with st.expander(f"Module {mod}", expanded=True):
                    # Aisle range
                    aisle_start = st.number_input(
                        f"Start Aisle for {mod}",
                        min_value=1,
                        value=200,
                        step=1,
                        key=f"aisle_start_{mod_idx}"
                    )
                    aisle_end = st.number_input(
                        f"End Aisle for {mod}",
                        min_value=aisle_start,
                        value=aisle_start,
                        step=1,
                        key=f"aisle_end_{mod_idx}"
                    )

                    # Slot ranges per aisle
                    slot_ranges = {}
                    aisles = list(range(aisle_start, aisle_end + 1))
                    for aisle in aisles:
                        st.markdown(f"**Slot Range for Aisle {aisle}**")
                        slot_start = st.number_input(
                            f"Start Slot for Aisle {aisle}",
                            min_value=1,
                            value=120,
                            step=10,
                            key=f"slot_start_{mod_idx}_{aisle}"
                        )
                        slot_end = st.number_input(
                            f"End Slot for Aisle {aisle}",
                            min_value=slot_start,
                            value=slot_start,
                            step=10,
                            key=f"slot_end_{mod_idx}_{aisle}"
                        )
                        slot_ranges[aisle] = (slot_start, slot_end)

                    mod_groups.append({
                        "mod": mod,
                        "aisle_start": aisle_start,
                        "aisle_end": aisle_end,
                        "slot_ranges": slot_ranges
                    })

    # Parse facing aisles
    facing_pairs = []
    if facing_aisles_input:
        pairs = [pair.strip() for pair in facing_aisles_input.split(",") if pair.strip()]
        for pair in pairs:
            try:
                left, right = pair.split("/")
                left_mod, left_aisle = left.rsplit("-", 1)
                right_mod, right_aisle = right.rsplit("-", 1)
                facing_pairs.append((left_mod, int(left_aisle), right_mod, int(right_aisle)))
            except ValueError:
                st.warning(f"Invalid facing aisle pair: {pair}. Expected format: P-1-A-200/P-1-A-201")

    # Check for duplicate aisles
    if mod_groups:
        duplicate_errors = check_duplicate_aisles(mod_groups)
        with st.expander("Duplicate Errors", expanded=bool(duplicate_errors)):
            if duplicate_errors:
                for error in duplicate_errors:
                    st.warning(error)
            else:
                st.info("No duplicate aisles detected.")

    # Generate signage
    if mod_groups and not duplicate_errors:
        signage_data = []
        processed_aisles = set()

        # Handle user-specified facing aisles
        for left_mod, left_aisle, right_mod, right_aisle in facing_pairs:
            left_slots, right_slots = None, None
            for mod_group in mod_groups:
                if mod_group["mod"] == left_mod and left_aisle in mod_group["slot_ranges"]:
                    left_slots = mod_group["slot_ranges"][left_aisle]
                if mod_group["mod"] == right_mod and right_aisle in mod_group["slot_ranges"]:
                    right_slots = mod_group["slot_ranges"][right_aisle]
            if left_slots and right_slots:
                signage_data.append({
                    "Left.Mod": left_mod,
                    "Left.Aisle": left_aisle,
                    "Left.Slots": f"{left_slots[0]}-{left_slots[1]}",
                    "Right.Mod": right_mod,
                    "Right.Aisle": right_aisle,
                    "Right.Slots": f"{right_slots[0]}-{right_slots[1]}",
                    "Deployment Location": f"Low End of Aisle {left_aisle}/{right_aisle}"
                })
                processed_aisles.add(f"{left_mod}-{left_aisle}")
                processed_aisles.add(f"{right_mod}-{right_aisle}")

        # Handle module aisles
        for mod_group in mod_groups:
            mod = mod_group["mod"]
            aisles = sorted(mod_group["slot_ranges"].keys())
            for i in range(len(aisles)):
                aisle = aisles[i]
                aisle_key = f"{mod}-{aisle}"
                if aisle_key in processed_aisles:
                    continue
                slot_start, slot_end = mod_group["slot_ranges"][aisle]

                if i == 0:  # Starting aisle (one-sided)
                    if aisle % 2 == 0:  # Even
                        low_side = "Right"
                        high_side = "Left"
                    else:  # Odd
                        low_side = "Left"
                        high_side = "Right"

                    # Low End
                    signage_data.append({
                        "Left.Mod": mod if low_side == "Left" else "",
                        "Left.Aisle": aisle if low_side == "Left" else "",
                        "Left.Slots": f"{slot_start}-{slot_end}" if low_side == "Left" else "",
                        "Right.Mod": mod if low_side == "Right" else "",
                        "Right.Aisle": aisle if low_side == "Right" else "",
                        "Right.Slots": f"{slot_start}-{slot_end}" if low_side == "Right" else "",
                        "Deployment Location": f"Low End of Aisle {aisle}"
                    })

                    # High End
                    signage_data.append({
                        "Left.Mod": mod if high_side == "Left" else "",
                        "Left.Aisle": aisle if high_side == "Left" else "",
                        "Left.Slots": f"{slot_end}-{slot_start}" if high_side == "Left" else "",
                        "Right.Mod": mod if high_side == "Right" else "",
                        "Right.Aisle": aisle if high_side == "Right" else "",
                        "Right.Slots": f"{slot_end}-{slot_start}" if high_side == "Right" else "",
                        "Deployment Location": f"High End of Aisle {aisle}"
                    })
                    processed_aisles.add(aisle_key)

                elif i < len(aisles) - 1:  # Pair with next aisle (two-sided)
                    next_aisle = aisles[i + 1]
                    next_aisle_key = f"{mod}-{next_aisle}"
                    if next_aisle_key in processed_aisles:
                        continue
                    next_slot_start, next_slot_end = mod_group["slot_ranges"][next_aisle]

                    # Low End
                    signage_data.append({
                        "Left.Mod": mod,
                        "Left.Aisle": aisle,
                        "Left.Slots": f"{slot_start}-{slot_end}",
                        "Right.Mod": mod,
                        "Right.Aisle": next_aisle,
                        "Right.Slots": f"{next_slot_start}-{next_slot_end}",
                        "Deployment Location": f"Low End of Aisle {aisle}/{next_aisle}"
                    })

                    # High End
                    signage_data.append({
                        "Left.Mod": mod,
                        "Left.Aisle": next_aisle,
                        "Left.Slots": f"{next_slot_end}-{next_slot_start}",
                        "Right.Mod": mod,
                        "Right.Aisle": aisle,
                        "Right.Slots": f"{slot_end}-{slot_start}",
                        "Deployment Location": f"High End of Aisle {aisle}/{next_aisle}"
                    })

                    processed_aisles.add(aisle_key)
                    processed_aisles.add(next_aisle_key)

                else:  # Last aisle (one-sided, if unpaired)
                    if aisle % 2 == 0:  # Even
                        low_side = "Right"
                        high_side = "Left"
                    else:  # Odd
                        low_side = "Left"
                        high_side = "Right"

                    # Low End
                    signage_data.append({
                        "Left.Mod": mod if low_side == "Left" else "",
                        "Left.Aisle": aisle if low_side == "Left" else "",
                        "Left.Slots": f"{slot_start}-{slot_end}" if low_side == "Left" else "",
                        "Right.Mod": mod if low_side == "Right" else "",
                        "Right.Aisle": aisle if low_side == "Right" else "",
                        "Right.Slots": f"{slot_start}-{slot_end}" if low_side == "Right" else "",
                        "Deployment Location": f"Low End of Aisle {aisle}"
                    })

                    # High End
                    signage_data.append({
                        "Left.Mod": mod if high_side == "Left" else "",
                        "Left.Aisle": aisle if high_side == "Left" else "",
                        "Left.Slots": f"{slot_end}-{slot_start}" if high_side == "Left" else "",
                        "Right.Mod": mod if high_side == "Right" else "",
                        "Right.Aisle": aisle if high_side == "Right" else "",
                        "Right.Slots": f"{slot_end}-{slot_start}" if high_side == "Right" else "",
                        "Deployment Location": f"High End of Aisle {aisle}"
                    })
                    processed_aisles.add(aisle_key)

        # Preview signage data
        if signage_data:
            st.subheader("Preview Signage Data")
            df_preview = pd.DataFrame(signage_data)
            st.dataframe(df_preview)

            if st.button("Generate EOA Signage Excel", key="generate_eoa_excel"):
                with st.spinner("Generating Excel file..."):
                    output = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "EOA Signage"

                    # Headers
                    ws.merge_cells("A1:C1")
                    ws["A1"] = "Left Side of Sign"
                    ws["A2"] = "Mod"
                    ws["B2"] = "Aisle"
                    ws["C2"] = "Slots"
                    ws.merge_cells("E1:G1")
                    ws["E1"] = "Right Side of Sign"
                    ws["E2"] = "Mod"
                    ws["F2"] = "Aisle"
                    ws["G2"] = "Slots"
                    ws["H2"] = "Deployment Location"

                    # Styling
                    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
                    white_font = Font(color="FFFFFF", bold=True)
                    center_align = Alignment(horizontal="center", vertical="center")
                    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                        top=Side(style='thin'), bottom=Side(style='thin'))

                    # Apply black background to all cells in A1:H2
                    for row in ws["A1:H2"]:
                        for cell in row:
                            cell.fill = black_fill
                            cell.border = thin_border
                            # Apply font and alignment to cells with content
                            if cell in [ws["A1"], ws["A2"], ws["B2"], ws["C2"], ws["E1"], ws["E2"], ws["F2"], ws["G2"], ws["H2"]]:
                                cell.font = white_font
                                cell.alignment = center_align

                    # Data
                    for row_idx, row_data in enumerate(signage_data, start=3):
                        ws[f"A{row_idx}"] = row_data["Left.Mod"]
                        ws[f"B{row_idx}"] = row_data["Left.Aisle"]
                        ws[f"C{row_idx}"] = row_data["Left.Slots"]
                        ws[f"E{row_idx}"] = row_data["Right.Mod"]
                        ws[f"F{row_idx}"] = row_data["Right.Aisle"]
                        ws[f"G{row_idx}"] = row_data["Right.Slots"]
                        ws[f"H{row_idx}"] = row_data["Deployment Location"]
                        for col in ["A", "B", "C", "E", "F", "G", "H"]:
                            cell = ws[f"{col}{row_idx}"]
                            cell.alignment = center_align
                            cell.border = thin_border

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.success("✅ EOA Signage Excel generated successfully!")
                    st.download_button(
                        label="📥 Download EOA Signage Excel",
                        data=output,
                        file_name="eoa_signage.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_eoa_excel"
                    )

    if not mod_groups:
        st.warning("⚠️ Please define at least one module with valid aisle and slot ranges.")
