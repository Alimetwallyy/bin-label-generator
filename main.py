import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
import seaborn as sns
import string
import re
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# Custom CSS for improved UI
st.markdown("""
    <style>
    .stTextInput > div > div > input {
        border: 1px solid #ccc;
    }
    .stTextInput > div > div > input.invalid {
        border: 2px solid red;
    }
    .stTextArea > div > div > textarea {
        border: 1px solid #ccc;
    }
    .stTextArea > div > div > textarea.invalid {
        border: 2px solid red;
    }
    .stContainer > div {
        padding: 10px;
        border-radius: 5px;
    }
    .error-container {
        background-color: #fff5f5;
        border: 1px solid #ffcccc;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
    }
    .stButton > button {
        margin-right: 10px;
    }
    label {
        font-weight: bold;
        color: #333;
    }
    </style>
""", unsafe_allow_html=True)

def generate_bin_labels_table(group_name, bay_ids, shelves, bins_per_shelf):
    data = []
    for bay in bay_ids:
        try:
            base_label = bay.replace("BAY-", "")
            base_number = int(base_label[-3:])
            aisle = base_label[9:12] if len(base_label) >= 12 else ""

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

def validate_bay_ids(bay_ids):
    pattern = r'^BAY-\d{3}-\d{3}-\d{3}$'
    invalid_ids = []
    for bay_id in bay_ids:
        if not re.match(pattern, bay_id.strip()):
            invalid_ids.append(bay_id.strip())
    return invalid_ids

# --- Streamlit App ---
st.title("📦 Bin Label Generator")
st.markdown("Define bay groups, shelves, and bins to generate bin labels. Bay IDs must be unique (e.g., BAY-001-001-001).")

# Toggle for showing all diagrams
if 'show_all_diagrams' not in st.session_state:
    st.session_state['show_all_diagrams'] = False

bay_groups = []
duplicate_errors = []
num_groups = st.number_input(
    "How many bay groups do you want to define?",
    min_value=1,
    max_value=10,
    value=1,
    help="Select the number of bay groups to configure."
)

for group_idx in range(num_groups):
    # Initialize session state
    if f"group_name_{group_idx}" not in st.session_state:
        st.session_state[f"group_name_{group_idx}"] = f"Bay Group {group_idx + 1}"
    if f"bays_{group_idx}" not in st.session_state:
        st.session_state[f"bays_{group_idx}"] = ""
    if f"shelf_count_{group_idx}" not in st.session_state:
        st.session_state[f"shelf_count_{group_idx}"] = 3
    for shelf in string.ascii_uppercase[:st.session_state[f"shelf_count_{group_idx}"]]:
        if f"bins_{group_idx}_{shelf}" not in st.session_state:
            st.session_state[f"bins_{group_idx}_{shelf}"] = 5

    # Reset group callback
    def reset_group(group_idx=group_idx):
        st.session_state[f"group_name_{group_idx}"] = f"Bay Group {group_idx + 1}"
        st.session_state[f"bays_{group_idx}"] = ""
        st.session_state[f"shelf_count_{group_idx}"] = 3
        for shelf in string.ascii_uppercase[:26]:
            if f"bins_{group_idx}_{shelf}" in st.session_state:
                st.session_state[f"bins_{group_idx}_{shelf}"] = 5

    # Update group name callback
    def update_group_name(group_idx=group_idx):
        st.session_state[f"group_name_{group_idx}"] = st.session_state[f"group_name_input_{group_idx}"]

    # Header
    header = st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}"

    with st.expander(header, expanded=True):
        # Group name input
        st.text_input(
            "Group Name",
            value=st.session_state[f"group_name_{group_idx}"],
            key=f"group_name_input_{group_idx}",
            on_change=update_group_name,
            help="Enter a name for this bay group (e.g., Warehouse A)."
        )

        # File uploader for bay IDs
        uploaded_file = st.file_uploader(
            f"Upload bay IDs (CSV or TXT, one per line)",
            type=["csv", "txt"],
            key=f"file_uploader_{group_idx}",
            help="Upload a CSV (single column) or TXT file with bay IDs like BAY-001-001-001."
        )
        if uploaded_file:
            if uploaded_file.type == "text/csv":
                df = pd.read_csv(uploaded_file, header=None)
                bay_ids = df.iloc[:, 0].dropna().astype(str).tolist()
            else:  # TXT
                bay_ids = uploaded_file.read().decode("utf-8").splitlines()
            bay_ids = [bid.strip() for bid in bay_ids if bid.strip()]
            st.session_state[f"bays_{group_idx}"] = "\n".join(bay_ids)

        # Bay IDs input
        bays_input = st.text_area(
            "Enter bay IDs (one per line, e.g., BAY-001-001-001)",
            value=st.session_state[f"bays_{group_idx}"],
            key=f"bays_{group_idx}",
            help="Enter unique bay IDs in the format BAY-XXX-XXX-XXX."
        )
        bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
        
        # Validate bay IDs
        if bay_list:
            invalid_ids = validate_bay_ids(bay_list)
            if invalid_ids:
                st.markdown(
                    f'<style>textarea[key="bays_{group_idx}"] {{ border: 2px solid red; }}</style>',
                    unsafe_allow_html=True
                )
                st.error(f"Invalid bay IDs: {', '.join(invalid_ids)}. Use format BAY-XXX-XXX-XXX (e.g., BAY-001-001-001).")
            else:
                st.markdown(
                    f'<style>textarea[key="bays_{group_idx}"] {{ border: 1px solid #ccc; }}</style>',
                    unsafe_allow_html=True
                )

        # Shelf and bin inputs
        shelf_count = st.number_input(
            "How many shelves?",
            min_value=1,
            max_value=26,
            value=st.session_state[f"shelf_count_{group_idx}"],
            key=f"shelf_count_{group_idx}",
            help="Select the number of shelves (A-Z)."
        )
        shelves = list(string.ascii_uppercase[:shelf_count])

        bins_per_shelf = {}
        for shelf in shelves:
            count = st.number_input(
                f"Number of bins in shelf {shelf}",
                min_value=1,
                max_value=100,
                value=st.session_state[f"bins_{group_idx}_{shelf}"],
                key=f"bins_{group_idx}_{shelf}",
                help=f"Set the number of bins for shelf {shelf}."
            )
            bins_per_shelf[shelf] = count

        # Reset button
        if st.button("Reset Group", key=f"reset_{group_idx}"):
            reset_group()

        # Duplicate check
        if bay_list and not invalid_ids:
            bay_groups.append({
                "name": st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}",
                "bays": bay_list,
                "shelves": shelves,
                "bins_per_shelf": bins_per_shelf
            })
            temp_errors = check_duplicate_bay_ids(bay_groups)
            if temp_errors:
                with st.container():
                    st.markdown('<div class="error-container">**Errors in this group:**</div>', unsafe_allow_html=True)
                    for error in temp_errors:
                        st.warning(error)
            bay_groups.pop()  # Remove for global check later

# Global duplicate errors
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

# Generate button
if st.button("Generate Bin Labels", disabled=bool(duplicate_errors or not bay_groups or any(validate_bay_ids(group["bays"]) for group in bay_groups))):
    with st.spinner("Generating bin labels and diagrams..."):
        progress_bar = st.progress(0)
        total_steps = sum(len(group["bays"]) for group in bay_groups) + len(bay_groups)  # Excel + diagrams
        current_step = 0

        # Excel generation
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for group in bay_groups:
                    df = generate_bin_labels_table(group["name"], group["bays"], group["shelves"], group["bins_per_shelf"])
                    if not df.empty:
                        df.to_excel(writer, index=False, startrow=1, sheet_name=group["name"])
                        style_excel(writer, group["name"], df, group["shelves"])
                    current_step += 1
                    progress_bar.progress(min(current_step / total_steps, 1.0))
            output.seek(0)

            st.success("✅ Bin labels generated successfully!")
            st.download_button(
                label="📥 Download Excel File",
                data=output,
                file_name="bin_labels.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error generating Excel: {str(e)}")

        # Diagram generation
        st.subheader("🖼️ Interactive Bin Layout Diagrams")
        if st.button("Toggle All Diagrams", key="toggle_diagrams"):
            st.session_state['show_all_diagrams'] = not st.session_state['show_all_diagrams']

        for group in bay_groups:
            for bay_id in group['bays']:
                shelves = group['shelves']
                bins_per_shelf = group['bins_per_shelf']
                try:
                    base_label = bay_id.replace("BAY-", "")
                    base_number = int(base_label[-3:])
                    fig = plot_bin_diagram(bay_id, shelves, bins_per_shelf, base_number)
                    if fig:
                        with st.expander(f"Diagram for {bay_id}", expanded=st.session_state['show_all_diagrams']):
                            st.plotly_chart(fig, use_container_width=True)
                            # PNG download
                            img_bytes = pio.to_image(fig, format="png")
                            st.download_button(
                                label="📷 Download Diagram as PNG",
                                data=img_bytes,
                                file_name=f"bin_layout_{bay_id}.png",
                                mime="image/png",
                                key=f"download_png_{bay_id}"
                            )
                    current_step += 1
                    progress_bar.progress(min(current_step / total_steps, 1.0))
                except Exception as e:
                    st.error(f"Error processing bay ID '{bay_id}': {str(e)}")
