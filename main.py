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
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.enum.shapes import MSO_SHAPE

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

# --- All Helper Functions ---

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
                row = {'BAY TYPE': group_name, 'AISLE': aisle, 'BAY ID': bay}
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
                fig.add_shape(type="rect", x0=x0, y0=y0, x1=x1, y1=y1, fillcolor=shelf_colors.get(shelf, "lightblue"), line=dict(color="black"))
                fig.add_trace(go.Scatter(x=[(x0 + x1) / 2], y=[(y0 + y1) / 2], text=[bin_label], mode="text", hoverinfo="text", showlegend=False))
        fig.update_layout(title=f"Bin Layout for {bay_id}", xaxis=dict(tickmode="array", tickvals=list(range(len(shelves))), ticktext=shelves if shelves else ["No Shelves"], showgrid=False, zeroline=False), yaxis=dict(showgrid=False, zeroline=False, autorange="reversed"), showlegend=bool(shelves), legend_title_text="Shelves", width=200 * (len(shelves) if shelves else 1), height=100 * (max(bins_per_shelf.values(), default=1) if bins_per_shelf else 1), margin=dict(l=20, r=20, t=50, b=20))
        for shelf in shelves:
            fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", name=shelf, marker=dict(size=10, color=shelf_colors.get(shelf, "lightblue"))))
        return fig
    except Exception as e:
        st.error(f"Error generating diagram for '{bay_id}': {str(e)}")
        return None

def style_excel(writer, sheet_name, df, shelves):
    ws = writer.sheets[sheet_name]
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    hex_colors = ["339900", "9B30FF", "FFFF00", "00FFFF", "CC0000", "F88017", "FF00FF", "996600", "00FF00", "FF6565", "9999FE"]
    styling_colors = ["FFFFFF"] + hex_colors
    if shelves:
        ws.merge_cells('A1:C1')
        ws['A1'] = "HEX COLOR CODES ->"
        ws['A1'].fill = yellow_fill
        ws['A1'].font = bold_font
        ws['A1'].alignment = center_align
        ws['A1'].border = border
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

def generate_elevation_powerpoint(bay_types_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6] 
    
    for bay_type in bay_types_data:
        slide = prs.slides.add_slide(blank_slide_layout)
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.5))
        title_text = title_shape.text_frame
        p = title_text.paragraphs[0]
        p.text = bay_type['name']
        p.font.bold = True
        p.font.size = Pt(24)

        start_x = Inches(2.0)
        start_y = Inches(7.0) 
        scale = Cm(0.2)
        current_y = start_y
        
        for i, shelf_name in enumerate(reversed(bay_type['shelves'])):
            shelf_info = bay_type['shelf_details'][shelf_name]
            num_bins = shelf_info['num_bins']
            bin_h = shelf_info['h'] * scale
            bin_w = shelf_info['w'] * scale
            shelf_height = bin_h
            shelf_width = num_bins * bin_w

            slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start_x, current_y - shelf_height, shelf_width, shelf_height)
            for j in range(num_bins):
                bin_x = start_x + (j * bin_w)
                slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, bin_x, current_y - shelf_height, bin_w, shelf_height)

            label_box = slide.shapes.add_textbox(start_x - Inches(0.6), current_y - shelf_height, Inches(0.5), shelf_height)
            label_box.text_frame.text = shelf_name
            label_box.text_frame.paragraphs[0].font.size = Pt(10)

            is_first_of_dim = True
            if i > 0:
                prev_shelf_name = list(reversed(bay_type['shelves']))[i-1]
                if bay_type['shelf_details'][prev_shelf_name] == shelf_info:
                    is_first_of_dim = False
            
            if is_first_of_dim:
                dim_x = start_x + shelf_width + Inches(0.2)
                dim_y = current_y - (shelf_height / 2)
                dim_text = f"H: {shelf_info['h']}cm\nW: {shelf_info['w']}cm\nD: {shelf_info['d']}cm"
                dim_box = slide.shapes.add_textbox(dim_x, dim_y - Inches(0.25), Inches(1.5), Inches(0.5))
                tf = dim_box.text_frame
                tf.text = dim_text
                for para in tf.paragraphs:
                    para.font.size = Pt(9)

            current_y -= (shelf_height + Cm(0.2))

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

# --- Main App ---
st.set_page_config(layout="wide")

st.title("Space Launch Quick Tools")
st.markdown("A collection of tools for space launch operations.")

if 'eoa_placement_rule' not in st.session_state:
    st.session_state.eoa_placement_rule = "Odd on Left / Even on Right"

tab1, tab2, tab3, tab4 = st.tabs(["Bin Label Generator", "Bin Bay Mapping", "EOA Generator", "Bay Elevation Generator"])

with tab1:
    st.header("Bin Label Generator 🏷️", divider='rainbow')
    st.markdown("Define bay groups, shelves, and bins per shelf to generate structured bin labels.")
    
    bay_groups_tab1 = []
    num_groups_tab1 = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=50, value=1, key="num_groups_bin_label")

    for group_idx in range(num_groups_tab1):
        if f"group_name_{group_idx}" not in st.session_state:
            st.session_state[f"group_name_{group_idx}"] = f"Bay Group {group_idx + 1}"
        def update_group_name(idx=group_idx):
            st.session_state[f"group_name_{idx}"] = st.session_state[f"group_name_input_{idx}"]
        header = st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}"
        
        with st.expander(header, expanded=True):
            st.text_input("Group Name", value=st.session_state[f"group_name_{group_idx}"], key=f"group_name_input_{group_idx}", on_change=update_group_name, args=(group_idx,))
            bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")
            shelf_count = st.number_input("How many shelves?", min_value=1, max_value=26, value=3, key=f"shelf_count_{group_idx}")
            shelves = list(string.ascii_uppercase[:shelf_count])
            st.divider()
            bins_per_shelf = {}
            st.markdown("**Bins per Shelf**")
            for shelf in shelves:
                bins_per_shelf[shelf] = st.number_input(f"Number of bins in shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
            
            if bays_input:
                bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
                if bay_list:
                    bay_groups_tab1.append({"name": st.session_state[f"group_name_{group_idx}"].strip() or f"Bay Group {group_idx + 1}", "bays": bay_list, "shelves": shelves, "bins_per_shelf": bins_per_shelf})

    if st.button("Generate Bin Labels", key="generate_bin_labels"):
        # Generation Logic for Tab 1
        with st.spinner("Generating bin labels and diagrams..."):
            total_labels_generated = 0
            total_bays_processed = 0
            output = io.BytesIO()
            try:
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for group in bay_groups_tab1:
                        df = generate_bin_labels_table(group["name"], group["bays"], group["shelves"], group["bins_per_shelf"])
                        if not df.empty:
                            shelf_cols = [col for col in df.columns if col not in ['BAY TYPE', 'AISLE', 'BAY ID']]
                            total_labels_generated += df[shelf_cols].count().sum()
                            total_bays_processed += df['BAY ID'].nunique()
                            df.to_excel(writer, index=False, startrow=1, sheet_name=group["name"])
                            style_excel(writer, group["name"], df, group["shelves"])
                output.seek(0)
                st.success(f"✅ Success! Generated {total_labels_generated} labels for {total_bays_processed} bays across {len(bay_groups_tab1)} groups.")
                st.download_button(label="📥 Download Excel File", data=output, file_name="bin_labels.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.subheader("🖼️ Interactive Bin Layout Diagrams")
                st.caption("Click on a bay to expand its visual layout.")
                for group in bay_groups_tab1:
                    for bay_id in group['bays']:
                        with st.expander(f"View Diagram for **{bay_id}**"):
                            fig = plot_bin_diagram(bay_id, group['shelves'], group['bins_per_shelf'], int(bay_id.replace("BAY-", "")[-3:]))
                            if fig: st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Error generating output: {str(e)}")

with tab2:
    st.header("Bin Bay Mapping ↔️", divider='rainbow')
    st.markdown("Define bay definition groups and map bin IDs to bay types.")
    
    bay_groups_tab2 = []
    num_groups_tab2 = st.number_input("How many bay definition groups?", min_value=1, max_value=50, value=1, key="num_groups_bin_mapping")

    for group_idx in range(num_groups_tab2):
        if f"bin_group_name_{group_idx}" not in st.session_state:
            st.session_state[f"bin_group_name_{group_idx}"] = f"Bay Definition Group {group_idx + 1}"
        def update_bin_group_name(idx=group_idx):
            st.session_state[f"bin_group_name_{idx}"] = st.session_state[f"bin_group_name_input_{idx}"]
        header = st.session_state[f"bin_group_name_{group_idx}"].strip() or f"Bay Definition Group {group_idx + 1}"
        
        with st.expander(header, expanded=True):
            st.text_input("Group Name", value=st.session_state[f"bin_group_name_{group_idx}"], key=f"bin_group_name_input_{group_idx}", on_change=update_bin_group_name, args=(group_idx,))
            # ... The rest of the inputs for Tab 2
            
    if st.button("Generate Excel", key="generate_bin_mapping_excel"):
        # Generation Logic for Tab 2
        pass

with tab3:
    st.header("EOA Generator 🪧", divider='rainbow')
    # ... The rest of the UI and logic for Tab 3 from the previous version ...

with tab4:
    st.header("Bay Elevation Generator 📐", divider='rainbow')
    st.markdown("Define bay types, their shelves, and bin configurations to generate a PowerPoint elevation drawing.")
    st.info("ℹ️ **Note:** This feature requires `python-pptx`. Please add `python-pptx>=0.6.23` to your requirements.txt file.")

    num_bay_types = st.number_input("How many bay types do you want to define?", min_value=1, max_value=20, value=1, key="num_bay_types")
    
    bay_types_data = []

    for i in range(num_bay_types):
        if f"bay_type_name_{i}" not in st.session_state:
            st.session_state[f"bay_type_name_{i}"] = f"Bay Type {i + 1}"

        def update_bay_type_name(idx=i):
            st.session_state[f"bay_type_name_{idx}"] = st.session_state[f"bay_type_name_input_{idx}"]

        header = st.session_state[f"bay_type_name_{i}"].strip() or f"Bay Type {i + 1}"

        with st.expander(header, expanded=True):
            bay_name = st.text_input("Bay Type Name", value=st.session_state[f"bay_type_name_{i}"], key=f"bay_type_name_input_{i}", on_change=update_bay_type_name, args=(i,))
            
            shelf_count = st.number_input("Number of shelves in this bay?", min_value=1, max_value=26, value=3, key=f"elevation_shelf_count_{i}")
            shelf_names = list(string.ascii_uppercase[:shelf_count])
            
            shelf_details = {}
            for shelf_name in shelf_names:
                st.divider()
                st.markdown(f"**Configuration for Shelf {shelf_name}**")
                
                num_bins = st.number_input("Number of bins in this shelf?", min_value=1, max_value=50, value=5, key=f"num_bins_{i}_{shelf_name}")
                
                st.markdown(f"**Bin Dimensions for all bins in Shelf {shelf_name} (cm)**")
                c1, c2, c3 = st.columns(3)
                with c1:
                    bin_h = st.number_input("Height", min_value=1.0, value=10.0, key=f"bin_h_{i}_{shelf_name}")
                with c2:
                    bin_w = st.number_input("Width", min_value=1.0, value=10.0, key=f"bin_w_{i}_{shelf_name}")
                with c3:
                    bin_d = st.number_input("Depth", min_value=1.0, value=10.0, key=f"bin_d_{i}_{shelf_name}")
                
                shelf_details[shelf_name] = {'num_bins': num_bins, 'h': bin_h, 'w': bin_w, 'd': bin_d}

            if bay_name.strip():
                bay_types_data.append({
                    "name": bay_name.strip(),
                    "shelves": shelf_names,
                    "shelf_details": shelf_details
                })

    if st.button("Generate PowerPoint Elevation", key="generate_ppt"):
        if any(not bay['name'] for bay in bay_types_data):
            st.error("Please provide a name for all bay types.")
        else:
            with st.spinner("Generating PowerPoint file..."):
                try:
                    ppt_buffer = generate_elevation_powerpoint(bay_types_data)
                    st.success(f"✅ Success! Generated a PowerPoint with {len(bay_types_data)} slides.")
                    st.download_button(
                        label="📥 Download PowerPoint File",
                        data=ppt_buffer,
                        file_name="bay_elevations.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"An error occurred during PowerPoint generation: {e}")
