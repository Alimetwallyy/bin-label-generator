import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

def generate_bin_labels_table(bay_groups):
    data = []

    for group in bay_groups:
        bay_ids = group['bays']
        shelves = group['shelves']
        bins_per_shelf = group['bins_per_shelf']

        for bay in bay_ids:
            base_label = bay.replace("BAY-", "")
            base_number = int(base_label[-3:])

            max_bins = max(bins_per_shelf.get(shelf, 0) for shelf in shelves)

            for i in range(max_bins):
                row = {'Bay_ID': bay}
                for shelf in shelves:
                    shelf_bin_count = bins_per_shelf.get(shelf, 0)
                    if i < shelf_bin_count:
                        bin_label = base_label[:-4] + shelf + f"{base_number + i:03d}"
                        row[f"Shelf_{shelf}"] = bin_label
                    else:
                        row[f"Shelf_{shelf}"] = None
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

# --- 🖥️ Streamlit UI ---
st.set_page_config(page_title="Bin Label Generator", layout="centered")
st.title("📦 Bin Label Generator")
st.markdown("Define bay groups, shelves, and bins per shelf to generate structured bin labels.")

bay_groups = []
num_groups = st.number_input("How many bay groups do you want to define?", min_value=1, max_value=10, value=1)

for group_idx in range(num_groups):
    st.header(f"🧱 Bay Group {group_idx + 1}")

    bays_input = st.text_area(f"Enter bay IDs (one per line)", key=f"bays_{group_idx}")
    shelves_input = st.text_input(f"Enter shelf labels (comma-separated like A,B,C)", key=f"shelves_{group_idx}")

    shelves = [s.strip() for s in shelves_input.split(",") if s.strip()]
    bins_per_shelf = {}
    for shelf in shelves:
        count = st.number_input(f"Number of bins in shelf {shelf}", min_value=1, max_value=100, value=5, key=f"bins_{group_idx}_{shelf}")
        bins_per_shelf[shelf] = count

    if bays_input:
        bay_list = [b.strip() for b in bays_input.splitlines() if b.strip()]
        bay_groups.append({
            "bays": bay_list,
            "shelves": shelves,
            "bins_per_shelf": bins_per_shelf
        })

if st.button("Generate Bin Labels"):
    if bay_groups:
        df = generate_bin_labels_table(bay_groups)
        st.success("✅ Bin labels generated successfully!")
        st.dataframe(df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

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
