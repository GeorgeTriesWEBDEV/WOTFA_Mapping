import streamlit as st
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz

st.set_page_config(page_title="Asset Mapper", layout="centered")

st.title("Excel Asset Mapper")
st.write("""
1. **Upload Asset Master (File 1):**  
    Columns: Asset Name, Type, Schedule Sub Type, Dashboard Name  
2. **Upload Schedule File (File 2):**  
    Columns: Schedule Name, Type (empty), Schedule Sub Type (empty), Dashboard Name (empty)  
3. **Get back a file with Type, Schedule Sub Type, and Dashboard Name inferred.
""")

@st.cache_data
def process_files(file1, file2, threshold):
    # Read files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    asset_names = df1['Asset Name'].astype(str).tolist()

    for idx, row in df2.iterrows():
        schedule_name = str(row['Schedule Name'])
        match_idx, score = None, 0

        # First: substring match
        for i, asset in enumerate(asset_names):
            if asset in schedule_name:
                match_idx = i
                score = 100
                break

        # If no substring match found, do fuzzy match
        if match_idx is None:
            best_score = 0
            best_idx = None
            for i, asset in enumerate(asset_names):
                s = fuzz.partial_ratio(schedule_name.lower(), asset.lower())
                if s > best_score:
                    best_score = s
                    best_idx = i
            match_idx = best_idx
            score = best_score

        if match_idx is not None and score >= threshold:
            df2.at[idx, 'Type'] = df1.at[match_idx, 'Type']
            df2.at[idx, 'Schedule Sub Type'] = df1.at[match_idx, 'Schedule Sub Type']
            df2.at[idx, 'Dashboard Name'] = df1.at[match_idx, 'Dashboard Name']
        else:
            df2.at[idx, 'Type'] = ''
            df2.at[idx, 'Schedule Sub Type'] = ''
            df2.at[idx, 'Dashboard Name'] = ''

    return df2

# --- File Upload UI ---
file1 = st.file_uploader("Upload Asset Master (File 1)", type=["xlsx"])
file2 = st.file_uploader("Upload Schedule File (File 2)", type=["xlsx"])
threshold = st.slider("Fuzzy match threshold", 0, 100, 75, step=1, help="Similarity required for best effort mapping (higher is stricter)")

if file1 and file2:
    if st.button("Map and Download Results"):
        result = process_files(file1, file2, threshold)

        output = BytesIO()
        result.to_excel(output, index=False)
        output.seek(0)

        st.success("Processing complete. Download your file below.")
        st.download_button(
            label="Download mapped file",
            data=output,
            file_name="schedule_mapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(result.head())

    with st.expander("See preview of uploaded files"):
        st.subheader("File 1 (Asset Master) Preview")
        st.write(pd.read_excel(file1).head())
        st.subheader("File 2 (Schedule File) Preview")
        st.write(pd.read_excel(file2).head())