import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Merger & Processor")

# 1. File uploaders
file1 = st.file_uploader("Upload first Excel file", type=["xls","xlsx"])
file2 = st.file_uploader("Upload second Excel file", type=["xls","xlsx"])

# 2. Once both are uploaded, read into pandas
if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # 3. Processing / filtering / merging
    # (→ replace below with your client’s logic)
    merged = pd.merge(df1, df2, on="common_key", how="inner")
    filtered = merged[merged["some_column"] > 0]

    st.write("Preview of processed data:")
    st.dataframe(filtered)

    # 4. Download button
    towrite = BytesIO()
    filtered.to_excel(towrite, index=False, engine="openpyxl")
    towrite.seek(0)
    st.download_button(
        label="Download result as Excel",
        data=towrite,
        file_name="processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
