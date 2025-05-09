import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Caroptic Tool")

# 1. File uploaders now only accept CSV
file1 = st.file_uploader("Importez votre premier CSV", type=["csv"])
file2 = st.file_uploader("Importez votre second CSV", type=["csv"])

if file1 and file2:
    # 2. Read CSVs instead of Excel
    df1 = pd.read_csv(file1)
    df2 = pd.read_csv(file2)

    # 3. Your merge/filter logic (example)
    merged = pd.merge(df1, df2, on="common_key", how="inner")
    filtered = merged[merged["some_column"] > 0]

    st.write("Preview of processed data:")
    st.dataframe(filtered)

    # 4. Write to an in-memory Excel file
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Results")
    towrite.seek(0)

    st.download_button(
        label="Télécharger l'Excel généré",
        data=towrite,
        file_name="processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )