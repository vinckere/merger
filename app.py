import streamlit as st
import pandas as pd
from io import BytesIO
import csv

st.title("CSV Merger & Excel Exporter with Robust CSV Parsing")

# Let user optionally pick a delimiter, or let pandas sniff
use_auto = st.checkbox("Auto-detect delimiter (using pandas)", value=True)
if not use_auto:
    sep = st.selectbox("Select delimiter", [",", ";", "\t", "|", " "], index=0)
else:
    sep = None

file1 = st.file_uploader("Upload first CSV file", type=["csv"])
file2 = st.file_uploader("Upload second CSV file", type=["csv"])

def read_csv_robust(file, sep):
    file.seek(0)
    if sep:
        # user‐specified delimiter
        return pd.read_csv(
            file,
            sep=sep,
            engine="python",           # fallback parser
            on_bad_lines="warn",       # skip malformed rows with a warning
            encoding="utf-8",          # adjust if your files use another encoding
        )
    else:
        # let pandas sniff both delimiter and quoting
        return pd.read_csv(
            file,
            sep=None,                  # auto-detect
            engine="python",
            on_bad_lines="warn",
            encoding="utf-8",
        )

if file1 and file2:
    try:
        df1 = read_csv_robust(file1, sep)
        df2 = read_csv_robust(file2, sep)
    except Exception as e:
        st.error(f"Could not parse CSV: {e}")
        st.stop()

    # --- your merge/filter logic ---
    merged = pd.merge(df1, df2, on="common_key", how="inner")
    filtered = merged[merged["some_column"] > 0]

    st.write("Preview of processed data:")
    st.dataframe(filtered)

    # --- write to Excel in memory ---
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Results")
    towrite.seek(0)

    st.download_button(
        label="Download result as Excel",
        data=towrite,
        file_name="processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
