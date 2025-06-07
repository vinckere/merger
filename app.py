import streamlit as st
import pandas as pd

def extract_audio_ca(df):
    """Extracts store name and CA Encaissé Audio, rounds CA"""
    store_col = next(col for col in df.columns if df[col].astype(str).str.contains("brest|tours", case=False).any())
    result = df[[store_col, "CA EncaissÈ Audio"]].copy()
    result.columns = ["Store", "CA Audio"]
    result["CA Audio"] = result["CA Audio"].round().astype(int)
    return result

def save_to_excel(df):
    """Convert DataFrame to Excel in memory"""
    from io import BytesIO
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

st.title("Extract CA Audio")

uploaded_file = st.file_uploader("Upload 1 CSV file", type="csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    try:
        cleaned_df = extract_audio_ca(df)
        st.write(cleaned_df)
        excel_file = save_to_excel(cleaned_df)
        st.download_button("Download Excel", excel_file, "output_audio_ca.xlsx")
    except Exception as e:
        st.error(f"Error processing file: {e}")
