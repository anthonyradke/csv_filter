import streamlit as st
import pandas as pd
from datetime import timedelta
from io import StringIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import base64

def round_to_15_min(dt):
    discard = timedelta(minutes=dt.minute % 15, seconds=dt.second, microseconds=dt.microsecond)
    dt -= discard
    if discard >= timedelta(minutes=7.5):
        dt += timedelta(minutes=15)
    return dt

def simplify_name(full_name):
    if isinstance(full_name, str):
        if "|dac-" in full_name:
            return full_name.split("|")[0].split(".")[-1]
        elif ".FLN_" in full_name:
            return full_name.split(".FLN_", 1)[1]
        else:
            return full_name.split(".")[-1]
    return full_name

def ensure_unique_columns(df):
    seen = {}
    new_cols = []
    for col in df.columns:
        if col not in seen:
            seen[col] = 1
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    df.columns = new_cols
    return df

def process_file(file):
    try:
        df_raw = pd.read_csv(file, header=None, skiprows=[0, 2])
        clean_df = pd.DataFrame()
        for i in range(0, df_raw.shape[1], 4):
            try:
                time_col = pd.to_datetime(df_raw.iloc[1:, i].astype(str), errors='coerce')
                values = df_raw.iloc[1:, i + 1]
                time_rounded = time_col.map(round_to_15_min)
                valid = time_rounded.notna()
                title = str(df_raw.iloc[0, i])

                temp_df = pd.DataFrame({
                    "datetime": time_rounded[valid],
                    title: values[valid]
                })
                clean_df = temp_df if clean_df.empty else pd.merge(clean_df, temp_df, on="datetime", how="outer")
            except:
                continue

        if "datetime" not in clean_df.columns:
            return None

        clean_df = clean_df.groupby("datetime", as_index=False).first().sort_values("datetime").reset_index(drop=True)
        clean_df = clean_df.rename(columns={col: simplify_name(col) for col in clean_df.columns if col != "datetime"})
        return ensure_unique_columns(clean_df)
    except:
        return None

def to_excel_download_link(df, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned Data"
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 31.00
    for cell in ws['A'][1:]:
        cell.number_format = 'm/d/yyyy h:mm:ss AM/PM'

    from tempfile import NamedTemporaryFile
    tmp = NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.seek(0)
    data = tmp.read()
    tmp.close()
    b64 = base64.b64encode(data).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">📥 Download {filename}</a>'

st.set_page_config(page_title="Andrew Tool", layout="wide")
st.title("📊 CSV Batch Cleaner - Streamlit Version")

uploaded_files = st.file_uploader("Upload one or more CSV files", type=["csv"], accept_multiple_files=True)
out_format = st.selectbox("Choose Output Format", ["xlsx", "csv"])

if uploaded_files:
    processed_files = {}
    for uploaded_file in uploaded_files:
        st.markdown(f"### Processing: {uploaded_file.name}")
        df = process_file(uploaded_file)
        if df is not None:
            st.success(f"✅ {uploaded_file.name} cleaned successfully!")
            st.dataframe(df.head())
            processed_files[uploaded_file.name] = df
        else:
            st.error(f"❌ Failed to process {uploaded_file.name}")

    if processed_files:
        mode = st.selectbox("Output Mode", [
            "Keep files separate",
            "Merge into one file with one sheet"
        ])

        if mode == "Keep files separate":
            for name, df in processed_files.items():
                if out_format == "csv":
                    csv = df.to_csv(index=False).encode()
                    st.download_button(
                        label=f"Download {name}_Filtered.csv",
                        data=csv,
                        file_name=f"{name}_Filtered.csv",
                        mime='text/csv'
                    )
                else:
                    st.markdown(to_excel_download_link(df, f"{name}_Filtered.xlsx"), unsafe_allow_html=True)
        else:
            merged_df = pd.concat(processed_files.values(), axis=0, ignore_index=True)
            merged_df = merged_df.groupby("datetime", as_index=False).first().sort_values("datetime").reset_index(drop=True)
            merged_df = ensure_unique_columns(merged_df)

            if out_format == "csv":
                csv = merged_df.to_csv(index=False).encode()
                st.download_button(
                    label="Download Merged File",
                    data=csv,
                    file_name="Merged_File.csv",
                    mime='text/csv'
                )
            else:
                st.markdown(to_excel_download_link(merged_df, "Merged_File.xlsx"), unsafe_allow_html=True)
