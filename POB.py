import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import zipfile
import os
import openpyxl
import pytz

st.set_page_config(page_title="Excel Processor", layout="wide")

FOLDER_PATH = "saved_files"

# Bikin folder kalau belum ada
if not os.path.exists(FOLDER_PATH):
    os.makedirs(FOLDER_PATH)

# Function untuk load ulang file dari disk saat app start
def load_files():
    files = []
    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith(".xlsx"):
            with open(os.path.join(FOLDER_PATH, filename), "rb") as f:
                files.append({
                    "name": filename,
                    "data": f.read()
                })
    return files

if 'files' not in st.session_state:
    st.session_state.files = load_files()

if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False

if 'confirm_delete_all' not in st.session_state:
    st.session_state.confirm_delete_all = False

tab1, tab2 = st.tabs(["Upload Excel", "Overview Files"])

with tab1:
    st.header("Upload Excel File")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, header=None)
        dist = df.iloc[1, 1]
        area = df.iloc[2, 1]
        cabang = df.iloc[3, 1]
        bulan = df.iloc[4, 1]

        bulan_mapping = {
            "January": "Januari", "February": "Februari", "March": "Maret",
            "April": "April", "May": "Mei", "June": "Juni", "July": "Juli",
            "August": "Agustus", "September": "September", "October": "Oktober",
            "November": "November", "December": "Desember"
        }

        bulan_minus = {
            "Januari": "Desember", "Februari": "Januari", "Maret": "Februari",
            "April": "Maret", "Mei": "April", "Juni": "Mei", "Juli": "Juni",
            "Agustus": "Juli", "September": "Agustus", "Oktober": "September",
            "November": "Oktober", "Desember": "November"
        }

        bulan_plus = {
            "Januari": "Februari", "Februari": "Maret", "Maret": "April",
            "April": "Mei", "Mei": "Juni", "Juni": "Juli", "Juli": "Augustus",
            "Agustus": "September", "September": "Oktober", "Oktober": "November",
            "November": "Desember", "Desember": "Januari"
        }

        bulan_plus2 = {
            "Januari": "Maret", "Februari": "April", "Maret": "Mei",
            "April": "Juni", "Mei": "Juli", "Juni": "Agustus",
            "Juli": "September", "Agustus": "Oktober", "September": "November",
            "Oktober": "Desember", "November": "Januari", "Desember": "Februari"
        }

        tahun = datetime.now().year
        bulan_str = bulan.strftime('%B') if isinstance(bulan, datetime) else str(bulan)
        nama_bulan = bulan_mapping.get(bulan_str, bulan_str)
        bulan_minus_1 = bulan_minus.get(nama_bulan, nama_bulan)
        bulan_plus_fix= bulan_plus.get(nama_bulan, nama_bulan)
        bulan_plus2_fix = bulan_plus2.get(nama_bulan, nama_bulan)

        item_products = df.iloc[9:, 1].dropna().str.strip()
        sales = df.iloc[9:, 10]
        avg = df.iloc[9:, 7]
        sellout = df.iloc[9:, 12]
        target = df.iloc[9:, 11]
        pob = df.iloc[9:, 82]
        forecast_1 = df.iloc[9:, 80]
        forecast_2 = df.iloc[9:, 124]
        total_final = df.iloc[9:, 92]

        item_products = item_products.iloc[:-5].reset_index(drop=True)
        sales = sales.iloc[:-5].reset_index(drop=True)
        avg = avg.iloc[:-5].reset_index(drop=True)
        sellout = sellout.iloc[:-5].reset_index(drop=True)
        target = target.iloc[:-5].reset_index(drop=True)
        pob = pob.iloc[:-5].reset_index(drop=True)
        forecast_1 = forecast_1[:-5].reset_index(drop=True)
        forecast_2 = forecast_2[:-5].reset_index(drop=True)
        total_final = total_final[:-5].reset_index(drop=True)

        result_df = pd.DataFrame({
            'Dist': [dist] * len(item_products),
            'Area': [area] * len(item_products),
            'Cabang': [cabang] * len(item_products),
            'Bulan': [bulan] * len(item_products),
            'Item Product': item_products,
            'Total Final POB Adjust RM-AM / DISt': total_final,
            f'POB {nama_bulan}-{tahun}': pob,
            f'Forecast {bulan_plus_fix}-{tahun}': forecast_1,
            f'Forecast {bulan_plus2_fix}-{tahun}': forecast_2          
        })

        st.dataframe(result_df)

        if st.button("Submit & Save to Overview"):
            filename = f"PO {nama_bulan} {cabang} {tahun}.xlsx"
            file_path = os.path.join(FOLDER_PATH, filename)
            # Save to disk
            result_df.to_excel(file_path, index=False)
            # Save to session state
            with open(file_path, "rb") as f:
                st.session_state.files.append({
                    "name": filename,
                    "data": f.read()
                })
            st.success("✅ Dataset sudah diolah dan disimpan di Overview!")

with tab2:
    st.header("Overview Excel Files")

    if st.session_state.files:
        selected_files = []
        select_all = st.checkbox("Select All")
        for file in st.session_state.files:
            checked = st.checkbox(file['name'], key=file['name'], value=select_all)
            if checked:
                selected_files.append(file['name'])

        col1, col2, _ = st.columns([1, 1, 5])
        with col1:
            if selected_files:
                if st.button("📥 Download Selected as ZIP"):
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zipf:
                        for file in st.session_state.files:
                            if file['name'] in selected_files:
                                zipf.writestr(file['name'], file['data'])
                    st.download_button(
                        label="Download ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="datasets.zip",
                        mime="application/zip",
                        key="download_zip_btn"
                    )

        with col2:
            if selected_files and not st.session_state.confirm_delete:
                if st.button("🗑️ Delete Selected"):
                    st.session_state.confirm_delete = True
                    st.rerun()

        if st.session_state.confirm_delete:
            st.warning("⚠️ Yakin ingin mendelete file yang dipilih?")
            col_ok, col_cancel = st.columns(2)
            with col_ok:
                if st.button("✅ Ya, Delete"):
                    for fname in selected_files:
                        file_path = os.path.join(FOLDER_PATH, fname)
                        if os.path.exists(file_path):
                            os.remove(file_path)
                    st.session_state.files = [
                        file for file in st.session_state.files if file['name'] not in selected_files
                    ]
                    st.session_state.confirm_delete = False
                    st.rerun()
            with col_cancel:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete = False
                    st.rerun()

        st.divider()
        if not st.session_state.confirm_delete_all:
            if st.button("🗑️ Delete All Files"):
                st.session_state.confirm_delete_all = True
                st.rerun()

        if st.session_state.confirm_delete_all:
            st.warning("⚠️ Yakin ingin mendelete SEMUA file?")
            col_ok_all, col_cancel_all = st.columns(2)
            with col_ok_all:
                if st.button("✅ Ya, Delete All"):
                    # Delete all from disk
                    for file in os.listdir(FOLDER_PATH):
                        os.remove(os.path.join(FOLDER_PATH, file))
                    st.session_state.files = []
                    st.session_state.confirm_delete_all = False
                    st.rerun()
            with col_cancel_all:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete_all = False
                    st.rerun()
    else:
        st.info("Belum ada file yang disimpan.")
