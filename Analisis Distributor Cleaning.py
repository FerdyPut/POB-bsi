import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import zipfile
import os
from xlsxwriter import Workbook
import re

st.set_page_config(page_title="BSI - Support Information", layout="wide")

# Mulai Tabs
tab1, tab2, tab3 = st.tabs(["📦 POB", "Penggabungan Data - POB", "📝 RNL"])

with tab1:
    FOLDER_PATH = "saved_files"

    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)
    
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
    
    st.session_state.files = load_files()
    st.header("📊 Masukkan File POB")

    # Step 1: Pilih POB
    selected_pob = st.selectbox("Pilih POB", ('', 'POB - Dist', 'POB - SSO'))

    # Step 2: Jika sudah pilih POB, baru muncul pilihan MT/GT
    if selected_pob:
        selected_channel = st.selectbox("Pilih Channel", ('MT', 'GT'))

        # Step 3: Jika sudah pilih MT/GT, baru muncul upload
        if selected_channel:
            uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

            if uploaded_file is not None:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names

                all_data = {}  # Dictionary untuk menyimpan data per sheet
                all_results = []  # List untuk menggabungkan semua sheet

                for sheet_name in sheet_names:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)

#--------------------------------------------------------KODE UNTUK POB-DIST, MT

                    if selected_pob == "POB - Dist" and selected_channel == "MT":
                        # Ambil informasi dari file
                        dist = df.iloc[1, 1]
                        area = df.iloc[2, 1]
                        cabang = df.iloc[3, 1]
                        bulan = df.iloc[4, 1]

                        # Mapping bulan
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
                            "April": "Mei", "Mei": "Juni", "Juni": "Juli", "Juli": "Agustus",
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
                        bulan_plus_fix = bulan_plus.get(nama_bulan, nama_bulan)
                        bulan_plus2_fix = bulan_plus2.get(nama_bulan, nama_bulan)
                        try:
                            if isinstance(bulan, str):  # Pastikan bulan berupa string sebelum dikonversi
                                bulan_dt = datetime.strptime(bulan, "%B")  # Ubah ke datetime (format nama bulan lengkap)
                            else:
                                bulan_dt = bulan  # Jika sudah datetime, langsung gunakan

                            bulan_formatted = bulan_dt.strftime('%b-%y')  # Format singkat (contoh: "Mar-24")
                        except ValueError:
                            bulan_formatted = bulan  # Jika gagal parsing, tetap gunakan nilai aslinya


                        # Data Cleaning sesuai POB dan Channel
                        item_products = df.iloc[9:, 1].dropna().astype(str).str.strip()
                        total_final = df.iloc[9:, 95].replace("-", 0).fillna(0)
                        forecast_1 = df.iloc[9:, 117].replace("-", 0).fillna(0)
                        forecast_2 = df.iloc[9:, 127].replace("-", 0).fillna(0)

                        # Konversi ke numerik agar aman
                        total_final = pd.to_numeric(total_final, errors='coerce').fillna(0)
                        forecast_1 = pd.to_numeric(forecast_1, errors='coerce').fillna(0)
                        forecast_2 = pd.to_numeric(forecast_2, errors='coerce').fillna(0)

                        # Hapus baris dengan teks tidak relevan
                        remove_keywords = ["total", "MOHON DIPERHATIKAN", "Jika", "adjust average", "Stock ED"]
                        mask_valid_products = ~item_products.astype(str).str.contains('|'.join(remove_keywords), case=False, na=False)
                        
                        item_products = item_products[mask_valid_products].reset_index(drop=True)
                        total_final = total_final[mask_valid_products].reset_index(drop=True)
                        forecast_1 = forecast_1[mask_valid_products].reset_index(drop=True)
                        forecast_2 = forecast_2[mask_valid_products].reset_index(drop=True)

                        # Hapus baris terakhir (biasanya total)
                        item_products = item_products.iloc[:-1].reset_index(drop=True)
                        total_final = total_final.iloc[:-1].reset_index(drop=True)
                        forecast_1 = forecast_1.iloc[:-1].reset_index(drop=True)
                        forecast_2 = forecast_2.iloc[:-1].reset_index(drop=True)

                        # Buat dataframe dari sheet ini
                        result_df = pd.DataFrame({
                            'POB': [selected_pob] * len(item_products),
                            'Channel': [selected_channel] * len(item_products),
                            'Dist': [dist] * len(item_products),
                            'Area': [area] * len(item_products),
                            'Cabang': [cabang] * len(item_products),
                            #'Nama Sheet': [sheet_name] * len(item_products),  # Tambahkan nama sheet
                            'Bulan': [bulan_formatted] * len(item_products),
                            'Item Product': item_products,
                            'Total Final POB Adjust RM-AM / DISt': total_final,
                            f'Forecast {bulan_plus_fix}-{tahun}': forecast_1,
                            f'Forecast {bulan_plus2_fix}-{tahun}': forecast_2          
                        })
#-------------------------------------------------------KODE UNTUK POB-SSO, MT

                    elif selected_pob == "POB - SSO" and selected_channel == "MT":
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        now = datetime.now()
                        nama_bulan = now.strftime("%B")  # Nama bulan dalam format "January", "February", dll.
                        tahun = now.year  # Tahun dalam format 2025, 2026, dll.
                        df[0] = df[0].astype(str).str.strip().str.upper()  # Kolom A
                        df[1] = df[1].astype(str).str.strip().str.upper()  # Kolom B

                        total_index = df[(df[0] == "TOTAL") | (df[1] == "TOTAL")].index
                        if total_index.empty:
                            st.warning(f"Tidak ditemukan baris 'TOTAL' di sheet '{sheet_name}'. Lewati sheet ini!")
                        
                        total_index = total_index[0]  # Ambil indeks pertama

                        # **Cabang**
                        try:
                            cabang = str(df.iloc[0, 1]).strip()  # Ambil cabang di baris pertama kolom ke-2
                            if pd.isna(cabang) or cabang == "":
                                cabang = "Unknown"
                        except IndexError:
                            cabang = "Unknown"

                        # **Produk**
                        products = df.iloc[2:total_index, 1].dropna().tolist()

                        # **Distributor**
                        distributors = []
                        col_start = 2  # Mulai dari kolom C
                        while col_start < df.shape[1]:
                            try:
                                dist_name = str(df.iloc[0, col_start]).strip()
                                if pd.notna(df.iloc[0, col_start + 1]):
                                    dist_name += " " + str(df.iloc[0, col_start + 1]).strip()
                                distributors.append(dist_name)
                            except IndexError:
                                break  # Jika kolom tidak ada, hentikan loop
                            col_start += 6  # Lompat 6 kolom

                        # **Buat Data**
                        data = []
                        col_start = 2  # Mulai dari kolom C
                        for distributor in distributors:
                            for product_idx, product in enumerate(products):
                                for week_idx, week in enumerate(["W1", "W2", "W3", "W4", "W5"]):
                                    try:
                                        nilai = df.iloc[2 + product_idx, col_start + week_idx]  # Ambil nilai berdasarkan Week
                                        nilai = 0 if pd.isna(nilai) else nilai  # Ganti NaN jadi 0
                                        data.append([timestamp, sheet_name, product, distributor, cabang, week, nilai])
                                    except IndexError:
                                        data.append([timestamp, sheet_name, product, distributor, cabang, week, 0])  # Jika out of bounds, isi 0
                            col_start += 6  # Lompat ke distributor berikutnya

                        # **Buat DataFrame**
                        result_df = pd.DataFrame(data, columns=["Timestamp", "Sheet", "Produk", "Distributor", "Cabang", "Week", "Nilai"])

                        # **Hapus angka "0 0 0 0" dari distributor secara aman**
                        result_df["Distributor"] = result_df["Distributor"].str.replace(r"(\s0+)+$", "", regex=True)

                    elif selected_pob == "POB - SSO" and selected_channel == "GT":
                        1
#---------------------------------------RESULT DAN PENYIMPANAN

                   # Simpan data yang sudah dibersihkan
                    all_data[sheet_name] = result_df
                    all_results.append(result_df)

                # Gabungkan semua sheet jadi satu DataFrame
                if all_results:
                    final_result = pd.concat(all_results, ignore_index=True)

                    # Pilih sheet untuk ditampilkan
                    selected_sheet = st.selectbox("Pilih untuk Melihat Sheet!", sheet_names)

                    if selected_sheet in all_data:
                        st.write(f"### Data dari {selected_sheet}")
                        st.dataframe(all_data[selected_sheet])

                    if st.button("Proses dan Simpan!"):
                        safe_branch_name = cabang.replace("/", "_")

                        # Tentukan format nama file berdasarkan sheet yang dipilih
                        if "POB - SSO" in selected_sheet and "MT" in selected_sheet:
                            filename = f"POB - SSO - {nama_bulan} - {tahun}.xlsx"
                        elif "POB - DIST" in selected_sheet and "MT" in selected_sheet:
                            filename = f"PO {nama_bulan} {safe_branch_name} {tahun}.xlsx"
                        elif "POB - SSO" in selected_sheet and "GT" in selected_sheet:
                            filename = f"POB - SSO - GT - {nama_bulan} - {tahun}.xlsx"
                        elif "POB - DIST" in selected_sheet and "GT" in selected_sheet:
                            filename = f"PO {nama_bulan} {safe_branch_name} GT {tahun}.xlsx"
                        else:
                            filename = f"PO {nama_bulan} {safe_branch_name} {tahun}.xlsx"  # Default jika tidak ada kondisi yang cocok


                        def get_unique_filename(folder_path, filename):
                            base, ext = os.path.splitext(filename)
                            counter = 1
                            new_filename = filename

                            while os.path.exists(os.path.join(folder_path, new_filename)):
                                new_filename = f"{base} ({counter}){ext}"
                                counter += 1

                            return new_filename  # Harus ada return ini!

                        # Panggil fungsi untuk mendapatkan nama file unik
                        filename = get_unique_filename(FOLDER_PATH, filename)
                        file_path = os.path.join(FOLDER_PATH, filename)

                        # Simpan ke Excel
                        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                            final_result.to_excel(writer, sheet_name="POB Combined", index=False)

                        st.success(f"✅ File telah dicleaning dan berhasil disimpan di Overview!: `{filename}`")
                        st.rerun()  # Paksa Streamlit untuk refresh tanpa manual reload                       
                
    st.subheader("📂 Overview Saved Files")

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
                if len(selected_files) == 1:
                    # Download langsung file XLSX tanpa membaca CSV
                    file_name = selected_files[0]
                    file_path = os.path.join(FOLDER_PATH, file_name)

                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()

                        st.download_button(
                            label="📥 Download File",
                            data=file_bytes,
                            file_name=file_name,
                            mime="application/octet-stream",
                            key="download_single_btn"
                        )
                else:
                    # Jika lebih dari satu file, buat ZIP
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

with tab2:
    st.header("🔗 Merge POB Files")
    
    if st.session_state.files:
        selected_files = st.multiselect("Pilih file untuk digabungkan:", [file["name"] if isinstance(file, dict) else file for file in st.session_state.files])
        
        if st.button("🔄 Merge Files"):
            if selected_files:
                merged_data = []
                
                for file in selected_files:
                    file_path = os.path.join(FOLDER_PATH, file)
                    df = pd.read_excel(file_path)
                    merged_data.append(df)
                
                final_df = pd.concat(merged_data, ignore_index=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='Merged POB', index=False)
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Merged File",
                    data=output,
                    file_name="Merged_POB.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("✅ File berhasil digabung!")
            else:
                st.warning("⚠️ Pilih minimal satu file untuk digabungkan.")
    else:
        st.info("Belum ada file yang tersedia untuk digabungkan.")
with tab3:
    # Folder tempat file disimpan
    FOLDER_PATH = "saved_files"
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)
    
    def get_unique_filename(folder_path, filename):
        base, ext = os.path.splitext(filename)
    
        # Ambil ulang daftar file di folder dan di session_state
        existing_files = set(os.listdir(folder_path))
        session_files = {file["name"] for file in st.session_state.files} if "files" in st.session_state else set()
    
        all_existing_files = existing_files.union(session_files)  # Gabungkan daftar file dari dua sumber
    
        if filename not in all_existing_files:
            return filename  # Jika nama file belum ada, langsung pakai nama aslinya
    
        counter = 1
        while f"{base} ({counter}){ext}" in all_existing_files:
            counter += 1
    
        return f"{base} ({counter}){ext}"



    
    # Fungsi untuk membersihkan data
    def process_excel(file, sheet_name):
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # Ambil metadata dari file
        data_type = df.iloc[1, 1]
        area = df.iloc[2, 1]
        cabang = df.iloc[3, 1]
        bulan = datetime.now().strftime('%B')
        tahun = datetime.now().year
        
        # Lokasi awal data
        start_row = 13  # Baris ke-14
        start_col = 2    # Kolom C
        
        # Cari batas akhir First Value (sebelum "listing by toko")
        end_col = df.shape[1]
        for col in range(start_col, df.shape[1]):
            if "listing by toko" in str(df.iloc[start_row-1, col]).lower():
                end_col = col
                break
        
        # Lokasi awal dan akhir Second Value
        second_start_col = 49  # Sesuai dengan kolom AX
        second_end_col = df.shape[1]
        
        # Ambil nama produk
        first_products = df.iloc[start_row-1, start_col:end_col].dropna().values.tolist()
        second_products = df.iloc[start_row-1, second_start_col:second_end_col].dropna().values.tolist()
        
        # Ambil data customer dan kode
        data_rows = df.iloc[start_row:, :].dropna(how='all')
        kode_list = data_rows.iloc[:, 0].astype(str).tolist()
        customer_list = data_rows.iloc[:, 1].astype(str).tolist()
        
        # Ambil nilai produk
        first_values = data_rows.iloc[:, start_col:end_col].apply(pd.to_numeric, errors='coerce').values
        second_values = data_rows.iloc[:, second_start_col:second_end_col]
        
        records = []
        for i, kode in enumerate(kode_list):
            customer = customer_list[i]
            for j, product in enumerate(first_products):
                value = first_values[i][j] if j < len(first_values[i]) else None
                if pd.notna(value) and value > 1:
                    records.append({
                        "Type": data_type,
                        "Area": area,
                        "Cabang": cabang,
                        "Bulan": bulan,
                        "Tahun": tahun,
                        "Kode": kode,
                        "Customer_Name": customer,
                        "Produk": product,
                        "Value": value,
                        "Real": None
                    })
            for j, product in enumerate(second_products):
                if j < second_values.shape[1]:
                    value = second_values.iloc[i, j]
                    if pd.notna(value):
                        is_numeric = isinstance(value, (int, float)) and value > 1
                        is_alphanumeric = isinstance(value, str) and bool(re.search(r'[A-Za-z]', value)) and bool(re.search(r'\d', value))
                        is_string = isinstance(value, str) and not is_numeric
                        records.append({
                            "Type": data_type,
                            "Area": area,
                            "Cabang": cabang,
                            "Bulan": bulan,
                            "Tahun": tahun,
                            "Kode": kode,
                            "Customer_Name": customer,
                            "Produk": product,
                            "Value": value if is_numeric else None,
                            "Real": value if (is_alphanumeric or is_string) or (isinstance(value, (int, float)) and value > 1) else None
                        })
        
        df_records = pd.DataFrame(records)
        df_records = df_records.groupby(["Type", "Area", "Cabang", "Bulan", "Tahun", "Kode", "Customer_Name", "Produk"], as_index=False).first()
        # Menghapus baris yang memiliki None di kolom 'Value' atau 'Real'
        df_records = df_records.dropna(subset=["Value", "Real"])
        df_records = df_records[df_records['Customer_Name'] != 'nan']
        df_records = df_records.reset_index(drop=True)
        
        return df_records
    
    # Streamlit App
    st.title("Data Cleaning & Transformation")
    
    # Upload File
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        # Tampilkan pilihan sheet
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        sheet_name = st.selectbox("Pilih Sheet untuk Dibersihkan", sheet_names)
    
        # Proses dan tampilkan hasilnya
        cleaned_df = process_excel(uploaded_file, sheet_name)
        
        # Simpan data yang sudah dibersihkan hanya ketika tombol simpan diklik
        if 'cleaned_data' not in st.session_state:
            st.session_state.cleaned_data = None
    
        # Tampilkan data yang telah dibersihkan
        st.dataframe(cleaned_df)
    
        if st.button("Simpan Data"):
            # Gunakan fungsi get_unique_filename untuk mendapatkan nama file yang unik
            file_name = f"cleaned_{sheet_name}.csv"
            file_name = get_unique_filename(FOLDER_PATH, file_name)
            file_path = os.path.join(FOLDER_PATH, file_name)
        
            # Simpan file cleaned_df di dalam folder
            cleaned_df.to_csv(file_path, index=False)
        
            # Pastikan 'files' tetap menyimpan semua file yang ada
            if 'files' not in st.session_state:
                st.session_state.files = []
        
            # Tambahkan file baru tanpa menggantikan file lama
            st.session_state.files.append({"name": file_name, "data": open(file_path, "rb").read()})
        
            st.success(f"Data telah disimpan dengan nama {file_name}")
    
    FOLDER_PATH = "saved_files"  # Pastikan folder ini ada
    
    # Buat folder jika belum ada
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)
    
    # File Management in session state
    if "files" not in st.session_state:
        st.session_state.files = [{"name": f} for f in os.listdir(FOLDER_PATH)]  # Load file dari folder
    
    st.subheader("📂 Overview Saved Files")
    
    # Pastikan daftar file selalu diperbarui dengan isi folder
    st.session_state.files = [{"name": f} for f in os.listdir(FOLDER_PATH)]
    
    if st.session_state.files:
        selected_files = []
        select_all = st.checkbox("Select All")
    
        # Menampilkan checkbox untuk memilih file
        for file in st.session_state.files:
            checked = st.checkbox(file["name"], key=file["name"], value=select_all)
            if checked:
                selected_files.append(file["name"])
    
        col1, col2, col3 = st.columns([1, 1, 1])
    
        with col1:
            if selected_files:
                if len(selected_files) == 1:
                    file_name = selected_files[0]
                    file_path = os.path.join(FOLDER_PATH, file_name)
    
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as f:
                            file_bytes = f.read()
    
                        st.download_button(
                            label="📥 Download File",
                            data=file_bytes,
                            file_name=file_name,
                            mime="application/octet-stream",
                            key="download_single_btn",
                        )
    
                else:
                    if st.button("📥 Download Selected as ZIP"):
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zipf:
                            for file in selected_files:
                                file_path = os.path.join(FOLDER_PATH, file)
                                if os.path.exists(file_path):
                                    with open(file_path, "rb") as f:
                                        zipf.writestr(file, f.read())
    
                        st.download_button(
                            label="Download ZIP",
                            data=zip_buffer.getvalue(),
                            file_name="datasets.zip",
                            mime="application/zip",
                            key="download_zip_btn",
                        )
    
        with col2:
            if selected_files:
                if st.button("🗑️ Delete Selected"):
                    st.session_state.confirm_delete = True
                    st.rerun()
    
        with col3:
            if st.button("🔄 Refresh File List"):
                st.session_state.files = [{"name": f} for f in os.listdir(FOLDER_PATH)]
                st.success("Daftar file telah diperbarui!")
                st.rerun()
    
        if st.session_state.get("confirm_delete", False):
            st.warning("⚠️ Yakin ingin mendelete file yang dipilih?")
            col_ok, col_cancel = st.columns(2)
    
            with col_ok:
                if st.button("✅ Ya, Delete"):
                    for fname in selected_files:
                        file_path = os.path.join(FOLDER_PATH, fname)
                        if os.path.exists(file_path):
                            os.remove(file_path)
    
                    st.session_state.files = [{"name": f} for f in os.listdir(FOLDER_PATH)]
                    st.session_state.confirm_delete = False
                    st.rerun()
    
            with col_cancel:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete = False
                    st.rerun()
    
        st.divider()
    
        if not st.session_state.get("confirm_delete_all", False):
            if st.button("🗑️ Delete All Files"):
                st.session_state.confirm_delete_all = True
                st.rerun()
    
        if st.session_state.get("confirm_delete_all", False):
            st.warning("⚠️ Yakin ingin mendelete SEMUA file?")
            col_ok_all, col_cancel_all = st.columns(2)
    
            with col_ok_all:
                if st.button("✅ Ya, Delete All"):
                    for file in os.listdir(FOLDER_PATH):
                        os.remove(os.path.join(FOLDER_PATH, file))
    
                    st.session_state.files = []  # Reset daftar file di session state
                    st.session_state.confirm_delete_all = False
                    st.rerun()
    
            with col_cancel_all:
                if st.button("❌ Kembali"):
                    st.session_state.confirm_delete_all = False
                    st.rerun()
    
    else:
        st.info("Belum ada file yang disimpan.")

