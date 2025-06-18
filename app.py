import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import xlwt
from datetime import datetime
import io
import zipfile

st.title("📊 Aplikasi Gabung Data Excel Harga IPH")

# Dropdown Tahun dan Bulan
tahun = st.selectbox("Pilih Tahun", options=[2024, 2025], index=1)
bulan = st.selectbox("Pilih Bulan", options=[f"{i:02d}" for i in range(1, 13)], index=0)
indeks_kolom = []
indeks_kolom_prop  = []


# Upload file Excel
uploaded_files = st.file_uploader("Upload beberapa file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# Tombol proses
if st.button("Proses dan Unduh .xls") and uploaded_files:

    semua_data = []

    semua_data_prop = []
    # Indeks kolom yang ingin diambil dari file Excel (0-based)
    if tahun == 2025:
        indeks_kolom_prop = [0, 1, 2, 3, 4, 5]
        indeks_kolom = [0, 2, 3, 4, 5, 8, 9, 10]
    elif tahun == 2024:
        indeks_kolom = [0, 1, 2, 3, 4, 7, 8, 9]

#    print(indeks_kolom)

    def extract_minggu_from_filename(filename):
        for i in range(1, 6):
            if f"M{i}" in filename.upper():
                return i
        return None

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(uploaded_file, data_only=True)            
            nama_file = uploaded_file.name
            minggu = extract_minggu_from_filename(nama_file)

            #baca sheet pertama, data kabupaten
            ws = wb.worksheets[0] 
            for row in ws.iter_rows(min_row=2, values_only=True):
                row = list(row)
                kode = row[0] if len(row) > 0 else None

                if isinstance(kode, (int, str)) and str(kode).startswith("18"):
                    selected = []
                    for i in indeks_kolom:
                        selected.append(row[i] if i < len(row) else None)
                    semua_data.append((minggu, selected))

            #baca sheet kedua, data provinsi
            if len(wb.worksheets) > 1:
                ws2 = wb.worksheets[1]
                # proses seperti biasa
            else:
                st.warning(f"File {nama_file} hanya memiliki 1 sheet, sheet kedua (provinsi) dilewati.")
            #ws2 = wb.worksheets[1]
            for row2 in ws2.iter_rows(min_row=2, values_only=True):
                row2 = list(row2)
                kode2 = row2[0] if len(row2) > 0 else None

                if isinstance(kode2, (int, str)):
                    selected = []
                    for i in indeks_kolom_prop:
                        selected.append(row2[i] if i < len(row2) else None)
                    semua_data_prop.append((minggu, selected))

        except Exception as e:
            st.error(f"Gagal memproses file {uploaded_file.name}: {e}")

    if semua_data or semua_data_prop:
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:

            if semua_data:
                book = xlwt.Workbook()
                sheet = book.add_sheet("Gabungan_Kabupaten")

                headers = [
                    "id", "tahun", "bulan", "minggu", "kode_kab", "prov", "kab", "nilai_iph",
                    "komoditas", "fluktuasi_harga_tertinggi", "nilai_fluktuasi_tertinggi",
                    "disparitas_harga_antar_wilayah", "date_created"
                ]

                for col_idx, val in enumerate(headers):
                    sheet.write(0, col_idx, val)

                today = datetime.today().strftime("%Y-%m-%d")

                for idx, (minggu, data_row) in enumerate(semua_data, start=1):
                    replaced_komoditas = data_row[4].replace(",",";")
                    baris = [
                        idx,
                        str(tahun),
                        bulan,
                        minggu,
                        data_row[0],
                        data_row[1],
                        data_row[2],
                        data_row[3],
                        replaced_komoditas,
                        data_row[5],
                        data_row[6],
                        data_row[7],
                        today
                    ]
                    for col_idx, val in enumerate(baris):
                        sheet.write(idx, col_idx, val)

                output_kab = io.BytesIO()
                book.save(output_kab)
                output_kab.seek(0)
                zip_file.writestr(f"gabungan_{bulan}_{tahun}_kabupaten.xls", output_kab.read())

            if semua_data_prop:
                book = xlwt.Workbook()
                sheet = book.add_sheet("Gabungan_Provinsi")

                headers = [
                    "id", "tahun", "bulan", "minggu", "kode_prov", "prov", "nilai_iph",
                    "komoditas", "fluktuasi_harga_tertinggi", "nilai_fluktuasi_tertinggi",
                    "disparitas_harga_antar_wilayah", "date_created"
                ]

                for col_idx, val in enumerate(headers):
                    sheet.write(0, col_idx, val)

                today = datetime.today().strftime("%Y-%m-%d")

                for idx, (minggu, data_row) in enumerate(semua_data_prop, start=1):
                    replaced_komoditas = data_row[3].replace(",",";")
                    baris = [
                        idx,
                        str(tahun),
                        bulan,
                        minggu,
                        data_row[0],
                        data_row[1],
                        data_row[2],
                        replaced_komoditas,
                        data_row[4],
                        data_row[5],
                        "",
                        today
                    ]
                    for col_idx, val in enumerate(baris):
                        sheet.write(idx, col_idx, val)

                output_prov = io.BytesIO()
                book.save(output_prov)
                output_prov.seek(0)
                zip_file.writestr(f"gabungan_{bulan}_{tahun}_provinsi.xls", output_prov.read())

        zip_buffer.seek(0)
        st.download_button(
            "📦 Unduh Gabungan File (Kabupaten + Provinsi)",
            data=zip_buffer,
            file_name=f"gabungan_IPH_{bulan}_{tahun}.zip",
            mime="application/zip"
        )
    else:
        st.warning("❗ Tidak ada data yang dapat disimpan.")
