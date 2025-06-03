import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import xlwt
from datetime import datetime
import io

st.title("📊 Aplikasi Gabung Data Excel Harga IPH")

# Input tahun dan bulan
tahun = st.selectbox("Pilih Tahun", options=[2024, 2025], index=5)
bulan = st.selectbox("Pilih Bulan", options=[f"{i:02d}" for i in range(1, 13)], index=0)

# Upload file
uploaded_files = st.file_uploader("Upload beberapa file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# Tombol proses
if st.button("Proses dan Unduh .xls") and uploaded_files:

    semua_data = []

    # Kolom yang diambil
    indeks_kolom = []

    # Indeks kolom yang ingin diambil dari file Excel (0-based)
    if tahun == '2025':
        indeks_kolom = [0, 2, 3, 4, 5, 8, 9, 10]
    elif tahun == '2024':
        indeks_kolom = [0, 1, 2, 3, 4, 7, 8, 9]

    # Fungsi ambil minggu dari nama file
    def extract_minggu_from_filename(filename):
        for i in range(1, 6):
            if f"M{i}" in filename.upper():
                return i
        return None

    for uploaded_file in uploaded_files:
        try:
            wb = load_workbook(uploaded_file, data_only=True)
            ws = wb.active
            nama_file = uploaded_file.name
            minggu = extract_minggu_from_filename(nama_file)

            for row in ws.iter_rows(min_row=2, values_only=True):
                row = list(row)
                kode = row[0] if len(row) > 0 else None

                if isinstance(kode, (int, str)) and str(kode).startswith("18"):
                    selected = []
                    for i in indeks_kolom:
                        selected.append(row[i] if i < len(row) else None)
                    semua_data.append((minggu, selected))
        except Exception as e:
            st.error(f"Gagal memproses file {uploaded_file.name}: {e}")

    if semua_data:
        book = xlwt.Workbook()
        sheet = book.add_sheet("Gabungan")

        headers = [
            "id", "tahun", "bulan", "minggu", "kode_kab", "prov", "kab", "nilai_iph",
            "komoditas", "fluktuasi_harga_tertinggi", "nilai_fluktuasi_tertinggi",
            "disparitas_harga_antar_wilayah", "date_created"
        ]

        # Tulis header
        for col_idx, val in enumerate(headers):
            sheet.write(0, col_idx, val)

        today = datetime.today().strftime("%Y-%m-%d")

        for idx, (minggu, data_row) in enumerate(semua_data, start=1):
            baris = [
                idx,
                str(tahun),
                bulan,
                minggu,
                data_row[0],
                data_row[1],
                data_row[2],
                data_row[3],
                data_row[4],
                data_row[5],
                data_row[6],
                data_row[7],
                today
            ]
            for col_idx, val in enumerate(baris):
                sheet.write(idx, col_idx, val)

        # Simpan ke buffer .xls
        output = io.BytesIO()
        book.save(output)
        output.seek(0)

        filename = f"gabungan_{bulan}_{tahun}.xls"
        st.success("✅ File berhasil diproses.")
        st.download_button("📥 Unduh File .xls", data=output, file_name=filename, mime="application/vnd.ms-excel")
    else:
        st.warning("❗ Tidak ada data yang sesuai filter kode diawali '18'.")
