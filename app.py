import streamlit as st
from openpyxl import load_workbook, Workbook
import xlwt
import io
import zipfile

st.title("📊 Aplikasi Gabung Data Excel Harga IPH")

# Pilih Tahun & Bulan
tahun = st.selectbox("Pilih Tahun", options=[2023, 2024, 2025], index=2)
bulan_nama = st.selectbox(
    "Pilih Bulan",
    options=[
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ],
    index=0
)

map_bulan = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12",
}
bulan = map_bulan[bulan_nama]

uploaded_files = st.file_uploader(
    "Upload beberapa file Excel (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("Proses & Unduh .zip") and uploaded_files:

    semua_data_kab = []
    semua_data_prov = []

    indeks_kolom_kab = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    indeks_kolom_prov = [0, 1, 2, 3, 4, 5]

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for uploaded_file in uploaded_files:
            try:
                wb = load_workbook(uploaded_file, data_only=True)
                nama_file = uploaded_file.name

                sheet_kab = wb["360 KabKota"] if "360 KabKota" in wb.sheetnames else None
                sheet_prov = wb["Provinsi"] if "Provinsi" in wb.sheetnames else None

                if sheet_kab:
                    wb.remove(sheet_kab)
                    wb._sheets.insert(0, sheet_kab)
                if sheet_prov:
                    wb.remove(sheet_prov)
                    wb._sheets.insert(1, sheet_prov)

                if sheet_kab:
                    for row in sheet_kab.iter_rows(min_row=2, values_only=True):
                        if row[0] and str(row[0]).startswith("18"):
                            selected = [row[i] if i < len(row) else None for i in indeks_kolom_kab]
                            semua_data_kab.append(selected)

                if sheet_prov:
                    for row in sheet_prov.iter_rows(min_row=2, values_only=True):
                        if row[0]:
                            selected = [row[i] if i < len(row) else None for i in indeks_kolom_prov]
                            semua_data_prov.append(selected)

                # Only create cleaned workbook if there is data
                if sheet_kab or sheet_prov:
                    wb_clean = Workbook()
                    made_clean = False
                    if sheet_kab:
                        ws_clean_kab = wb_clean.active
                        ws_clean_kab.title = "360 KabKota"
                        for row in sheet_kab.iter_rows(values_only=True):
                            if any(row):
                                ws_clean_kab.append(row)
                                made_clean = True
                    if sheet_prov:
                        ws_clean_prov = wb_clean.create_sheet("Provinsi")
                        for row in sheet_prov.iter_rows(values_only=True):
                            if any(row):
                                ws_clean_prov.append(row)
                                made_clean = True
                    if made_clean:
                        clean_buffer = io.BytesIO()
                        wb_clean.save(clean_buffer)
                        clean_buffer.seek(0)
                        zip_file.writestr(f"original_cleaned_{bulan}_{tahun}.xlsx", clean_buffer.read())

            except Exception as e:
                st.error(f"❌ Gagal proses file {uploaded_file.name}: {e}")

        # Kabupaten
        if semua_data_kab:
            book_kab = xlwt.Workbook()
            sheet_kab = book_kab.add_sheet("Gabungan_Kabupaten")
            headers_kab = [
                "kode_kab", "pulau", "nama_prov", "nama_kab",
                "NON HK", "Perubahan IPH", "Komoditas Andil Besar",
                "Fluktuasi Harga Tertinggi Minggu Berjalan",
                "Nilai CV (Nilai Fluktuasi)", "status"
            ]
            for col, val in enumerate(headers_kab):
                sheet_kab.write(0, col, val)
            for idx, row in enumerate(semua_data_kab, start=1):
                komoditas = str(row[6]).replace(",", ";")
                baris = [
                    row[0], row[1], row[2], row[3],
                    row[4], row[5], komoditas,
                    row[7], row[8], row[9]
                ]
                for col, val in enumerate(baris):
                    sheet_kab.write(idx, col, val)

            output_kab = io.BytesIO()
            book_kab.save(output_kab)
            output_kab.seek(0)
            zip_file.writestr(f"gabungan_{bulan}_{tahun}_kabupaten.xls", output_kab.read())

        # Provinsi
        if semua_data_prov:
            book_prov = xlwt.Workbook()
            sheet_prov = book_prov.add_sheet("Gabungan_Provinsi")
            headers_prov = [
                "kode_prov", "nama_prov", "Perubahan IPH",
                "Komoditas Andil Terbesar",
                "Fluktuasi Harga Tertinggi Minggu Berjalan",
                "Nilai CV (Nilai Fluktuasi)"
            ]
            for col, val in enumerate(headers_prov):
                sheet_prov.write(0, col, val)
            for idx, row in enumerate(semua_data_prov, start=1):
                komoditas = str(row[3]).replace(",", ";")
                baris = [
                    row[0], row[1], row[2],
                    komoditas, row[4], row[5]
                ]
                for col, val in enumerate(baris):
                    sheet_prov.write(idx, col, val)

            output_prov = io.BytesIO()
            book_prov.save(output_prov)
            output_prov.seek(0)
            zip_file.writestr(f"gabungan_{bulan}_{tahun}_provinsi.xls", output_prov.read())

    zip_buffer.seek(0)
    st.success("✅ Selesai! File ZIP siap diunduh.")
    st.download_button(
        "📥 Unduh Gabungan (.zip)",
        data=zip_buffer,
        file_name=f"gabungan_IPH_{bulan}_{tahun}.zip",
        mime="application/zip"
    )
