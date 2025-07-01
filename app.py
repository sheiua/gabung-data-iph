import streamlit as st
from openpyxl import load_workbook, Workbook
import xlwt
import io
import zipfile

st.title("📊 Aplikasi Gabung Data IPH")

tahun = st.selectbox("Pilih Tahun", [2023, 2024, 2025], index=2)
bulan = st.selectbox(
    "Pilih Bulan",
    ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
     "Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    index=0
)
map_bulan = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12",
}
bulan_num = map_bulan[bulan]

uploaded_files = st.file_uploader("Upload file Excel (.xlsx)", type="xlsx", accept_multiple_files=True)

if st.button("Proses & Unduh ZIP") and uploaded_files:
    semua_kab, semua_prov = [], []
    indeks_kolom_kab = list(range(10))
    indeks_kolom_prov = list(range(6))

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:

        for f in uploaded_files:
            wb = load_workbook(f, data_only=True)
            sheets = wb.sheetnames

            sheet_kab = wb["360 KabKota"] if "360 KabKota" in sheets else None
            sheet_prov = wb["Provinsi"] if "Provinsi" in sheets else None

            if sheet_kab:
                wb.remove(sheet_kab)
                wb._sheets.insert(0, sheet_kab)
            if sheet_prov:
                wb.remove(sheet_prov)
                wb._sheets.insert(1, sheet_prov)

            if sheet_kab:
                for r in sheet_kab.iter_rows(min_row=2, values_only=True):
                    if r[0] and str(r[0]).startswith("18"):
                        semua_kab.append([r[i] if i < len(r) else None for i in indeks_kolom_kab])

            if sheet_prov:
                for r in sheet_prov.iter_rows(min_row=2, values_only=True):
                    if r[0]:
                        semua_prov.append([r[i] if i < len(r) else None for i in indeks_kolom_prov])

            # Save cleaned original
            if sheet_kab or sheet_prov:
                wb_clean = Workbook()
                has_data = False
                if sheet_kab:
                    ws_kab = wb_clean.active
                    ws_kab.title = "360 KabKota"
                    for r in sheet_kab.iter_rows(values_only=True):
                        if any(r):
                            ws_kab.append(r)
                            has_data = True
                if sheet_prov:
                    ws_prov = wb_clean.create_sheet("Provinsi")
                    for r in sheet_prov.iter_rows(values_only=True):
                        if any(r):
                            ws_prov.append(r)
                            has_data = True
                if has_data:
                    buf = io.BytesIO()
                    wb_clean.save(buf)
                    buf.seek(0)
                    zf.writestr(f"{f.name}_CLEANED.xlsx", buf.read())

        # Buat XLS kabupaten
        if semua_kab:
            bk = xlwt.Workbook()
            sk = bk.add_sheet("Gabungan_Kabupaten")
            for i, col in enumerate([
                "kode_kab", "pulau", "nama_prov", "nama_kab",
                "NON HK", "Perubahan IPH", "Komoditas Andil Besar",
                "Fluktuasi Harga Tertinggi Minggu Berjalan",
                "Nilai CV (Nilai Fluktuasi)", "status"
            ]):
                sk.write(0, i, col)
            for i, row in enumerate(semua_kab, 1):
                for j, val in enumerate(row):
                    sk.write(i, j, val)
            buf = io.BytesIO()
            bk.save(buf)
            buf.seek(0)
            zf.writestr(f"kabupaten_{bulan_num}_{tahun}.xls", buf.read())

        # Buat XLS provinsi
        if semua_prov:
            bp = xlwt.Workbook()
            sp = bp.add_sheet("Gabungan_Provinsi")
            for i, col in enumerate([
                "kode_prov", "nama_prov", "Perubahan IPH",
                "Komoditas Andil Terbesar",
                "Fluktuasi Harga Tertinggi Minggu Berjalan",
                "Nilai CV (Nilai Fluktuasi)"
            ]):
                sp.write(0, i, col)
            for i, row in enumerate(semua_prov, 1):
                for j, val in enumerate(row):
                    sp.write(i, j, val)
            buf = io.BytesIO()
            bp.save(buf)
            buf.seek(0)
            zf.writestr(f"provinsi_{bulan_num}_{tahun}.xls", buf.read())

    zip_buffer.seek(0)
    st.download_button(
        "📥 Unduh ZIP",
        data=zip_buffer,
        file_name=f"gabungan_IPH_{bulan_num}_{tahun}.zip",
        mime="application/zip"
    )
