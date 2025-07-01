import streamlit as st
from openpyxl import load_workbook, Workbook
import xlwt
import io
import zipfile

st.title("📊 Aplikasi Gabung Data IPH (Tanpa Kolom Tambahan)")

# Pilih tahun & bulan
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
    header_kab, header_prov = [], []

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in uploaded_files:
            wb = load_workbook(f, data_only=True)
            sheets = wb.sheetnames

            sheet_kab = wb["360 KabKota"] if "360 KabKota" in sheets else None
            sheet_prov = wb["Provinsi"] if "Provinsi" in sheets else None

            # ==== Sheet KabKota ====
            if sheet_kab:
                if not header_kab:
                    header_kab = [cell.value for cell in next(sheet_kab.iter_rows(min_row=1, max_row=1))]
                for r in sheet_kab.iter_rows(min_row=2, values_only=True):
                    if r[0] and str(r[0]).startswith("18"):
                        semua_kab.append(list(r))

            # ==== Sheet Provinsi ====
            if sheet_prov:
                if not header_prov:
                    header_prov = [cell.value for cell in next(sheet_prov.iter_rows(min_row=1, max_row=1))]
                for r in sheet_prov.iter_rows(min_row=2, values_only=True):
                    if r[0]:
                        semua_prov.append(list(r))

        # ==== Simpan KABUPATEN ====
        if semua_kab:
            bk = xlwt.Workbook()
            sk = bk.add_sheet("Gabungan_Kabupaten")
            for i, col in enumerate(header_kab):
                sk.write(0, i, col)
            for i, row in enumerate(semua_kab, 1):
                for j, val in enumerate(row):
                    sk.write(i, j, val)
            buf = io.BytesIO()
            bk.save(buf)
            buf.seek(0)
            zf.writestr(f"kabupaten_{bulan_num}_{tahun}.xls", buf.read())

        # ==== Simpan PROVINSI ====
        if semua_prov:
            kol_dihapus = ["Upaya Pemda (Monev)", "Saran Kepada Pemda", "Disparitas Harga Antar Daerah"]
            idx_simpan = [i for i, h in enumerate(header_prov) if h not in kol_dihapus]
            header_final = [header_prov[i] for i in idx_simpan]
            data_final = [[row[i] for i in idx_simpan] for row in semua_prov]

            bp = xlwt.Workbook()
            sp = bp.add_sheet("Provinsi")
            for i, col in enumerate(header_final):
                sp.write(0, i, col)
            for i, row in enumerate(data_final, 1):
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
