import streamlit as st
from openpyxl import load_workbook
import xlwt
import io
import zipfile

st.title("📊 Aplikasi Gabung Data Excel Harga IPH")

# Pilih tahun & bulan
tahun = st.selectbox("Pilih Tahun", [2023, 2024, 2025], index=2, key="tahun")
bulan = st.selectbox(
    "Pilih Bulan",
    ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
     "Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    index=0,
    key="bulan"
)
map_bulan = {
    "Januari": "01", "Februari": "02", "Maret": "03", "April": "04",
    "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08",
    "September": "09", "Oktober": "10", "November": "11", "Desember": "12",
}
bulan_num = map_bulan[bulan]

uploaded_files = st.file_uploader(
    "Upload file Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if st.button("🔄 Proses & Unduh ZIP") and uploaded_files:
    semua_kab, semua_prov = [], []
    header_kab, header_prov = [], []

    kolom_dihapus = [
        "Upaya Pemda (Monev)",
        "Saran Kepada Pemda",
        "Disparitas Harga antar Daerah"
    ]

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in uploaded_files:
            wb = load_workbook(f, data_only=True)
            sheetnames = wb.sheetnames

            sheet_kab = wb["360 KabKota"] if "360 KabKota" in sheetnames else None
            sheet_prov = wb["Provinsi"] if "Provinsi" in sheetnames else None

            if sheet_kab:
                rows = list(sheet_kab.iter_rows(values_only=True))
                if not header_kab:
                    header_kab = [cell for cell in rows[0]]

                for row in rows[1:]:
                    if any("Row Label" in str(cell) or "Grand Total" in str(cell) for cell in row):
                        continue
                    kode_kab = str(row[0]).strip() if row[0] is not None else ""
                    if kode_kab.isdigit() and 4 <= len(kode_kab) <= 6:
                        semua_kab.append(list(row))

            if sheet_prov:
                rows = list(sheet_prov.iter_rows(values_only=True))
                if not header_prov:
                    header_prov = [cell for cell in rows[0]]

                for row in rows[1:]:
                    if any("Row Label" in str(cell) or "Grand Total" in str(cell) for cell in row):
                        continue
                    if row[0]:
                        semua_prov.append(list(row))

        if semua_kab:
            idx_simpan_kab = [i for i, h in enumerate(header_kab) if h not in kolom_dihapus]
            header_kab_final = [header_kab[i] for i in idx_simpan_kab]
            data_kab_final = [[row[i] for i in idx_simpan_kab] for row in semua_kab]

            bk = xlwt.Workbook()
            sk = bk.add_sheet("Gabungan_Kabupaten")

            for i, col in enumerate(header_kab_final):
                sk.write(0, i, col)
                sk.col(i).width = 5000  # Atur lebar default

            for i, row in enumerate(data_kab_final, 1):
                for j, val in enumerate(row):
                    sk.write(i, j, val)
                    # Update lebar kolom jika data panjang
                    lebar = max(len(str(val)) * 256, sk.col(j).width)
                    sk.col(j).width = min(lebar, 10000)  # Batas max 10000

            buf = io.BytesIO()
            bk.save(buf)
            buf.seek(0)
            zf.writestr(f"kabupaten_{bulan_num}_{tahun}.xls", buf.read())

        if semua_prov:
            idx_simpan_prov = [i for i, h in enumerate(header_prov) if h not in kolom_dihapus]
            header_prov_final = [header_prov[i] for i in idx_simpan_prov]
            data_prov_final = [[row[i] for i in idx_simpan_prov] for row in semua_prov]

            bp = xlwt.Workbook()
            sp = bp.add_sheet("Gabungan_Provinsi")

            for i, col in enumerate(header_prov_final):
                sp.write(0, i, col)
                sp.col(i).width = 5000

            for i, row in enumerate(data_prov_final, 1):
                for j, val in enumerate(row):
                    sp.write(i, j, val)
                    lebar = max(len(str(val)) * 256, sp.col(j).width)
                    sp.col(j).width = min(lebar, 10000)

            buf = io.BytesIO()
            bp.save(buf)
            buf.seek(0)
            zf.writestr(f"provinsi_{bulan_num}_{tahun}.xls", buf.read())

    zip_buffer.seek(0)
    st.download_button(
        "📥 Unduh Gabungan IPH",
        data=zip_buffer,
        file_name=f"gabungan_IPH_{bulan_num}_{tahun}.zip",
        mime="application/zip"
    )
