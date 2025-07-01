import streamlit as st
from openpyxl import load_workbook, Workbook
import xlwt
import io
import zipfile
import re

st.title("📊 Aplikasi Gabung Data IPH")

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

def ekstrak_komoditas_tertinggi(row, header):
    try:
        idx = header.index("Komoditas Andil Besar")
        data = row[idx]
    except (ValueError, IndexError):
        return "", ""

    if not data:
        return "", ""

    komoditas_data = str(data).split(';')
    max_komoditas = ""
    max_value = 0.0
    for item in komoditas_data:
        match = re.search(r"(.+?)\((-?\d+\.?\d*)\)", item.strip())
        if match:
            nama = match.group(1).strip()
            try:
                nilai = float(match.group(2))
                if abs(nilai) > abs(max_value):
                    max_value = nilai
                    max_komoditas = nama
            except:
                continue
    return max_komoditas, round(abs(max_value), 6)

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

            if sheet_kab:
                if not header_kab:
                    header_kab = [cell.value for cell in next(sheet_kab.iter_rows(min_row=1, max_row=1))]
                for r in sheet_kab.iter_rows(min_row=2, values_only=True):
                    if r[0] and str(r[0]).startswith("18"):
                        semua_kab.append(list(r))

            if sheet_prov:
                if not header_prov:
                    header_prov = [cell.value for cell in next(sheet_prov.iter_rows(min_row=1, max_row=1))]
                for r in sheet_prov.iter_rows(min_row=2, values_only=True):
                    if r[0]:
                        semua_prov.append(list(r))

        # ==== Sheet KABUPATEN ====
        if semua_kab:
            if "Fluktuasi Harga Tertinggi Minggu Berjalan" not in header_kab:
                header_kab.append("Fluktuasi Harga Tertinggi Minggu Berjalan")
            if "Nilai CV (Nilai fluktuasi)" not in header_kab:
                header_kab.append("Nilai CV (Nilai fluktuasi)")

            for row in semua_kab:
                komoditas, nilai = ekstrak_komoditas_tertinggi(row, header_kab)
                while len(row) < len(header_kab):
                    row.append("")
                row[header_kab.index("Fluktuasi Harga Tertinggi Minggu Berjalan")] = komoditas
                row[header_kab.index("Nilai CV (Nilai fluktuasi)")] = nilai

            # Hapus kolom kosong dari sheet Kab (tidak wajib di sini)
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

        # ==== Sheet PROVINSI ====
        if semua_prov:
            # Kolom akhir yang ingin disimpan
            kolom_diambil = [
                "kode_prov", "nama_prov", "Perubahan IPH",
                "Komoditas Andil Besar", "Fluktuasi Harga Tertinggi Minggu Berjalan", "Nilai CV (Nilai fluktuasi)"
            ]

            # Tambahkan kolom jika belum ada
            for kol in kolom_diambil[-2:]:
                if kol not in header_prov:
                    header_prov.append(kol)

            for row in semua_prov:
                komoditas, nilai = ekstrak_komoditas_tertinggi(row, header_prov)
                while len(row) < len(header_prov):
                    row.append("")
                row[header_prov.index("Fluktuasi Harga Tertinggi Minggu Berjalan")] = komoditas
                row[header_prov.index("Nilai CV (Nilai fluktuasi)")] = nilai

            # Cek indeks kolom yang ada
            idx_diambil = [i for i, h in enumerate(header_prov) if h in kolom_diambil]
            header_prov_bersih = [header_prov[i] for i in idx_diambil]
            semua_prov_bersih = [
                [row[i] if i < len(row) else "" for i in idx_diambil] for row in semua_prov
            ]

            bp = xlwt.Workbook()
            sp = bp.add_sheet("Provinsi")
            for i, col in enumerate(header_prov_bersih):
                sp.write(0, i, col)
            for i, row in enumerate(semua_prov_bersih, 1):
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
