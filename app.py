# ==================== GABUNG KABUPATEN ====================
if semua_kab:
    idx_simpan_kab = [i for i, h in enumerate(header_kab) if h not in kolom_dihapus]
    header_kab_final = [header_kab[i] for i in idx_simpan_kab]
    data_kab_final = [[row[i] for i in idx_simpan_kab] for row in semua_kab]

    bk = xlwt.Workbook()
    sk = bk.add_sheet("Gabungan_Kabupaten")

    header_styles = [
        xlwt.easyxf('pattern: pattern solid, fore_colour red; font: bold on, colour white;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour light_blue; font: bold on, colour white;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour light_green; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour gray25; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour orange; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour pink; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour turquoise; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour gold; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour lime; font: bold on, colour black;')
    ]

    # Hitung panjang maksimum per kolom (header + data)
    max_lens = [len(str(col)) for col in header_kab_final]
    for row in data_kab_final:
        for i, val in enumerate(row):
            panjang = len(str(val)) if val is not None else 0
            if panjang > max_lens[i]:
                max_lens[i] = panjang

    # Tulis header dan set lebar kolom sesuai max panjang + padding
    for i, col in enumerate(header_kab_final):
        style = header_styles[i % len(header_styles)]
        sk.write(0, i, col, style)
        sk.col(i).width = (max_lens[i] + 2) * 256  # +2 padding

    # Tulis data
    for i, row in enumerate(data_kab_final, 1):
        for j, val in enumerate(row):
            sk.write(i, j, val)


    buf = io.BytesIO()
    bk.save(buf)
    buf.seek(0)
    zf.writestr(f"kabupaten_{bulan_num}_{tahun}.xls", buf.read())

# ==================== GABUNG PROVINSI ====================
if semua_prov:
    idx_simpan_prov = [i for i, h in enumerate(header_prov) if h not in kolom_dihapus]
    header_prov_final = [header_prov[i] for i in idx_simpan_prov]
    data_prov_final = [[row[i] for i in idx_simpan_prov] for row in semua_prov]

    bp = xlwt.Workbook()
    sp = bp.add_sheet("Gabungan_Provinsi")

    header_styles_prov = [
        xlwt.easyxf('pattern: pattern solid, fore_colour blue; font: bold on, colour white;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour green; font: bold on, colour white;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour orange; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour pink; font: bold on, colour black;'),
        xlwt.easyxf('pattern: pattern solid, fore_colour turquoise; font: bold on, colour black;'),
    ]

    # Hitung panjang maksimum per kolom (header + data)
    max_lens_prov = [len(str(col)) for col in header_prov_final]
    for row in data_prov_final:
        for i, val in enumerate(row):
            panjang = len(str(val)) if val is not None else 0
            if panjang > max_lens_prov[i]:
                max_lens_prov[i] = panjang

    # Tulis header dan set lebar kolom sesuai max panjang + padding
    for i, col in enumerate(header_prov_final):
        style = header_styles_prov[i % len(header_styles_prov)]
        sp.write(0, i, col, style)
        sp.col(i).width = (max_lens_prov[i] + 2) * 256  # +2 padding

    # Tulis data
    for i, row in enumerate(data_prov_final, 1):
        for j, val in enumerate(row):
            sp.write(i, j, val)

    buf = io.BytesIO()
    bp.save(buf)
    buf.seek(0)
    zf.writestr(f"provinsi_{bulan_num}_{tahun}.xls", buf.read())
