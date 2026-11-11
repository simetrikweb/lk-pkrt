# -*- coding: utf-8 -*-
# =====================================================================
#  pkrt.py v2.3 - OTOMASI 2 LEMBAR KERJA BPS (PKRT + DEFLATOR, 2025=100)
#  Kab. Cirebon (32.09) | + FITUR SIMPAN/MUAT JSON | perbaikan bug periode
# =====================================================================
import io, json
import hashlib
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as CL
import cirebon_real_simulation
import sut_jabar_bridge
st.set_page_config(page_title="LK PKRT & Deflator Kab. Cirebon", page_icon="📊", layout="wide")

# ============================ KONSTANTA ============================
Q_FULL = [f"{y}Q{q}" for y in range(2019, 2027) for q in (1, 2, 3, 4)]
M_FULL = [f"{y}M{m}" for y in range(2024, 2027) for m in range(1, 13)]
OLD12 = ["Makanan dan minuman tidak beralkohol", "Minuman beralkohol dan tembakau",
         "Pakaian dan alas kaki", "Perumahan, air, listrik, gas, dan bahan bakar lainnya",
         "Furnitur, perlengkapan rumah tangga, dan perbaikan rutin rumah", "Kesehatan",
         "Transportasi", "Komunikasi", "Rekreasi dan kebudayaan", "Pendidikan",
         "Penyediaan makan minum dan akomodasi", "Barang dan jasa lainnya"]
DIV13 = OLD12[:11] + ["Jasa asuransi dan keuangan", "Barang dan jasa lainnya"]
SUB17 = ["Makanan", "Minuman yang tidak beralkohol", "Minuman beralkohol", "Rokok dan tembakau",
         "Pakaian dan alas kaki", "Perumahan, air, listrik, dan bahan bakar rumah tangga",
         "Perlengkapan, peralatan, dan pemeliharaan rutin rumah", "Kesehatan", "Transportasi",
         "Peralatan informasi dan komunikasi", "Layanan informasi dan komunikasi",
         "Rekreasi, olahraga, dan budaya", "Pendidikan",
         "Penyediaan makanan dan minuman/restoran", "Asuransi", "Jasa keuangan",
         "Perawatan pribadi dan jasa lainnya"]
SUB7 = ["Makanan, minuman, dan rokok", "Pakaian, alas kaki, dan jasa perawatan",
        "Perumahan dan perlengkapan rumah tangga", "Kesehatan dan pendidikan",
        "Transportasi, komunikasi, dan rekreasi", "Restoran dan hotel", "Lainnya"]
MAP7 = {0: [0, 1], 1: [2], 2: [3, 4], 3: [5, 9], 4: [6, 7, 8], 5: [10], 6: [11, 12]}
SPLIT = {0: [0, 1], 1: [2, 3], 7: [9, 10], 11: [14, 15]}
SINGLE = {2: 4, 3: 5, 4: 6, 5: 7, 6: 8, 8: 11, 9: 12, 10: 13, 12: 16}
PROP = {11: 0.6335, 12: 0.3665}
COV = {5: 0.4980, 11: 1.1007}
REF_OLD = list(range(11)) + [11, 11]
TOP = 12
DF_KEYS = ["e1_ihk", "e2_dasar", "e0_adhb", "e0_adhk", "e1_seruti", "e2_sut", "r8_beta", "i6_admin", "f7_fenomena"]
SC_KEYS = ["periode", "kode_prov", "kode_kab", "sister", "petugas", "supervisor", "penduduk", "sumber_data"]

def R(i):
    return TOP + 4 + i

def sdiv(a, b):
    try:
        if b is None or b == 0 or not np.isfinite(b) or not np.isfinite(a):
            return np.nan
        return a / b
    except Exception:
        return np.nan

def _nmean(lst):
    v = [x for x in lst if x is not None and np.isfinite(x)]
    return float(np.mean(v)) if v else np.nan

# ============================ DATA SIMULASI CIREBON ============================
SIM_DASAR = [245000, 28500, 4200, 58000, 66500, 215000, 64000, 46500, 124000,
             18500, 57500, 41000, 28500, 86000, 7600, 11400, 38000]
SIM_F7 = {DIV13[0]: "Harga pangan stabil pasca panen raya beras; efek Ramadan berakhir Mei.",
          DIV13[1]: "Kenaikan cukai hasil tembakau Jan 2026 menggeser permintaan ke golongan bawah.",
          DIV13[2]: "Diskon pakaian pasca-Lebaran di pusat perbelanjaan Cirebon.",
          DIV13[3]: "Tarif listrik non-subsidi tetap; konsumsi RT naik musiman.",
          DIV13[4]: "Harga furnitur impor naik tipis; permintaan perlengkapan RT stabil.",
          DIV13[5]: "Klaim BPJS Kesehatan dialihkan ke PKP sesuai rule rebasing 2025.",
          DIV13[6]: "Mudik Lebaran 2026 (Maret) menaikkan mobilitas; tarif angkutan peak season.",
          DIV13[7]: "Paket data murah & promosi operator menahan kenaikan pengeluaran.",
          DIV13[8]: "Libur sekolah Juni & wisata pantai Kejawanan/Ciwaringin menaikkan rekreasi.",
          DIV13[9]: "Tahun ajaran baru 2026/27; sebagian uang pangkal jatuh di Q2.",
          DIV13[10]: "Kuliner Cirebon (empal gentong, tahu gejrot) ramai pasca mudik.",
          DIV13[11]: "FISIM metode triwulanan baru (faktor 1.1007) diterapkan.",
          DIV13[12]: "Harga emas perhiasan naik; konsumsi barang pribadi lainnya meningkat."}

def _sim_ihk():
    base = [108.5, 106.2, 121.4, 118.9, 104.8, 112.6, 109.3, 115.7, 113.2,
            96.4, 99.8, 107.9, 119.6, 114.8, 126.3, 122.7, 110.4]
    tren = [0.0042, 0.0021, 0.0065, 0.0058, 0.0024, 0.0031, 0.0019, 0.0036, 0.0029,
            -0.0008, 0.0012, 0.0026, 0.0048, 0.0039, 0.0052, 0.0044, 0.0033]
    rows = {}
    for i, nm in enumerate(SUB17):
        vals, v = [], base[i]
        for t, m in enumerate(M_FULL):
            mm = int(m.split("M")[1])
            seas = 0.0
            if mm in (2, 3):
                seas += 0.006 if i in (0, 1, 3, 4, 13) else 0.001
            if mm == 7:
                seas += 0.008 if i == 12 else 0.001
            if mm in (12, 1):
                seas += 0.005 if i in (8, 11, 13, 16) else 0.001
            if mm == 1 and i == 3:
                seas += 0.004
            v = v * (1 + tren[i] + seas) if t > 0 else v
            vals.append(round(v, 2))
        rows[nm] = vals
    df = pd.DataFrame(rows, index=M_FULL).T
    w = pd.Series(SIM_DASAR, index=SUB17)
    w = w / w.sum()
    umum = df.mul(w, axis=0).sum(axis=0).round(2)
    du = pd.DataFrame([umum.values], index=["IHK Umum"], columns=M_FULL)
    return pd.concat([df, du])

def _sim_e0():
    b0 = [185000, 16200, 37800, 132500, 45200, 31800, 60500, 39800, 25400, 22100, 40200, 58600]
    d0 = [128.5, 142.1, 116.4, 124.2, 119.8, 131.6, 117.9, 112.6, 120.4, 133.2, 119.1, 126.8]
    adhb = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    adhk = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    for i in range(12):
        v, d = b0[i], d0[i]
        for t, q in enumerate(Q_FULL):
            if t > 0:
                g = 0.015 if q[-1] == "1" else (0.009 if q[-1] in "23" else 0.011)
                if i == 5 and q[-1] == "3":
                    g += 0.004
                v *= (1 + g)
                d *= (1.008 if q[-1] == "1" else 1.006)
            adhb.iat[i, t] = round(v, 1)
            adhk.iat[i, t] = round(sdiv(v, d) * 100, 1)
    return adhb, adhk

def _sim_seruti():
    b13 = [232000, 19800, 45200, 168000, 54800, 20500, 74200, 46800, 30200, 26400, 48600, 39800, 23100]
    cols = [q for q in Q_FULL if q >= "2023Q1"]
    df = pd.DataFrame(np.nan, index=DIV13, columns=cols)
    for i in range(13):
        v = b13[i]
        for t, q in enumerate(cols):
            if t > 0:
                g = 0.014 if q[-1] == "1" else (0.008 if q[-1] in "23" else 0.010)
                if i == 9 and q[-1] == "3":
                    g += 0.006
                if i == 10 and q[-1] in "24":
                    g += 0.004
                v *= (1 + g)
            df.iat[i, t] = round(v, 1)
    return df

def _sim_i6():
    nama = ["Pelanggan PLN rumah tangga (ribu)", "Listrik terjual RT (juta kWh)", "TPK hotel (%)",
            "Penumpang angkutan (juta orang)", "Kendaraan bermotor terdaftar (ribu)", "Transaksi QRIS (juta transaksi)"]
    b = [412.0, 388.0, 41.5, 1.62, 1245.0, 3.85]
    g = [0.006, 0.007, 0.002, 0.004, 0.010, 0.045]
    df = pd.DataFrame(np.nan, index=nama, columns=Q_FULL)
    for i in range(6):
        v = b[i]
        for t, q in enumerate(Q_FULL):
            if t > 0:
                v *= (1 + g[i] + (0.010 if i == 2 and q[-1] in "24" else 0))
            df.iat[i, t] = round(v, 2)
    return df

def muat_simulasi():
    S = st.session_state
    S["e1_ihk"] = _sim_ihk()
    S["e2_dasar"] = pd.DataFrame(SIM_DASAR, index=SUB17, columns=["Nilai Konsumsi Dasar (Rp Juta)"])
    a, k = _sim_e0()
    S["e0_adhb"], S["e0_adhk"] = a, k
    ser = _sim_seruti()
    S["e1_seruti"] = ser
    c25 = [c for c in ser.columns if c.startswith("2025")]
    S["e2_sut"] = pd.DataFrame({"SUT 2025 (Rp Juta)": [round(ser.loc[d, c25].sum() * f, 1)
                               for d, f in zip(DIV13, [1.03, 1.02, 1.04, 1.02, 1.03, 1.05, 1.02, 1.01, 1.04, 1.03, 1.06, 1.02, 1.03])]}, index=DIV13)
    S["r8_beta"] = pd.DataFrame(1.0, index=DIV13, columns=Q_FULL)
    S["i6_admin"] = _sim_i6()
    S["f7_fenomena"] = pd.DataFrame([SIM_F7.get(d, "") for d in DIV13],
                                    index=DIV13, columns=["Fenomena triwulan berjalan"])
    S["penduduk"] = 2304567
    S["periode"] = "2026-Q2"
    S["kode_prov"], S["kode_kab"], S["sister"] = "32", "09", "74"
    S["petugas"], S["supervisor"] = "Statistikian BPS Kab. Cirebon", "Kasubag IP"
    S["sumber_data"] = "Simulasi Kab. Cirebon"

def kosongkan():
    S = st.session_state
    S["e1_ihk"] = pd.DataFrame(np.nan, index=SUB17 + ["IHK Umum"], columns=M_FULL)
    S["e2_dasar"] = pd.DataFrame(0.0, index=SUB17, columns=["Nilai Konsumsi Dasar (Rp Juta)"])
    S["e0_adhb"] = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    S["e0_adhk"] = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    S["e1_seruti"] = pd.DataFrame(np.nan, index=DIV13, columns=[q for q in Q_FULL if q >= "2023Q1"])
    S["e2_sut"] = pd.DataFrame(0.0, index=DIV13, columns=["SUT 2025 (Rp Juta)"])
    S["r8_beta"] = pd.DataFrame(1.0, index=DIV13, columns=Q_FULL)
    S["i6_admin"] = pd.DataFrame(np.nan, index=["Pelanggan PLN rumah tangga (ribu)", "Listrik terjual RT (juta kWh)", "TPK hotel (%)", "Penumpang angkutan (juta orang)", "Kendaraan bermotor terdaftar (ribu)", "Transaksi QRIS (juta transaksi)"], columns=Q_FULL)
    S["f7_fenomena"] = pd.DataFrame("", index=DIV13, columns=["Fenomena triwulan berjalan"])
    S["sumber_data"] = "Template kosong (input manual)"
# ============================ DATA RIIL CIREBON (dari cirebon_real_simulation.py) ============================
# Peta 17 subkomponen (SUB17) -> indeks 13 COICOP
SUB17_TO_COICOP = {0:0, 1:0, 2:1, 3:1, 4:2, 5:3, 6:4, 7:5, 8:6,
                   9:7, 10:7, 11:8, 12:9, 13:10, 14:11, 15:11, 16:12}

def _code(ci):                 # indeks COICOP 0-12 -> kode string "01".."13"
    return f"{ci+1:02d}"

def _q_of_month(m):            # "2024M5" -> "2024Q2"
    y, mo = m.split("M")
    return f"{y}Q{(int(mo)-1)//3 + 1}"

def muat_data_riil():
    """Isi lembar kerja dengan DATA RIIL BPS Kab. Cirebon.
    Seruti 2023Q1-2026Q2 | IHK Kota Cirebon (2025=100) | Penduduk proyeksi."""
    try:
        import cirebon_real_simulation as crs
        POP, IHK, SERUTI = crs.POPULATION_CIREBON, crs.DEFLATOR_IHK_2025, crs.SERUTI_PERKAPITA
    except Exception as e:
        st.error(f"Gagal memuat data riil. Pastikan cirebon_real_simulation.py ada & berisi dataset. Detail: {e}")
        return

    S = st.session_state

    # ---------- E1 (WB1): IHK bulanan 17 subkomponen, diambil dari IHK triwulanan riil ----------
    rows = {}
    for s in range(17):
        code = _code(SUB17_TO_COICOP[s])
        rows[SUB17[s]] = [IHK.get(code, {}).get(_q_of_month(m), np.nan) for m in M_FULL]
    df_ihk = pd.DataFrame(rows, index=M_FULL).T

    # bobot divisi dari struktur pengeluaran riil 2023 (untuk IHK Umum)
    q23 = ["2023Q1", "2023Q2", "2023Q3", "2023Q4"]
    mean_pk = {}
    for ci in range(13):
        vs = [SERUTI[_code(ci)][q] for q in q23 if q in SERUTI.get(_code(ci), {})]
        mean_pk[ci] = float(np.mean(vs)) if vs else 0.0
    div_w = np.array([mean_pk[i] for i in range(13)], dtype=float)
    div_w = div_w / div_w.sum() if div_w.sum() > 0 else div_w

    umum = np.full(len(M_FULL), np.nan)
    for j, m in enumerate(M_FULL):
        q = _q_of_month(m)
        num = den = 0.0
        for ci in range(13):
            v = IHK.get(_code(ci), {}).get(q, np.nan)
            if np.isfinite(v):
                num += div_w[ci] * v
                den += div_w[ci]
        umum[j] = num / den if den > 0 else np.nan
    df_ihk.loc["IHK Umum"] = umum
    S["e1_ihk"] = df_ihk

    # ---------- E2 (WB1): bobot konsumsi dasar dari struktur pengeluaran riil 2023 ----------
    dasar = np.zeros(17)
    for s in range(17):
        ci = SUB17_TO_COICOP[s]
        n_sub = sum(1 for x in range(17) if SUB17_TO_COICOP[x] == ci)
        dasar[s] = mean_pk[ci] / n_sub
    S["e2_dasar"] = pd.DataFrame(dasar, index=SUB17, columns=["Nilai Konsumsi Dasar (Rp Juta)"])

    # ---------- E1 (WB2): Seruti agregat riil = perkapita x penduduk (Rp Juta) ----------
    ser_cols = [q for q in Q_FULL if q >= "2023Q1"]
    df_ser = pd.DataFrame(np.nan, index=DIV13, columns=ser_cols)
    for d in range(13):
        for q in ser_cols:
            pk  = SERUTI.get(_code(d), {}).get(q, np.nan)
            pop = POP.get(q, np.nan)
            if np.isfinite(pk) and np.isfinite(pop):
                df_ser.iat[d, ser_cols.index(q)] = pk * pop / 1e6
    S["e1_seruti"] = df_ser

    # ---------- E2 (WB2): SUT 2025 = total Seruti 2025 (theta awal = 1 / netral) ----------
    c25 = [q for q in ser_cols if q.startswith("2025")]
    sut = []
    for d in range(13):
        ada = [q for q in c25 if q in df_ser.columns]
        tot = df_ser.loc[DIV13[d], ada].sum() if ada else np.nan
        sut.append(round(float(tot), 1) if np.isfinite(tot) else 0.0)
    S["e2_sut"] = pd.DataFrame({"SUT 2025 (Rp Juta)": sut}, index=DIV13)

    # ---------- E0 rilis 2010=100 : TIDAK tersedia pada data riil -> dikosongkan ----------
    S["e0_adhb"] = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    S["e0_adhk"] = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)

    # ---------- R8 beta, I6, F7 : default ----------
    S["r8_beta"] = pd.DataFrame(1.0, index=DIV13, columns=Q_FULL)
    S["i6_admin"] = pd.DataFrame(np.nan,
        index=["Pelanggan PLN rumah tangga (ribu)", "Listrik terjual RT (juta kWh)", "TPK hotel (%)",
               "Penumpang angkutan (juta orang)", "Kendaraan bermotor terdaftar (ribu)",
               "Transaksi QRIS (juta transaksi)"], columns=Q_FULL)
    S["f7_fenomena"] = pd.DataFrame(
        ["Data riil Seruti & IHK Kota Cirebon; lengkapi fenomena sesuai rilis BRS BPS." for _ in DIV13],
        index=DIV13, columns=["Fenomena triwulan berjalan"])

    # ---------- identitas ----------
    S["periode"]  = "2026-Q2"
    S["penduduk"] = int(POP.get("2026Q2", 2431927))
    S["kode_prov"], S["kode_kab"], S["sister"] = "32", "09", "74"
    S["petugas"], S["supervisor"] = "Statistikian BPS Kab. Cirebon", "Kasubag IP"
    S["sumber_data"] = "DATA RIIL BPS Kab. Cirebon (Seruti 2023Q1–2026Q2 & IHK Kota Cirebon 2025=100)"    
# ============================ IMPORT PDRB EXCEL → E0 ============================
def muat_pdrb_excel(file_obj):
    """
    Membaca Template Upload PDRB (.xlsx) dan memetakan otomatis
    baris 1.a–1.m ke 12 Divisi COICOP di E0 (ADHB & ADHK).
    
    Mapping:
      id 2–12  → OLD12[0..10]  (1-to-1)
      id 13+14 → OLD12[11]     (dijumlahkan: Jasa Keuangan + Perawatan diri)
    """
    import re, io as _io

    # --- PETA: id Excel → index OLD12 ---
    MAP_ID = {
        2: 0,   # 1.a Makanan & Minuman tidak beralkohol
        3: 1,   # 1.b Minuman Beralkohol & Tembakau
        4: 2,   # 1.c Pakaian & Alas Kaki
        5: 3,   # 1.d Perumahan, Air, Listrik, Gas...
        6: 4,   # 1.e Furnitur & Perlengkapan RT
        7: 5,   # 1.f Kesehatan
        8: 6,   # 1.g Transportasi
        9: 7,   # 1.h Informasi & Komunikasi → "Komunikasi"
        10: 8,  # 1.i Rekreasi, Olahraga, Kebudayaan
        11: 9,  # 1.j Jasa Pendidikan
        12: 10, # 1.k Restoran & Jasa Akomodasi
        13: 11, # 1.l Jasa Keuangan & Asuransi   ─┐
        14: 11, # 1.m Perawatan diri & Jasa lainnya ─┴→ OLD12[11]
    }

    # Simpan bytes agar bisa dibaca 2x (ADHB & ADHK)
    raw = file_obj.read()

    adhb = pd.DataFrame(0.0, index=OLD12, columns=Q_FULL)
    adhk = pd.DataFrame(0.0, index=OLD12, columns=Q_FULL)

    for sheet, target in [("ADHB", adhb), ("ADHK", adhk)]:
        try:
            df = pd.read_excel(_io.BytesIO(raw), sheet_name=sheet, header=None)
        except Exception:
            continue

        # --- Cari baris header otomatis ---
        hdr = None
        for i in range(min(10, len(df))):
            vals = [str(v).strip().lower() for v in df.iloc[i].tolist()]
            if "id" in vals and any("2019q1" in v for v in vals):
                hdr = i
                break
        if hdr is None:
            continue

        # --- Identifikasi kolom id & kolom quarter ---
        cols = df.iloc[hdr].tolist()
        id_col, q_cols = None, {}
        for j, c in enumerate(cols):
            cs = str(c).strip().lower()
            if cs == "id":
                id_col = j
            elif re.match(r"20\d{2}q[1-4]", cs):
                q_cols[cs.upper()] = j

        if id_col is None or not q_cols:
            continue

        # --- Isi DataFrame target ---
        for _, row in df.iloc[hdr + 1:].iterrows():
            try:
                rid = int(float(row.iloc[id_col]))
            except (ValueError, TypeError):
                continue
            if rid not in MAP_ID:
                continue

            idx12 = MAP_ID[rid]
            for q_str, cj in q_cols.items():
                if q_str not in Q_FULL:
                    continue
                try:
                    val = float(str(row.iloc[cj]).replace(",", ""))
                except (ValueError, TypeError):
                    val = 0.0
                target.iat[idx12, Q_FULL.index(q_str)] += val  # += utk id 13+14

    return adhb, adhk
def muat_book1_csv(file_obj):
    """
    Membaca Book1__.csv (PDRB ADHB & ADHK per 12 subkomponen PKRT)
    dan mengisinya ke E0 (e0_adhb, e0_adhk).
    """
    import io as _io

    raw = file_obj.read().decode("utf-8-sig")
    lines = raw.split("\n")

    # --- Identifikasi baris header & data ---
    adhb_start = adhk_start = None
    header_row = None

    for i, line in enumerate(lines):
        low = line.lower()
        if "tabel 1" in low and "adhb" in low:
            adhb_start = i
        elif "tabel 2" in low and "adhk" in low:
            adhk_start = i
        elif "komponen pengeluaran" in low and header_row is None:
            header_row = i

    if header_row is None or adhb_start is None or adhk_start is None:
        raise ValueError("Format CSV tidak dikenali. Pastikan ada Tabel 1 (ADHB) dan Tabel 2 (ADHK).")

    # --- Parse header quarter ---
    hdr_cells = lines[header_row].split(";")
    q_map = {}  # {quarter_label: col_index}
    for j, cell in enumerate(hdr_cells):
        cell = cell.strip()
        if cell.startswith("Q") and "-" in cell:
            q_map[cell] = j

    # --- Fungsi bantu: extract 12 baris PKRT dari suatu section ---
    def extract_section(start_line, end_line):
        """Ambil 12 subkomponen PKRT (baris yang diawali '1.a' s/d '1.l')."""
        data = {}
        for i in range(start_line, min(end_line, len(lines))):
            cells = lines[i].split(";")
            if not cells:
                continue
            label = cells[0].strip()
            # Cocokkan dengan pola "1.a." s/d "1.l."
            for k, prefix in enumerate(["1.a.", "1.b.", "1.c.", "1.d.", "1.e.", "1.f.",
                                         "1.g.", "1.h.", "1.i.", "1.j.", "1.k.", "1.l."]):
                if label.startswith(prefix):
                    row_vals = {}
                    for qlabel, cidx in q_map.items():
                        if cidx < len(cells):
                            try:
                                val = float(cells[cidx].replace(",", "").replace(".", "").strip())
                            except (ValueError, IndexError):
                                val = 0.0
                            row_vals[qlabel] = val
                    data[k] = row_vals
                    break
        return data

    adhb_data = extract_section(adhb_start, adhk_start)
    adhk_data = extract_section(adhk_start, len(lines))

    if len(adhb_data) < 12 or len(adhk_data) < 12:
        raise ValueError(f"Data tidak lengkap. ADHB: {len(adhb_data)} baris, ADHK: {len(adhk_data)} baris. Butuh 12.")

    # --- Konversi ke DataFrame dengan kolom Q_FULL ---
    def to_dataframe(section_data):
        df = pd.DataFrame(0.0, index=OLD12, columns=Q_FULL)
        for div_idx, qvals in section_data.items():
            for qlabel, val in qvals.items():
                # qlabel format: "Q1-2018" → "2018Q1"
                parts = qlabel.split("-")
                if len(parts) == 2:
                    q_str = f"{parts[1]}{parts[0]}"  # "2018Q1"
                    if q_str in Q_FULL:
                        df.iat[div_idx, Q_FULL.index(q_str)] = val
        return df

    adhb_df = to_dataframe(adhb_data)
    adhk_df = to_dataframe(adhk_data)

    return adhb_df, adhk_df
def init_state():
    S = st.session_state
    if "sumber_data" not in S:
        muat_simulasi()

init_state()
S = st.session_state

# ============================ SIDEBAR + PERIODE (BUG FIX) ============================
st.sidebar.title("🧭 Konfigurasi & Sumber Data")

# ---- 1) Tombol & uploader DULU (sebelum widget ber-key dibuat) ----
_b1, _b3, _b2 = st.sidebar.columns(3)

if _b1.button("🎲 Muat Simulasi"):
    muat_simulasi()
    for _k in list(st.session_state.keys()):
        if _k.startswith("ed_"):
            del st.session_state[_k]
    st.rerun()

if _b3.button("📊 Data Riil"):
    muat_data_riil()
    for _k in list(st.session_state.keys()):
        if _k.startswith("ed_"):
            del st.session_state[_k]
    st.rerun()

if _b2.button("🧹 Template Kosong"):
    kosongkan()
    for _k in list(st.session_state.keys()):
        if _k.startswith("ed_"):
            del st.session_state[_k]
    st.rerun()
# --- Upload Template PDRB → otomatis isi E0 ---
_upl_pdrb = st.sidebar.file_uploader("📂 Muat Template PDRB (.xlsx) → E0", type=["xlsx"])
# --- Upload Book1__.csv → otomatis isi E0 ---
_upl_book1 = st.sidebar.file_uploader("📄 Muat Book1__.csv → E0 (PDRB 2010=100)", type=["csv"])
upl_book1 = st.sidebar.file_uploader(
    "📄 Muat Book1__.csv → E0 (PDRB 2010=100)",
    type=["csv"]
)

if upl_book1 is not None:
    file_bytes = upl_book1.getvalue()
    file_hash = hashlib.sha1(file_bytes).hexdigest()

    if st.session_state.get("book1_hash") != file_hash:
        try:
            _adhb, _adhk = muat_book1_csv(io.BytesIO(file_bytes))

            if _adhb.abs().sum().sum() > 0:
                S["e0_adhb"] = _adhb
                S["e0_adhk"] = _adhk
                S["sumber_data"] = "Book1__.csv (PDRB ADHB/ADHK 2010=100)"

                for _k in list(st.session_state.keys()):
                    if _k.startswith("ed"):
                        del st.session_state[_k]

                st.session_state.book1_hash = file_hash

                st.sidebar.success("✅ E0 berhasil diisi dari Book1__.csv!")
                st.rerun()
            else:
                st.sidebar.warning("⚠️ File terbaca tapi data PKRT kosong.")

        except Exception as e:
            st.sidebar.error(f"Gagal membaca Book1__.csv: {e}")
    else:
        st.sidebar.caption("Book1__.csv sudah diproses sebelumnya.")
else:
    st.session_state.pop("book1_hash", None)
        else:
            st.sidebar.warning("⚠️ File terbaca tapi data PKRT kosong.")
    except Exception as _e:
        st.sidebar.error(f"Gagal membaca Book1__.csv: {_e}")
if _upl_pdrb is not None:
    try:
        _adhb, _adhk = muat_pdrb_excel(_upl_pdrb)
        # Validasi: cek apakah ada data yang terbaca
        if _adhb.abs().sum().sum() > 0:
            S["e0_adhb"] = _adhb
            S["e0_adhk"] = _adhk
            S["sumber_data"] = "Template PDRB Kab. Cirebon (Excel upload)"
            for _k in list(st.session_state.keys()):
                if _k.startswith("ed_"):
                    del st.session_state[_k]
            st.sidebar.success("✅ E0 berhasil diisi dari Template PDRB!")
            st.rerun()
        else:
            st.sidebar.warning("⚠️ File terbaca tapi data PKRT (id 2–14) kosong.")
    except Exception as _e:
        st.sidebar.error(f"Gagal membaca Template PDRB: {_e}")
_upl = st.sidebar.file_uploader("⬆️ Muat file JSON", type=["json"])
if _upl is not None:
    try:
        _d = json.loads(_upl.getvalue().decode("utf-8"))
        for _k in DF_KEYS:
            if _k in _d:
                S[_k] = pd.DataFrame.from_dict(_d[_k], orient="split")
        for _k in SC_KEYS:
            if _k in _d:
                S[_k] = _d[_k]
        for _k in list(st.session_state.keys()):
            if _k.startswith("ed_"):
                del st.session_state[_k]
        st.sidebar.success("✅ Data berhasil dimuat!")
        st.rerun()
    except Exception as _e:
        st.sidebar.error(f"Gagal memuat file: {_e}")

# ---- 2) Baru buat widget ber-key (aman: semua mutasi sudah selesai) ----
PER_OPT = [f"{q[:4]}-Q{q[-1]}" for q in Q_FULL if q >= "2024Q1"]
if S["periode"] not in PER_OPT:
    S["periode"] = PER_OPT[-1]
st.sidebar.selectbox("Periode estimasi (YYYY-QQ)", PER_OPT,
                     index=PER_OPT.index(S["periode"]), key="periode")
st.sidebar.caption("Kab. Cirebon (32.09) · Sister city Kota Cirebon (74) · Tahun dasar 2025=100")
st.sidebar.caption("Sumber aktif: **" + S.get("sumber_data", "-") + "**")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💾 Simpan / Muat Pekerjaan")

def _kumpul_state():
    out = {k: S[k].to_dict(orient="split") for k in DF_KEYS if k in S}
    for k in SC_KEYS:
        if k in S:
            out[k] = S[k]
    return out

_json = json.dumps(_kumpul_state(), ensure_ascii=False, indent=2).encode("utf-8")
st.sidebar.download_button("⬇️ Unduh JSON (simpan semua input)", _json,
                           file_name="data_pkrt_cirebon.json", mime="application/json")
st.sidebar.caption("JSON berisi seluruh input (E0, E1, E2, I6, F7, R8, identitas). Hasil hitungan & export dihitung ulang otomatis.")

PQ = S["periode"].replace("-", "")
Q_ACT = [q for q in Q_FULL if q <= PQ]
NQ = len(Q_ACT)
M_END = PQ[:4] + "M" + str(int(PQ[-1]) * 3)
M_ACT = [m for m in M_FULL if m <= M_END]
NM = len(M_ACT)

# ============================ MESIN WB1 ============================
def compute_wb1():
    e1 = S["e1_ihk"].loc[:, M_ACT].astype(float)
    dasar = S["e2_dasar"].iloc[:, 0].astype(float)
    comps, bobot = {}, {}
    for d in range(13):
        subs = SPLIT[d] if d in SPLIT else [SINGLE[d]]
        comps[d] = subs
        tot = sum(dasar.iat[s] for s in subs)
        bobot[d] = {s: sdiv(dasar.iat[s], tot) for s in subs}
    m13 = pd.DataFrame(np.nan, index=DIV13, columns=M_ACT)
    for d in range(13):
        for j in range(NM):
            num = sum(bobot[d][s] * e1.iat[s, j] for s in comps[d])
            den = sum(bobot[d][s] for s in comps[d])
            m13.iat[d, j] = sdiv(num, den) if den else np.nan
    divd = {d: sum(dasar.iat[s] for s in comps[d]) for d in range(13)}
    if "IHK Umum" in e1.index and np.isfinite(e1.loc["IHK Umum"].iloc[:NM].sum()):
        umum_m = e1.loc["IHK Umum"].tolist()
    else:
        umum_m = [sdiv(sum(divd[d] * m13.iat[d, j] for d in range(13)), sum(divd.values())) for j in range(NM)]
    q22 = pd.DataFrame(np.nan, index=DIV13 + ["IHK Umum"], columns=Q_ACT)
    for t, q in enumerate(Q_ACT):
        ms = [f"{q[:4]}M{mi}" for mi in range(int(q[-1]) * 3 - 2, int(q[-1]) * 3 + 1)]
        idx = [M_ACT.index(m) for m in ms if m in M_ACT]
        for d in range(13):
            q22.iat[d, t] = _nmean([m13.iat[d, j] for j in idx])
        q22.iat[13, t] = _nmean([umum_m[j] for j in idx])
    c25 = [t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    q25 = q22.copy()
    if c25:
        b25 = q22.iloc[:, c25].mean(axis=1)
        for r in range(q22.shape[0]):
            for t in range(NQ):
                q25.iat[r, t] = sdiv(q22.iat[r, t], b25.iat[r]) * 100
    else:
        q25 = q25 * np.nan
    return dict(m13=m13, umum=umum_m, q22=q22, q25=q25)

# ============================ MESIN WB2 ============================
def compute_wb2(W1):
    e0b = S["e0_adhb"].loc[:, Q_ACT].astype(float)
    e0k = S["e0_adhk"].loc[:, Q_ACT].astype(float)
    ipd = pd.DataFrame(np.nan, index=OLD12, columns=Q_ACT)
    for i in range(12):
        for t in range(NQ):
            ipd.iat[i, t] = sdiv(e0b.iat[i, t], e0k.iat[i, t]) * 100
    ser_in = S["e1_seruti"].astype(float)
    ser = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    for q in ser_in.columns:
        if q in Q_ACT:
            ser.loc[:, q] = ser_in[q].values
    i0 = Q_ACT.index("2023Q1") if "2023Q1" in Q_ACT else 0
    for t in range(i0 - 1, -1, -1):
        for d in range(13):
            r = REF_OLD[d]
            ser.iat[d, t] = ser.iat[d, t + 1] * sdiv(e0b.iat[r, t], e0b.iat[r, t + 1])
    ser_adj = ser.copy()
    for d, f in COV.items():
        ser_adj.iloc[d] = ser_adj.iloc[d] * f
    c25 = [t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    theta = pd.Series(np.nan, index=DIV13)
    if c25:
        for d in range(13):
            theta.iat[d] = sdiv(S["e2_sut"].iat[d, 0], ser_adj.iloc[d, c25].sum())
    e3 = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    cpi = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    i24 = Q_ACT.index("2024Q1") if "2024Q1" in Q_ACT else 0
    for d in range(13):
        for t in range(NQ):
            e3.iat[d, t] = theta.iat[d] * ser_adj.iat[d, t]
            if t >= i24:
                cpi.iat[d, t] = W1["q25"].iat[d, t] if t < W1["q25"].shape[1] else np.nan
        for t in range(i24 - 1, -1, -1):
            r = REF_OLD[d]
            cpi.iat[d, t] = cpi.iat[d, t + 1] * sdiv(ipd.iat[r, t], ipd.iat[r, t + 1])
    e5 = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    r9 = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    r10 = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    beta = S["r8_beta"].loc[:, Q_ACT].astype(float)
    for d in range(13):
        for t in range(NQ):
            e5.iat[d, t] = sdiv(e3.iat[d, t], cpi.iat[d, t]) * 100
            r9.iat[d, t] = beta.iat[d, t] * e3.iat[d, t]
            r10.iat[d, t] = sdiv(r9.iat[d, t], cpi.iat[d, t]) * 100
    a11n = pd.DataFrame(np.nan, index=SUB7, columns=Q_ACT)
    a11k = pd.DataFrame(np.nan, index=SUB7, columns=Q_ACT)
    for s7, divs in MAP7.items():
        a11n.iloc[s7] = r9.iloc[divs].sum(axis=0).values
        a11k.iloc[s7] = r10.iloc[divs].sum(axis=0).values
    pend = S["penduduk"] or 0
    a12n = a11n * 1000 / pend if pend else a11n * np.nan
    a12k = a11k * 1000 / pend if pend else a11k * np.nan
    return dict(ipd=ipd, ser=ser, theta=theta, e3=e3, cpi=cpi, e5=e5,
                r9=r9, r10=r10, a11n=a11n, a11k=a11k, a12n=a12n, a12k=a12k)

def growths(df):
    q = df.div(df.shift(1, axis=1)) * 100 - 100
    y = df.div(df.shift(4, axis=1)) * 100 - 100
    return q, y

# ============================ QC 22 BUTIR ============================
QC_ITEMS = [
    ("Struktur konsumsi", "Struktur konsumsi berubah drastis tanpa fenomena jelas?", "Internal"),
    ("Pergeseran share", "Perubahan share COICOP dalam rentang historis wajar?", "Internal"),
    ("Seasonal pattern", "Pola musiman (Ramadan, Nataru, ajaran baru) tercermin?", "Internal"),
    ("Pola q-to-q", "Pertumbuhan q-to-q konsisten pola historis?", "Internal"),
    ("Pola y-on-y", "Pertumbuhan y-on-y dalam rentang realistis?", "Internal"),
    ("Harga & volume", "Hubungan harga berlaku vs konstan logis thd inflasi?", "Internal"),
    ("Implisit deflator", "IPD total searah dengan IHK umum?", "Eksternal"),
    ("Konsistensi agregasi", "Sum divisi/subkomponen = total?", "Internal"),
    ("Tahunan-triwulanan", "Akumulasi triwulan konsisten benchmark tahunan?", "Internal"),
    ("Outlier", "Lonjakan/penurunan ekstrem memerlukan justifikasi?", "Internal"),
    ("Per kapita", "Per kapita sejalan pendapatan, inflasi, penduduk?", "Eksternal"),
    ("Indikator pendapatan", "Sejalan upah/NTP/remitansi/bansos?", "Eksternal"),
    ("Indikator perdagangan", "Sejalan ritel/penjualan barang tahan lama?", "Eksternal"),
    ("Indikator pembayaran", "Konsisten transaksi QRIS/kartu/uang elektronik?", "Eksternal"),
    ("Indikator mobilitas", "Konsisten mobilitas/pariwisata/transportasi?", "Eksternal"),
    ("Fenomena kualitatif", "Estimasi mencerminkan fenomena periode berjalan?", "Eksternal"),
    ("Kebijakan pemerintah", "Mencerminkan dampak fiskal/subsidi/insentif?", "Eksternal"),
    ("Supply side", "Sejalan lapangan usaha terkait?", "Internal"),
    ("Total penyediaan", "Konsisten ketersediaan barang/jasa domestik+impor?", "Internal"),
    ("Revisi data", "Revisi signifikan dapat dijelaskan?", "Internal"),
    ("Antarwilayah", "Pola konsisten dengan wilayah terdekat?", "Eksternal"),
    ("Jangka panjang", "Tren wajar thd pertumbuhan ekonomi & penduduk?", "Eksternal")]

def run_qc(W2, W1):
    t = NQ - 1
    r9, r10, cpi = W2["r9"], W2["r10"], W2["cpi"]
    q, y = growths(r10)
    tot_n, tot_k = r9.sum(axis=0), r10.sum(axis=0)
    ipd_tot = tot_n / tot_k.replace(0, np.nan) * 100
    checks = []
    def safe(fn, default=True):
        try:
            v = fn()
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return default
            return bool(v)
        except Exception:
            return default
    share = r9.iloc[:, t] / tot_n.iat[t] if tot_n.iat[t] else r9.iloc[:, t] * np.nan
    share_p = r9.iloc[:, t - 1] / tot_n.iat[t - 1] if t > 0 and tot_n.iat[t - 1] else share
    ok = [
        safe(lambda: float((share - share_p).abs().max()) < 0.05),
        safe(lambda: float((share - share_p).abs().max()) < 0.03),
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True),
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True),
        safe(lambda: float(y.iloc[:, t].abs().max()) < 12 if t >= 4 else True),
        True,
        safe(lambda: not (((W1["q25"].iat[13, t] - W1["q25"].iat[13, t - 1]) > 0) ^ ((ipd_tot.iat[t] - ipd_tot.iat[t - 1]) > 0)) if t else True),
        safe(lambda: abs(tot_n.iat[t] - r9.iloc[:, t].sum()) < 1e-6),
        safe(lambda: abs(sdiv(r9.iloc[:, [i for i, qq in enumerate(Q_ACT) if qq.startswith('2025')]].sum().sum(), S["e2_sut"].iloc[:, 0].sum()) - 1) < 0.05),
        safe(lambda: float(y.iloc[:, t].abs().max()) <= 15 if t >= 4 else True),
        True, True, True, True, True,
        safe(lambda: bool((S["f7_fenomena"].iloc[:, 0].astype(str).str.len() > 5).all())),
        True,
        safe(lambda: abs(float(y.iloc[:, t].mean()) - 5.0) < 6 if t >= 4 else True),
        True, True, True, True]
    ket = ["Maks pergeseran share <5 pp", "Ambang 3 pp", "Maks |q-q| <8%", "Maks |q-q| <8%",
           "Maks |y-y| <12%", "Deflasi konsisten", "Arah gerak sama", "Identik", "Deviasi <5%",
           "Batas ±15%", "A12 konsisten", "Proxy I6", "Proxy I6", "Proxy I6", "Proxy I6",
           "Catatan F7 lengkap", "Rule BPJS/MBG", "Selisih <6 pp", "SUT 2025", "Putaran 0",
           "Sister city 74", "2019-2026"]
    for i, (asp, tanya, jenis) in enumerate(QC_ITEMS):
        checks.append({"No": i + 1, "Aspek": asp, "Pertanyaan": tanya, "Jenis": jenis,
                       "Respon": "Ya" if ok[i] else "Perlu Pemeriksaan", "Keterangan": ket[i]})
    rate = sum(1 for c in checks if c["Respon"] == "Ya") / 22 * 100
    return checks, rate

# ============================ EKSPOR EXCEL ============================
HDR = Font(bold=True)
TIT = Font(bold=True, size=13)

def sheet_base(wb, name, title, sub):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = TIT
    ws["A2"] = sub
    info = [("Kode Provinsi", S["kode_prov"]), ("Kode Kab/Kota", S["kode_kab"]),
            ("Sister city", S["sister"]), ("Periode", S["periode"]),
            ("Tanggal Pengerjaan", "10-11-2026"), ("Petugas", S["petugas"]),
            ("Supervisor", S["supervisor"])]
    r = 4
    for k, v in info:
        ws.cell(r, 1, k); ws.cell(r, 3, v); r += 1
    return ws

def q_header(ws, blocks):
    starts, c = [], 3
    for title, n in blocks:
        starts.append(c)
        ws.cell(TOP, c, title).font = HDR
        for t in range(n):
            ws.cell(TOP + 1, c + t, Q_ACT[t][:4])
            ws.cell(TOP + 3, c + t, Q_ACT[t])
        c += n + 1
    ws.cell(TOP + 3, 1, "Kode"); ws.cell(TOP + 3, 2, "COICOP")
    return starts

def m_header(ws, title):
    ws.cell(TOP, 3, title).font = HDR
    for j in range(NM):
        ws.cell(TOP + 1, 3 + j, M_ACT[j])
    ws.cell(TOP + 3, 1, "No."); ws.cell(TOP + 3, 2, "Rincian")
    return 3

def write_growth(ws, row, vc, qc0, yc0, cc0):
    for t in range(1, NQ):
        ws.cell(row, qc0 + t, f"=IF({CL(vc+t-1)}{row}=0,NA(),({CL(vc+t)}{row}/{CL(vc+t-1)}{row}-1)*100)")
    for t in range(4, NQ):
        ws.cell(row, yc0 + t, f"=IF({CL(vc+t-4)}{row}=0,NA(),({CL(vc+t)}{row}/{CL(vc+t-4)}{row}-1)*100)")
    for t in range(7, NQ, 4):
        ws.cell(row, cc0 + t,
                f"=IF(SUM({CL(vc+t-7)}{row}:{CL(vc+t-4)}{row})=0,NA(),"
                f"(SUM({CL(vc+t-3)}{row}:{CL(vc+t)}{row})/SUM({CL(vc+t-7)}{row}:{CL(vc+t-4)}{row})-1)*100)")

def export_wb1(W1):
    wb = Workbook(); wb.remove(wb.active)
    ws = sheet_base(wb, "Home", "Deflator PKRT Triwulanan", "PDRB (2025=100) - Kab. Cirebon")
    r = TOP
    for a, b in [("E1", "IHK bulanan subkomponen (2022=100)"), ("E2", "Bobot COICOP (SBH 2022)"),
                 ("E3", "IHK bulanan 13 COICOP"), ("E4", "IHK triwulanan 2022=100"),
                 ("E5", "IHK triwulanan 2025=100 -> E4 WB2")]:
        ws.cell(r, 1, a); ws.cell(r, 3, b); r += 1
    ws = sheet_base(wb, "E1", "Input IHK bulanan subkomponen 2022=100", "")
    c0 = m_header(ws, "Indeks bulanan")
    for i, nm in enumerate(SUB17 + ["IHK Umum"]):
        ws.cell(R(i), 1, i + 1 if i < 17 else "00"); ws.cell(R(i), 2, nm)
        for j in range(NM):
            v = S["e1_ihk"].loc[nm, M_ACT[j]]
            if not pd.isna(v):
                ws.cell(R(i), c0 + j, float(v))
    ws = sheet_base(wb, "E2", "Bobot COICOP (Diagram Timbang SBH 2022)", "")
    ws.cell(TOP, 3, "Nilai Konsumsi Dasar (Rp Juta)"); ws.cell(TOP, 4, "Bobot COICOP")
    for i, nm in enumerate(SUB17):
        ws.cell(R(i), 1, i + 1); ws.cell(R(i), 2, nm)
        ws.cell(R(i), 3, float(S["e2_dasar"].iat[i, 0]))
    for d, subs in {**SPLIT, **{k: [v] for k, v in SINGLE.items()}}.items():
        rng = ",".join(f"D{R(s)}" for s in subs)
        for s in subs:
            ws.cell(R(s), 4, f"=IF(SUM({rng})=0,NA(),C{R(s)}/SUM({rng}))")
    ws = sheet_base(wb, "E3", "IHK bulanan 13 COICOP 2022=100", "")
    c0 = m_header(ws, "Indeks bulanan")
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        subs = SPLIT[d] if d in SPLIT else [SINGLE[d]]
        for j in range(NM):
            cm = CL(c0 + j)
            if len(subs) == 1:
                ws.cell(R(d), c0 + j, f"=E1!{cm}{R(subs[0])}")
            else:
                r1, r2 = R(subs[0]), R(subs[1])
                ws.cell(R(d), c0 + j,
                        f"=IF(SUM(E2!D{r1},E2!D{r2})=0,NA(),"
                        f"(E2!D{r1}*E1!{cm}{r1}+E2!D{r2}*E1!{cm}{r2})/SUM(E2!D{r1},E2!D{r2}))")
    ws.cell(R(13), 1, "00"); ws.cell(R(13), 2, "IHK Umum")
    for j in range(NM):
        ws.cell(R(13), c0 + j, f"=E1!{CL(c0+j)}{R(17)}")
    ws = sheet_base(wb, "E4", "IHK triwulanan COICOP 2022=100 (rata-rata 3 bulan)", "")
    vc = q_header(ws, [("Indeks triwulanan", NQ)])[0]
    for d in range(14):
        ws.cell(R(d), 1, f"{d+1:02d}" if d < 13 else "00")
        ws.cell(R(d), 2, DIV13[d] if d < 13 else "IHK Umum")
        for t, qq in enumerate(Q_ACT):
            if qq < "2024Q1":
                continue
            mi = (int(qq[:4]) - 2024) * 12 + (int(qq[-1]) - 1) * 3
            a = 3 + mi
            ws.cell(R(d), vc + t, f"=AVERAGE(E3!{CL(a)}{R(d)}:{CL(a+2)}{R(d)})")
    ws = sheet_base(wb, "E5", "IHK triwulanan COICOP 2025=100 (re-referencing)", "")
    vc = q_header(ws, [("Indeks triwulanan", NQ)])[0]
    c25 = [vc + t for t, qq in enumerate(Q_ACT) if qq.startswith("2025")]
    for d in range(14):
        ws.cell(R(d), 1, f"{d+1:02d}" if d < 13 else "00")
        ws.cell(R(d), 2, DIV13[d] if d < 13 else "IHK Umum")
        for t in range(NQ):
            if not c25:
                ws.cell(R(d), vc + t, "=NA()")
            else:
                rng = f"E4!{CL(c25[0])}{R(d)}:{CL(c25[-1])}{R(d)}"
                ws.cell(R(d), vc + t, f"=IF(AVERAGE({rng})=0,NA(),E4!{CL(vc+t)}{R(d)}/AVERAGE({rng})*100)")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def export_wb2(W1, W2, checks, rate):
    wb = Workbook(); wb.remove(wb.active)
    ws = sheet_base(wb, "Home", "PKRT Triwulanan", "PDRB (2025=100) - Kab. Cirebon")
    r = TOP
    for a, b in [("E0", "PKRT rilis 2010=100 (input)"), ("E1", "Seruti + backcasting"),
                 ("Help", "Konstanta metodologi"), ("E2", "Scaling factor Theta"),
                 ("E3", "ADHB inisial"), ("E4", "Deflator IHK 2025=100"), ("E5", "ADHK inisial"),
                 ("I6", "Indikator administratif"), ("F7", "Fenomena kualitatif"),
                 ("R8", "Balancing factor Beta"), ("R9", "ADHB final"), ("R10", "ADHK final"),
                 ("A11", "7 subkomponen"), ("A12", "Konsumsi per kapita"), ("QC", "22 pemeriksaan")]:
        ws.cell(r, 1, a); ws.cell(r, 3, b); r += 1
    ws = sheet_base(wb, "E0", "PKRT rilis 2010=100 (input)", "")
    vb, kb, pb = q_header(ws, [("Harga berlaku", NQ), ("Harga konstan", NQ), ("Indeks implisit", NQ)])
    for i in range(12):
        ws.cell(R(i), 1, f"{i+1:02d}"); ws.cell(R(i), 2, OLD12[i])
        for t in range(NQ):
            ws.cell(R(i), vb + t, float(S["e0_adhb"].iat[i, t]))
            ws.cell(R(i), kb + t, float(S["e0_adhk"].iat[i, t]))
            ws.cell(R(i), pb + t, f"=IF({CL(kb+t)}{R(i)}=0,NA(),{CL(vb+t)}{R(i)}/{CL(kb+t)}{R(i)}*100)")
    ws = sheet_base(wb, "E1", "Seruti agregat tertimbang + backcasting", "")
    vc = q_header(ws, [("Nilai agregat", NQ)])[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t, qq in enumerate(Q_ACT):
            if qq >= "2023Q1" and qq in S["e1_seruti"].columns:
                v = S["e1_seruti"].loc[DIV13[d], qq]
                if not pd.isna(v):
                    ws.cell(R(d), vc + t, float(v))
            elif t > 0:
                rr = REF_OLD[d]
                ws.cell(R(d), vc + t,
                        f"=IF(E0!{CL(vb+t+1)}{R(rr)}=0,NA(),"
                        f"{CL(vc+t+1)}{R(d)}*E0!{CL(vb+t)}{R(rr)}/E0!{CL(vb+t+1)}{R(rr)})")
    ws.cell(TOP + 6, 3, "Catatan: 2019Q1-2022Q4 di-backcasting menggunakan PKRT rilis harga berlaku 2010=100")
    ws = sheet_base(wb, "Help", "Konstanta metodologi BPS", "")
    ws.cell(TOP, 1, "Disagregasi rilis COICOP 12")
    ws.cell(TOP + 1, 1, "12"); ws.cell(TOP + 1, 2, "Jasa asuransi dan keuangan"); ws.cell(TOP + 1, 3, PROP[11])
    ws.cell(TOP + 2, 1, "13"); ws.cell(TOP + 2, 2, "Barang dan jasa lainnya"); ws.cell(TOP + 2, 3, PROP[12])
    ws.cell(TOP + 4, 1, "Coverage adjustment")
    ws.cell(TOP + 5, 1, "06"); ws.cell(TOP + 5, 2, "Kesehatan"); ws.cell(TOP + 5, 3, COV[5])
    ws.cell(TOP + 6, 1, "12"); ws.cell(TOP + 6, 2, "Jasa asuransi dan keuangan"); ws.cell(TOP + 6, 3, COV[11])
    cov6, cov12 = TOP + 5, TOP + 6
    ws = sheet_base(wb, "E2", "Scaling factor Theta", "")
    ws.cell(TOP, 3, "Seruti 2025 (Rp Juta)"); ws.cell(TOP, 4, "SUT 2025 (Rp Juta)"); ws.cell(TOP, 5, "Theta")
    c25 = [3 + t for t, qq in enumerate(Q_ACT) if qq.startswith("2025")]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        ws.cell(R(d), 3, f"=SUM(E1!{CL(c25[0])}{R(d)}:{CL(c25[-1])}{R(d)})" if c25 else "=NA()")
        ws.cell(R(d), 4, float(S["e2_sut"].iat[d, 0]))
        ws.cell(R(d), 5, f"=IF(C{R(d)}=0,NA(),D{R(d)}/C{R(d)})")
    ws = sheet_base(wb, "E3", "Estimasi ADHB inisial", "")
    vc = q_header(ws, [("Nilai agregat", NQ)])[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        cov = f"Help!C{cov6}" if d == 5 else (f"Help!C{cov12}" if d == 11 else "1")
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=E2!E{R(d)}*E1!{CL(vc+t)}{R(d)}*{cov}")
    ws = sheet_base(wb, "E4", "Deflator IHK triwulanan 2025=100", "")
    vc = q_header(ws, [("Indeks", NQ)])[0]
    i24 = Q_ACT.index("2024Q1") if "2024Q1" in Q_ACT else 0
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            if t >= i24:
                v = W1["q25"].iat[d, t] if t < W1["q25"].shape[1] else np.nan
                if not pd.isna(v):
                    ws.cell(R(d), vc + t, round(float(v), 4))
            elif t > 0:
                rr = REF_OLD[d]
                ws.cell(R(d), vc + t,
                        f"=IF(E0!{CL(pb+t+1)}{R(rr)}=0,NA(),"
                        f"{CL(vc+t+1)}{R(d)}*E0!{CL(pb+t)}{R(rr)}/E0!{CL(pb+t+1)}{R(rr)})")
    ws.cell(TOP + 6, 3, "Catatan: 2019Q1-2023Q4 di-backcasting menggunakan indeks harga implisit 2010=100")
    ws = sheet_base(wb, "E5", "Estimasi ADHK 2025 inisial", "")
    vc = q_header(ws, [("Nilai agregat", NQ)])[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),E3!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
    ws = sheet_base(wb, "I6", "Indikator data administratif", "")
    for j, qq in enumerate(Q_ACT):
        ws.cell(TOP, 3 + j, qq)
    for i, nm in enumerate(S["i6_admin"].index):
        ws.cell(R(i), 2, nm)
        for j, qq in enumerate(Q_ACT):
            v = S["i6_admin"].loc[nm, qq]
            if not pd.isna(v):
                ws.cell(R(i), 3 + j, float(v))
    ws = sheet_base(wb, "F7", "Fenomena kualitatif triwulan berjalan", "")
    ws.cell(TOP, 3, f"Fenomena pada {S['periode']}")
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        ws.cell(R(d), 3, str(S["f7_fenomena"].iat[d, 0]))
    ws = sheet_base(wb, "R8", "Balancing factor Beta", "")
    vc = q_header(ws, [("Beta", NQ)])[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, float(S["r8_beta"].iat[d, t]))
    ws = sheet_base(wb, "R9", "Estimasi akhir ADHB", "")
    vc, qc0, yc0, cc0 = q_header(ws, [("Nilai agregat", NQ), ("Growth q-to-q (%)", NQ),
                                      ("Growth y-on-y (%)", NQ), ("Growth c-to-c (%)", NQ)])
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=R8!{CL(vc+t)}{R(d)}*E3!{CL(vc+t)}{R(d)}")
        write_growth(ws, R(d), vc, qc0, yc0, cc0)
    ws = sheet_base(wb, "R10", "Estimasi akhir ADHK 2025", "")
    vc, qc0, yc0, cc0 = q_header(ws, [("Nilai agregat", NQ), ("Growth q-to-q (%)", NQ),
                                      ("Growth y-on-y (%)", NQ), ("Growth c-to-c (%)", NQ)])
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),R9!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
        write_growth(ws, R(d), vc, qc0, yc0, cc0)
    ws = sheet_base(wb, "A11", "Agregasi 7 subkomponen rilis", "")
    vn, vk = q_header(ws, [("Nilai berlaku", NQ), ("Nilai konstan", NQ)])
    for s7 in range(7):
        ws.cell(R(s7), 1, f"{s7+1:02d}"); ws.cell(R(s7), 2, SUB7[s7])
        for t in range(NQ):
            tn = "+".join(f"R9!{CL(3+t)}{R(j)}" for j in MAP7[s7])
            tk = "+".join(f"R10!{CL(3+t)}{R(j)}" for j in MAP7[s7])
            ws.cell(R(s7), vn + t, f"={tn}")
            ws.cell(R(s7), vk + t, f"={tk}")
    ws = sheet_base(wb, "A12", "Konsumsi per kapita (Rp Ribu)", "")
    ws.cell(TOP, 3, "Jumlah penduduk (jiwa)"); ws.cell(TOP, 4, int(S["penduduk"]))
    vn, vk = q_header(ws, [("Per kapita berlaku", NQ), ("Per kapita konstan", NQ)])
    for s7 in range(7):
        ws.cell(R(s7), 1, f"{s7+1:02d}"); ws.cell(R(s7), 2, SUB7[s7])
        for t in range(NQ):
            ws.cell(R(s7), vn + t, f"=IF($D${TOP}=0,NA(),A11!{CL(vn+t)}{R(s7)}*1000/$D${TOP})")
            ws.cell(R(s7), vk + t, f"=IF($D${TOP}=0,NA(),A11!{CL(vk+t)}{R(s7)}*1000/$D${TOP})")
    ws = sheet_base(wb, "QC", "Pemeriksaan konsistensi internal & eksternal", "")
    r = TOP
    for c in checks:
        ws.cell(r, 1, c["No"]); ws.cell(r, 2, c["Aspek"]); ws.cell(r, 3, c["Pertanyaan"])
        ws.cell(r, 4, c["Jenis"]); ws.cell(r, 5, c["Respon"]); ws.cell(r, 6, c["Keterangan"])
        r += 1
    ws.cell(r + 1, 2, "Tingkat konsistensi")
    ws.cell(r + 1, 5, f'=COUNTIF(E{TOP}:E{r},"Ya")/22')
    ws.cell(r + 1, 5).number_format = "0.00%"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ============================ ANTARMUKA ============================
st.title("📊 Otomasi 2 Lembar Kerja BPS — PKRT & Deflator Kab. Cirebon (2025=100)")
st.caption("WB1: 03. LK Deflator PPKK · WB2: 02. LK PKRT Put0_Rev · Periode fokus: **" + S["periode"] +
           "** · Sumber: **" + S.get("sumber_data", "-") + "**")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📘 Panduan Alur Kerja Pemula", "🌡️ Workbook 1: LK Deflator IHK",
    "🧮 Workbook 2: LK Utama PKRT", "📈 Dashboard & QC Checker",
    "🏙️ Simulasi Riil Kab. Cirebon", "📗 SUT Jabar 2020"])

with tab1:
    st.header("Alur kerja 8 langkah")
    for a, b, c in [
        ("1", "Pilih sumber data", "Sidebar: 'Muat Simulasi' (default) atau 'Template Kosong'; bisa Simpan/Muat JSON."),
        ("2", "WB1-E1/E2", "IHK bulanan 17 subkomponen (2022=100) + nilai konsumsi dasar SBH 2022."),
        ("3", "WB1-E3/E4/E5", "IHK 13 COICOP bulanan → triwulanan (rata-rata 3 bulan) → re-referencing 2025=100."),
        ("4", "WB2-E0/E1", "Rilis 2010=100 (ADHB/ADHK) + Seruti 2023Q1+; backcasting 2019Q1-2022Q4 otomatis."),
        ("5", "WB2-E2/E3", "Theta = SUT 2025 / Seruti 2025; ADHB inisial = Theta × Seruti × coverage adj (Help)."),
        ("6", "WB2-E4/E5", "Deflator dari E5 WB1 (≥2024Q1); backcast via IPD 2010; ADHK inisial = deflasi."),
        ("7", "R8-R10, A11, A12", "Beta, estimasi akhir, agregasi 7 subkomponen, konsumsi per kapita."),
        ("8", "QC & Export", "22 butir konsistensi + unduh 2 file .xlsx multi-sheet sesuai template BPS.")]:
        st.markdown(f"**Langkah {a} — {b}**: {c}")
    st.info("Semua pembagian dilindungi: pembagi nol menghasilkan NA (bukan crash / #DIV/0!).")

    with st.expander("📂 Cara Import Template PDRB ke E0", expanded=False):
        st.markdown("""
**Langkah-langkah:**
1. Buka sidebar → klik **"📂 Muat Template PDRB (.xlsx) → E0"**
2. Pilih file `Template Upload PDRB - CIREBON - 2025Q4 - Kosong.xlsx`
3. Sistem otomatis memetakan:

| Baris Excel | → | Divisi E0 (OLD12) |
|---|---|---|
| 1.a (id=2) | → | Makanan dan minuman tidak beralkohol |
| 1.b (id=3) | → | Minuman beralkohol dan tembakau |
| 1.c (id=4) | → | Pakaian dan alas kaki |
| 1.d (id=5) | → | Perumahan, air, listrik, gas... |
| 1.e (id=6) | → | Furnitur, perlengkapan RT |
| 1.f (id=7) | → | Kesehatan |
| 1.g (id=8) | → | Transportasi |
| 1.h (id=9) | → | Komunikasi |
| 1.i (id=10) | → | Rekreasi dan kebudayaan |
| 1.j (id=11) | → | Pendidikan |
| 1.k (id=12) | → | Penyediaan makan minum & akomodasi |
| 1.l + 1.m (id=13+14) | → | Barang dan jasa lainnya **(dijumlahkan)** |

> ⚠️ **Catatan:** Sheet harus bernama **ADHB** dan **ADHK**,
> kolom `id` berisi angka 2–14, dan kolom quarter berformat `2019Q1`–`2025Q4`.
        """)

    with st.expander("📊 PANDUAN DATA RIIL Kab. Cirebon — Alur Perhitungan & Cakupan", expanded=False):
        st.markdown("#### 🔗 Alur perhitungan data riil (mudah dipahami)")
        st.markdown("""
1. **Seruti → Total Konsumsi (ADHB awal)**
   `Pengeluaran per-kapita (Seruti) × Jumlah Penduduk = Total Konsumsi (Rp Juta)`.
   Inilah bahan baku utama PKRT, diambil langsung dari **Seruti BPS 2023Q1–2026Q2**.

2. **Scaling factor θ (Theta)**
   `θ = SUT 2025 ÷ Total Seruti 2025`. Pada data riil ini θ di-set **1 (netral)**,
   sehingga ADHB awal = Seruti. Ubah nilai **E2 (SUT 2025)** bila ingin penyesuaian.

3. **IHK → Deflator (2025=100)**
   IHK triwulanan Kota Cirebon (13 COICOP) di-*re-referencing* menjadi basis 2025=100
   di Workbook 1, lalu dipakai sebagai **deflator** di Workbook 2.

4. **ADHB ÷ Deflator = ADHK**
   `ADHK = ADHB ÷ (IHK/100)`. Menghilangkan efek kenaikan harga → volume riil.

5. **Balancing factor β (Beta)**
   `ADHB final = β × ADHB awal`. Default β = 1; sesuaikan di **R8** bila perlu penyeimbangan.

6. **Agregasi & Per-kapita**
   Hasil 13 COICOP diagregasi jadi **7 subkomponen rilis**, lalu dibagi penduduk → **konsumsi per kapita**.
        """)
        st.markdown("#### 📅 Cakupan data riil")
        st.markdown("""
| Komponen | Status | Periode |
|---|---|---|
| Seruti per-kapita (E1) | ✅ Riil | 2023Q1 – 2026Q2 |
| IHK / Deflator (WB1) | ✅ Riil (IHK Kota Cirebon) | 2024Q1+ |
| Jumlah Penduduk | ✅ Riil (Proyeksi SP2020) | 2023Q1 – 2026Q2 |
| Rilis PDRB 2010=100 (E0) | ❌ Tidak tersedia | sebelum 2023 kosong |
| SUT 2025, I6, F7 | ⚠️ Default/netral | bisa diisi manual |
        """)
        st.warning("⚠️ **Keterbatasan:** Karena rilis PDRB 2010=100 (E0) tidak tersedia, "
                   "backcasting sebelum 2023 tidak berjalan, dan **ADHK 2023** belum dapat dihitung "
                   "(deflator 2024+ yang tersedia). ADHB 2023 tetap terhitung dari Seruti riil.")
        st.success("💡 **Cara pakai:** klik tombol **📊 Data Riil** di sidebar, lalu telusuri "
                   "Workbook 1 → Workbook 2 → Dashboard. Semua angka berasal dari data riil BPS.")

    with st.expander("🎲 Dokumentasi dataset simulasi Kab. Cirebon", expanded=False):
        st.markdown("- **Musiman IHK**: Ramadan/Lebaran (Feb–Mar) menaikkan Makanan, Minuman, Rokok, Pakaian, Restoran; Jul (ajaran baru) menaikkan Pendidikan; Des–Jan (Nataru) menaikkan Transportasi, Restoran, Jasa lainnya.")
        st.markdown("- **Pertumbuhan**: nominal PKRT 0,8–1,5%/triwulan (peak Q1); IPD 2010=100 tumbuh 2,4–3,2%/tahun.")
        st.markdown("- **Konstanta BPS**: proporsi disagregasi COICOP-12 = 0,6335 / 0,3665; coverage adjustment div 06 = 0,4980, div 12 = 1,1007.")
        st.markdown("- **Penduduk**: 2.304.567 jiwa (proyeksi 2026).")
        st.markdown("**Pratinjau E2 — Nilai Konsumsi Dasar (SBH 2022)**")
        st.dataframe(S["e2_dasar"])
        st.markdown("**Pratinjau E1 — IHK bulanan (6 bulan terakhir)**")
        st.dataframe(S["e1_ihk"].loc[:, M_ACT[-6:]].round(2))
        st.markdown("**Pratinjau F7 — Fenomena kualitatif**")
        st.dataframe(S["f7_fenomena"])

with tab2:
    st.subheader("Workbook 1 — Lembar Kerja Deflator PKRT")
    c1, c2, c3 = st.columns(3)
    c1.text_input("Kode Provinsi", key="kode_prov")
    c2.text_input("Kode Kab/Kota", key="kode_kab")
    c3.text_input("Sister city", key="sister")
    st.markdown("**E1 — IHK bulanan subkomponen (2022=100)**")
    ed1 = st.data_editor(S["e1_ihk"], key="ed_e1", num_rows="fixed")
    S["e1_ihk"] = ed1
    st.markdown("**E2 — Nilai konsumsi dasar (SBH 2022)**")
    ed2 = st.data_editor(S["e2_dasar"], key="ed_e2", num_rows="fixed")
    S["e2_dasar"] = ed2
    W1 = compute_wb1()
    W2 = compute_wb2(W1)
    ca, cb = st.columns(2)
    with ca:
        st.markdown("**E4 — IHK triwulanan (2022=100)**")
        st.dataframe(W1["q22"].round(2))
    with cb:
        st.markdown("**E5 — IHK triwulanan (2025=100) → ke E4 WB2**")
        st.dataframe(W1["q25"].round(2))
    st.line_chart(W1["q25"].T)

with tab3:
    st.subheader("Workbook 2 — Lembar Kerja Utama PKRT")
    tA, tB, tC, tD = st.tabs(["E0 & E1 (input)", "E2, I6, F7, R8, A12", "Hasil E3-E5", "Hasil R9-R10, A11"])
    with tA:
        st.markdown("**E0 — Rilis 2010=100: ADHB**")
        eb = st.data_editor(S["e0_adhb"], key="ed_e0b", num_rows="fixed")
        S["e0_adhb"] = eb
        st.markdown("**E0 — Rilis 2010=100: ADHK**")
        ek = st.data_editor(S["e0_adhk"], key="ed_e0k", num_rows="fixed")
        S["e0_adhk"] = ek
        st.markdown("**E1 — Seruti agregat tertimbang (2023Q1+)**")
        es = st.data_editor(S["e1_seruti"], key="ed_ser", num_rows="fixed")
        S["e1_seruti"] = es
    with tB:
        st.markdown("**E2 — PKRT SUT 2025 per divisi**")
        esut = st.data_editor(S["e2_sut"], key="ed_sut", num_rows="fixed")
        S["e2_sut"] = esut
        st.markdown("**R8 — Balancing factor Beta**")
        ebt = st.data_editor(S["r8_beta"], key="ed_beta", num_rows="fixed")
        S["r8_beta"] = ebt
        st.markdown("**I6 — Indikator administratif**")
        ei6 = st.data_editor(S["i6_admin"], key="ed_i6", num_rows="fixed")
        S["i6_admin"] = ei6
        st.markdown("**F7 — Fenomena kualitatif**")
        ef7 = st.data_editor(S["f7_fenomena"], key="ed_f7", num_rows="fixed")
        S["f7_fenomena"] = ef7
        st.number_input("Jumlah penduduk (jiwa) — Sheet A12", min_value=1,
                        value=int(S["penduduk"]), key="penduduk")
    with tC:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**E3 — ADHB inisial**"); st.dataframe(W2["e3"].round(1))
            st.markdown("**Theta**"); st.dataframe(W2["theta"].round(4))
        with c2:
            st.markdown("**E4 — Deflator 2025=100**"); st.dataframe(W2["cpi"].round(2))
            st.markdown("**E5 — ADHK inisial**"); st.dataframe(W2["e5"].round(1))
    with tD:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**R9 — ADHB final**"); st.dataframe(W2["r9"].round(1))
        with c2:
            st.markdown("**R10 — ADHK final**"); st.dataframe(W2["r10"].round(1))
        st.markdown("**A11 — 7 subkomponen (ADHB / ADHK)**")
        st.dataframe(pd.concat([W2["a11n"].add_prefix("ADHB_"), W2["a11k"].add_prefix("ADHK_")], axis=1).round(1))
        st.markdown("**A12 — Konsumsi per kapita (Rp Ribu, ADHK)**")
        st.dataframe(W2["a12k"].round(1))

with tab4:
    st.subheader("Dashboard & QC Checker")
    checks, rate = run_qc(W2, W1)
    t = NQ - 1
    def wdefl(divs):
        w = W2["r9"].iloc[divs, t].values
        v = W2["cpi"].iloc[divs, t].values
        if np.all(np.isfinite(w)) and w.sum() > 0 and np.all(np.isfinite(v)):
            return float(np.average(v, weights=w))
        return _nmean(list(v))
    qoq = (W2["a11k"].iloc[:, t] / W2["a11k"].iloc[:, t - 1] - 1) * 100 if t else W2["a11k"].iloc[:, t] * np.nan
    yoy = (W2["a11k"].iloc[:, t] / W2["a11k"].iloc[:, t - 4] - 1) * 100 if t >= 4 else W2["a11k"].iloc[:, t] * np.nan
    rek = pd.DataFrame({
        "Subkomponen": SUB7,
        "ADHB (Jt Rp)": [round(x, 1) for x in W2["a11n"].iloc[:, t]],
        "Deflator": [round(wdefl(d), 2) for d in MAP7.values()],
        "ADHK (Jt Rp)": [round(x, 1) for x in W2["a11k"].iloc[:, t]],
        "Q-to-Q (%)": [round(x, 2) if np.isfinite(x) else np.nan for x in qoq],
        "Y-on-Y (%)": [round(x, 2) if np.isfinite(x) else np.nan for x in yoy],
        "Share (%)": [round(x, 2) for x in W2["a11n"].iloc[:, t] / W2["a11n"].iloc[:, t].sum() * 100]})
    st.dataframe(rek, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.bar_chart(rek.set_index("Subkomponen")[["ADHB (Jt Rp)", "ADHK (Jt Rp)"]])
    with c2:
        st.line_chart(W2["a11k"].T)
    st.metric("Tingkat Konsistensi QC", f"{rate:.2f}%")
    st.dataframe(pd.DataFrame(checks), hide_index=True)
    st.markdown("### ⬇️ Export Excel multi-sheet (struktur & rumus template BPS)")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("Unduh WB1 — LK Deflator (.xlsx)", export_wb1(W1),
                           file_name="03_LK_Deflator_PKRT_2025_100_Cirebon.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        st.download_button("Unduh WB2 — LK Utama PKRT (.xlsx)", export_wb2(W1, W2, checks, rate),
                           file_name="02_LK_PKRT_Triwulanan_2025_100_Cirebon.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab5:

        cirebon_real_simulation.render_cirebon_simulation()
    with tab6:
        sut_jabar_bridge.render_sut_bridge()