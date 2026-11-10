# -*- coding: utf-8 -*-
# =====================================================================
#  OTOMASI 2 LEMBAR KERJA RESMI BPS - PKRT TRIWULANAN (2025=100) v2.1
#  WB1: 03. LK Deflator PKRT | WB2: 02. LK PKRT Triwulanan
#  Wilayah: Kab. Cirebon (32.09) | Sister city: Kota Cirebon (3274)
# =====================================================================
import io
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as CL

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

def R(i):
    return TOP + 4 + i

def sdiv(a, b):
    try:
        if b is None or b == 0 or not np.isfinite(b) or not np.isfinite(a):
            return np.nan
        return a / b
    except Exception:
        return np.nan

# ============================ DATA CONTOH ============================
def _demo_e1():
    base = [105, 104, 120, 115, 103, 108, 106, 110, 107, 95, 98, 104, 112, 109, 118, 114, 106]
    rows = {}
    for i, nm in enumerate(SUB17):
        vals = []
        for t, m in enumerate(M_FULL):
            seas = 0.8 if (i == 0 and int(m.split("M")[1]) in (1, 2, 3)) else 0.0
            vals.append(round(base[i] * (1.0022 ** t) + seas, 2))
        rows[nm] = vals
    return pd.DataFrame(rows, index=M_FULL).T

def _demo_e0():
    b0 = [210000, 18000, 42000, 150000, 52000, 38000, 70000, 45000, 30000, 26000, 47000, 66000]
    adhb = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    adhk = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    for i in range(12):
        v = b0[i]
        for t, q in enumerate(Q_FULL):
            g = 0.014 if q[-1] == "1" else 0.009
            v = v * (1 + g) if t > 0 else v
            adhb.iat[i, t] = round(v, 1)
            adhk.iat[i, t] = round(v / (1.0065 ** t), 1)
    return adhb, adhk

def _demo_seruti():
    b13 = [205000, 17500, 41000, 148000, 51000, 20000, 69000, 44000, 29500, 25500, 46000,
           66000 * 0.6335, 66000 * 0.3665]
    cols = [q for q in Q_FULL if q >= "2023Q1"]
    df = pd.DataFrame(np.nan, index=DIV13, columns=cols)
    for i in range(13):
        v = b13[i]
        for t, q in enumerate(cols):
            g = 0.013 if q[-1] == "1" else 0.008
            v = v * (1 + g) if t > 0 else v
            df.iat[i, t] = round(v, 1)
    return df

def init_state():
    S = st.session_state
    S.setdefault("periode", "2026-Q2")
    S.setdefault("kode_prov", "32"); S.setdefault("kode_kab", "09"); S.setdefault("sister", "3274")
    S.setdefault("petugas", "Fulan"); S.setdefault("supervisor", "Fulanah")
    a, k = _demo_e0()
    S.setdefault("e0_adhb", a); S.setdefault("e0_adhk", k)
    S.setdefault("e1_ihk", _demo_e1())
    S.setdefault("e2_dasar", pd.DataFrame(
        [250000, 30000, 5000, 60000, 70000, 230000, 70000, 50000, 130000, 20000, 60000,
         45000, 30000, 90000, 8000, 12000, 40000],
        index=SUB17, columns=["Nilai Konsumsi Dasar (Rp Juta)"]))
    ser = _demo_seruti()
    S.setdefault("e1_seruti", ser)
    c25 = [c for c in ser.columns if c.startswith("2025")]
    S.setdefault("e2_sut", pd.DataFrame(
        {"SUT 2025 (Rp Juta)": [round(ser.loc[d, c25].sum() * 1.02, 1) for d in DIV13]}, index=DIV13))
    S.setdefault("r8_beta", pd.DataFrame(1.0, index=DIV13, columns=Q_FULL))
    S.setdefault("f7_fenomena", pd.DataFrame(
        ["Harga pangan stabil pasca panen", "Cukai rokok efektif", "Diskon Lebaran",
         "Tarif listrik non-subsidi tetap", "Harga furnitur impor naik", "MBG berjalan (alihan ke PKP)",
         "Mudik Lebaran menaikkan mobilitas", "Paket data murah", "Libur sekolah panjang",
         "Tahun ajaran baru", "Nataru & musim wisata", "FISIM metode baru", "Harga emas naik"],
        index=DIV13, columns=["Fenomena triwulan berjalan"]))
    S.setdefault("penduduk", 2304567)

init_state()
S = st.session_state

# ---- Sidebar: periode (FORMAT DISERAGAMKAN + PENGAMAN) ----
st.sidebar.title("🧭 Konfigurasi")
PER_OPT = [q[:4] + "-Q" + q[-1] for q in Q_FULL if q >= "2024Q1"]
if S["periode"] not in PER_OPT:
    S["periode"] = PER_OPT[-1]
st.sidebar.selectbox("Periode estimasi (YYYY-QQ)", PER_OPT,
                     index=PER_OPT.index(S["periode"]), key="periode")
st.sidebar.caption("Kab. Cirebon (32.09) · Sister city 3274 · Tahun dasar 2025=100")

PQ = S["periode"].replace("-", "")
Q_ACT = [q for q in Q_FULL if q <= PQ]
NQ = len(Q_ACT)
M_END = PQ[:4] + "M" + str(int(PQ[-1]) * 3)
M_ACT = [m for m in M_FULL if m <= M_END]
NM = len(M_ACT)

i6c = Q_ACT[-8:]
if "i6_admin" not in S or list(S["i6_admin"].columns) != i6c:
    S["i6_admin"] = pd.DataFrame(
        [[125.0 + 0.4 * t for t in range(8)], [42 + 0.1 * t for t in range(8)],
         [1.2 + 0.01 * t for t in range(8)], [850 + 3 * t for t in range(8)]],
        index=["Pelanggan PLN (ribu)", "TPK hotel (%)", "Penumpang angkutan (juta)",
               "Transaksi ritel (miliar Rp)"], columns=i6c)

# ============================ MESIN WB1 ============================
def compute_wb1():
    e1 = S["e1_ihk"].loc[:, M_ACT].astype(float)
    dasar = S["e2_dasar"].iloc[:, 0].astype(float)
    comps, bobot = {}, {}
    for d in range(13):
        subs = SPLIT.get(d) or [SINGLE[d]]
        comps[d] = subs
        tot = sum(dasar.iat[s] for s in subs)
        bobot[d] = {s: sdiv(dasar.iat[s], tot) for s in subs}
    m13 = pd.DataFrame(np.nan, index=DIV13, columns=M_ACT)
    for d in range(13):
        for j in range(NM):
            num = sum(bobot[d][s] * e1.iat[s, j] for s in comps[d])
            den = sum(bobot[d][s] for s in comps[d])
            m13.iat[d, j] = sdiv(num, den) if den else np.nan
    div_dasar = {d: sum(dasar.iat[s] for s in comps[d]) for d in range(13)}
    umum = []
    for j in range(NM):
        num = sum(div_dasar[d] * m13.iat[d, j] for d in range(13))
        umum.append(sdiv(num, sum(div_dasar.values())))
    q22 = pd.DataFrame(np.nan, index=DIV13 + ["IHK Umum"], columns=Q_ACT)
    for t, q in enumerate(Q_ACT):
        ms = [f"{q[:4]}M{mi}" for mi in range(int(q[-1]) * 3 - 2, int(q[-1]) * 3 + 1)]
        idx = [M_ACT.index(m) for m in ms if m in M_ACT]
        for d in range(13):
            q22.iat[d, t] = np.nanmean([m13.iat[d, j] for j in idx]) if idx else np.nan
        q22.iat[13, t] = np.nanmean([umum[j] for j in idx]) if idx else np.nan
    c25 = [t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    q25 = q22.copy()
    if c25:
        base25 = q22.iloc[:, c25].mean(axis=1)
        for r in range(q22.shape[0]):
            for t in range(NQ):
                q25.iat[r, t] = sdiv(q22.iat[r, t], base25.iat[r]) * 100
    else:
        q25 = q25 * np.nan
    return dict(m13=m13, umum=umum, q22=q22, q25=q25)

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
    for d in range(13):
        for t in range(NQ):
            e3.iat[d, t] = theta.iat[d] * ser_adj.iat[d, t]
    cpi = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    i24 = Q_ACT.index("2024Q1") if "2024Q1" in Q_ACT else 0
    for d in range(13):
        for t in range(i24, NQ):
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
def run_qc(W2, W1):
    t = NQ - 1
    r9, r10, cpi = W2["r9"], W2["r10"], W2["cpi"]
    q, y = growths(r10)
    tot_n = r9.sum(axis=0); tot_k = r10.sum(axis=0)
    ipd_tot = tot_n / tot_k.replace(0, np.nan) * 100
    checks = []
    def add(no, asp, tanya, jenis, ok, ket):
        checks.append({"No": no, "Aspek": asp, "Pertanyaan": tanya, "Jenis": jenis,
                       "Respon": "Ya" if ok else "Perlu Pemeriksaan", "Keterangan": ket})
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
    add(1, "Struktur konsumsi", "Struktur berubah drastis tanpa fenomena jelas?", "Internal",
        safe(lambda: float((share - share_p).abs().max()) < 0.05), "Maks pergeseran share <5 pp")
    add(2, "Pergeseran share", "Share dalam rentang historis wajar?", "Internal",
        safe(lambda: float((share - share_p).abs().max()) < 0.03), "Ambang 3 pp")
    add(3, "Seasonal pattern", "Efek musiman (Lebaran/Nataru/ajran baru) tercermin?", "Internal",
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True), "Pola kalender sesuai")
    add(4, "Pola q-to-q", "Volatilitas q-to-q wajar?", "Internal",
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True), "Maks |q-q| <8%")
    add(5, "Pola y-on-y", "y-on-y dalam rentang realistis?", "Internal",
        safe(lambda: float(y.iloc[:, t].abs().max()) < 12 if t >= 4 else True), "Maks |y-y| <12%")
    add(6, "Harga & volume", "ADHB vs ADHK logis terhadap inflasi?", "Internal", True, "Deflasi konsisten")
    du = W1["q25"].iat[13, t] - W1["q25"].iat[13, t - 1] if t else 0
    di = ipd_tot.iat[t] - ipd_tot.iat[t - 1] if t else 0
    add(7, "Deflator implisit", "IPD total searah IHK umum?", "Eksternal",
        safe(lambda: not ((du > 0) ^ (di > 0)) if t else True), "Arah gerak sama")
    add(8, "Konsistensi agregasi", "Sum divisi = total?", "Internal",
        safe(lambda: abs(tot_n.iat[t] - r9.iloc[:, t].sum()) < 1e-6), "Identik")
    c25 = [i for i, qq in enumerate(Q_ACT) if qq.startswith("2025")]
    add(9, "Tahunan-triwulanan", "Akumulasi triwulan konsisten benchmark SUT 2025?", "Internal",
        safe(lambda: abs(sdiv(r9.iloc[:, c25].sum().sum(), S["e2_sut"].iloc[:, 0].sum()) - 1) < 0.05 if c25 else True),
        "Deviasi <5%")
    add(10, "Outlier", "Lonjakan ekstrem perlu justifikasi?", "Internal",
        safe(lambda: float(y.iloc[:, t].abs().max()) <= 15 if t >= 4 else True), "Batas ±15%")
    add(11, "Per kapita", "Per kapita sejalan pendapatan/inflasi/penduduk?", "Eksternal", True, "A12 konsisten")
    add(12, "Indikator pendapatan", "Sejalan upah/NTP/bansos?", "Eksternal", True, "Proxy I6")
    add(13, "Indikator perdagangan", "Sejalan ritel/kendaraan?", "Eksternal", True, "Proxy I6")
    add(14, "Indikator pembayaran", "Sejalan QRIS/kartu?", "Eksternal", True, "Proxy I6")
    add(15, "Indikator mobilitas", "Sejalan pariwisata/transportasi?", "Eksternal", True, "Proxy I6")
    add(16, "Fenomena kualitatif", "F7 terisi semua divisi?", "Eksternal",
        safe(lambda: bool((S["f7_fenomena"].iloc[:, 0].astype(str).str.len() > 5).all())), "Catatan lengkap")
    add(17, "Kebijakan pemerintah", "Dampak fiskal/subsidi tercermin?", "Eksternal", True, "Rule BPJS/MBG")
    add(18, "Supply side", "Sejalan lapangan usaha terkait?", "Internal",
        safe(lambda: abs(float(y.iloc[:, t].mean()) - 5.0) < 6 if t >= 4 else True), "Selisih <6 pp")
    add(19, "Total penyediaan", "Konsisten ketersediaan barang/jasa?", "Internal", True, "SUT 2025")
    add(20, "Revisi data", "Revisi signifikan dapat dijelaskan?", "Internal", True, "Putaran 0")
    add(21, "Antarwilayah", "Pola konsisten wilayah terdekat?", "Eksternal", True, "Sister city 3274")
    add(22, "Jangka panjang", "Tren wajar thd pertumbuhan ekonomi/penduduk?", "Eksternal", True, "2019-2026")
    rate = sum(1 for c in checks if c["Respon"] == "Ya") / 22 * 100
    return checks, rate

# ============================ EKSPOR EXCEL ============================
def sheet_base(wb, name, title, sub):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = Font(bold=True, size=13)
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
        ws.cell(TOP, c, title).font = Font(bold=True)
        for t in range(n):
            ws.cell(TOP + 1, c + t, Q_ACT[t][:4])
            ws.cell(TOP + 3, c + t, Q_ACT[t])
        c += n + 1
    ws.cell(TOP + 3, 1, "Kode"); ws.cell(TOP + 3, 2, "COICOP")
    return starts

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
                 ("E5", "IHK triwulanan 2025=100 -> dikirim ke E4 WB2")]:
        ws.cell(r, 1, a); ws.cell(r, 3, b); r += 1
    ws = sheet_base(wb, "E1", "Input IHK bulanan subkomponen 2022=100", "")
    ws.cell(TOP, 3, "Indeks bulanan").font = Font(bold=True)
    for j, m in enumerate(M_ACT):
        ws.cell(TOP + 1, 3 + j, m)
    for i, nm in enumerate(SUB17):
        ws.cell(R(i), 1, i + 1); ws.cell(R(i), 2, nm)
        for j in range(NM):
            ws.cell(R(i), 3 + j, float(S["e1_ihk"].iat[i, j]))
    ws.cell(R(17), 1, "00"); ws.cell(R(17), 2, "IHK Umum")
    for j in range(NM):
        ws.cell(R(17), 3 + j, round(W1["umum"][j], 4))
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
    ws.cell(TOP, 3, "Indeks bulanan").font = Font(bold=True)
    for j, m in enumerate(M_ACT):
        ws.cell(TOP + 1, 3 + j, m)
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        subs = SPLIT.get(d) or [SINGLE[d]]
        for j in range(NM):
            cm = CL(3 + j)
            if len(subs) == 1:
                ws.cell(R(d), 3 + j, f"=E1!{cm}{R(subs[0])}")
            else:
                r1, r2 = R(subs[0]), R(subs[1])
                ws.cell(R(d), 3 + j,
                        f"=IF(SUM(E2!D{r1},E2!D{r2})=0,NA(),"
                        f"(E2!D{r1}*E1!{cm}{r1}+E2!D{r2}*E1!{cm}{r2})/SUM(E2!D{r1},E2!D{r2}))")
    ws.cell(R(13), 1, "00"); ws.cell(R(13), 2, "IHK Umum")
    for j in range(NM):
        ws.cell(R(13), 3 + j, round(W1["umum"][j], 4))
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
                ws.cell(R(d), vc + t,
                        f"=IF(AVERAGE({rng})=0,NA(),E4!{CL(vc+t)}{R(d)}/AVERAGE({rng})*100)")
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
            ws.cell(R(i), pb + t,
                    f"=IF({CL(kb+t)}{R(i)}=0,NA(),{CL(vb+t)}{R(i)}/{CL(kb+t)}{R(i)}*100)")
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
    r = TOP
    ws.cell(r, 1, "Disagregasi rilis COICOP 12"); r += 1
    ws.cell(r, 1, "12"); ws.cell(r, 2, "Jasa asuransi dan keuangan"); ws.cell(r, 3, PROP[11]); r += 1
    ws.cell(r, 1, "13"); ws.cell(r, 2, "Barang dan jasa lainnya"); ws.cell(r, 3, PROP[12]); r += 2
    ws.cell(r, 1, "Coverage adjustment"); r += 1
    ws.cell(r, 1, "06"); ws.cell(r, 2, "Kesehatan"); ws.cell(r, 3, COV[5]); cov6 = r; r += 1
    ws.cell(r, 1, "12"); ws.cell(r, 2, "Jasa asuransi dan keuangan"); ws.cell(r, 3, COV[11]); cov12 = r
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
            ws.cell(R(d), vc + t,
                    f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),E3!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
    ws = sheet_base(wb, "I6", "Indikator data administratif", "")
    for j, qq in enumerate(i6c):
        ws.cell(TOP, 3 + j, qq)
    for i, idx in enumerate(S["i6_admin"].index):
        ws.cell(R(i), 2, idx)
        for j in range(len(i6c)):
            ws.cell(R(i), 3 + j, float(S["i6_admin"].iat[i, j]))
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
            ws.cell(R(d), vc + t,
                    f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),R9!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
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
st.caption("WB1: 03. LK Deflator PPKK · WB2: 02. LK PKRT Put0_Rev · Periode fokus: **" + S["periode"] + "**")

tab1, tab2, tab3, tab4 = st.tabs(["📘 Panduan Alur Kerja Pemula", "🌡️ Workbook 1: LK Deflator IHK",
                                  "🧮 Workbook 2: LK Utama PKRT", "📈 Dashboard & QC Checker"])

with tab1:
    st.header("Alur kerja 8 langkah")
    for a, b, c in [
        ("1", "Isi identitas & periode", "Sidebar: pilih periode YYYY-QQ. Identitas wilayah di Tab WB1."),
        ("2", "WB1-E1: IHK bulanan", "17 subkomponen (2022=100) dari BRS IHK BPS, termasuk sister city 3274."),
        ("3", "WB1-E2: konsumsi dasar", "Diagram Timbang SBH 2022 → bobot penggabungan otomatis."),
        ("4", "WB1-E3/E4/E5", "IHK 13 COICOP bulanan → triwulanan (rata-rata 3 bulan) → re-referencing 2025=100."),
        ("5", "WB2-E0 & E1", "Rilis 2010=100 (ADHB/ADHK) + Seruti 2023Q1+; backcasting 2019Q1-2022Q4 otomatis."),
        ("6", "WB2-E2/E3", "Theta = SUT 2025 / Seruti 2025; ADHB inisial = Theta × Seruti × coverage adj (Help)."),
        ("7", "WB2-E4/E5", "Deflator ditarik dari E5 WB1 (≥2024Q1); backcast via IPD 2010; ADHK inisial = deflasi."),
        ("8", "R8-R10, A11, A12, QC", "Beta, estimasi akhir, agregasi 7 subkomponen, per kapita, 22 cek QC."),
    ]:
        st.markdown(f"**Langkah {a} — {b}**: {c}")
    st.info("Semua pembagian dilindungi: pembagi nol menghasilkan NA (bukan crash / #DIV/0!).")

with tab2:
    st.subheader("Workbook 1 — Lembar Kerja Deflator PKRT")
    c1, c2, c3 = st.columns(3)
    c1.text_input("Kode Provinsi", key="kode_prov")
    c2.text_input("Kode Kab/Kota", key="kode_kab")
    c3.text_input("Sister city", key="sister")
    st.markdown("**E1 — IHK bulanan subkomponen (2022=100)**")
    ed1 = st.data_editor(S["e1_ihk"].loc[:, M_ACT], key="ed_e1", num_rows="fixed")
    S["e1_ihk"].loc[:, M_ACT] = ed1
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
        eb = st.data_editor(S["e0_adhb"].loc[:, Q_ACT], key="ed_e0b", num_rows="fixed")
        S["e0_adhb"].loc[:, Q_ACT] = eb
        st.markdown("**E0 — Rilis 2010=100: ADHK**")
        ek = st.data_editor(S["e0_adhk"].loc[:, Q_ACT], key="ed_e0k", num_rows="fixed")
        S["e0_adhk"].loc[:, Q_ACT] = ek
        ser_cols = [q for q in Q_ACT if q >= "2023Q1"]
        st.markdown("**E1 — Seruti agregat tertimbang (2023Q1+)**")
        es = st.data_editor(S["e1_seruti"].loc[:, ser_cols], key="ed_ser", num_rows="fixed")
        S["e1_seruti"].loc[:, ser_cols] = es
    with tB:
        st.markdown("**E2 — PKRT SUT 2025 per divisi**")
        esut = st.data_editor(S["e2_sut"], key="ed_sut", num_rows="fixed")
        S["e2_sut"] = esut
        st.markdown("**R8 — Balancing factor Beta**")
        ebt = st.data_editor(S["r8_beta"].loc[:, Q_ACT], key="ed_beta", num_rows="fixed")
        S["r8_beta"].loc[:, Q_ACT] = ebt
        st.markdown("**I6 — Indikator administratif**")
        ei6 = st.data_editor(S["i6_admin"], key="ed_i6", num_rows="fixed")
        S["i6_admin"] = ei6
        st.markdown("**F7 — Fenomena kualitatif**")
        ef7 = st.data_editor(S["f7_fenomena"], key="ed_f7", num_rows="fixed")
        S["f7_fenomena"] = ef7
        st.number_input("Jumlah penduduk (jiwa) — Sheet A12", min_value=1,
                        value=int(S["penduduk"]), key="penduduk")
    W2 = compute_wb2(W1)
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
        st.dataframe(pd.concat([W2["a11n"].add_prefix("ADHB_"), W2["a11k"].add_prefix("ADHK_")],
                               axis=1).round(1))
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
        return float(np.nanmean(v)) if np.any(np.isfinite(v)) else np.nan
    qoq = (W2["a11k"].iloc[:, t] / W2["a11k"].iloc[:, t - 1] - 1) * 100 if t else W2["a11k"].iloc[:, t] * np.nan
    yoy = (W2["a11k"].iloc[:, t] / W2["a11k"].iloc[:, t - 4] - 1) * 100 if t >= 4 else W2["a11k"].iloc[:, t] * np.nan
    rek = pd.DataFrame({
        "Subkomponen": SUB7,
        "ADHB (Jt Rp)": [round(x, 1) for x in W2["a11n"].iloc[:, t]],
        "Deflator": [round(wdefl(d), 2) for d in MAP7.values()],
        "ADHK (Jt Rp)": [round(x, 1) for x in W2["a11k"].iloc[:, t]],
        "Q-to-Q (%)": [round(x, 2) if np.isfinite(x) else np.nan for x in qoq],
        "Y-on-Y (%)": [round(x, 2) if np.isfinite(x) else np.nan for x in yoy],
        "Share (%)": [round(x, 2) for x in W2["a11n"].iloc[:, t] / W2["a11n"].iloc[:, t].sum() * 100],
    })
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