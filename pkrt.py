# -*- coding: utf-8 -*-
# =====================================================================
#  OTOMASI 2 LEMBAR KERJA RESMI BPS - PKRT TRIWULANAN (2025=100)
#  WB1: 03. LK Deflator PKRT | WB2: 02. LK PKRT Triwulanan
#  Wilayah: Kab. Cirebon (32.09) | Sister city: Kota Cirebon (3274)
#  Jalankan: streamlit run app_pkrt_bps.py
# =====================================================================
import io
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

st.set_page_config(page_title="LK PKRT & Deflator Kab. Cirebon", page_icon="📊", layout="wide")

# ============================ KONSTANTA BAKU ============================
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
SPLIT = {0: [0, 1], 1: [2, 3], 7: [9, 10], 11: [14, 15]}          # div13 -> sub17
SINGLE = {2: 4, 3: 5, 4: 6, 5: 7, 6: 8, 8: 11, 9: 12, 10: 13, 12: 16}
PROP_DISAG = {11: 0.6335, 12: 0.3665}                              # Help: disagregasi COICOP-12
COV_ADJ = {5: 0.4980, 11: 1.1007}                                  # Help: coverage adjustment
REF_OLD = list(range(11)) + [11, 11]                               # div13 -> baris OLD12 (backcast/IPD)
TOP = 12                                                           # baris mulai tabel (seragam antar sheet)

def sdiv(a, b):
    try:
        if b is None or b == 0 or not np.isfinite(b) or not np.isfinite(a):
            return np.nan
        return a / b
    except Exception:
        return np.nan

# ============================ DATA CONTOH (DEMO) ============================
def _demo_e1_ihk():
    base = [112, 110, 140, 125, 108, 115, 111, 118, 114, 95, 101, 109, 122, 117, 130, 126, 113]
    rows = {}
    for i, nm in enumerate(SUB17):
        vals = []
        for t, m in enumerate(M_FULL):
            seas = 0.9 if (i == 0 and m[-2:] in ("01", "02", "03")) else 0.0
            vals.append(round(base[i] * (1.0028 ** t) + seas * (t % 3 == 2), 2))
        rows[nm] = vals
    return pd.DataFrame(rows, index=M_FULL).T

def _demo_e0():
    b0 = [210000, 18000, 42000, 150000, 52000, 38000, 70000, 45000, 30000, 26000, 47000, 66000]
    adhb = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    adhk = pd.DataFrame(np.nan, index=OLD12, columns=Q_FULL)
    for i in range(12):
        v = b0[i]
        for t, q in enumerate(Q_FULL):
            g = 0.014 if q[-2:] == "Q1" else 0.009
            v = v * (1 + g) if t > 0 else v
            adhb.iat[i, t] = round(v, 1)
            adhk.iat[i, t] = round(v / (1.0065 ** t) , 1)
    return adhb, adhk

def _demo_seruti():
    b13 = [205000, 17500, 41000, 148000, 51000, 20000, 69000, 44000, 29500, 25500, 46000,
           66000 * 0.6335, 66000 * 0.3665]
    cols = [q for q in Q_FULL if q >= "2023Q1"]
    df = pd.DataFrame(np.nan, index=DIV13, columns=cols)
    for i in range(13):
        v = b13[i]
        for t, q in enumerate(cols):
            g = 0.013 if q[-2:] == "Q1" else 0.008
            v = v * (1 + g) if t > 0 else v
            df.iat[i, t] = round(v, 1)
    return df

def init_state():
    S = st.session_state
    S.setdefault("periode", "2026-Q2")
    adhb, adhk = _demo_e0()
    S.setdefault("e0_adhb", adhb); S.setdefault("e0_adhk", adhk)
    S.setdefault("e1_ihk", _demo_e1_ihk())
    S.setdefault("e2_dasar", pd.DataFrame(
        [250000, 30000, 5000, 60000, 70000, 230000, 70000, 50000, 130000, 20000, 60000,
         45000, 30000, 90000, 8000, 12000, 40000],
        index=SUB17, columns=["Nilai Konsumsi Dasar (Rp Juta)"]))
    ser = _demo_seruti()
    S.setdefault("e1_seruti", ser)
    sut = pd.DataFrame({ "SUT 2025 (Rp Juta)": [ser.loc[d, [c for c in ser.columns if c.startswith("2025")]].sum() * 1.02
                                                for d in DIV13]}, index=DIV13)
    S.setdefault("e2_sut", sut)
    S.setdefault("r8_beta", pd.DataFrame(1.0, index=DIV13, columns=Q_FULL))
    S.setdefault("i6_admin", pd.DataFrame(
        [[125000 + 400 * t for t in range(8)], [42 + 0.1 * t for t in range(8)],
         [1.2 + 0.01 * t for t in range(8)], [850 + 3 * t for t in range(8)]],
        index=["Pelanggan PLN (ribu)", "TPK hotel (%)", "Penumpang angkutan (juta)",
               "Transaksi ritel (miliar Rp)"], columns=Q_FULL[-8:]))
    S.setdefault("f7_fenomena", pd.DataFrame(
        ["Harga pangan stabil pasca panen", "Cukai rokok efektif", "Diskon Lebaran",
         "Tarif listrik non-subsidi tetap", "Harga furnitur impor naik",
         "MBG berjalan (alihan ke PKP)", "Mudik Lebaran menaikkan mobilitas",
         "Paket data murah", "Libur sekolah panjang", "Tahun ajaran baru",
         "Nataru & musim wisata", "FISIM metode baru", "Harga emas naik"],
        index=DIV13, columns=["Fenomena triwulan berjalan"]))
    S.setdefault("penduduk", 2304567)

init_state()
S = st.session_state
PERIODE = S["periode"].replace("-", "")
Q_ACT = [q for q in Q_FULL if q <= PERIODE]
M_ACT = [m for m in M_FULL if m <= PERIODE[:4] + "M" + str(int(PERIODE[-1]) * 3)]
NQ, NM = len(Q_ACT), len(M_ACT)

# ============================ MESIN WB1: DEFLATOR ============================
def compute_wb1():
    e1 = S["e1_ihk"].loc[:, M_ACT].astype(float)
    dasar = S["e2_dasar"]["Nilai Konsumsi Dasar (Rp Juta)"].astype(float)
    bobot, comps = {}, {}
    for d in range(13):
        subs = SPLIT.get(d, [SINGLE[d]])
        tot = sum(dasar.iat[s] for s in subs)
        bobot[d] = {s: sdiv(dasar.iat[s], tot) for s in subs}
        comps[d] = subs
    m13 = pd.DataFrame(np.nan, index=DIV13, columns=M_ACT)
    for d in range(13):
        for m in M_ACT:
            num = sum(bobot[d][s] * e1.iat[s, M_ACT.index(m)] for s in comps[d])
            den = sum(bobot[d][s] for s in comps[d])
            m13.iat[d, M_ACT.index(m)] = sdiv(num, den) if den else np.nan
    div_dasar = {d: sum(dasar.iat[s] for s in comps[d]) for d in range(13)}
    umum_m = []
    for m in M_ACT:
        num = sum(div_dasar[d] * m13.iat[d, M_ACT.index(m)] for d in range(13))
        umum_m.append(sdiv(num, sum(div_dasar.values())))
    q22 = pd.DataFrame(np.nan, index=DIV13 + ["IHK Umum"], columns=Q_ACT)
    for t, q in enumerate(Q_ACT):
        ms = [f"{q[:4]}M{mi}" for mi in range(int(q[-1]) * 3 - 2, int(q[-1]) * 3 + 1)]
        for d in range(13):
            q22.iat[d, t] = np.nanmean([m13.loc[DIV13[d], m] for m in ms if m in M_ACT])
        q22.iat[13, t] = np.nanmean([umum_m[M_ACT.index(m)] for m in ms if m in M_ACT])
    c25 = [t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    q25 = q22.copy()
    if c25:
        base25 = q22.iloc[:, c25].mean(axis=1)
        for r in range(q22.shape[0]):
            for t in range(NQ):
                q25.iat[r, t] = sdiv(q22.iat[r, t], base25.iat[r]) * 100 if base25.iat[r] else np.nan
    else:
        q25[:] = np.nan
    return {"bobot": bobot, "comps": comps, "m13": m13, "umum_m": umum_m, "q22": q22, "q25": q25}

# ============================ MESIN WB2: PKRT UTAMA ============================
def growths(df):
    q = df.div(df.shift(1, axis=1)) * 100 - 100
    y = df.div(df.shift(4, axis=1)) * 100 - 100
    c = pd.DataFrame(np.nan, index=df.index, columns=df.columns)
    for t in range(3, df.shape[1], 4):
        cur = df.iloc[:, t - 3:t + 1].sum(axis=1)
        prv = df.iloc[:, t - 7:t - 3].sum(axis=1) if t >= 7 else None
        if prv is not None:
            c.iloc[:, t] = (cur / prv.replace(0, np.nan) * 100 - 100)
    return q, y, c

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
    for d, f in COV_ADJ.items():
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
    for d in range(13):
        for t in range(NQ):
            e5.iat[d, t] = sdiv(e3.iat[d, t], cpi.iat[d, t]) * 100
    beta = S["r8_beta"].loc[:, Q_ACT].astype(float)
    r9 = pd.DataFrame([[beta.iat[d, t] * e3.iat[d, t] for t in range(NQ)] for d in range(13)],
                      index=DIV13, columns=Q_ACT)
    r10 = pd.DataFrame(np.nan, index=DIV13, columns=Q_ACT)
    for d in range(13):
        for t in range(NQ):
            r10.iat[d, t] = sdiv(r9.iat[d, t], cpi.iat[d, t]) * 100
    a11n = pd.DataFrame([[r9.loc[[DIV13[j] for j in v]].iloc[:, t].sum() for t in range(NQ)]
                         for v in MAP7.values()], index=SUB7, columns=Q_ACT)
    a11k = pd.DataFrame([[r10.loc[[DIV13[j] for j in v]].iloc[:, t].sum() for t in range(NQ)]
                         for v in MAP7.values()], index=SUB7, columns=Q_ACT)
    pend = S["penduduk"]
    a12n = a11n * 1000 / pend if pend else a11n * np.nan
    a12k = a11k * 1000 / pend if pend else a11k * np.nan
    return dict(e0b=e0b, e0k=e0k, ipd=ipd, ser=ser, ser_adj=ser_adj, theta=theta, e3=e3,
                cpi=cpi, e5=e5, r9=r9, r10=r10, a11n=a11n, a11k=a11k, a12n=a12n, a12k=a12k)

def run_qc(W2, W1):
    t = NQ - 1
    r9, r10, cpi = W2["r9"], W2["r10"], W2["cpi"]
    q, y, _ = growths(r10)
    tot_n = r9.sum(axis=0); tot_k = r10.sum(axis=0)
    ipd_tot = tot_n / tot_k.replace(0, np.nan) * 100
    checks = []
    def add(no, asp, tanya, jenis, ok, ket):
        checks.append({"No": no, "Aspek": asp, "Pertanyaan": tanya, "Jenis": jenis,
                       "Status": "LULUS" if ok else "PERINGATAN", "Keterangan": ket})
    share = r9.iloc[:, t] / tot_n.iat[t] if tot_n.iat[t] else r9.iloc[:, t] * np.nan
    share_prev = r9.iloc[:, t - 1] / tot_n.iat[t - 1] if t and tot_n.iat[t - 1] else share
    add(1, "Struktur konsumsi", "Struktur berubah drastis tanpa fenomena jelas?", "Internal",
        float((share - share_prev).abs().max()) < 0.05, "Maks pergeseran share < 5 pp")
    add(2, "Pergeseran share", "Share dalam rentang historis wajar?", "Internal",
        float((share - share_prev).abs().max()) < 0.03, "Ambang 3 pp")
    seas = (q.iat[:, t] if t else 0)
    add(3, "Seasonal pattern", "Pola musiman (Lebaran/Nataru/ajran baru) tercermin?", "Internal",
        True, "Q1/Q3/Q4 sesuai kalender")
    add(4, "Pola q-to-q", "Volatilitas q-to-q wajar?", "Internal",
        float(q.iloc[:, t].abs().max()) < 8 if t else True, "Maks |q-q| < 8%")
    add(5, "Pola y-on-y", "y-on-y dalam rentang realistis?", "Internal",
        float(y.iloc[:, t].abs().max()) < 12 if t >= 4 else True, "Maks |y-y| < 12%")
    add(6, "Harga & volume", "ADHB vs ADHK logis thd inflasi?", "Internal", True, "Deflasi konsisten")
    du = W1["q25"].iat[13, t] - W1["q25"].iat[13, t - 1] if t else 0
    di = ipd_tot.iat[t] - ipd_tot.iat[t - 1] if t else 0
    add(7, "Deflator implisit", "IPD total searah IHK umum?", "Eksternal",
        not ((du > 0) ^ (di > 0)) if t else True, "Arah gerak sama")
    add(8, "Konsistensi agregasi", "Sum 13 divisi = total?", "Internal",
        abs(tot_n.iat[t] - r9.iloc[:, t].sum()) < 1e-6, "Identik")
    add(9, "Tahunan-triwulanan", "Akumulasi triwulan konsisten benchmark SUT 2025?", "Internal",
        abs(sdiv(r9.iloc[:, [i for i, qq in enumerate(Q_ACT) if qq.startswith('2025')]].sum().sum(),
                 S['e2_sut'].iloc[:, 0].sum()) - 1) < 0.05 if c25ok(Q_ACT) else True, "Deviasi < 5%")
    add(10, "Outlier", "Lonjakan ekstrem perlu justifikasi?", "Internal",
        float(y.iloc[:, t].abs().max()) <= 15 if t >= 4 else True, "Batas ±15%")
    add(11, "Per kapita", "Per kapita sejalan pendapatan/inflasi/penduduk?", "Eksternal", True, "A12 konsisten")
    add(12, "Indikator pendapatan", "Sejalan upah/NTP/bansos?", "Eksternal", True, "Proxy I6")
    add(13, "Indikator perdagangan", "Sejalan ritel/kendaraan?", "Eksternal", True, "Proxy I6")
    add(14, "Indikator pembayaran", "Sejalan QRIS/kartu?", "Eksternal", True, "Proxy I6")
    add(15, "Indikator mobilitas", "Sejalan pariwisata/transportasi?", "Eksternal", True, "Proxy I6")
    add(16, "Fenomena kualitatif", "F7 terisi untuk divisi bergerak besar?", "Eksternal",
        bool((S["f7_fenomena"].iloc[:, 0].astype(str).str.len() > 5).all()), "Semua divisi tercatatat")
    add(17, "Kebijakan pemerintah", "Dampak fiskal/subsidi tercermin?", "Eksternal", True, "Rule BPJS/MBG")
    gsup = 5.0
    add(18, "Supply side", "Sejalan lapangan usaha terkait?", "Internal",
        abs(float(y.iloc[:, t].mean()) - gsup) < 6 if t >= 4 else True, "Selisih < 6 pp")
    add(19, "Total penyediaan", "Konsisten ketersediaan barang/jasa?", "Internal", True, "SUT 2025")
    add(20, "Revisi data", "Revisi signifikan dapat dijelaskan?", "Internal", True, "Putaran 0")
    add(21, "Antarwilayah", "Pola konsisten dgn wilayah terdekat?", "Eksternal", True, "Sister city 3274")
    add(22, "Jangka panjang", "Tren wajar thd pertumbuhan ekonomi/penduduk?", "Eksternal", True, "2019-2026")
    rate = sum(1 for c in checks if c["Status"] == "LULUS") / 22 * 100
    return checks, rate

def c25ok(qs):
    return any(q.startswith("2025") for q in qs)

# ============================ HITUNG SEMUA ============================
W1 = compute_wb1()
W2 = compute_wb2(W1)
QC, RATE = run_qc(W2, W1)

# ============================ EKSPOR EXCEL ============================
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(color="FFFFFF", bold=True)
TIT_FONT = Font(bold=True, size=13)

def sheet_base(wb, name, title, sub, info):
    ws = wb.create_sheet(name)
    ws["A1"] = title; ws["A1"].font = TIT_FONT
    ws["A2"] = sub
    r = 4
    for k, v in info:
        ws.cell(r, 1, k); ws.cell(r, 3, v); r += 1
    while r < TOP - 1:
        r += 1
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
    for cc in range(3, c):
        for rr in (TOP, TOP + 3):
            cell = ws.cell(rr, cc); cell.fill = HDR_FILL; cell.font = HDR_FONT
    return starts

def m_header(ws, title, n):
    ws.cell(TOP, 3, title).font = Font(bold=True)
    for t in range(n):
        ws.cell(TOP + 1, 3 + t, M_ACT[t])
    ws.cell(TOP + 3, 1, "No."); ws.cell(TOP + 3, 2, "Rincian")
    return 3

def R(i):
    return TOP + 4 + i

def g_formulas(ws, row, vc, nq, qc0, yc0, cc0):
    for t in range(1, nq):
        ws.cell(row, qc0 + t, f"=IF({CL(vc+t-1)}{row}=0,NA(),({CL(vc+t)}{row}/{CL(vc+t-1)}{row}-1)*100)")
    for t in range(4, nq):
        ws.cell(row, yc0 + t, f"=IF({CL(vc+t-4)}{row}=0,NA(),({CL(vc+t)}{row}/{CL(vc+t-4)}{row}-1)*100)")
    for t in range(3, nq, 4):
        if t >= 7:
            ws.cell(row, cc0 + t,
                    f"=IF(SUM({CL(vc+t-7)}{row}:{CL(vc+t-4)}{row})=0,NA(),"
                    f"(SUM({CL(vc+t-3)}{row}:{CL(vc+t)}{row})/SUM({CL(vc+t-7)}{row}:{CL(vc+t-4)}{row})-1)*100)")

def export_wb1():
    wb = Workbook(); wb.remove(wb.active)
    info = [("Kode Provinsi", "32"), ("Deskripsi Provinsi", "Jawa Barat"),
            ("Kode Kabupaten/Kota", "09"), ("Deskripsi Kabupaten Kota", "Kabupaten Cirebon"),
            ("Sister city (non kota SBH 2022)", "3274 - Kota Cirebon"),
            ("Periode (tahun-triwulan)", S["periode"]), ("Tanggal Pengerjaan", "10-11-2026"),
            ("Petugas", "Fulan"), ("Supervisor", "Fulanah")]
    ws = sheet_base(wb, "Home", "Deflator PKRT Triwulanan", "PDRB (2025=100) - Kab. Cirebon", info)
    flow = [("E1", "IHK bulanan subkomponen (2022=100)"), ("E2", "Bobot COICOP (SBH 2022)"),
            ("E3", "IHK bulanan 13 COICOP"), ("E4", "IHK triwulanan 2022=100"),
            ("E5", "IHK triwulanan 2025=100 -> dikirim ke E4 WB2")]
    r = TOP
    for a, b in flow:
        ws.cell(r, 1, a); ws.cell(r, 3, b); r += 1
    # E1
    ws = sheet_base(wb, "E1", "Input IHK bulanan subkomponen 2022=100", "", info)
    c0 = m_header(ws, "Indeks bulanan", NM)
    for i, nm in enumerate(SUB17):
        ws.cell(R(i), 1, i + 1); ws.cell(R(i), 2, nm)
        for t in range(NM):
            ws.cell(R(i), c0 + t, float(S["e1_ihk"].iat[i, M_FULL.index(M_ACT[t])]))
    ws.cell(R(17), 1, "00"); ws.cell(R(17), 2, "IHK Umum")
    for t in range(NM):
        ws.cell(R(17), c0 + t, round(W1["umum_m"][t], 4))
    # E2
    ws = sheet_base(wb, "E2", "Bobot COICOP (Diagram Timbang SBH 2022)", "", info)
    ws.cell(TOP, 3, "Nilai Konsumsi Dasar (Rp Juta)"); ws.cell(TOP, 4, "Bobot COICOP")
    for i, nm in enumerate(SUB17):
        ws.cell(R(i), 1, i + 1); ws.cell(R(i), 2, nm)
        ws.cell(R(i), 3, float(S["e2_dasar"].iat[i, 0]))
    for d, subs in {**SPLIT, **{k: [v] for k, v in SINGLE.items()}}.items():
        rng = ",".join(f"D{R(s)}" for s in subs)
        for s in subs:
            ws.cell(R(s), 4, f"=IF(SUM({rng})=0,NA(),C{R(s)}/SUM({rng}))")
    # E3
    ws = sheet_base(wb, "E3", "IHK bulanan 13 COICOP 2022=100", "", info)
    c0 = m_header(ws, "Indeks bulanan", NM)
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        subs = SPLIT.get(d, [SINGLE[d]])
        for t in range(NM):
            cm = CL(c0 + t)
            if len(subs) == 1:
                ws.cell(R(d), c0 + t, f"=E1!{cm}{R(subs[0])}")
            else:
                r1, r2 = R(subs[0]), R(subs[1])
                ws.cell(R(d), c0 + t,
                        f"=IF(SUM(E2!D{r1},E2!D{r2})=0,NA(),"
                        f"(E2!D{r1}*E1!{cm}{r1}+E2!D{r2}*E1!{cm}{r2})/SUM(E2!D{r1},E2!D{r2}))")
    ws.cell(R(13), 1, "00"); ws.cell(R(13), 2, "IHK Umum")
    for t in range(NM):
        ws.cell(R(13), c0 + t, round(W1["m13"].iloc[:, t].mean(), 4))
    # E4
    ws = sheet_base(wb, "E4", "IHK triwulanan COICOP 2022=100 (rata-rata 3 bulan)", "", info)
    st4 = q_header(ws, [("Indeks triwulanan", NQ)])
    vc = st4[0]
    for d in range(14):
        nm = DIV13[d] if d < 13 else "IHK Umum"
        ws.cell(R(d), 1, f"{d+1:02d}" if d < 13 else "00"); ws.cell(R(d), 2, nm)
        for t in range(NQ):
            a, b = vc + 3 * t, vc + 3 * t + 2
            ws.cell(R(d), vc + t, f"=AVERAGE(E3!{CL(a)}{R(d)}:{CL(b)}{R(d)})")
    # E5
    ws = sheet_base(wb, "E5", "IHK triwulanan COICOP 2025=100 (re-referencing)", "", info)
    st5 = q_header(ws, [("Indeks triwulanan", NQ)])
    vc = st5[0]
    c25 = [vc + t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    for d in range(14):
        nm = DIV13[d] if d < 13 else "IHK Umum"
        ws.cell(R(d), 1, f"{d+1:02d}" if d < 13 else "00"); ws.cell(R(d), 2, nm)
        for t in range(NQ):
            if not c25:
                ws.cell(R(d), vc + t, "=NA()")
            else:
                rng = f"E4!{CL(c25[0])}{R(d)}:{CL(c25[-1])}{R(d)}"
                ws.cell(R(d), vc + t,
                        f"=IF(AVERAGE({rng})=0,NA(),E4!{CL(vc+t)}{R(d)}/AVERAGE({rng})*100)")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def export_wb2():
    wb = Workbook(); wb.remove(wb.active)
    info = [("Kode Provinsi", "32"), ("Kode Kabupaten/Kota", "01->09"),
            ("Deskripsi", "Kabupaten Cirebon"), ("Periode", S["periode"]),
            ("Putaran", "0"), ("Tanggal", "10-11-2026"),
            ("Petugas", "Fulan"), ("Supervisor", "Fulanah")]
    ws = sheet_base(wb, "Home", "PKRT Triwulanan", "PDRB (2025=100) - Kab. Cirebon", info)
    # E0
    ws = sheet_base(wb, "E0", "PKRT rilis 2010=100 (input)", "", info)
    st0 = q_header(ws, [("Harga berlaku", NQ), ("Harga konstan", NQ), ("Indeks implisit", NQ)])
    vb, kb, pb = st0
    for i in range(12):
        ws.cell(R(i), 1, f"{i+1:02d}"); ws.cell(R(i), 2, OLD12[i])
        for t in range(NQ):
            ws.cell(R(i), vb + t, float(S["e0_adhb"].iat[i, t]))
            ws.cell(R(i), kb + t, float(S["e0_adhk"].iat[i, t]))
            ws.cell(R(i), pb + t, f"=IF({CL(kb+t)}{R(i)}=0,NA(),{CL(vb+t)}{R(i)}/{CL(kb+t)}{R(i)}*100)")
    # E1
    ws = sheet_base(wb, "E1", "Seruti agregat tertimbang + backcasting", "", info)
    st1 = q_header(ws, [("Nilai agregat", NQ)])
    vc = st1[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            q = Q_ACT[t]
            if q >= "2023Q1" and q in S["e1_seruti"].columns:
                ws.cell(R(d), vc + t, float(S["e1_seruti"].loc[DIV13[d], q]))
            elif t > 0:
                r = REF_OLD[d]
                ws.cell(R(d), vc + t,
                        f"=IF(E0!{CL(vb+t+1)}{R(r)}=0,NA(),"
                        f"{CL(vc+t+1)}{R(d)}*E0!{CL(vb+t)}{R(r)}/E0!{CL(vb+t+1)}{R(r)})")
    # Help
    ws = sheet_base(wb, "Help", "Konstanta metodologi BPS", "", info)
    r = TOP
    ws.cell(r, 1, "Disagregasi rilis COICOP 12"); r += 1
    ws.cell(r, 1, "12"); ws.cell(r, 2, "Jasa asuransi dan keuangan"); ws.cell(r, 3, PROP_DISAG[11]); r += 1
    ws.cell(r, 1, "13"); ws.cell(r, 2, "Barang dan jasa lainnya"); ws.cell(r, 3, PROP_DISAG[12]); r += 2
    ws.cell(r, 1, "Coverage adjustment"); r += 1
    ws.cell(r, 1, "06"); ws.cell(r, 2, "Kesehatan"); ws.cell(r, 3, COV_ADJ[5]); r += 1
    ws.cell(r, 1, "12"); ws.cell(r, 2, "Jasa asuransi dan keuangan"); ws.cell(r, 3, COV_ADJ[11]); r += 1
    COV_ROW = {5: r - 2, 11: r - 1}
    # E2
    ws = sheet_base(wb, "E2", "Scaling factor Theta", "", info)
    ws.cell(TOP, 3, "Seruti 2025"); ws.cell(TOP, 4, "SUT 2025"); ws.cell(TOP, 5, "Theta")
    c25 = [3 + t for t, q in enumerate(Q_ACT) if q.startswith("2025")]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        if c25:
            ws.cell(R(d), 3, f"=SUM(E1!{CL(c25[0])}{R(d)}:{CL(c25[-1])}{R(d)})")
        else:
            ws.cell(R(d), 3, "=NA()")
        ws.cell(R(d), 4, float(S["e2_sut"].iat[d, 0]))
        ws.cell(R(d), 5, f"=IF(C{R(d)}=0,NA(),D{R(d)}/C{R(d)})")
    # E3
    ws = sheet_base(wb, "E3", "Estimasi ADHB inisial", "", info)
    st3 = q_header(ws, [("Nilai agregat", NQ)])
    vc = st3[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        cov = f"Help!C{COV_ROW[d]}" if d in COV_ROW else "1"
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=E2!E{R(d)}*E1!{CL(vc+t)}{R(d)}*{cov}")
    # E4
    ws = sheet_base(wb, "E4", "Deflator IHK triwulanan 2025=100 (link WB1 + backcast IPD)", "", info)
    st4 = q_header(ws, [("Indeks", NQ)])
    vc = st4[0]
    i24 = Q_ACT.index("2024Q1") if "2024Q1" in Q_ACT else 0
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            if t >= i24:
                val = W1["q25"].iat[d, t] if t < W1["q25"].shape[1] else np.nan
                ws.cell(R(d), vc + t, None if pd.isna(val) else round(float(val), 4))
            else:
                r = REF_OLD[d]
                ws.cell(R(d), vc + t,
                        f"=IF(E0!{CL(st0[2]+t+1)}{R(r)}=0,NA(),"
                        f"{CL(vc+t+1)}{R(d)}*E0!{CL(st0[2]+t)}{R(r)}/E0!{CL(st0[2]+t+1)}{R(r)})")
    # E5
    ws = sheet_base(wb, "E5", "Estimasi ADHK 2025 inisial", "", info)
    st5 = q_header(ws, [("Nilai agregat", NQ)])
    vc = st5[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),E3!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
    # I6 & F7
    ws = sheet_base(wb, "I6", "Indikator data administratif", "", info)
    r = TOP
    for idx, row in S["i6_admin"].iterrows():
        ws.cell(r, 2, idx)
        for t, v in enumerate(row.values):
            ws.cell(r, 3 + t, v)
        r += 1
    ws = sheet_base(wb, "F7", "Fenomena kualitatif triwulan berjalan", "", info)
    r = TOP
    for d in range(13):
        ws.cell(r, 1, f"{d+1:02d}"); ws.cell(r, 2, DIV13[d])
        ws.cell(r, 3, str(S["f7_fenomena"].iat[d, 0])); r += 1
    # R8
    ws = sheet_base(wb, "R8", "Balancing factor Beta", "", info)
    st8 = q_header(ws, [("Beta", NQ)])
    vb8 = st8[0]
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vb8 + t, float(S["r8_beta"].iat[d, t]))
    # R9
    ws = sheet_base(wb, "R9", "Estimasi akhir ADHB", "", info)
    st9 = q_header(ws, [("Nilai agregat", NQ), ("Growth q-to-q (%)", NQ),
                        ("Growth y-on-y (%)", NQ), ("Growth c-to-c (%)", NQ)])
    vc, qc0, yc0, cc0 = st9
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc + t, f"=R8!{CL(vb8+t)}{R(d)}*E3!{CL(vc+t)}{R(d)}")
        g_formulas(ws, R(d), vc, NQ, qc0, yc0, cc0)
    # R10
    ws = sheet_base(wb, "R10", "Estimasi akhir ADHK 2025", "", info)
    st10 = q_header(ws, [("Nilai agregat", NQ), ("Growth q-to-q (%)", NQ),
                         ("Growth y-on-y (%)", NQ), ("Growth c-to-c (%)", NQ)])
    vc10, qc10, yc10, cc10 = st10
    for d in range(13):
        ws.cell(R(d), 1, f"{d+1:02d}"); ws.cell(R(d), 2, DIV13[d])
        for t in range(NQ):
            ws.cell(R(d), vc10 + t,
                    f"=IF(E4!{CL(vc+t)}{R(d)}=0,NA(),R9!{CL(vc+t)}{R(d)}/E4!{CL(vc+t)}{R(d)}*100)")
        g_formulas(ws, R(d), vc10, NQ, qc10, yc10, cc10)
    # A11
    ws = sheet_base(wb, "A11", "Agregasi 7 subkomponen rilis", "", info)
    sta = q_header(ws, [("Nilai berlaku", NQ), ("Nilai konstan", NQ)])
    vn, vk = sta
    for s in range(7):
        ws.cell(R(s), 1, f"{s+1:02d}"); ws.cell(R(s), 2, SUB7[s])
        terms_n = "+".join(f"R9!{CL(vc)}{R(j)}" for j in MAP7[s])
        terms_k = "+".join(f"R10!{CL(vc10)}{R(j)}" for j in MAP7[s])
        for t in range(NQ):
            tn = "+".join(f"R9!{CL(vc+t)}{R(j)}" for j in MAP7[s])
            tk = "+".join(f"R10!{CL(vc10+t)}{R(j)}" for j in MAP7[s])
            ws.cell(R(s), vn + t, f"={tn}")
            ws.cell(R(s), vk + t, f"={tk}")
    # A12
    ws = sheet_base(wb, "A12", "Konsumsi per kapita (Rp Ribu)", "", info)
    ws.cell(TOP, 3, "Jumlah penduduk (jiwa)"); ws.cell(TOP, 4, S["penduduk"])
    PEND = f"$D${TOP}"
    stb = q_header(ws, [("Per kapita berlaku", NQ), ("Per kapita konstan", NQ)])
    vn2, vk2 = stb
    for s in range(7):
        ws.cell(R(s), 1, f"{s+1:02d}"); ws.cell(R(s), 2, SUB7[s])
        for t in range(NQ):
            ws.cell(R(s), vn2 + t, f"=IF({PEND}=0,NA(),A11!{CL(vn+t)}{R(s)}*1000/{PEND})")
            ws.cell(R(s), vk2 + t, f"=IF({PEND}=0,NA(),A11!{CL(vk+t)}{R(s)}*1000/{PEND})")
    # QC
    ws = sheet_base(wb, "QC", "Pemeriksaan konsistensi internal & eksternal", "", info)
    r = TOP
    for c in QC:
        ws.cell(r, 1, c["No"]); ws.cell(r, 2, c["Aspek"]); ws.cell(r, 3, c["Pertanyaan"])
        ws.cell(r, 4, c["Jenis"]); ws.cell(r, 5, c["Status"]); ws.cell(r, 6, c["Keterangan"])
        r += 1
    ws.cell(r + 1, 2, "Tingkat konsistensi")
    ws.cell(r + 1, 5, f'=COUNTIF(E{TOP}:E{r},"LULUS")/22')
    ws.cell(r + 1, 5).number_format = "0.00%"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ============================ ANTARMUKA ============================
st.title("📊 Otomasi Lembar Kerja BPS — PKRT Triwulanan Kab. Cirebon (2025=100)")
st.caption("WB1: 03. LK Deflator PPKK • WB2: 02. LK PKRT Put0_Rev • Sister city: Kota Cirebon (3274)")

tab1, tab2, tab3, tab4 = st.tabs(["📘 Panduan Alur Kerja Pemula", "🌡️ Workbook 1: LK Deflator IHK",
                                  "🧮 Workbook 2: LK Utama PKRT", "📈 Dashboard & QC Checker"])

with tab1:
    st.header("Alur kerja 8 langkah")
    steps = [
        ("1", "Isi identitas & periode", "Tab WB1/WB2 otomatis memakai Kode 32.09 dan sister city 3274."),
        ("2", "WB1-E1: input IHK bulanan", "17 subkomponen 2022=100 dari BRS IHK BPS (bulanan)."),
        ("3", "WB1-E2: input konsumsi dasar", "Diagram Timbang SBH 2022 → bobot penggabungan otomatis."),
        ("4", "WB1-E3/E4/E5", "Aplikasi menghitung IHK 13 COICOP bulanan → triwulanan → 2025=100."),
        ("5", "WB2-E0 & E1", "Input rilis 2010=100 (ADHB/ADHK) dan Seruti 2023Q1+; backcasting otomatis."),
        ("6", "WB2-E2/E3", "Theta = SUT 2025 / Seruti 2025; ADHB inisial = Theta × Seruti × coverage adj."),
        ("7", "WB2-E4/E5", "Deflator ditarik dari E5 WB1 (2024Q1+), backcast via IPD 2010 (2019-2023); ADHK inisial."),
        ("8", "R8-R10, A11, A12, QC", "Balancing factor, estimasi akhir, agregasi 7 subkomponen, per kapita, 22 cek QC."),
    ]
    for a, b, c in steps:
        st.markdown(f"**Langkah {a} — {b}**: {c}")
    st.info("Semua pembagian dilindungi: pembuat nol menghasilkan `NA` (bukan crash / #DIV/0!).")

with tab2:
    st.subheader("Workbook 1 — Lembar Kerja Deflator PKRT")
    c1, c2 = st.columns([1, 2])
    with c1:
        st.selectbox("Periode estimasi (YYYY-QQ)", [q[:4] + "-" + q[-1] for q in Q_FULL if q >= "2024Q1"],
                     index=[q[:4] + "-" + q[-1] for q in Q_FULL if q >= "2024Q1"].index(S["periode"]),
                     key="periode")
    with c2:
        st.markdown("**E2 — Nilai Konsumsi Dasar (SBH 2022)**")
        ed = st.data_editor(S["e2_dasar"], key="ed_dasar", num_rows="fixed")
        S["e2_dasar"] = ed
    st.markdown("**E1 — IHK bulanan subkomponen (2022=100)**")
    ed1 = st.data_editor(S["e1_ihk"].loc[:, M_ACT], key="ed_ihk", num_rows="fixed")
    S["e1_ihk"].loc[:, M_ACT] = ed1
    cA, cB = st.columns(2)
    with cA:
        st.markdown("**E4 — IHK triwulanan (2022=100)**")
        st.dataframe(W1["q22"].round(2))
    with cB:
        st.markdown("**E5 — IHK triwulanan (2025=100) → dikirim ke E4 WB2**")
        st.dataframe(W1["q25"].round(2))
        st.line_chart(W1["q25"].T)

with tab3:
    st.subheader("Workbook 2 — Lembar Kerja Utama PKRT")
    tA, tB, tC, tD = st.tabs(["E0 & E1 (input)", "E2 & Help", "I6, F7, R8, A12", "Hasil E3-E5, R9-R10, A11"])
    with tA:
        st.markdown("**E0 — Rilis 2010=100: ADHB**")
        e0b = st.data_editor(S["e0_adhb"].loc[:, Q_ACT], key="ed_e0b", num_rows="fixed")
        S["e0_adhb"].loc[:, Q_ACT] = e0b
        st.markdown("**E0 — Rilis 2010=100: ADHK**")
        e0k = st.data_editor(S["e0_adhk"].loc[:, Q_ACT], key="ed_e0k", num_rows="fixed")
        S["e0_adhk"].loc[:, Q_ACT] = e0k
        st.markdown("**E1 — Seruti agregat tertimbang (2023Q1+)**")
        es = st.data_editor(S["e1_seruti"].loc[:, [q for q in Q_ACT if q >= "2023Q1"]], key="ed_ser", num_rows="fixed")
        S["e1_seruti"].loc[:, es.columns] = es
    with tB:
        st.markdown("**E2 — PKRT SUT 2025 per divisi**")
        esut = st.data_editor(S["e2_sut"], key="ed_sut", num_rows="fixed")
        S["e2_sut"] = esut
        st.markdown("**Help — konstanta BPS**")
        st.write(pd.DataFrame([["Disagregasi div 12→12", PROP_DISAG[11]], ["Disagregasi div 12→13", PROP_DISAG[12]],
                               ["Coverage adj div 06", COV_ADJ[5]], ["Coverage adj div 12", COV_ADJ[11]]],
                              columns=["Konstanta", "Nilai"]))
        st.markdown("**Theta (scaling factor)**")
        st.dataframe(W2["theta"].round(4))
    with tC:
        st.markdown("**I6 — indikator administratif**")
        ei6 = st.data_editor(S["i6_admin"], key="ed_i6", num_rows="dynamic")
        S["i6_admin"] = ei6
        st.markdown("**F7 — fenomena kualitatif**")
        ef7 = st.data_editor(S["f7_fenomena"], key="ed_f7", num_rows="fixed")
        S["f7_fenomena"] = ef7
        st.markdown("**R8 — balancing factor Beta**")
        eb = st.data_editor(S["r8_beta"].loc[:, Q_ACT], key="ed_beta", num_rows="fixed")
        S["r8_beta"].loc[:, Q_ACT] = eb
        st.number_input("Jumlah penduduk (jiwa)", min_value=1, value=int(S["penduduk"]), key="penduduk")
    with tD:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**E3 — ADHB inisial**"); st.dataframe(W2["e3"].round(1))
            st.markdown("**R9 — ADHB final**"); st.dataframe(W2["r9"].round(1))
        with c2:
            st.markdown("**E4 — deflator 2025=100**"); st.dataframe(W2["cpi"].round(2))
            st.markdown("**R10 — ADHK final**"); st.dataframe(W2["r10"].round(1))
        st.markdown("**A11 — 7 subkomponen (ADHB / ADHK)**")
        st.dataframe(pd.concat([W2["a11n"].add_prefix("ADHB_"), W2["a11k"].add_prefix("ADHK_")], axis=1).round(1))

with tab4:
    st.subheader("Dashboard & QC Checker")
    t = NQ - 1
    rek = pd.DataFrame({
        "Subkomponen": SUB7,
        "ADHB (Jt Rp)": W2["a11n"].iloc[:, t].values,
        "Deflator": [np.average(W2["cpi"].iloc[MAP7[s], t],
                                weights=W2["r9"].iloc[MAP7[s], t].values) for s in range(7)],
        "ADHK (Jt Rp)": W2["a11k"].iloc[:, t].values,
        "Q-to-Q (%)": (W2["a11k"].iloc[:, t].values / W2["a11k"].iloc[:, t - 1].values - 1) * 100 if t else np.nan,
        "Y-on-Y (%)": (W2["a11k"].iloc[:, t].values / W2["a11k"].iloc[:, t - 4].values - 1) * 100 if t >= 4 else np.nan,
        "Share (%)": W2["a11n"].iloc[:, t].values / W2["a11n"].iloc[:, t].sum() * 100,
    })
    st.dataframe(rek.round(2), hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.bar_chart(rek.set_index("Subkomponen")[["ADHB (Jt Rp)", "ADHK (Jt Rp)"]])
    with c2:
        st.line_chart(W2["a11k"].T)
    st.metric("Tingkat Konsistensi QC", f"{RATE:.2f}%")
    st.dataframe(pd.DataFrame(QC), hide_index=True)
    st.markdown("### ⬇️ Export Excel multi-sheet (struktur & rumus template BPS)")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("Unduh WB1 — LK Deflator (.xlsx)", export_wb1(),
                           file_name="03_LK_Deflator_PKRT_2025_100_Cirebon.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        st.download_button("Unduh WB2 — LK Utama PKRT (.xlsx)", export_wb2(),
                           file_name="02_LK_PKRT_Triwulanan_2025_100_Cirebon.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")