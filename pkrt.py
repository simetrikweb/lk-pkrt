# -*- coding: utf-8 -*-
# =====================================================================
#  APLIKASI LEMBAR KERJA PK-RT PDRB KABUPATEN CIREBON (REBASING 2025=100)
#  Mesin: PK-RT Expert & Engine Calculator | Metodologi BPS - SNA 2008 & SUT 2025
#  Jalankan: streamlit run app.py
# =====================================================================
import streamlit as st
import pandas as pd
import numpy as np

# ----------------------------- KONFIGURASI UMUM -----------------------------
st.set_page_config(page_title="PK-RT PDRB Kab. Cirebon (2025=100)",
                   page_icon="📊", layout="wide")

st.markdown("""
<style>
.hero{background:linear-gradient(90deg,#0b3d91,#1565c0);color:#fff;
      padding:18px 22px;border-radius:12px;margin-bottom:6px}
.hero h1{font-size:1.45rem;margin:0}
.hero p{margin:4px 0 0;opacity:.92;font-size:.92rem}
.smallnote{color:#666;font-size:.82rem}
</style>
""", unsafe_allow_html=True)

# ----------------------------- KONSTANTA WILAYAH & PERIODE -----------------------------
KODE_WILAYAH = "3209"
NAMA_WILAYAH = "Kabupaten Cirebon"
TAHUN_DASAR  = 2025
PERIODE = ["2025-Q1", "2025-Q2", "2025-Q3", "2025-Q4",
           "2026-Q1", "2026-Q2", "2026-Q3", "2026-Q4"]
GCOLS = PERIODE[1:]                      # kolom input laju indikator (7 triwulan)
BASE_COL = "ADHB 2025-Q1 (Jt Rp)"

SUBKOMPONEN = [
    "Makanan, Minuman, & Tembakau",
    "Pakaian & Alas Kaki",
    "Perumahan, Air, Listrik, Gas & BB Lainnya",
    "Kesehatan (Non-BPJS) & Pendidikan",
    "Transportasi, Informasi, & Komunikasi",
    "Restoran & Hotel",
    "Lainnya (Barang & Jasa Pribadi/Lainnya)",
]
NAMA_PENDEK = ["Makanan & Minuman", "Pakaian", "Perumahan",
               "Kesehatan & Pendidikan", "Transportasi",
               "Restoran & Hotel", "Lainnya"]

DIVISI = [
    "Makanan & Minuman", "Tembakau", "Pakaian & Alas Kaki",
    "Perumahan, Air, Listrik, Gas & BB Lainnya",
    "Perlengkapan & Pemeliharaan Rutinan RT", "Kesehatan",
    "Transportasi", "Informasi & Komunikasi",
    "Rekreasi, Olahraga & Budaya", "Pendidikan",
    "Penyediaan Makanan & Minuman (Restoran)",
    "Jasa Akomodasi (Hotel)", "Perawatan Pribadi & Jasa Lainnya",
]

# Bobot IHK umum (persentase, total = 100)
BOBOT_IHK = {DIVISI[0]: 24.0, DIVISI[1]: 5.0, DIVISI[2]: 6.0, DIVISI[3]: 20.0,
             DIVISI[4]: 6.0, DIVISI[5]: 4.5, DIVISI[6]: 11.0, DIVISI[7]: 5.0,
             DIVISI[8]: 4.0, DIVISI[9]: 2.5, DIVISI[10]: 7.0,
             DIVISI[11]: 1.5, DIVISI[12]: 3.5}

# Pemetaan Divisi COICOP 2018 -> 7 Subkomponen Rilis (Rule 3)
MAPPING = [
    {DIVISI[0]: 0.85, DIVISI[1]: 0.15},
    {DIVISI[2]: 1.00},
    {DIVISI[3]: 1.00},
    {DIVISI[5]: 0.55, DIVISI[9]: 0.45},
    {DIVISI[6]: 0.60, DIVISI[7]: 0.40},
    {DIVISI[10]: 0.75, DIVISI[11]: 0.25},
    {DIVISI[4]: 0.35, DIVISI[8]: 0.35, DIVISI[12]: 0.30},
]

# ----------------------------- DATA CONTOH (PROXY KAB. CIREBON) -----------------------------
def default_pdrb_in():
    data = {
        BASE_COL: [2_870_000, 478_000, 1_546_000, 515_000, 957_000, 589_000, 405_000],
        "2025-Q2": [0.8, 0.5, 0.9, 1.0, 1.0, 1.2, 1.0],
        "2025-Q3": [1.2, 0.6, 0.9, 2.5, 1.0, 1.2, 1.0],
        "2025-Q4": [0.9, 1.5, 1.0, 1.0, 2.2, 2.6, 1.1],
        "2026-Q1": [2.8, 3.2, 1.0, 1.0, 1.8, 2.0, 1.2],
        "2026-Q2": [0.7, 0.4, 0.9, 1.0, 1.0, 1.2, 1.0],
        "2026-Q3": [1.1, 0.6, 0.9, 2.4, 1.0, 1.2, 1.0],
        "2026-Q4": [0.9, 1.4, 1.0, 1.0, 2.1, 2.5, 1.1],
    }
    return pd.DataFrame(data, index=SUBKOMPONEN)

def default_susenas():
    rows = [
        ("Padi-padian, Umbi-umbian & Hasilnya", 950_000),
        ("Ikan, Daging, Telur, Susu & Kacang-kacangan", 1_420_000),
        ("Sayur-sayuran & Buah-buahan", 1_080_000),
        ("Minyak, Bumbu, Gula & Minuman", 1_210_000),
        ("Rokok & Tembakau", 780_000),
        ("Makanan & Minuman Jadi (termasuk makan di luar)", 1_475_000),
        ("Perumahan, Utilitas & Bahan Bakar RT", 2_330_000),
        ("Barang Tahan Lama & Perlengkapan RT", 870_000),
        ("Pakaian, Alas Kaki & Tutup Kepala", 465_000),
        ("Transportasi & Komunikasi", 1_150_000),
        ("Pendidikan & Kesehatan (Non-BPJS)", 620_000),
        ("Perawatan Pribadi & Jasa Lainnya", 455_000),
    ]
    return pd.DataFrame(rows, columns=["Kelompok Pengeluaran", "Rp/Kapita/Tahun"])

def default_ihk():
    base = {DIVISI[0]: 0.45, DIVISI[1]: 0.60, DIVISI[2]: 0.35, DIVISI[3]: 0.40,
            DIVISI[4]: 0.30, DIVISI[5]: 0.50, DIVISI[6]: 0.55, DIVISI[7]: 0.10,
            DIVISI[8]: 0.40, DIVISI[9]: 0.70, DIVISI[10]: 0.50,
            DIVISI[11]: 0.60, DIVISI[12]: 0.45}
    rows = {}
    for div, b in base.items():
        vals = []
        for t in range(8):
            v = b
            q = PERIODE[t][-2:]
            if q == "Q1" and div in (DIVISI[0], DIVISI[1], DIVISI[2], DIVISI[6]):
                v += 0.35          # efek Ramadan/Lebaran
            if q == "Q3" and div == DIVISI[9]:
                v += 0.50          # tahun ajaran baru
            if q == "Q4" and div in (DIVISI[6], DIVISI[10], DIVISI[11]):
                v += 0.30          # Nataru
            vals.append(round(v, 2))
        rows[div] = vals
    df = pd.DataFrame(rows).T
    df.columns = PERIODE
    return df

def init_defaults():
    S = st.session_state
    S.setdefault("mode_periode", "Triwulanan")
    S.setdefault("fokus_tw", "2026-Q1")
    S.setdefault("fokus_th", "Tahunan 2026")
    S.setdefault("penduduk", 2_304_567)
    S.setdefault("pdrb_in", default_pdrb_in())
    S.setdefault("susenas_in", default_susenas())
    S.setdefault("ihk_in", default_ihk())
    S.setdefault("rule_bpjs", True)
    S.setdefault("rule_stik", True)
    S.setdefault("bpjs_miliar", 185.0)
    S.setdefault("stik_miliar", 120.0)
    S.setdefault("ood_pct", 38.0)
    S.setdefault("fisim_miliar", 210.0)
    S.setdefault("asuransi_miliar", 45.0)
    S.setdefault("prox_penumpang", 6.5)
    S.setdefault("prox_tpk", 42.3)
    S.setdefault("prox_kwh", 4.8)
    S.setdefault("prox_industri", 5.2)
    S.setdefault("catatan_fenomena", "")

def reset_all():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

# ----------------------------- UTILITAS FORMAT ANGKA -----------------------------
def fnum(x, dec=1):
    """Format angka gaya Indonesia (titik ribuan, koma desimal)."""
    try:
        if x is None or pd.isna(x):
            return "–"
    except Exception:
        pass
    s = f"{x:,.{dec}f}"
    return s.replace(",", "·").replace(".", ",").replace("·", ".")

# ----------------------------- MESIN PENGHITUNG (TAHAP 4) -----------------------------
def hitung_level_ihk(ihk_df):
    """Level IHK per divisi (2025-Q1 = 100) dan IHK umum."""
    ihk = ihk_df.astype(float)
    level = {}
    for div in DIVISI:
        infl = ihk.loc[div].tolist()
        lv = [100.0]
        for t in range(1, 8):
            lv.append(lv[-1] * (1.0 + infl[t] / 100.0))
        level[div] = lv
    umum = [sum(BOBOT_IHK[d] * level[d][t] for d in DIVISI) / 100.0 for t in range(8)]
    return level, umum

def hitung_semua():
    S = st.session_state
    dfp = S["pdrb_in"]
    base = dfp.iloc[:, 0].astype(float).tolist()
    growth = dfp.iloc[:, 1:].astype(float)

    level, umum = hitung_level_ihk(S["ihk_in"])

    # Deflator per subkomponen = rata-rata tertimbang IHK divisi (Rule 3)
    defl = {}
    for i, sk in enumerate(SUBKOMPONEN):
        w = MAPPING[i]
        ws = sum(w.values())
        defl[sk] = [sum(w[d] * level[d][t] for d in w) / ws for t in range(8)]

    # ADHB: ekstrapolasi dari nilai basis  ADHB_t = ADHB_{t-1} x (1 + %ΔInd/100)
    adhb_raw = {}
    for i, sk in enumerate(SUBKOMPONEN):
        ser = [base[i]]
        for q in range(7):
            ser.append(ser[-1] * (1.0 + float(growth.iloc[i, q]) / 100.0))
        adhb_raw[sk] = ser

    # Penyesuaian rebasing (Rule 1 & Rule 2) -> dialihkan ke PKP
    bpjs_q = (S["bpjs_miliar"] * 1000.0 / 4.0) if S["rule_bpjs"] else 0.0
    stik_q = (S["stik_miliar"] * 1000.0 / 4.0) if S["rule_stik"] else 0.0
    adhb = {sk: list(adhb_raw[sk]) for sk in SUBKOMPONEN}
    for t in range(8):
        adhb[SUBKOMPONEN[0]][t] -= stik_q    # STIK/MBG keluar dari makanan
        adhb[SUBKOMPONEN[3]][t] -= bpjs_q    # BPJS keluar dari kesehatan

    # ADHK 2025 = ADHB / Deflator x 100
    adhk = {}
    for sk in SUBKOMPONEN:
        adhk[sk] = []
        for t in range(8):
            d = defl[sk][t] if defl[sk][t] != 0 else 1e-9
            adhk[sk].append(adhb[sk][t] / d * 100.0)

    tot_hb = [sum(adhb[sk][t] for sk in SUBKOMPONEN) for t in range(8)]
    tot_hk = [sum(adhk[sk][t] for sk in SUBKOMPONEN) for t in range(8)]

    qoq, yoy = {}, {}
    for sk in SUBKOMPONEN:
        q = [None] * 8
        y = [None] * 8
        for t in range(1, 8):
            q[t] = (adhk[sk][t] / adhk[sk][t - 1] - 1.0) * 100.0
        for t in range(4, 8):
            y[t] = (adhk[sk][t] / adhk[sk][t - 4] - 1.0) * 100.0
        qoq[sk], yoy[sk] = q, y

    qoq_tot = [None] + [(tot_hk[t] / tot_hk[t - 1] - 1.0) * 100.0 for t in range(1, 8)]
    yoy_tot = [None] * 4 + [(tot_hk[t] / tot_hk[t - 4] - 1.0) * 100.0 for t in range(4, 8)]
    ctc = (sum(tot_hk[4:8]) / sum(tot_hk[0:4]) - 1.0) * 100.0

    S["hasil"] = {
        "defl": defl, "ihk_umum": umum, "adhb": adhb, "adhb_raw": adhb_raw,
        "adhk": adhk, "total_adhb": tot_hb, "total_adhk": tot_hk,
        "qoq": qoq, "yoy": yoy, "qoq_tot": qoq_tot, "yoy_tot": yoy_tot, "ctc": ctc,
        "bpjs_q": bpjs_q, "stik_q": stik_q,
    }

# ----------------------------- MODUL QC — 22 CHECKPOINT (TAHAP 5) -----------------------------
def jalankan_qc():
    S, H = st.session_state, st.session_state["hasil"]
    checks = []

    def add(no, nama, status, detail):
        checks.append({"No": no, "Checkpoint": nama, "Status": status, "Detail": detail})

    # 1–7 : Arah deflator vs arah IHK umum (triwulan terakhir 2026-Q4)
    t_ref = 7
    dU = H["ihk_umum"][t_ref] - H["ihk_umum"][t_ref - 1]
    arahU = "naik" if dU > 0 else ("turun" if dU < 0 else "flat")
    for i, sk in enumerate(SUBKOMPONEN):
        d = H["defl"][sk][t_ref] - H["defl"][sk][t_ref - 1]
        arah = "naik" if d > 0 else ("turun" if d < 0 else "flat")
        if (d > 0 and dU < 0) or (d < 0 and dU > 0):
            add(i + 1, f"Anomali deflator — {NAMA_PENDEK[i]}", "⚠️ PERINGATAN",
                f"Deflator {arah} saat IHK umum {arahU} (2026-Q4 vs 2026-Q3).")
        else:
            add(i + 1, f"Anomali deflator — {NAMA_PENDEK[i]}", "✅ LULUS",
                f"Pergerakan deflator searah IHK umum ({arahU}).")

    # 8–12 : Pola musiman
    def qoq_riil(sk, label):
        t = PERIODE.index(label)
        return (H["adhk"][sk][t] / H["adhk"][sk][t - 1] - 1.0) * 100.0

    musiman = [
        (8,  "2026-Q1", 0, 1.0, "Q1 (Ramadan/Lebaran): Makanan, Minuman & Tembakau wajib tumbuh positif tinggi"),
        (9,  "2026-Q1", 1, 1.0, "Q1 (Ramadan/Lebaran): Pakaian & Alas Kaki wajib tumbuh positif tinggi"),
        (10, "2025-Q3", 3, 1.2, "Q3 (Tahun Ajaran Baru): Kesehatan & Pendidikan wajib melonjak"),
        (11, "2025-Q4", 5, 1.2, "Q4 (Nataru): Restoran & Hotel wajib meningkat"),
        (12, "2025-Q4", 4, 0.8, "Q4 (Nataru): Transportasi wajib meningkat"),
    ]
    for no, label, i, thr, teks in musiman:
        g = qoq_riil(SUBKOMPONEN[i], label)
        ok = g >= thr
        add(no, teks, "✅ LULUS" if ok else "⚠️ PERINGATAN",
            f"Q-to-Q riil {label} = {fnum(g, 2)}% (ambang minimal {fnum(thr, 1)}%).")

    # 13–19 : Outlier pertumbuhan Y-on-Y (2026-Q4)
    catatan = S.get("catatan_fenomena", "").strip()
    for i, sk in enumerate(SUBKOMPONEN):
        y = H["yoy"][sk][7]
        if y is None or pd.isna(y):
            add(13 + i, f"Outlier YoY — {NAMA_PENDEK[i]}", "✅ LULUS", "YoY belum dapat dihitung.")
        elif y > 15.0 or y < -10.0:
            if catatan:
                add(13 + i, f"Outlier YoY — {NAMA_PENDEK[i]}", "✅ LULUS (tercatat)",
                    f"YoY 2026-Q4 = {fnum(y, 2)}% — fenomena khusus tercatat oleh pengguna.")
            else:
                add(13 + i, f"Outlier YoY — {NAMA_PENDEK[i]}", "⚠️ PERINGATAN",
                    f"YoY 2026-Q4 = {fnum(y, 2)}% melewati batas wajar (>15% atau <-10%) tanpa catatan.")
        else:
            add(13 + i, f"Outlier YoY — {NAMA_PENDEK[i]}", "✅ LULUS",
                f"YoY 2026-Q4 = {fnum(y, 2)}% dalam rentang wajar.")

    # 20 : Integritas agregat
    tot = H["total_adhb"][7]
    shares = [H["adhb"][sk][7] / tot * 100.0 for sk in SUBKOMPONEN]
    ok = abs(sum(shares) - 100.0) < 0.01
    add(20, "Integritas agregat (Σ share = 100%)", "✅ LULUS" if ok else "❌ GAGAL",
        f"Σ share subkomponen = {fnum(sum(shares), 3)}%.")

    # 21 : Benchmark Susenas x Penduduk
    perkap = S["susenas_in"]["Rp/Kapita/Tahun"].astype(float).sum()
    bench = perkap * S["penduduk"] / 1e6          # juta Rp / tahun
    tot25 = sum(H["total_adhb"][0:4])
    dev = (tot25 - bench) / bench * 100.0
    add(21, "Benchmark Susenas × Penduduk vs ADHB 2025",
        "✅ LULUS" if abs(dev) <= 15 else "⚠️ PERINGATAN",
        f"ADHB 2025 = {fnum(tot25, 0)} jt vs benchmark {fnum(bench, 0)} jt (deviasi {fnum(dev, 2)}%; toleransi ±15%).")

    # 22 : Nilai negatif / kosong
    neg = []
    for sk in SUBKOMPONEN:
        for t, v in enumerate(H["adhb"][sk]):
            if (not np.isfinite(v)) or v < 0:
                neg.append(f"{NAMA_PENDEK[SUBKOMPONEN.index(sk)]} {PERIODE[t]}")
    add(22, "Tidak ada nilai negatif/kosong", "✅ LULUS" if not neg else "❌ GAGAL",
        "Seluruh series nilai valid." if not neg else "Temuan: " + ", ".join(neg[:6]))
    return checks

REF_QC_SK = {0: [1, 8, 13], 1: [2, 9, 14], 2: [3, 15], 3: [4, 10, 16],
             4: [5, 12, 17], 5: [6, 11, 18], 6: [7, 19]}

def status_qc_per_sk(checks):
    byno = {c["No"]: c for c in checks}
    stat = {}
    for i in range(7):
        hasil = "✅ Valid"
        for n in REF_QC_SK[i]:
            c = byno.get(n)
            if c and "❌" in c["Status"]:
                hasil = "❌ Gagal"
            elif c and "⚠️" in c["Status"] and hasil == "✅ Valid":
                hasil = "⚠️ Perlu Cek"
        stat[i] = hasil
    return stat

# ----------------------------- REKAPITULASI (SHEET R8/R10) -----------------------------
def buat_rekap(fokus, stat_sk):
    H = st.session_state["hasil"]
    rows = []
    if fokus.startswith("Tahunan"):
        tahun = int(fokus.split()[1])
        idxs = [0, 1, 2, 3] if tahun == 2025 else [4, 5, 6, 7]
        for i, sk in enumerate(SUBKOMPONEN):
            hb = sum(H["adhb"][sk][t] for t in idxs)
            df_ = sum(H["defl"][sk][t] for t in idxs) / 4.0
            hk = sum(H["adhk"][sk][t] for t in idxs)
            if tahun == 2026:
                hk_prev = sum(H["adhk"][sk][t] for t in [0, 1, 2, 3])
                yoy = (hk / hk_prev - 1.0) * 100.0
            else:
                yoy = None
            rows.append([i + 1, sk, hb, df_, hk, None, yoy, stat_sk[i]])
        tot_hb = sum(r[2] for r in rows)
        tot_hk = sum(r[4] for r in rows)
        defl_avg = sum(r[3] * r[2] for r in rows) / tot_hb
        if tahun == 2026:
            yoy_t = (tot_hk / sum(H["total_adhk"][0:4]) - 1.0) * 100.0
        else:
            yoy_t = None
        total_row = ["TOTAL", "PDRB PK-RT KABUPATEN CIREBON", tot_hb, defl_avg,
                     tot_hk, None, yoy_t, "–"]
        return pd.DataFrame(rows + [total_row],
                            columns=["No", "Subkomponen PK-RT (COICOP 2018)",
                                     "ADHB (Jt Rp)", "Deflator (2025=100)", "ADHK (Jt Rp)",
                                     "Q-to-Q (%)", "Y-on-Y (%)", "Status QC"])
    else:
        t = PERIODE.index(fokus)
        for i, sk in enumerate(SUBKOMPONEN):
            rows.append([i + 1, sk, H["adhb"][sk][t], H["defl"][sk][t], H["adhk"][sk][t],
                         H["qoq"][sk][t], H["yoy"][sk][t], stat_sk[i]])
        tot_hb = H["total_adhb"][t]
        tot_hk = H["total_adhk"][t]
        defl_w = sum(H["defl"][sk][t] * H["adhb"][sk][t] for sk in SUBKOMPONEN) / tot_hb
        total_row = ["TOTAL", "PDRB PK-RT KABUPATEN CIREBON", tot_hb, defl_w, tot_hk,
                     H["qoq_tot"][t], H["yoy_tot"][t], "–"]
        return pd.DataFrame(rows + [total_row],
                            columns=["No", "Subkomponen PK-RT (COICOP 2018)",
                                     "ADHB (Jt Rp)", "Deflator (2025=100)", "ADHK (Jt Rp)",
                                     "Q-to-Q (%)", "Y-on-Y (%)", "Status QC"])

def rekap_display(df):
    d = df.copy()
    d["ADHB (Jt Rp)"] = d["ADHB (Jt Rp)"].apply(lambda x: fnum(x, 1))
    d["Deflator (2025=100)"] = d["Deflator (2025=100)"].apply(lambda x: fnum(x, 2))
    d["ADHK (Jt Rp)"] = d["ADHK (Jt Rp)"].apply(lambda x: fnum(x, 1))
    d["Q-to-Q (%)"] = d["Q-to-Q (%)"].apply(lambda x: "–" if x is None or pd.isna(x) else fnum(x, 2))
    d["Y-on-Y (%)"] = d["Y-on-Y (%)"].apply(lambda x: "–" if x is None or pd.isna(x) else fnum(x, 2))
    return d

# =====================================================================
#                              ANTARMUKA
# =====================================================================
init_defaults()
hitung_semua()
S = st.session_state
fokus = S["fokus_th"] if S["mode_periode"] == "Tahunan" else S["fokus_tw"]

st.markdown(f"""
<div class='hero'>
  <h1>📊 Lembar Kerja Interaktif PK‑RT PDRB {NAMA_WILAYAH}</h1>
  <p>Rebasing 2025 = 100 · SNA 2008 & SUT 2025 · Kode Wilayah {KODE_WILAYAH} · Fokus: <b>{fokus}</b></p>
</div>
""", unsafe_allow_html=True)

# ----------------------------- SIDEBAR -----------------------------
st.sidebar.title("🧭 Navigasi Tahapan")
page = st.sidebar.radio(
    "Pilih tahap:",
    ["🏁 Tahap 1 — Konfigurasi", "📥 Tahap 2 — Input Data",
     "🧹 Tahap 3 — Penyesuaian Rebasing", "🧮 Tahap 4 — Penghitungan",
     "🛡️ Tahap 5 — QC (22 Checkpoint)", "📊 Dashboard Hasil"],
    label_visibility="collapsed")
st.sidebar.markdown("---")
st.sidebar.caption(f"**{NAMA_WILAYAH}** ({KODE_WILAYAH})\n\nTahun Dasar **2025 = 100**")
if st.sidebar.button("🔄 Muat Ulang Data Contoh", use_container_width=True):
    reset_all()
st.sidebar.markdown("---")
st.sidebar.caption("👥 *Untuk pengguna awam*: setiap langkah disertai penjelasan **mengapa** data dibutuhkan dan **di mana** memperolehnya.")

# =====================================================================
# TAHAP 1 — KONFIGURASI
# =====================================================================
if page.startswith("🏁"):
    st.subheader("🏁 Tahap 1 · Konfigurasi Periode & Wilayah")
    c1, c2, c3 = st.columns(3)
    c1.metric("Kode Wilayah", KODE_WILAYAH)
    c2.metric("Wilayah (wajib)", NAMA_WILAYAH)
    c3.metric("Tahun Dasar", "2025 = 100")
    st.markdown("###### Pengaturan Periode Penghitungan")
    col_a, col_b = st.columns([1, 2])
    with col_a:
        st.radio("Mode periode", ["Triwulanan", "Tahunan"], key="mode_periode")
    with col_b:
        if S["mode_periode"] == "Triwulanan":
            st.selectbox("Triwulan fokus analisis", PERIODE, index=4, key="fokus_tw",
                         help="Semua seri tetap dihitung 2025-Q1 s.d. 2026-Q4; pilihan ini menentukan triwulan yang ditampilkan pada rekap & QC musiman.")
        else:
            st.selectbox("Tahun fokus analisis", ["Tahunan 2025", "Tahunan 2026"],
                         index=1, key="fokus_th")
    st.number_input("Jumlah penduduk pertengahan tahun (jiwa) — proyeksi BPS Kab. Cirebon",
                    min_value=100_000, value=2_304_567, step=1000, key="penduduk",
                    help="Sumber: Proyeksi Penduduk BPS Kabupaten Cirebon (hasil SP2020). Dipakai untuk benchmark PK-RT = pengeluaran per kapita × penduduk.")
    st.info("ℹ️ **Mengapa tahun dasar 2025=100?** Rebasing memperbarui struktur ekonomi ke tahun referensi terbaru sehingga volume ADHK mencerminkan pola konsumsi terkini (SNA 2008). Indeks 100 berarti harga triwulan I-2025 menjadi titik nol pembanding.")

# =====================================================================
# TAHAP 2 — INPUT DATA
# =====================================================================
elif page.startswith("📥"):
    st.subheader("📥 Tahap 2 · Input Indikator & Data Dasar (Sheet E0–E5 & I6)")
    tab0, tab1, tab2, tab3, tab4 = st.tabs(
        ["💰 Basis ADHB & Indikator", "🍚 Susenas", "📈 IHK (COICOP 2018)",
         "🏭 Proxy Sektoral", "🏠 Imputasi Khusus"])

    with tab0:
        st.markdown("**Kolom pertama**: nilai ADHB 2025‑Q1 per subkomponen (juta Rp) dari lembar kerja sebelumnya / hasil SUT 2025. "
                     "**Kolom berikutnya**: laju perubahan (%) indikator volume/proxy tiap triwulan untuk ekstrapolasi.")
        cfg = {BASE_COL: st.column_config.NumberColumn(format="%.0f", min_value=0)}
        for c in GCOLS:
            cfg[c] = st.column_config.NumberColumn(format="%.2f")
        edited = st.data_editor(S["pdrb_in"], key="ed_pdrb", column_config=cfg, num_rows="fixed")
        S["pdrb_in"] = edited
        st.caption(f"Σ basis 2025‑Q1 = **{fnum(edited.iloc[:,0].sum(), 0)} juta Rp** "
                   f"(≈ {fnum(edited.iloc[:,0].sum()/1e6, 2)} triliun Rp)")
        with st.expander("💡 Bingung mengisi laju indikator? Gunakan saran proxy historis Kab. Cirebon"):
            st.markdown(
                "- **Makanan** → pertumbuhan pengeluaran per kapita Susenas + efek musim (Q1 Ramadan: +2,5–3,5%).\n"
                "- **Pakaian** → penjualan eceran sandang; melonjak menjelang Lebaran.\n"
                "- **Perumahan** → pertumbuhan pelanggan & kWh listrik PLN + UMR/sewa.\n"
                "- **Transportasi** → jumlah penumpang angkutan & kendaraan terdaftar (Samsat).\n"
                "- **Restoran & Hotel** → TPK hotel BPS + omzet usaha penyediaan makan-minum.\n"
                "- **Pendidikan** → kalender ajaran baru (lonjakan Q3).\n"
                "- Tren historis Kab. Cirebon: pertumbuhan triwulanan riil umumnya 0,5–1,5% (non-musiman).")

    with tab1:
        st.markdown("**Pengeluaran per kapita Susenas** menurut kelompok makanan & non-makanan.")
        st.info("📌 *Sumber*: BPS — **Susenas Maret** (BPS Provinsi Jabar / BPS Kab. Cirebon). "
                "*Mengapa dibutuhkan*: sebagai benchmark total PK-RT = per kapita × penduduk (QC #21).")
        cfg_s = {"Rp/Kapita/Tahun": st.column_config.NumberColumn(format="%.0f", min_value=0)}
        ed_s = st.data_editor(S["susenas_in"], key="ed_susenas", column_config=cfg_s)
        S["susenas_in"] = ed_s
        perkap = ed_s["Rp/Kapita/Tahun"].astype(float).sum()
        bench = perkap * S["penduduk"] / 1e6
        m1, m2 = st.columns(2)
        m1.metric("Total pengeluaran per kapita / tahun", f"Rp {fnum(perkap, 0)}")
        m2.metric("Benchmark PK-RT setahun (juta Rp)", fnum(bench, 0))

    with tab2:
        st.markdown("**Inflasi triwulanan (%) per 13 Divisi COICOP 2018** — basis indeks 2025‑Q1 = 100.")
        st.info("📌 *Sumber*: **Berita Resmi Statistik IHK** BPS (Kota Cirebon sebagai rujukan IHK eks-Karesidenan). "
                "*Mengapa dibutuhkan*: IHK adalah **deflator** untuk mengubah ADHB menjadi ADHK 2025. "
                "Kolom 2025‑Q1 hanya informatif (basis); penghitungan memakai 2025‑Q2 s.d. 2026‑Q4.")
        cfg_i = {c: st.column_config.NumberColumn(format="%.2f") for c in PERIODE}
        ed_i = st.data_editor(S["ihk_in"], key="ed_ihk", column_config=cfg_i)
        S["ihk_in"] = ed_i
        _, umum = hitung_level_ihk(ed_i)
        st.markdown("**Hasil: IHK Umum (2025‑Q1 = 100)**")
        st.line_chart(pd.DataFrame({"IHK Umum": umum}, index=PERIODE))

    with tab3:
        st.markdown("**Indikator volume / proxy sektoral** (referensi pendukung pengisian laju indikator).")
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("Pertumbuhan penumpang transportasi (% yoy)", value=6.5, step=0.1, key="prox_penumpang")
            st.number_input("TPK Hotel & Akomodasi (%)", value=42.3, step=0.1, key="prox_tpk")
        with c2:
            st.number_input("Pertumbuhan kWh listrik PLN terjual (%)", value=4.8, step=0.1, key="prox_kwh")
            st.number_input("Pertumbuhan PDRB Industri Pengolahan Pangan (%)", value=5.2, step=0.1, key="prox_industri")
        st.info("📌 *Sumber*: Dinas Perhubungan, Samsat (kendaraan terdaftar), BPS (TPK), **Laporan PLN** (pelanggan & kWh), Dinas Perindustrian. "
                "*Cara pakai*: terjemahkan arah & besaran indikator ini ke kolom laju (%) pada tab **Basis ADHB & Indikator**.")

    with tab4:
        st.markdown("**Imputasi khusus rebasing** — komponen yang dihitung dengan pendekatan tidak langsung.")
        c1, c2, c3 = st.columns(3)
        c1.number_input("Sewa rumah imputasi / OOD — share subkomponen perumahan (%)",
                        value=38.0, step=0.5, key="ood_pct")
        c2.number_input("FISIM (Rp miliar/triwulan)", value=210.0, step=5.0, key="fisim_miliar")
        c3.number_input("Premi-neto asuransi (Rp miliar/triwulan)", value=45.0, step=5.0, key="asuransi_miliar")
        st.info("📌 *Sumber*: tabel SUT 2025 & publikasi neraca kelembagaan BPS. "
                "*Mengapa*: OOD (rumah milik sendiri dihuni sendiri) tetap dicatat sebagai konsumsi; "
                "FISIM & asuransi dialokasikan ke PK-RT sesuai SNA 2008. Nilai-nilai ini telah tercakup dalam nilai basis ADHB.")

# =====================================================================
# TAHAP 3 — PENYESUAIAN REBASING
# =====================================================================
elif page.startswith("🧹"):
    st.subheader("🧹 Tahap 3 · Penyesuaian Metodologis Rebasing 2025")
    st.markdown("Tiga aturan statistik diterapkan otomatis. Geser toggle untuk mengaktifkan/menonaktifkan.")

    st.markdown("#### Rule 1 — BPJS Kesehatan & JKK → PKP")
    st.caption("Seluruh klaim/manfaat BPJS Kesehatan & JKK dikeluarkan dari COICOP 06 (Kesehatan PK-RT) dan dialihkan ke Pengeluaran Konsumsi Pemerintah.")
    c1, c2 = st.columns([1, 1])
    c1.checkbox("Terapkan Rule 1", value=True, key="rule_bpjs")
    c2.number_input("Nilai klaim BPJS Kes + JKK (Rp miliar/tahun)", min_value=0.0, value=185.0, step=5.0, key="bpjs_miliar")

    st.markdown("#### Rule 2 — Social Transfers in Kind (STIK) → PKP")
    st.caption("Bantuan fisik barang pemerintah kepada rumah tangga (mis. **Makan Bergizi Gratis / MBG**) dipindahkan dari PK-RT ke PKP. Pada aplikasi ini STIK mengurangi subkomponen Makanan.")
    c1, c2 = st.columns([1, 1])
    c1.checkbox("Terapkan Rule 2", value=True, key="rule_stik")
    c2.number_input("Nilai STIK/MBG (Rp miliar/tahun)", min_value=0.0, value=120.0, step=5.0, key="stik_miliar")

    st.markdown("#### Rule 3 — Rekonsiliasi COICOP 2018 → 7 Subkomponen Rilis")
    map_df = pd.DataFrame([
        ["1", "Makanan, Minuman, & Tembakau", "Divisi 01, 02"],
        ["2", "Pakaian & Alas Kaki", "Divisi 03"],
        ["3", "Perumahan, Air, Listrik, Gas & BB Lainnya", "Divisi 04"],
        ["4", "Kesehatan (Non-BPJS) & Pendidikan", "Divisi 06, 10"],
        ["5", "Transportasi, Informasi, & Komunikasi", "Divisi 07, 08"],
        ["6", "Restoran & Hotel", "Divisi 11, 12"],
        ["7", "Lainnya (Barang & Jasa Pribadi/Lainnya)", "Divisi 05, 09, 13"],
    ], columns=["No", "Subkomponen Rilis", "Divisi COICOP 2018"])
    st.dataframe(map_df, hide_index=True, use_container_width=False)

    st.markdown("#### Dampak penyesuaian pada ADHB (triwulan terakhir 2026‑Q4)")
    H = S["hasil"]
    t = 7
    rek = pd.DataFrame([
        {"Subkomponen": SUBKOMPONEN[0], "ADHB Sebelum (Jt Rp)": H["adhb_raw"][SUBKOMPONEN[0]][t],
         "Penyesuaian (Jt Rp)": -H["stik_q"], "ADHB Sesudah (Jt Rp)": H["adhb"][SUBKOMPONEN[0]][t],
         "Keterangan": "STIK/MBG → PKP" if S["rule_stik"] else "Rule 2 nonaktif"},
        {"Subkomponen": SUBKOMPONEN[3], "ADHB Sebelum (Jt Rp)": H["adhb_raw"][SUBKOMPONEN[3]][t],
         "Penyesuaian (Jt Rp)": -H["bpjs_q"], "ADHB Sesudah (Jt Rp)": H["adhb"][SUBKOMPONEN[3]][t],
         "Keterangan": "BPJS Kes & JKK → PKP" if S["rule_bpjs"] else "Rule 1 nonaktif"},
    ])
    for col in ["ADHB Sebelum (Jt Rp)", "Penyesuaian (Jt Rp)", "ADHB Sesudah (Jt Rp)"]:
        rek[col] = rek[col].apply(lambda x: fnum(x, 1))
    st.dataframe(rek, hide_index=True)

# =====================================================================
# TAHAP 4 — PENGHITUNGAN
# =====================================================================
elif page.startswith("🧮"):
    st.subheader("🧮 Tahap 4 · Penghitungan Otomatis")
    with st.expander("📐 Lihat formula yang digunakan", expanded=False):
        st.latex(r"""
        \text{ADHB}_{t} = \text{ADHB}_{t-1}\times\left(1+\frac{\%\Delta\,\text{Indikator}}{100}\right)
        \qquad
        \text{ADHK}_{t} = \frac{\text{ADHB}_{t}}{\text{Deflator}_{t}}\times 100
        """)
        st.latex(r"""
        Q\text{-to-}Q = \left(\frac{\text{ADHK}_{t}}{\text{ADHK}_{t-1}}-1\right)\!\times100\%
        \qquad
        Y\text{-on-}Y = \left(\frac{\text{ADHK}_{t}}{\text{ADHK}_{t-4}}-1\right)\!\times100\%
        """)
    hitung_semua()
    H = S["hasil"]
    t = PERIODE.index(fokus) if fokus in PERIODE else 7

    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"Total ADHB {fokus}", f"{fnum(H['total_adhb'][t], 0)} jt Rp")
    m2.metric(f"Total ADHK {fokus}", f"{fnum(H['total_adhk'][t], 0)} jt Rp",
              delta=None if H["qoq_tot"][t] is None else f"{fnum(H['qoq_tot'][t], 2)}% q-to-q")
    m3.metric("Y-on-Y", "–" if H["yoy_tot"][t] is None else f"{fnum(H['yoy_tot'][t], 2)}%")
    m4.metric("C-to-C 2026 vs 2025", f"{fnum(H['ctc'], 2)}%")

    def tabel_seri(d, dec, judul):
        df = pd.DataFrame(d, index=SUBKOMPONEN, columns=PERIODE).T
        disp = df.applymap(lambda x: fnum(x, dec))
        st.markdown(f"**{judul}**")
        st.dataframe(disp)

    tabel_seri({sk: H["adhb"][sk] for sk in SUBKOMPONEN}, 1, "ADHB — Juta Rp (setelah Rule 1 & 2)")
    tabel_seri({sk: H["defl"][sk] for sk in SUBKOMPONEN}, 2, "Deflator (2025‑Q1 = 100)")
    tabel_seri({sk: H["adhk"][sk] for sk in SUBKOMPONEN}, 1, "ADHK 2025 — Juta Rp")
    st.line_chart(pd.DataFrame({"Total ADHB": H["total_adhb"], "Total ADHK": H["total_adhk"]},
                               index=PERIODE))

# =====================================================================
# TAHAP 5 — QC 22 CHECKPOINT
# =====================================================================
elif page.startswith("🛡️"):
    st.subheader("🛡️ Tahap 5 · Quality Control Otomatis — 22 Checkpoints")
    st.text_area("Catatan fenomena khusus (opsional) — diisi jika ada kejadian luar biasa yang menjelaskan outlier, mis. 'bencana banjir Q2' atau 'program MBG mulai masif'. Jika terisi, peringatan outlier YoY dianggap tercatat.",
                 key="catatan_fenomena", height=70)
    checks = jalankan_qc()
    lulus = sum(1 for c in checks if "✅" in c["Status"])
    pering = sum(1 for c in checks if "⚠️" in c["Status"])
    gagal = sum(1 for c in checks if "❌" in c["Status"])
    m1, m2, m3 = st.columns(3)
    m1.metric("✅ Lulus", lulus)
    m2.metric("⚠️ Peringatan", pering)
    m3.metric("❌ Gagal", gagal)
    dfc = pd.DataFrame(checks)
    st.dataframe(dfc, hide_index=True)
    if pering == 0 and gagal == 0:
        st.success("🎉 Seluruh 22 checkpoint LULUS. Angka siap ditelaah untuk rilis.")
    else:
        st.warning("Terdapat temuan QC. Telusuri detail di atas — periksa kembali sumber data indikator terkait.")

# =====================================================================
# DASHBOARD HASIL (REKAP R8/R10)
# =====================================================================
else:
    st.subheader("📊 Dashboard Hasil — Rekapitulasi PK-RT (Sheet R8/R10)")
    checks = jalankan_qc()
    stat_sk = status_qc_per_sk(checks)
    rekap = buat_rekap(fokus, stat_sk)
    st.markdown(f"**Periode fokus: {fokus}** · Satuan: Juta Rupiah · Deflator 2025=100")
    st.dataframe(rekap_display(rekap), hide_index=True)

    total_status = "✅ VALID" if not any("⚠️" in c["Status"] or "❌" in c["Status"] for c in checks) else "⚠️ PERLU REVIEW"
    st.markdown(f"**Status agregat: {total_status}**")

    H = S["hasil"]
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Struktur ADHB (triwulan terakhir)**")
        share = pd.DataFrame({"Share (%)": [H["adhb"][sk][7] / H["total_adhb"][7] * 100 for sk in SUBKOMPONEN]},
                             index=NAMA_PENDEK)
        st.bar_chart(share)
    with c2:
        st.markdown("**Laju ADHK 2026‑Q4 (%)**")
        gr = pd.DataFrame({"Q-to-Q": [H["qoq"][sk][7] for sk in SUBKOMPONEN],
                           "Y-on-Y": [H["yoy"][sk][7] for sk in SUBKOMPONEN]},
                          index=NAMA_PENDEK)
        st.bar_chart(gr)

    st.markdown("**Tren Total PK-RT (2025‑Q1 s.d. 2026‑Q4)**")
    st.line_chart(pd.DataFrame({"ADHB": H["total_adhb"], "ADHK": H["total_adhk"]}, index=PERIODE))

    st.markdown("### ⬇️ Unduh Hasil")
    dl1, dl2 = st.columns(2)
    with dl1:
        csv1 = rekap.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Unduh Rekap R8/R10 (CSV)", csv1,
                           file_name="rekap_pkrt_cirebon.csv", mime="text/csv")
    with dl2:
        rows = []
        for sk in SUBKOMPONEN:
            rows.append(["ADHB", sk] + H["adhb"][sk])
            rows.append(["ADHK", sk] + H["adhk"][sk])
            rows.append(["Deflator", sk] + H["defl"][sk])
        rows.append(["ADHB", "TOTAL"] + H["total_adhb"])
        rows.append(["ADHK", "TOTAL"] + H["total_adhk"])
        dfx = pd.DataFrame(rows, columns=["Seri", "Subkomponen"] + PERIODE)
        csv2 = dfx.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Unduh Series Lengkap (CSV)", csv2,
                           file_name="series_pkrt_cirebon.csv", mime="text/csv")

    st.caption("⚠️ Aplikasi ini alat bantu hitung & QC. Angka final tetap mengikuti hasil Quality Control resmi BPS. "
               "Data contoh bersifat proxy historis Kabupaten Cirebon — ganti dengan data real Anda pada Tahap 2.")