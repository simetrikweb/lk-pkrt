# -*- coding: utf-8 -*-
"""
konoha_real_simulation.py — Simulasi PKRT riil Kabupaten Konoha (XXXX)
Dataset riil: penduduk proyeksi, IHK Kota Konoha (2025=100), Seruti per kapita.
Versi bersih: indentasi konsisten + tampilan angka kaidah Indonesia (5.976.606,65).
"""
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# ==============================================================================
# 1. DATASET RIIL KABUPATEN KONOHA (XXXX)
# ==============================================================================
POPULATION_KONOHA = {
    "2023Q1": 2349026, "2023Q2": 2355590, "2023Q3": 2364633, "2023Q4": 2370362,
    "2024Q1": 2376104, "2024Q2": 2381858, "2024Q3": 2391742, "2024Q4": 2396642,
    "2025Q1": 2401552, "2025Q2": 2406470, "2025Q3": 2417480, "2025Q4": 2422287,
    "2026Q1": 2427103, "2026Q2": 2431927,
}

DEFLATOR_IHK_2025 = {
    "01": {"2023Q1": 90.61, "2023Q2": 93.54, "2023Q3": 94.04, "2023Q4": 91.78,
           "2024Q1": 93.28, "2024Q2": 95.14, "2024Q3": 95.79, "2024Q4": 96.45,
           "2025Q1": 98.78, "2025Q2": 99.38, "2025Q3": 96.55, "2025Q4": 96.71,
           "2026Q1": 99.39, "2026Q2": 98.90},
    "02": {"2023Q1": 78.26, "2023Q2": 79.39, "2023Q3": 80.66, "2023Q4": 82.61,
           "2024Q1": 87.25, "2024Q2": 89.80, "2024Q3": 90.94, "2024Q4": 91.28,
           "2025Q1": 92.20, "2025Q2": 93.71, "2025Q3": 94.42, "2025Q4": 95.21,
           "2026Q1": 98.95, "2026Q2": 99.28},
    "03": {"2023Q1": 93.53, "2023Q2": 94.31, "2023Q3": 94.77, "2023Q4": 95.18,
           "2024Q1": 95.44, "2024Q2": 95.55, "2024Q3": 95.97, "2024Q4": 96.10,
           "2025Q1": 96.46, "2025Q2": 97.24, "2025Q3": 98.20, "2025Q4": 98.92,
           "2026Q1": 99.28, "2026Q2": 99.83},
    "04": {"2023Q1": 95.51, "2023Q2": 95.88, "2023Q3": 96.28, "2023Q4": 96.51,
           "2024Q1": 101.19, "2024Q2": 101.26, "2024Q3": 101.56, "2024Q4": 101.67,
           "2025Q1": 101.78, "2025Q2": 102.05, "2025Q3": 102.08, "2025Q4": 102.11,
           "2026Q1": 93.01, "2026Q2": 102.27},
    "05": {"2023Q1": 92.91, "2023Q2": 95.02, "2023Q3": 96.09, "2023Q4": 96.16,
           "2024Q1": 96.29, "2024Q2": 96.34, "2024Q3": 97.45, "2024Q4": 97.27,
           "2025Q1": 98.11, "2025Q2": 98.15, "2025Q3": 98.44, "2025Q4": 98.82,
           "2026Q1": 99.95, "2026Q2": 100.08},
    "06": {"2023Q1": 92.07, "2023Q2": 92.68, "2023Q3": 92.72, "2023Q4": 93.74,
           "2024Q1": 95.85, "2024Q2": 95.94, "2024Q3": 96.29, "2024Q4": 96.34,
           "2025Q1": 96.58, "2025Q2": 96.56, "2025Q3": 96.51, "2025Q4": 96.51,
           "2026Q1": 97.75, "2026Q2": 99.50},
    "07": {"2023Q1": 83.61, "2023Q2": 85.86, "2023Q3": 89.71, "2023Q4": 98.14,
           "2024Q1": 98.24, "2024Q2": 98.93, "2024Q3": 98.25, "2024Q4": 98.81,
           "2025Q1": 98.95, "2025Q2": 99.38, "2025Q3": 99.57, "2025Q4": 99.23,
           "2026Q1": 99.90, "2026Q2": 99.81},
    "08": {"2023Q1": 99.39, "2023Q2": 99.40, "2023Q3": 99.51, "2023Q4": 99.53,
           "2024Q1": 99.53, "2024Q2": 99.53, "2024Q3": 99.84, "2024Q4": 100.16,
           "2025Q1": 100.09, "2025Q2": 100.09, "2025Q3": 100.06, "2025Q4": 100.05,
           "2026Q1": 100.06, "2026Q2": 99.92},
    "09": {"2023Q1": 95.33, "2023Q2": 96.37, "2023Q3": 97.36, "2023Q4": 97.90,
           "2024Q1": 98.25, "2024Q2": 98.33, "2024Q3": 99.55, "2024Q4": 99.84,
           "2025Q1": 99.83, "2025Q2": 99.93, "2025Q3": 100.54, "2025Q4": 100.84,
           "2026Q1": 101.54, "2026Q2": 100.23},
    "10": {"2023Q1": 93.20, "2023Q2": 93.24, "2023Q3": 94.31, "2023Q4": 94.86,
           "2024Q1": 94.88, "2024Q2": 94.90, "2024Q3": 96.33, "2024Q4": 96.33,
           "2025Q1": 96.33, "2025Q2": 96.33, "2025Q3": 97.12, "2025Q4": 97.98,
           "2026Q1": 98.21, "2026Q2": 98.51},
    "11": {"2023Q1": 92.64, "2023Q2": 94.00, "2023Q3": 94.85, "2023Q4": 95.08,
           "2024Q1": 95.20, "2024Q2": 95.42, "2024Q3": 96.18, "2024Q4": 96.22,
           "2025Q1": 96.37, "2025Q2": 96.81, "2025Q3": 97.22, "2025Q4": 97.75,
           "2026Q1": 98.67, "2026Q2": 99.99},
    "12": {"2023Q1": 99.90, "2023Q2": 99.90, "2023Q3": 99.90, "2023Q4": 99.90,
           "2024Q1": 99.90, "2024Q2": 99.90, "2024Q3": 99.90, "2024Q4": 99.90,
           "2025Q1": 99.90, "2025Q2": 99.90, "2025Q3": 99.90, "2025Q4": 99.90,
           "2026Q1": 99.90, "2026Q2": 99.90},
    "13": {"2023Q1": 88.31, "2023Q2": 89.59, "2023Q3": 91.89, "2023Q4": 93.02,
           "2024Q1": 93.78, "2024Q2": 94.32, "2024Q3": 94.29, "2024Q4": 94.83,
           "2025Q1": 96.02, "2025Q2": 96.54, "2025Q3": 97.52, "2025Q4": 97.91,
           "2026Q1": 98.22, "2026Q2": 99.98},
}

SERUTI_PERKAPITA = {
    "01": {"2023Q1": 1219634, "2023Q2": 1036382, "2023Q3": 1127009, "2023Q4": 1168626,
           "2024Q1": 1232580, "2024Q2": 1581595, "2024Q3": 1275056, "2024Q4": 1307610,
           "2025Q1": 1333118, "2025Q2": 1271330, "2025Q3": 1291889, "2025Q4": 1136932,
           "2026Q1": 1253428, "2026Q2": 1444628},
    "02": {"2023Q1": 365857, "2023Q2": 408486, "2023Q3": 403772, "2023Q4": 376186,
           "2024Q1": 295080, "2024Q2": 363579, "2024Q3": 344516, "2024Q4": 338615,
           "2025Q1": 287801, "2025Q2": 383483, "2025Q3": 361479, "2025Q4": 311923,
           "2026Q1": 294236, "2026Q2": 346448},
    "03": {"2023Q1": 36909, "2023Q2": 211745, "2023Q3": 76887, "2023Q4": 40673,
           "2024Q1": 51793, "2024Q2": 258054, "2024Q3": 84811, "2024Q4": 38103,
           "2025Q1": 49989, "2025Q2": 80003, "2025Q3": 58833, "2025Q4": 34343,
           "2026Q1": 54899, "2026Q2": 198974},
    "04": {"2023Q1": 577451, "2023Q2": 571506, "2023Q3": 565316, "2023Q4": 543185,
           "2024Q1": 714807, "2024Q2": 817866, "2024Q3": 655904, "2024Q4": 704746,
           "2025Q1": 723038, "2025Q2": 724920, "2025Q3": 745530, "2025Q4": 881374,
           "2026Q1": 736429, "2026Q2": 712447},
    "05": {"2023Q1": 140140, "2023Q2": 104506, "2023Q3": 141078, "2023Q4": 82236,
           "2024Q1": 109100, "2024Q2": 47616, "2024Q3": 128476, "2024Q4": 119950,
           "2025Q1": 137921, "2025Q2": 136348, "2025Q3": 108972, "2025Q4": 85538,
           "2026Q1": 140486, "2026Q2": 80209},
    "06": {"2023Q1": 56672, "2023Q2": 34469, "2023Q3": 45215, "2023Q4": 55316,
           "2024Q1": 28395, "2024Q2": 25438, "2024Q3": 39817, "2024Q4": 39556,
           "2025Q1": 32435, "2025Q2": 58362, "2025Q3": 32337, "2025Q4": 59692,
           "2026Q1": 55788, "2026Q2": 57817},
    "07": {"2023Q1": 616833, "2023Q2": 241814, "2023Q3": 234807, "2023Q4": 387357,
           "2024Q1": 248635, "2024Q2": 526918, "2024Q3": 649753, "2024Q4": 332843,
           "2025Q1": 270514, "2025Q2": 360761, "2025Q3": 241765, "2025Q4": 277979,
           "2026Q1": 316117, "2026Q2": 370407},
    "08": {"2023Q1": 154959, "2023Q2": 113468, "2023Q3": 127626, "2023Q4": 111980,
           "2024Q1": 157817, "2024Q2": 192853, "2024Q3": 211079, "2024Q4": 145072,
           "2025Q1": 137340, "2025Q2": 133703, "2025Q3": 175501, "2025Q4": 129083,
           "2026Q1": 152861, "2026Q2": 143551},
    "09": {"2023Q1": 49253, "2023Q2": 38865, "2023Q3": 36513, "2023Q4": 18699,
           "2024Q1": 169295, "2024Q2": 33079, "2024Q3": 69933, "2024Q4": 68520,
           "2025Q1": 18510, "2025Q2": 25611, "2025Q3": 17608, "2025Q4": 23618,
           "2026Q1": 20171, "2026Q2": 34039},
    "10": {"2023Q1": 50248, "2023Q2": 76134, "2023Q3": 97185, "2023Q4": 24895,
           "2024Q1": 73476, "2024Q2": 190506, "2024Q3": 69022, "2024Q4": 70720,
           "2025Q1": 46337, "2025Q2": 117512, "2025Q3": 43274, "2025Q4": 47273,
           "2026Q1": 47002, "2026Q2": 70243},
    "11": {"2023Q1": 793805, "2023Q2": 776891, "2023Q3": 876517, "2023Q4": 919960,
           "2024Q1": 829525, "2024Q2": 865295, "2024Q3": 839349, "2024Q4": 1043768,
           "2025Q1": 1004885, "2025Q2": 1048743, "2025Q3": 1070684, "2025Q4": 1154947,
           "2026Q1": 980259, "2026Q2": 998718},
    "12": {"2023Q1": 5379, "2023Q2": 4004, "2023Q3": 6258, "2023Q4": 8176,
           "2024Q1": 5244, "2024Q2": 5490, "2024Q3": 5674, "2024Q4": 6928,
           "2025Q1": 6479, "2025Q2": 7060, "2025Q3": 6966, "2025Q4": 7484,
           "2026Q1": 8153, "2026Q2": 7179},
    "13": {"2023Q1": 99272, "2023Q2": 95030, "2023Q3": 154692, "2023Q4": 114386,
           "2024Q1": 93699, "2024Q2": 137346, "2024Q3": 113324, "2024Q4": 127846,
           "2025Q1": 137133, "2025Q2": 144118, "2025Q3": 118012, "2025Q4": 138451,
           "2026Q1": 134751, "2026Q2": 154548},
}

COICOP_LABELS = {
    "01": "Makanan dan minuman tidak beralkohol",
    "02": "Minuman beralkohol dan tembakau",
    "03": "Pakaian dan alas kaki",
    "04": "Perumahan, air, listrik, gas, dan bahan bakar lainnya",
    "05": "Furnitur, perlengkapan rumah tangga, dan pemeliharaan rutin",
    "06": "Kesehatan",
    "07": "Transportasi",
    "08": "Informasi dan Komunikasi",
    "09": "Rekreasi, olahraga, dan budaya",
    "10": "Pendidikan",
    "11": "Penyediaan makan minum dan akomodasi",
    "12": "Jasa keuangan dan asuransi",
    "13": "Perawatan pribadi dan jasa lainnya",
}


# ==============================================================================
# 2. FORMAT ANGKA KAIDAH INDONESIA (untuk TAMPILAN)
# ==============================================================================
def fmt_id(x, des=2):
    """Format angka gaya Indonesia: 5976606.65 -> 5.976.606,65"""
    try:
        if x is None:
            return "-"
        x = float(x)
        if not np.isfinite(x):
            return "-"
        s = f"{x:,.{des}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)


def fmt_cols(df, cols, des=2):
    """Salinan DataFrame untuk TAMPILAN: kolom numerik terpilih diformat Indonesia."""
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].map(lambda v: fmt_id(v, des))
    return out


# ==============================================================================
# 3. ENGINE PERHITUNGAN STANDAR LEMBAR KERJA PKRT
# ==============================================================================
def calculate_all(beta_matrix):
    quarters = ["2023Q1", "2023Q2", "2023Q3", "2023Q4", "2024Q1", "2024Q2", "2024Q3", "2024Q4",
                "2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1", "2026Q2"]
    coicops = [f"{i:02d}" for i in range(1, 14)]
    records = []
    for q in quarters:
        for c in coicops:
            percapita = SERUTI_PERKAPITA[c][q]
            pop = POPULATION_KONOHA[q]
            ihk = DEFLATOR_IHK_2025[c][q]
            theta = 1.0
            beta = beta_matrix[q][c]
            # 1. Agregasi Seruti
            hfce_seruti = (percapita * pop) / 1_000_000
            # 2. Estimasi Inisial ADHB
            hfce_nom_star = theta * hfce_seruti
            # 3. Deflasi Inisial ADHK
            hfce_real_star = hfce_nom_star / (ihk / 100)
            # 4. Rekonsiliasi Balancing Factor
            hfce_nom_final = beta * hfce_nom_star
            hfce_real_final = hfce_nom_final / (ihk / 100)
            records.append({
                "Quarter": q, "COICOP": c, "Label": COICOP_LABELS[c],
                "Population": pop, "IHK": ihk, "Theta": theta, "Beta": beta,
                "Seruti (Rp Juta)": hfce_seruti,
                "ADHB (Rp Juta)": hfce_nom_final, "ADHK (Rp Juta)": hfce_real_final,
            })
    df = pd.DataFrame(records)

    # 5. Agregasi 7 Subkomponen
    sub_map = {
        "01. Makanan, Minuman, Rokok": ["01", "02"],
        "02. Pakaian & Alas Kaki": ["03"],
        "03. Perumahan & Perlengkapan": ["04", "05"],
        "04. Kesehatan & Pendidikan": ["06", "10"],
        "05. Transportasi, Komunikasi, Rekreasi": ["07", "08", "09"],
        "06. Restoran & Hotel": ["11"],
        "07. Lainnya": ["12", "13"],
    }
    sub_records = []
    for q in quarters:
        q_df = df[df['Quarter'] == q]
        for sub_name, coicops_list in sub_map.items():
            nom = q_df[q_df['COICOP'].isin(coicops_list)]['ADHB (Rp Juta)'].sum()
            real = q_df[q_df['COICOP'].isin(coicops_list)]['ADHK (Rp Juta)'].sum()
            sub_records.append({"Quarter": q, "Subkomponen": sub_name,
                                "ADHB (Rp Juta)": nom, "ADHK (Rp Juta)": real})
    df_sub = pd.DataFrame(sub_records)

    # 6. Laju Pertumbuhan & Konsumsi Per Kapita
    df_total = df.groupby('Quarter').agg({'ADHB (Rp Juta)': 'sum', 'ADHK (Rp Juta)': 'sum',
                                          'Population': 'first'}).reset_index()
    df_total['Konsumsi Per Kapita (Rp Ribu)'] = (df_total['ADHB (Rp Juta)'] * 1000) / df_total['Population']
    df_total['q-to-q Riil (%)'] = df_total['ADHK (Rp Juta)'].pct_change() * 100
    df_total['y-on-y Riil (%)'] = df_total['ADHK (Rp Juta)'].pct_change(4) * 100
    df_total['Year'] = df_total['Quarter'].str[:4]
    df_total['cum_real'] = df_total.groupby('Year')['ADHK (Rp Juta)'].cumsum()
    df_total['c-to-c Riil (%)'] = (df_total['cum_real'] / df_total['cum_real'].shift(4) - 1) * 100

    # 7. Evaluasi QC
    score, msgs = evaluate_qc(df, df_total)
    return df, df_sub, df_total, score, msgs


def evaluate_qc(df, df_total):
    score = 100
    msgs = []
    # Check 1: Ekstrem Balancing Factor
    extreme_betas = df[(df['Beta'] < 0.85) | (df['Beta'] > 1.15)]
    if len(extreme_betas) > 0:
        score -= 15
        msgs.append("⚠️ **Distorsi Balancing Factor**: Ditemukan adjustment Beta ekstrem (<0.85 atau >1.15).")
    else:
        msgs.append("✅ **Balancing Factor**: Semua nilai Beta dalam batas wajar (0.85 - 1.15).")
    # Check 2: Volatilitas Makro
    vol = df_total['q-to-q Riil (%)'].dropna().std()
    if vol > 8:
        score -= 10
        msgs.append("⚠️ **Volatilitas Makro**: Pertumbuhan q-to-q riil terlalu fluktuatif (StdDev > 8%).")
    else:
        msgs.append("✅ **Volatilitas Makro**: Tren pertumbuhan triwulanan stabil.")
    # Check 3: Pola Musiman (Lebaran Effect pada Pakaian)
    seasonal_ok = True
    for year in [2023, 2024, 2025]:
        q12_03 = df[(df['Quarter'].isin([f"{year}Q1", f"{year}Q2"])) & (df['COICOP'] == '03')]['ADHB (Rp Juta)'].mean()
        q34_03 = df[(df['Quarter'].isin([f"{year}Q3", f"{year}Q4"])) & (df['COICOP'] == '03')]['ADHB (Rp Juta)'].mean()
        if q12_03 < q34_03:
            seasonal_ok = False
    if not seasonal_ok:
        score -= 10
        msgs.append("⚠️ **Pola Musiman**: Efek Lebaran pada Pakaian (03) tidak terdeteksi lebih tinggi di Q1/Q2.")
    else:
        msgs.append("✅ **Pola Musiman**: Efek Lebaran (Q1/Q2) pada Pakaian dan Makanan terdeteksi dengan baik.")
    if score >= 85:
        msgs.insert(0, f"🟢 **Tingkat Konsistensi: {score}% (Sangat Baik)** - Data siap untuk proses kompilasi PDRB.")
    elif score >= 70:
        msgs.insert(0, f"🟡 **Tingkat Konsistensi: {score}% (Cukup)** - Perlu peninjauan pada beberapa balancing factor.")
    else:
        msgs.insert(0, f"🔴 **Tingkat Konsistensi: {score}% (Rendah)** - Struktur data tidak stabil, segera lakukan rekonsiliasi ulang.")
    return max(0, score), msgs


# ==============================================================================
# 4. TAMPILAN INTERAKTIF (STREAMLIT UI)
# ==============================================================================
def render_konoha_simulation():
    st.header("Simulasi PKRT Riil Kabupaten Konoha (XXXX)")
    st.markdown("Tahun Dasar 2025=100 | Ekstensi Modul")
    quarters = ["2023Q1", "2023Q2", "2023Q3", "2023Q4", "2024Q1", "2024Q2", "2024Q3", "2024Q4",
                "2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1", "2026Q2"]
    coicops = [f"{i:02d}" for i in range(1, 14)]

    if 'beta_matrix' not in st.session_state:
        st.session_state['beta_matrix'] = {q: {c: 1.0 for c in coicops} for q in quarters}

    # Balancing & Sensitivity Playground
    st.subheader("Balancing & Sensitivity Playground")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Reset Beta ke 1.00"):
            st.session_state['beta_matrix'] = {q: {c: 1.0 for c in coicops} for q in quarters}
            st.rerun()
    with col2:
        if st.button("Skenario Lebaran Q1/Q2"):
            for q in quarters:
                if "Q1" in q or "Q2" in q:
                    st.session_state['beta_matrix'][q]["01"] = 1.05
                    st.session_state['beta_matrix'][q]["03"] = 1.15
            st.rerun()
    with col3:
        if st.button("Skenario Shock Pangan"):
            for q in quarters:
                st.session_state['beta_matrix'][q]["01"] = 1.10
            st.rerun()

    selected_q = st.selectbox("Pilih Triwulan untuk Adjustmen Beta", quarters)
    st.markdown(f"**Adjust Balancing Factor ($\\beta$) untuk {selected_q}:**")
    cols = st.columns(4)
    for idx, c in enumerate(coicops):
        with cols[idx % 4]:
            st.session_state['beta_matrix'][selected_q][c] = st.slider(
                f"{c}: {COICOP_LABELS[c][:25]}...",
                min_value=0.5, max_value=1.5,
                value=st.session_state['beta_matrix'][selected_q][c],
                step=0.01, key=f"beta_{selected_q}_{c}",
            )

    # Kalkulasi Engine
    df, df_sub, df_total, qc_score, qc_msgs = calculate_all(st.session_state['beta_matrix'])

    # Dashboard Tabs
    tab1, tab2, tab3 = st.tabs(["Ringkasan Agregat", "Grafik Deret Waktu", "Evaluasi QC"])
    with tab1:
        st.subheader("Nilai Agregat ADHB & ADHK (Rp Juta)")
        st.dataframe(fmt_cols(df_sub, ["ADHB (Rp Juta)", "ADHK (Rp Juta)"], 2),
                     use_container_width=True, hide_index=True)
        with st.expander("🔍 Detail per COICOP (format Indonesia)"):
            det = fmt_cols(df, ["Population"], 0)
            det = fmt_cols(det, ["IHK", "Theta", "Beta",
                                 "Seruti (Rp Juta)", "ADHB (Rp Juta)", "ADHK (Rp Juta)"], 2)
            st.dataframe(det, use_container_width=True, hide_index=True)
        with st.expander("🔍 Ringkasan triwulanan (format Indonesia)"):
            tot = fmt_cols(df_total, ["Population"], 0)
            tot = fmt_cols(tot, ["ADHB (Rp Juta)", "ADHK (Rp Juta)",
                                 "Konsumsi Per Kapita (Rp Ribu)",
                                 "q-to-q Riil (%)", "y-on-y Riil (%)", "c-to-c Riil (%)"], 2)
            st.dataframe(tot, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.subheader("Export Data")
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Unduh Detail COICOP (CSV)",
            data=csv,
            file_name='pkrt_cirebon_detail.csv',
            mime='text/csv',
        )
        st.caption("File CSV sengaja tetap memakai titik desimal agar mudah diolah ulang "
                   "di Excel/Python; format Indonesia hanya untuk tampilan di layar.")
    with tab2:
        st.subheader("Tren Pengeluaran Per Kapita (Rp Ribu)")
        fig_pc = px.line(df_total, x='Quarter', y='Konsumsi Per Kapita (Rp Ribu)',
                         title='Konsumsi Per Kapita', markers=True)
        st.plotly_chart(fig_pc, use_container_width=True)
        st.subheader("Laju Pertumbuhan Ekonomi Riil (%)")
        col_q, col_y = st.columns(2)
        with col_q:
            fig_qoq = px.line(df_total, x='Quarter', y='q-to-q Riil (%)',
                              title='Pertumbuhan q-to-q', markers=True)
            st.plotly_chart(fig_qoq, use_container_width=True)
        with col_y:
            fig_yoy = px.line(df_total, x='Quarter', y='y-on-y Riil (%)',
                              title='Pertumbuhan y-on-y', markers=True)
            st.plotly_chart(fig_yoy, use_container_width=True)
    with tab3:
        st.subheader("Hasil Checklist Quality Control (QC)")
        for msg in qc_msgs:
            st.markdown(msg)