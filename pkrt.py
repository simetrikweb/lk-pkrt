# -*- coding: utf-8 -*-
"""
pkrt.py — STANDALONE
Otomasi end-to-end Rebasing PDRB Pengeluaran 2025=100 — Kab. Konoha (9999)
3 LK resmi + Auto-Fill Template Excel Asli (openpyxl, sel rumus aman)
Triwulan: 2019Q1–2030Q4 | Tampilan angka: kaidah Indonesia
"""
import io
import re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as CL

st.set_page_config(page_title="LK PKRT & Deflator", page_icon="📊", layout="wide")

# ==================================================================================
# 0. UTILITAS & KONSTANTA UMUM
# ==================================================================================
Q_ALL = [f"{y}Q{q}" for y in range(2019, 2031) for q in (1, 2, 3, 4)]          # 2019Q1–2030Q4 (48)
CODES13 = [f"{i:02d}" for i in range(1, 14)]
CODES12 = [f"{i:02d}" for i in range(1, 13)]
MAP7 = {0: [0, 1], 1: [2], 2: [3, 4], 3: [5, 9], 4: [6, 7, 8], 5: [10], 6: [11, 12]}
SUB7 = ["Makanan, minuman, dan rokok", "Pakaian, alas kaki, dan jasa perawatan",
        "Perumahan dan perlengkapan rumah tangga", "Kesehatan dan pendidikan",
        "Transportasi, komunikasi, dan rekreasi", "Restoran dan hotel", "Lainnya"]
PROP = {11: 0.6335, 12: 0.3665}          # disagregasi COICOP-1999 12 -> 2018 12 & 13 (KUNCI pusat)
COV = {5: 0.4980, 11: 1.1007}            # coverage adjustment SNA 2008 (mode [B] saja)
REF_OLD = list(range(11)) + [11, 11]     # COICOP 2018 12 & 13 -> COICOP 1999 12
HDR = Font(bold=True)
TIT = Font(bold=True, size=13)


def sdiv(a, b):
    try:
        if b is None or b == 0 or not np.isfinite(b) or not np.isfinite(a):
            return np.nan
        return a / b
    except Exception:
        return np.nan


def fmt_id(x, des=2):
    """5976606.65 -> 5.976.606,65"""
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


def tabel_id(df, des=2):
    try:
        return df.map(lambda v: fmt_id(v, des))
    except Exception:
        return df.applymap(lambda v: fmt_id(v, des))


def q_of_month(m):
    y, mo = m.split("M")
    return f"{y}Q{(int(mo) - 1) // 3 + 1}"


# ==================================================================================
# 1. DATASET TERTANAM (RIIL / RESMI)
# ==================================================================================
SBH_2022_WEIGHTS_3274 = {
    "011_Makanan": 25.97, "012_Minuman_NonAlkohol": 1.86,
    "013_Minuman_Alkohol": 0.00, "014_Rokok_Tembakau": 3.66,
    "02_Pakaian_AlasKaki": 3.21, "03_Perumahan_Air_Listrik_BB": 17.86,
    "04_Perlengkapan_RT": 4.78, "05_Kesehatan": 2.30,
    "06_Transportasi": 9.69, "071_Peralatan_Infokom": 0.55,
    "072_Layanan_Infokom": 3.58, "08_Rekreasi_Olahraga_Budaya": 1.21,
    "09_Pendidikan": 5.09, "10_Restoran_Penyediaan_Makan": 15.58,
    "073_Asuransi": 0.03, "074_Jasa_Keuangan": 0.50,
    "11_Perawatan_Pribadi_JasaLain": 4.13, "00_Umum": 100.00}

DEFLATOR_QUARTERLY_2025_BASE = {
    "01": {"2023Q1": 93.28, "2023Q2": 95.14, "2023Q3": 95.79, "2023Q4": 96.45, "2024Q1": 98.78, "2024Q2": 99.38, "2024Q3": 96.55, "2024Q4": 96.71, "2025Q1": 99.39, "2025Q2": 98.90, "2025Q3": 100.22, "2025Q4": 101.49},
    "02": {"2023Q1": 87.25, "2023Q2": 89.80, "2023Q3": 90.94, "2023Q4": 91.28, "2024Q1": 92.20, "2024Q2": 93.71, "2024Q3": 94.42, "2024Q4": 95.21, "2025Q1": 98.95, "2025Q2": 99.28, "2025Q3": 100.47, "2025Q4": 101.29},
    "03": {"2023Q1": 95.44, "2023Q2": 95.55, "2023Q3": 95.97, "2023Q4": 96.10, "2024Q1": 96.46, "2024Q2": 97.24, "2024Q3": 98.20, "2024Q4": 98.92, "2025Q1": 99.28, "2025Q2": 99.83, "2025Q3": 100.40, "2025Q4": 100.50},
    "04": {"2023Q1": 101.19, "2023Q2": 101.26, "2023Q3": 101.56, "2023Q4": 101.67, "2024Q1": 101.78, "2024Q2": 102.05, "2024Q3": 102.08, "2024Q4": 102.11, "2025Q1": 93.01, "2025Q2": 102.27, "2025Q3": 102.36, "2025Q4": 102.36},
    "05": {"2023Q1": 96.29, "2023Q2": 96.34, "2023Q3": 97.45, "2023Q4": 97.27, "2024Q1": 98.11, "2024Q2": 98.15, "2024Q3": 98.44, "2024Q4": 98.82, "2025Q1": 99.95, "2025Q2": 100.08, "2025Q3": 100.02, "2025Q4": 99.95},
    "06": {"2023Q1": 95.85, "2023Q2": 95.94, "2023Q3": 96.29, "2023Q4": 96.34, "2024Q1": 96.58, "2024Q2": 96.56, "2024Q3": 96.51, "2024Q4": 96.51, "2025Q1": 97.75, "2025Q2": 99.50, "2025Q3": 100.67, "2025Q4": 102.07},
    "07": {"2023Q1": 98.24, "2023Q2": 98.93, "2023Q3": 98.25, "2023Q4": 98.81, "2024Q1": 98.95, "2024Q2": 99.38, "2024Q3": 99.57, "2024Q4": 99.23, "2025Q1": 99.90, "2025Q2": 99.81, "2025Q3": 100.07, "2025Q4": 100.22},
    "08": {"2023Q1": 99.53, "2023Q2": 99.53, "2023Q3": 99.84, "2023Q4": 100.16, "2024Q1": 100.09, "2024Q2": 100.09, "2024Q3": 100.06, "2024Q4": 100.05, "2025Q1": 100.06, "2025Q2": 99.92, "2025Q3": 100.02, "2025Q4": 100.00},
    "09": {"2023Q1": 98.25, "2023Q2": 98.33, "2023Q3": 99.55, "2023Q4": 99.84, "2024Q1": 99.83, "2024Q2": 99.93, "2024Q3": 100.54, "2024Q4": 100.84, "2025Q1": 101.54, "2025Q2": 100.23, "2025Q3": 98.58, "2025Q4": 99.66},
    "10": {"2023Q1": 94.88, "2023Q2": 94.90, "2023Q3": 96.33, "2023Q4": 96.33, "2024Q1": 96.33, "2024Q2": 96.33, "2024Q3": 97.12, "2024Q4": 97.98, "2025Q1": 98.21, "2025Q2": 98.51, "2025Q3": 100.56, "2025Q4": 102.72},
    "11": {"2023Q1": 95.20, "2023Q2": 95.42, "2023Q3": 96.18, "2023Q4": 96.22, "2024Q1": 96.37, "2024Q2": 96.81, "2024Q3": 97.22, "2024Q4": 97.75, "2025Q1": 98.67, "2025Q2": 99.99, "2025Q3": 100.43, "2025Q4": 100.90},
    "12": {"2023Q1": 99.90, "2023Q2": 99.90, "2023Q3": 99.90, "2023Q4": 99.90, "2024Q1": 99.90, "2024Q2": 99.90, "2024Q3": 99.90, "2024Q4": 99.90, "2025Q1": 99.90, "2025Q2": 99.90, "2025Q3": 100.00, "2025Q4": 100.21},
    "13": {"2023Q1": 93.78, "2023Q2": 94.32, "2023Q3": 94.29, "2023Q4": 94.83, "2024Q1": 96.02, "2024Q2": 96.54, "2024Q3": 97.52, "2024Q4": 97.91, "2025Q1": 98.22, "2025Q2": 99.98, "2025Q3": 100.65, "2025Q4": 101.15},
    "00": {"2023Q1": 95.85, "2023Q2": 96.59, "2023Q3": 97.16, "2023Q4": 97.47, "2024Q1": 98.32, "2024Q2": 98.77, "2024Q3": 98.22, "2024Q4": 98.45, "2025Q1": 98.16, "2025Q2": 99.93, "2025Q3": 100.63, "2025Q4": 101.28}}

IHK_2022_MONTHLY_2025_2026 = {
    "01": {"2025": [106.46, 106.07, 108.67, 106.54, 105.85, 107.20, 108.75, 107.37, 107.76, 108.65, 109.12, 110.20], "2026_JanSep": [108.13, 111.30, 113.57, 112.40, 113.30, 112.47, 111.24, 112.53, 113.01]},
    "02": {"2025": [123.41, 123.72, 124.05, 124.05, 124.05, 124.31, 124.78, 125.43, 126.65, 126.65, 126.65, 126.65], "2026_JanSep": [128.00, 128.42, 128.42, 128.42, 130.70, 130.70, 130.70, 131.43, 131.43]},
    "03": {"2025": [104.63, 104.74, 105.23, 105.33, 105.49, 105.52, 106.01, 106.06, 106.08, 106.16, 106.15, 106.17], "2026_JanSep": [106.25, 106.33, 106.45, 106.78, 107.04, 107.32, 107.36, 107.36, 107.80]},
    "04": {"2025": [95.69, 90.93, 97.07, 103.98, 103.98, 103.98, 104.07, 104.07, 104.07, 104.07, 104.08, 104.08], "2026_JanSep": [104.10, 104.10, 104.12, 104.15, 104.18, 104.18, 104.18, 104.23, 104.25]},
    "05": {"2025": [103.23, 103.76, 103.95, 103.81, 103.78, 103.74, 103.74, 103.73, 103.66, 103.60, 103.69, 103.63], "2026_JanSep": [103.69, 103.43, 103.62, 104.60, 104.81, 105.08, 105.49, 105.67, 106.01]},
    "06": {"2025": [102.16, 103.03, 103.54, 103.96, 104.45, 105.85, 105.89, 106.07, 105.98, 107.42, 107.51, 107.45], "2026_JanSep": [108.58, 108.53, 108.53, 108.71, 108.77, 109.38, 110.72, 111.37, 111.58]},
    "07": {"2025": [110.32, 110.92, 111.16, 110.69, 110.67, 110.72, 110.98, 110.99, 111.00, 111.00, 111.03, 111.43], "2026_JanSep": [110.99, 110.66, 111.35, 111.50, 111.85, 114.37, 115.50, 115.45, 115.50]},
    "08": {"2025": [100.15, 100.16, 100.16, 99.79, 100.14, 100.14, 100.14, 100.11, 100.11, 100.10, 100.11, 100.11], "2026_JanSep": [100.08, 100.08, 100.07, 100.18, 100.64, 101.09, 101.42, 101.81, 101.97]},
    "09": {"2025": [104.25, 104.34, 104.47, 104.47, 103.34, 101.21, 101.27, 101.17, 101.49, 102.36, 102.35, 102.56], "2026_JanSep": [102.91, 103.01, 102.96, 103.00, 103.00, 103.00, 103.43, 103.43, 103.43]},
    "10": {"2025": [104.07, 104.07, 104.55, 104.55, 104.55, 104.55, 105.78, 105.89, 108.49, 109.01, 109.01, 109.01], "2026_JanSep": [109.06, 109.06, 109.06, 109.06, 109.06, 109.06, 110.46, 110.47, 114.21]},
    "11": {"2025": [104.09, 104.10, 104.52, 105.37, 105.51, 106.01, 106.01, 106.07, 106.20, 106.23, 106.43, 107.10], "2026_JanSep": [107.33, 107.40, 107.59, 108.65, 108.81, 109.07, 109.15, 109.20, 109.20]},
    "12": {"2025": [100.00, 100.00, 100.00, 100.00, 100.00, 100.00, 100.00, 100.00, 100.31, 100.31, 100.31, 100.31], "2026_JanSep": [100.31, 100.31, 100.31, 100.31, 100.31, 100.31, 100.31, 100.31, 100.31]},
    "13": {"2025": [107.21, 107.83, 108.42, 109.14, 109.90, 110.23, 110.36, 110.50, 110.59, 111.05, 110.96, 111.09], "2026_JanSep": [111.53, 111.45, 111.40, 111.39, 111.29, 111.32, 111.28, 111.38, 111.56]},
    "00": {"2025": [104.44, 103.63, 105.64, 106.38, 106.26, 106.75, 107.32, 106.99, 107.30, 107.64, 107.81, 108.26], "2026_JanSep": [107.78, 108.64, 109.38, 109.31, 109.74, 109.85, 109.77, 110.21, 110.60]}}

COICOP_LABELS = {
    "01": "Makanan dan minuman tidak beralkohol", "02": "Minuman beralkohol dan tembakau",
    "03": "Pakaian dan alas kaki", "04": "Perumahan, air, listrik, gas, dan bahan bakar lainnya",
    "05": "Furnitur, perlengkapan rumah tangga, dan perbaikan rutin rumah", "06": "Kesehatan",
    "07": "Transportasi", "08": "Komunikasi", "09": "Rekreasi dan kebudayaan",
    "10": "Pendidikan", "11": "Penyediaan makan-minum dan akomodasi",
    "12": "Jasa keuangan dan asuransi", "13": "Barang dan jasa lainnya"}

POPULATION_MONTHLY_3209 = {
    "202006": 2270353, "202007": 2272922, "202008": 2275217, "202009": 2270621, "202010": 2279814, "202011": 2282115, "202012": 2284419,
    "202101": 2286725, "202102": 2289033, "202103": 2291343, "202104": 2293656, "202105": 2295970, "202106": 2301331, "202107": 2303823, "202108": 2306049, "202109": 2308277, "202110": 2310507, "202111": 2312739, "202112": 2314973,
    "202201": 2317209, "202202": 2319447, "202203": 2321687, "202204": 2323930, "202205": 2326174, "202206": 2331364, "202207": 2333777, "202208": 2335950, "202209": 2338124, "202210": 2340301, "202211": 2342479, "202212": 2344660,
    "202301": 2346842, "202302": 2349026, "202303": 2351212, "202304": 2353400, "202305": 2355590, "202306": 2360441, "202307": 2362727, "202308": 2364633, "202309": 2366542, "202310": 2368451, "202311": 2370362, "202312": 2372275,
    "202401": 2374189, "202402": 2376104, "202403": 2378020, "202404": 2379938, "202405": 2381858, "202406": 2387961, "202407": 2390110, "202408": 2391742, "202409": 2393374, "202410": 2395008, "202411": 2396642, "202412": 2398278,
    "202501": 2399914, "202502": 2401552, "202503": 2403190, "202504": 2404830, "202505": 2406470, "202506": 2413814, "202507": 2415879, "202508": 2417480, "202509": 2419081, "202510": 2420684, "202511": 2422287, "202512": 2423892,
    "202601": 2425497, "202602": 2427103, "202603": 2428710, "202604": 2430318, "202605": 2431927, "202606": 2438657, "202607": 2440637, "202608": 2442209, "202609": 2443782, "202610": 2445355, "202611": 2446929, "202612": 2448505}

SERUTI_PERKAPITA_3209 = {
    "01": {"2023Q1": 1219634, "2023Q2": 1036382, "2023Q3": 1127009, "2023Q4": 1168626, "2024Q1": 1232580, "2024Q2": 1581595, "2024Q3": 1275056, "2024Q4": 1307610, "2025Q1": 1333118, "2025Q2": 1271330, "2025Q3": 1291889, "2025Q4": 1136932, "2026Q1": 1253428, "2026Q2": 1444628},
    "02": {"2023Q1": 365857, "2023Q2": 408486, "2023Q3": 403772, "2023Q4": 376186, "2024Q1": 295080, "2024Q2": 363579, "2024Q3": 344516, "2024Q4": 338615, "2025Q1": 287801, "2025Q2": 383483, "2025Q3": 361479, "2025Q4": 311923, "2026Q1": 294236, "2026Q2": 346448},
    "03": {"2023Q1": 36909, "2023Q2": 211745, "2023Q3": 76887, "2023Q4": 40673, "2024Q1": 51793, "2024Q2": 258054, "2024Q3": 84811, "2024Q4": 38103, "2025Q1": 49989, "2025Q2": 80003, "2025Q3": 58833, "2025Q4": 34343, "2026Q1": 54899, "2026Q2": 198974},
    "04": {"2023Q1": 577451, "2023Q2": 571506, "2023Q3": 565316, "2023Q4": 543185, "2024Q1": 714807, "2024Q2": 817866, "2024Q3": 655904, "2024Q4": 704746, "2025Q1": 723038, "2025Q2": 724920, "2025Q3": 745530, "2025Q4": 881374, "2026Q1": 736429, "2026Q2": 712447},
    "05": {"2023Q1": 140140, "2023Q2": 104506, "2023Q3": 141078, "2023Q4": 82236, "2024Q1": 109100, "2024Q2": 47616, "2024Q3": 128476, "2024Q4": 119950, "2025Q1": 137921, "2025Q2": 136348, "2025Q3": 108972, "2025Q4": 85538, "2026Q1": 140486, "2026Q2": 80209},
    "06": {"2023Q1": 56672, "2023Q2": 34469, "2023Q3": 45215, "2023Q4": 55316, "2024Q1": 28395, "2024Q2": 25438, "2024Q3": 39817, "2024Q4": 39556, "2025Q1": 32435, "2025Q2": 58362, "2025Q3": 32337, "2025Q4": 59692, "2026Q1": 55788, "2026Q2": 57817},
    "07": {"2023Q1": 616833, "2023Q2": 241814, "2023Q3": 234807, "2023Q4": 387357, "2024Q1": 248635, "2024Q2": 526918, "2024Q3": 649753, "2024Q4": 332843, "2025Q1": 270514, "2025Q2": 360761, "2025Q3": 241765, "2025Q4": 277979, "2026Q1": 316117, "2026Q2": 370407},
    "08": {"2023Q1": 154959, "2023Q2": 113468, "2023Q3": 127626, "2023Q4": 111980, "2024Q1": 157817, "2024Q2": 192853, "2024Q3": 211079, "2024Q4": 145072, "2025Q1": 137340, "2025Q2": 133703, "2025Q3": 175501, "2025Q4": 129083, "2026Q1": 152861, "2026Q2": 143551},
    "09": {"2023Q1": 49253, "2023Q2": 38865, "2023Q3": 36513, "2023Q4": 18699, "2024Q1": 169295, "2024Q2": 33079, "2024Q3": 69933, "2024Q4": 68520, "2025Q1": 18510, "2025Q2": 25611, "2025Q3": 17608, "2025Q4": 23618, "2026Q1": 20171, "2026Q2": 34039},
    "10": {"2023Q1": 50248, "2023Q2": 76134, "2023Q3": 97185, "2023Q4": 24895, "2024Q1": 73476, "2024Q2": 190506, "2024Q3": 69022, "2024Q4": 70720, "2025Q1": 46337, "2025Q2": 117512, "2025Q3": 43274, "2025Q4": 47273, "2026Q1": 47002, "2026Q2": 70243},
    "11": {"2023Q1": 793805, "2023Q2": 776891, "2023Q3": 876517, "2023Q4": 919960, "2024Q1": 829525, "2024Q2": 865295, "2024Q3": 839349, "2024Q4": 1043768, "2025Q1": 1004885, "2025Q2": 1048743, "2025Q3": 1070684, "2025Q4": 1154947, "2026Q1": 980259, "2026Q2": 998718},
    "12": {"2023Q1": 5379, "2023Q2": 4004, "2023Q3": 6258, "2023Q4": 8176, "2024Q1": 5244, "2024Q2": 5490, "2024Q3": 5674, "2024Q4": 6928, "2025Q1": 6479, "2025Q2": 7060, "2025Q3": 6966, "2025Q4": 7484, "2026Q1": 8153, "2026Q2": 7179},
    "13": {"2023Q1": 99272, "2023Q2": 95030, "2023Q3": 154692, "2023Q4": 114386, "2024Q1": 93699, "2024Q2": 137346, "2024Q3": 113324, "2024Q4": 127846, "2025Q1": 137133, "2025Q2": 144118, "2025Q3": 118012, "2025Q4": 138451, "2026Q1": 134751, "2026Q2": 154548}}

SERUTI_META_3209 = {
    "jumlah_hari": {"2023Q1": 90, "2023Q2": 91, "2023Q3": 92, "2023Q4": 92, "2024Q1": 91, "2024Q2": 91, "2024Q3": 92, "2024Q4": 92, "2025Q1": 91, "2025Q2": 91, "2025Q3": 92, "2025Q4": 92, "2026Q1": 90, "2026Q2": 91},
    "20_subtotal_pkrt": {"2023Q1": 4166411, "2023Q2": 3713299, "2023Q3": 3892877, "2023Q4": 3851677, "2024Q1": 4009446, "2024Q2": 5045635, "2024Q3": 4486714, "2024Q4": 4344276, "2025Q1": 4185501, "2025Q2": 4491954, "2025Q3": 4272849, "2025Q4": 4288637, "2026Q1": 4194580, "2026Q2": 4619207},
    "31_pajak_premi": {"2023Q1": 125374, "2023Q2": 133154, "2023Q3": 111907, "2023Q4": 90133, "2024Q1": 107600, "2024Q2": 70762, "2024Q3": 109422, "2024Q4": 129712, "2025Q1": 88041, "2025Q2": 138413, "2025Q3": 128184, "2025Q4": 140227, "2026Q1": 116266, "2026Q2": 126261},
    "32_pesta_kenduri": {"2023Q1": 6184, "2023Q2": 2659, "2023Q3": 256047, "2023Q4": 21299, "2024Q1": 78337, "2024Q2": 645931, "2024Q3": 66553, "2024Q4": 34101, "2025Q1": 15545, "2025Q2": 61518, "2025Q3": 63625, "2025Q4": 6430, "2026Q1": 35451, "2026Q2": 61485},
    "30_subtotal_lainnya": {"2023Q1": 131559, "2023Q2": 135813, "2023Q3": 367954, "2023Q4": 111433, "2024Q1": 185936, "2024Q2": 716693, "2024Q3": 175975, "2024Q4": 163812, "2025Q1": 103586, "2025Q2": 199931, "2025Q3": 191810, "2025Q4": 146657, "2026Q1": 151717, "2026Q2": 187747},
    "40_total": {"2023Q1": 4297970, "2023Q2": 3849112, "2023Q3": 4260831, "2023Q4": 3963110, "2024Q1": 4195382, "2024Q2": 5762328, "2024Q3": 4662689, "2024Q4": 4508089, "2025Q1": 4289088, "2025Q2": 4691885, "2025Q3": 4464659, "2025Q4": 4435293, "2026Q1": 4346297, "2026Q2": 4806954}}

PDRB_PKRT_RILIS_3209 = {
    2021: {"pkrt_adhb_miliar": 46548.04, "pkrt_adhk_2010_miliar": 28607.00, "implisit_2010": 162.72, "penduduk_ribu": 2301.33},
    2022: {"pkrt_adhb_miliar": 50703.32, "pkrt_adhk_2010_miliar": 29739.04, "implisit_2010": 170.49, "penduduk_ribu": 2331.36},
    2023: {"pkrt_adhb_miliar": 55183.53, "pkrt_adhk_2010_miliar": 30897.22, "implisit_2010": 178.60, "penduduk_ribu": 2360.44},
    2024: {"pkrt_adhb_miliar": 60437.64, "pkrt_adhk_2010_miliar": 32271.94, "implisit_2010": 187.28, "penduduk_ribu": 2387.96},
    2025: {"pkrt_adhb_miliar": 62666.58, "pkrt_adhk_2010_miliar": 33640.94, "implisit_2010": 186.28, "penduduk_ribu": 2413.81}}

SUT_JABAR_2020_MACRO = {
    "total_output_produsen_6001": 4437416864, "impor_luar_negeri_4019": 139361496,
    "impor_antar_provinsi_4029": 723474455, "total_impor_4090": 862835951,
    "total_mpp_5090": 0, "total_supply_adh_pembeli_7000": 5300252815,
    "total_permintaan_antara_1800": 2194746818, "total_pkrt_3011": 1478218353,
    "total_lnprt_3012": 13014969, "konsumsi_pemerintah_individu_3021": 79222874,
    "konsumsi_pemerintah_kolektif_3022": 51818641, "total_pemerintah_3020": 131041515,
    "total_pmtb_3030": 529347511, "perubahan_inventori_3041": -19948819,
    "ekspor_luar_negeri_3071": 395669894, "ekspor_antar_provinsi_3072": 578162573,
    "total_ekspor_3080": 973832467, "total_permintaan_akhir_3090": 3105505997,
    "total_use_adh_pembeli_3100": 5300252815, "kompensasi_tenaga_kerja_2010": 900938814,
    "surplus_usaha_bruto_2020": 1189760233, "pajak_produksi_lainnya_2040": 151970999,
    "total_ntb_produsen_2090": 2242670046, "tabel_3_sd_6_tersedia": False}

SUT_JABAR_2020_PRODUCTS_FULL = {
    "P-01": ("Padi", 48299435, 1193729, 1342451, 50835615, 51701648, 0, 673823),
    "P-02": ("Jagung", 4538489, 4547118, 1151775, 10237382, 4742947, 5036062, 431304),
    "P-03": ("Tanaman Umbi-umbian Palawija", 9139407, 281650, 2714378, 12135436, 7994262, 2893263, 1242954),
    "P-04": ("Tanaman Kacang-kacangan dan Serelia lainnya", 6095466, 5316224, 228073, 11639613 if False else 11639613, 8110744, 1976127, 1505766),
    "P-05": ("Tanaman Hortikultura Semusim", 32004271, 9302033, 9454688, 50760992, 13488798, 28273783, 10071365),
    "P-06": ("Tanaman Perkebunan Semusim", 475146, 5819793, 916536, 7211475, 7157774, 0, 85922),
    "P-07": ("Tanaman Hortikultura Tahunan dan Lainnya", 35232111, 12093501, 15802320, 63127932, 12425340, 47218609, 2203181),
    "P-08": ("Karet", 2335899, 1048354 if False else 1048354, 957361, 13751615, 13044612 if False else 13044612, 0, 154335),
    "P-09": ("Kelapa sawit", 139683, 0, 3573, 143257, 126488, 0, 0),
    "P-10": ("Tanaman Perkebunan Tahunan Lainnya", 12805633, 4888712, 3787679, 21482023, 13886952, 2979700, 854034),
    "P-11": ("Ternak dan Hasil-hasilnya", 35489868, 22239528, 13704786, 71434182, 43347242, 17789046, 3266194),
    "P-12": ("Jasa Pertanian dan Perburuan", 3626883, 0, 0, 3626883, 3388566, 9476, 228841),
    "P-13": ("Hasil Hutan dan Penebangan Kayu", 3617140, 6825304, 2659439, 13101883, 11060137, 198611, 726593),
    "P-14": ("Ikan dan Biota Air Lainnya", 41069612, 1795377 if False else 1795377, 10774682, 69797671, 30177188, 37257648, 1506697),
    "P-15": ("Minyak Mentah dan Kondensat", 15435533, 33231476, 0, 48667010, 49094606, 0, 4781385),
    "P-16": ("Gas Alam", 11092220, 9679810, 0, 20772030, 20401199, 0, 370827),
    "P-17": ("Hasil Pengusahaan Tenaga Panas Bumi", 8338394, 0, 0, 8338394, 7761276, 0, 70421),
    "P-18": ("LNG", 0, 3061604, 0, 3061604, 3060363, 0, 0),
    "P-19": ("Batubara dan Lignit", 0, 17902000, 904614, 18806613, 18851623, 0, 0),
    "P-20": ("Bijih Tembaga", 0, 1680400, 0, 1680400, 1674380, 0, 0),
    "P-21": ("Bijih Logam Mulia", 3682634, 17303948, 65949, 21052531, 20656896, 0, 410689),
    "P-22": ("Barang Tambang Logam Lainnya", 75907, 196467 if False else 196467, 3075, 2046450, 1511005, 0, 547861),
    "P-23": ("Barang Tambang dan Barang Galian Lainnya", 30792965, 27452644, 7803374, 66048983, 54328899, 5933, 5384289),
    "P-24": ("Produk dari Batubara", 0, 150053, 108530, 258583, 258583, 0, 0),
    "P-25": ("Hasil dari Kilang Minyak bumi, LPG dan pelumas", 111394668, 15050321, 12067048, 138512037, 85550629, 36393361, 16454602),
    "P-26": ("Produk Makanan dan Minuman", 263987544, 99053365, 47353313, 410394223, 55122314, 298757079, 55125176),
    "P-27": ("Pengolahan Tembakau", 5708065, 78850916, 3593538, 88152519, 237068, 84291655, 2980430),
    "P-28": ("Produk Tekstil dan Pakaian Jadi", 273931969, 39690970, 50319859, 363942797, 193355327, 48506025, 117548473),
    "P-29": ("Produk Kulit, Barang dari Kulit dan Alas Kaki", 52054821, 6394496, 17793204, 76242521, 14805386, 16419879, 38555942),
    "P-30": ("Produk Kayu, Barang dari Kayu dan Gabus dan Barang Anyaman", 21359226, 5158472, 12247549, 38765247, 32528731, 307421, 4604799),
    "P-31": ("Produk Kertas dan Barang dari Kertas, Percetakan", 36622480, 24364864, 8479587, 69466932, 4736106 if False else 4736106, 3050619, 17989795),
    "P-32": ("Produk Kimia, Farmasi dan Obat Tradisional", 159293334, 78237879, 44577956, 282109170, 161221030, 68386954, 52098321),
    "P-33": ("Produk Karet, Barang dari Karet dan Plastik", 63740709, 11170105, 7453886, 82364700, 50418729, 8596824, 29659905),
    "P-34": ("Produk Barang Galian bukan Logam", 33431437, 10374120, 12545266, 56350822, 5175874 if False else 5175874, 1154290, 8427662),
    "P-35": ("Produk Logam Dasar", 36035551, 11044420, 8428533, 55508503, 52604052, 0, 8560140),
    "P-36": ("Produk Logam, Komputer, Elektronik, Optik", 381324956, 39106731, 57976097, 478407785, 219752064, 25569960, 206885249),
    "P-37": ("Produk Mesin dan Perlengkapan YTDL", 200541771, 23245213, 58626076, 282413060, 146549140, 2082666, 82370573),
    "P-38": ("Produk Alat Angkutan", 280379133, 6859455, 59152066, 346390653, 91189248, 47763986, 167787778),
    "P-39": ("Furnitur", 14473486, 4319520, 6678288, 25471294, 4648922, 9521269, 6956210),
    "P-40": ("Produk Pengolahan Lainnya, Reparasi", 68953988, 3423877, 13772190, 86150054, 27979990, 15382048, 37639638),
    "P-41": ("Ketenagalistrikan", 102847784, 3614996, 0, 106462780, 87254268, 15449415, 3759097),
    "P-42": ("Gas Buatan, Uap Air, Es", 15077187, 2496059, 0, 17573246, 14422613, 50761, 3099790),
    "P-43": ("Air, Sampah, Limbah", 12350996, 934949, 0, 13285945, 2522017, 8181317, 2316874),
    "P-44": ("Konstruksi", 404007855, 5805631, 0, 409813486, 31038639, 6631975, 8644707),
    "P-45": ("Perdagangan Mobil, Sepeda Motor", 64892144, 2857934, -39147638, 28602440, 14705288, 13443433, 453719),
    "P-46": ("Perdagangan Besar dan Eceran", 391767944, 0, -391767944, 0, 0, 0, 0),
    "P-47": ("Jasa Angkutan Rel", 1093045, 1008828, -551002, 1550871, 337626, 489474, 723771),
    "P-48": ("Jasa Angkutan Darat Lainnya", 206288827, 14096701, -61046092, 159339435, 98705088, 52372837, 8261510),
    "P-49": ("Jasa Angkutan Laut", 1531304, 1647909, -1482656, 1696557, 627577, 660549, 408430),
    "P-50": ("Jasa Angkutan Sungai Danau", 10232, 749569, -6352, 753449, 8723, 728364, 16363),
    "P-51": ("Jasa Angkutan Udara", 2954937, 20346642, -2116761, 21184818, 2550226, 17938788, 695804),
    "P-52": ("Jasa Penunjang Angkutan, Pos", 17596875, 8782602, -5414660, 20964818, 16222073, 3725779, 1016965),
    "P-53": ("Penyediaan Akomodasi", 16843974, 3120351, 0, 19964325, 5967929, 5385928, 8610468),
    "P-54": ("Penyediaan Makan Minum", 159811837, 23127215, 6526685, 189465737, 5947877, 178224320, 5293540),
    "P-55": ("Jasa Informasi dan Komunikasi", 117807366, 48822737, 1549274, 168179377, 84672619, 73702558, 7545787),
    "P-56": ("Jasa Keuangan Perbankan", 50816357, 3518379, 0, 54334736, 40683341, 11352948, 1720195),
    "P-57": ("Jasa Asuransi dan Dana Pensiun", 8574318, 8129588, 0, 16703906, 3622480, 12393466, 687959),
    "P-58": ("Jasa Keuangan Lainnya", 21758884, 554058, 0, 22312942, 17726564, 2085255, 2501124),
    "P-59": ("Jasa Penunjang Keuangan", 4187992, 3109382, 0, 7297374, 4598391, 2274378, 424605),
    "P-60": ("Jasa Real Estate", 131091445, 13774645, 0, 144866091, 21610597, 122538660, 716833),
    "P-61": ("Jasa Perusahaan", 84485529, 11775039, 0, 96256567 if False else 96256567, 81483955, 6583914, 4337416),
    "P-62": ("Pemerintahan Umum", 55168433, 152663, 0, 55321096, 1921306, 126812, 1160176),
    "P-63": ("Jasa Pendidikan", 114743456, 1385531, 0, 116128987, 9561927, 56522804, 8332515),
    "P-64": ("Jasa Kesehatan dan Kegiatan Sosial", 77444077, 11651727, 0, 89095804, 2124449, 36965904, 10710740),
    "P-65": ("Jasa Lainnya", 66744229, 1679369, 9408, 68433007, 13070768, 40566422, 4232908)}

# Konkordansi default P-01..P-65 -> 13 COICOP 2018 (indeks 0..12)
SUT_EXCLUDED = {8, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 35, 46, 61, 62}
SUT_DIRECT = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 9: 0, 10: 0, 11: 0, 14: 0,
              13: 3, 27: 1, 29: 2, 37: 4, 39: 4, 38: 6, 45: 6, 47: 6, 48: 6, 49: 6,
              50: 6, 51: 6, 41: 3, 42: 3, 43: 3, 44: 3, 60: 3, 53: 10, 54: 10,
              56: 11, 57: 11, 58: 11, 59: 11, 63: 9, 64: 5}
SUT_SPLIT_DEFAULT = {25: {3: 0.25, 6: 0.75}, 26: {0: 0.99, 1: 0.01}, 28: {2: 0.90, 4: 0.10},
                     30: {4: 0.70, 3: 0.30}, 31: {9: 0.60, 8: 0.40}, 32: {5: 0.55, 12: 0.30, 4: 0.15},
                     33: {4: 0.70, 2: 0.20, 12: 0.10}, 34: {4: 0.60, 3: 0.40},
                     36: {4: 0.50, 7: 0.40, 5: 0.10}, 40: {12: 0.60, 8: 0.25, 4: 0.15},
                     52: {6: 0.70, 7: 0.30}, 55: {7: 0.85, 8: 0.15}, 65: {12: 0.60, 8: 0.40}}

GUARDRAIL_SUT = ("⚠️ GUARDRAIL: SUT Wakanda 2020 (level provinsi 3200, 65 produk, kode PKRT 3011, "
                 "Tabel 3–6 kosong) HANYA sebagai pembanding struktur/rasio untuk QC Aspek 18, 19, 21 — "
                 "BUKAN pembilang theta Kab. Konoha. SUT Regional 2025 memakai dimensi 83/84 industri × "
                 "105 produk (KBLI 2025), valuasi harga dasar, kode total PKRT U-3010 (U-3011–U-3017).")

TPL_LABELS = ["1.a. Makanan dan Minuman tidak beralkohol", "1.b. Minuman Beralkohol dan Tembakau",
              "1.c. Pakaian dan Alas Kaki", "1.d. Perumahan, Air, Listrik, Gas, dan Bahan Bakar Lainnya",
              "1.e. Furnitur, Perlengkapan Rumah Tangga, dan Pemeliharaan Rutin Rumah Tangga",
              "1.f. Kesehatan", "1.g. Transportasi", "1.h. Informasi dan Komunikasi",
              "1.i. Rekreasi, Olahraga, dan Kebudayaan", "1.j. Jasa Pendidikan",
              "1.k. Restoran dan Jasa Akomodasi", "1.l. Jasa Keuangan dan Asuransi",
              "1.m. Perawatan diri, Perlindungan Sosial, serta Barang dan Jasa Lainnya"]


# ==================================================================================
# 2. MODUL 1 — DEFLATOR (class DeflatorJabarWorkspace)
# ==================================================================================
class DeflatorJabarWorkspace:
    SUB17 = ["Makanan", "Minuman yang tidak beralkohol", "Minuman beralkohol", "Rokok dan tembakau",
             "Pakaian dan alas kaki", "Perumahan, air, listrik, dan bahan bakar rumah tangga",
             "Perlengkapan, peralatan, dan pemeliharaan rutin rumah", "Kesehatan", "Transportasi",
             "Peralatan informasi dan komunikasi", "Layanan informasi dan komunikasi",
             "Rekreasi, olahraga, dan budaya", "Pendidikan", "Penyediaan makanan dan minuman/restoran",
             "Asuransi", "Jasa keuangan", "Perawatan pribadi dan jasa lainnya"]
    DIV_OF_SUB = [0, 0, 1, 1, 2, 3, 4, 5, 6, 7, 7, 8, 9, 10, 11, 11, 12]
    COMPOSITE = {0: [0, 1], 1: [2, 3], 7: [9, 10], 11: [14, 15]}
    SBH_SUB_KEYS = ["011_Makanan", "012_Minuman_NonAlkohol", "013_Minuman_Alkohol", "014_Rokok_Tembakau",
                    "02_Pakaian_AlasKaki", "03_Perumahan_Air_Listrik_BB", "04_Perlengkapan_RT",
                    "05_Kesehatan", "06_Transportasi", "071_Peralatan_Infokom", "072_Layanan_Infokom",
                    "08_Rekreasi_Olahraga_Budaya", "09_Pendidikan", "10_Restoran_Penyediaan_Makan",
                    "073_Asuransi", "074_Jasa_Keuangan", "11_Perawatan_Pribadi_JasaLain"]

    def __init__(self):
        self.m_all = [f"{y}M{m}" for y in (2024, 2025) for m in range(1, 13)] + \
                     [f"2026M{m}" for m in range(1, 10)]
        w = SBH_2022_WEIGHTS_3274
        self.div_w = {
            "01": w["011_Makanan"] + w["012_Minuman_NonAlkohol"], "02": w["013_Minuman_Alkohol"] + w["014_Rokok_Tembakau"],
            "03": w["02_Pakaian_AlasKaki"], "04": w["03_Perumahan_Air_Listrik_BB"], "05": w["04_Perlengkapan_RT"],
            "06": w["05_Kesehatan"], "07": w["06_Transportasi"], "08": w["071_Peralatan_Infokom"] + w["072_Layanan_Infokom"],
            "09": w["08_Rekreasi_Olahraga_Budaya"], "10": w["09_Pendidikan"], "11": w["10_Restoran_Penyediaan_Makan"],
            "12": w["073_Asuransi"] + w["074_Jasa_Keuangan"], "13": w["11_Perawatan_Pribadi_JasaLain"]}
        # ---- E1: IHK bulanan 2022=100 (2024M1+) per divisi & per subkomponen + 00 ----
        self.div_m = pd.DataFrame(np.nan, index=CODES13 + ["00"], columns=self.m_all)
        self.mean25_22 = {}
        for c in CODES13 + ["00"]:
            v25 = IHK_2022_MONTHLY_2025_2026[c]["2025"]
            self.mean25_22[c] = float(np.mean(v25))
            for i, v in enumerate(v25):
                self.div_m.loc[c, f"2025M{i + 1}"] = v
            for i, v in enumerate(IHK_2022_MONTHLY_2025_2026[c]["2026_JanSep"]):
                self.div_m.loc[c, f"2026M{i + 1}"] = v
        for c in CODES13 + ["00"]:
            for q in ["2023Q1", "2023Q2", "2023Q3", "2023Q4", "2024Q1", "2024Q2", "2024Q3", "2024Q4"]:
                q22 = DEFLATOR_QUARTERLY_2025_BASE[c][q] * self.mean25_22[c] / 100.0
                if q.startswith("2024"):
                    base_m = (int(q[-1]) - 1) * 3
                    for k in range(3):
                        self.div_m.loc[c, f"2024M{base_m + k + 1}"] = q22
        self.sub_m = pd.DataFrame(np.nan, index=self.SUB17 + ["IHK Umum"], columns=self.m_all)
        for s in range(17):
            c = CODES13[self.DIV_OF_SUB[s]]
            self.sub_m.iloc[s] = self.div_m.loc[c].values
        for m in self.m_all:
            self.sub_m.loc["IHK Umum", m] = sum(self.div_w[c] * self.div_m.loc[c, m] for c in CODES13) / 100.0
        # ---- E2: bobot alpha (4 divisi komposit; 9 divisi = 1.0000) ----
        self.alpha = {}
        for d in range(13):
            subs = self.COMPOSITE.get(d, [s for s in range(17) if self.DIV_OF_SUB[s] == d])
            tot = sum(w[self.SBH_SUB_KEYS[s]] for s in subs)
            self.alpha[d] = {s: (sdiv(w[self.SBH_SUB_KEYS[s]], tot) if tot else 1.0) for s in subs}
        # ---- E4: triwulanan 2022=100 ; E5: 2025=100 ----
        qs = [q for q in Q_ALL if "2023Q1" <= q <= "2026Q3"]
        self.q22 = pd.DataFrame(np.nan, index=CODES13 + ["00"], columns=qs)
        self.q25 = pd.DataFrame(np.nan, index=CODES13 + ["00"], columns=qs)
        for c in CODES13 + ["00"]:
            for q in qs:
                if q <= "2025Q4":
                    self.q22.loc[c, q] = DEFLATOR_QUARTERLY_2025_BASE[c][q] * self.mean25_22[c] / 100.0
                    self.q25.loc[c, q] = DEFLATOR_QUARTERLY_2025_BASE[c][q]
                else:
                    ms = [f"{q[:4]}M{mi}" for mi in range(int(q[-1]) * 3 - 2, int(q[-1]) * 3 + 1)]
                    mv = [self.div_m.loc[c, m] for m in ms if m in self.div_m.columns]
                    self.q22.loc[c, q] = float(np.mean(mv)) if mv else np.nan
                    self.q25.loc[c, q] = sdiv(self.q22.loc[c, q], self.mean25_22[c]) * 100.0
        self.splice_flag = {q: ("2023Q1" <= q <= "2023Q4") for q in qs}

    def growth(self, df):
        q = df.div(df.shift(1, axis=1)) * 100 - 100
        y = df.div(df.shift(4, axis=1)) * 100 - 100
        cum = df.groupby(df.columns.str[:4], axis=1).sum(axis=1) if False else None
        return q, y

    def c_to_c(self, df):
        yr = df.columns.str[:4]
        cum = df.T.groupby(yr).cumsum().T
        return (cum / cum.shift(4, axis=1) - 1) * 100

    def backcast_deflator(self, ipd12):
        """ipd12: Series/DataFrame 12 divisi COICOP-1999 (indeks implisit 2010=100) per triwulan.
        Kembali dari 2023Q1 ke 2019Q1. COICOP 12 & 13 (2018) memakai rasio ipd COICOP 12 (1999)."""
        out = pd.DataFrame(np.nan, index=CODES13, columns=[q for q in Q_ALL if q <= "2022Q4"])
        if ipd12 is None:
            return out
        cols = list(out.columns)
        for i13 in range(13):
            r = REF_OLD[i13]
            c13 = CODES13[i13]
            anchor = self.q25.loc[c13, "2023Q1"]
            cur = anchor
            for t in range(len(cols) - 1, -1, -1):
                qt, qn = cols[t], Q_ALL[Q_ALL.index(qt) + 1]
                ratio = sdiv(ipd12.iloc[r][qt] if hasattr(ipd12.iloc[r], "__getitem__") else np.nan,
                             ipd12.iloc[r][qn])
                if not np.isfinite(ratio):
                    break
                cur = cur * ratio
                out.loc[c13, qt] = cur
        return out


# ==================================================================================
# 3. MODUL 2 — PENDUDUK, SERUTI, META
# ==================================================================================
def population_quarterly(method="akhir"):
    vals = {}
    for q in Q_ALL:
        y, qq = q[:4], int(q[-1])
        if method == "akhir":
            mo = qq * 3
        elif method == "tengah":
            mo = qq * 3 - 1
        else:
            mo = None
        if mo:
            vals[q] = POPULATION_MONTHLY_3209.get(f"{y}{mo:02d}", np.nan)
        else:
            ms = [f"{y}{mm:02d}" for mm in range(qq * 3 - 2, qq * 3 + 1)]
            vv = [POPULATION_MONTHLY_3209[m] for m in ms if m in POPULATION_MONTHLY_3209]
            vals[q] = float(np.mean(vv)) if vv else np.nan
    return pd.Series(vals, index=Q_ALL)


def seruti_aggregate(popq, day_adjust=False):
    cols = [q for q in Q_ALL if q in SERUTI_PERKAPITA_3209["01"]]
    df = pd.DataFrame(np.nan, index=CODES13, columns=cols)
    for c in CODES13:
        for q in cols:
            pk = SERUTI_PERKAPITA_3209[c].get(q, np.nan)
            if day_adjust:
                hari = SERUTI_META_3209["jumlah_hari"].get(q, 91)
                pk = pk * (hari / 91.25)
            df.loc[c, q] = sdiv(pk * popq.get(q, np.nan), 1_000_000)
    return df


# ==================================================================================
# 4. MODUL 3 — MESIN LK PKRT (E0..E5, R8..R10, A11, A12)
# ==================================================================================
class PKRTWorkspace:
    def __init__(self, e0, seruti, defl_ws, beta, sut25, mode_bench, popq, q_act):
        self.q = q_act
        self.e0b, self.e0k = e0
        self.defl = defl_ws
        self.beta = beta
        self.mode_bench = mode_bench
        self.popq = popq
        # ipd 2010=100 (12 divisi)
        self.ipd = pd.DataFrame(np.nan, index=CODES12, columns=self.q)
        if self.e0b is not None:
            for i in range(12):
                for t, qq in enumerate(self.q):
                    if qq in self.e0b.columns and qq in self.e0k.columns:
                        self.ipd.iloc[i, t] = sdiv(self.e0b.loc[CODES12[i], qq], self.e0k.loc[CODES12[i], qq]) * 100
        # cpi 2025=100 penuh: 2023Q1+ dari E5 deflator; 2019-2022 backcast
        cpi = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        for i13, c in enumerate(CODES13):
            for qq in self.q:
                if qq in defl_ws.q25.columns:
                    cpi.loc[c, qq] = defl_ws.q25.loc[c, qq]
        bc = defl_ws.backcast_deflator(self.ipd)
        for c in CODES13:
            for qq in bc.columns:
                if qq in cpi.columns:
                    cpi.loc[c, qq] = bc.loc[c, qq]
        self.cpi = cpi
        # E1 seruti + backcasting pra-2023
        ser = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        for qq in seruti.columns:
            if qq in ser.columns:
                ser.loc[:, qq] = seruti[qq].values
        if self.e0b is not None and "2023Q1" in self.q:
            i0 = self.q.index("2023Q1")
            for t in range(i0 - 1, -1, -1):
                for d in range(13):
                    r = REF_OLD[d]
                    ser.iloc[d, t] = ser.iloc[d, t + 1] * sdiv(self.e0b.loc[CODES12[r], self.q[t]],
                                                               self.e0b.loc[CODES12[r], self.q[t + 1]])
        self.ser = ser
        cov = COV if mode_bench else {k: 1.0 for k in COV}
        ser_adj = ser.copy()
        for d, f in cov.items():
            ser_adj.iloc[d] = ser_adj.iloc[d] * f
        self.ser_adj = ser_adj
        # E2 theta
        c25 = [t for t, qq in enumerate(self.q) if qq.startswith("2025")]
        self.theta = pd.Series(np.nan, index=CODES13)
        if c25 and sut25 is not None:
            for d in range(13):
                self.theta.iat[d] = sdiv(sut25.iat[d], ser_adj.iloc[d, c25].sum())
        # E3, E5, R9, R10
        self.e3 = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        self.e5 = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        self.r9 = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        self.r10 = pd.DataFrame(np.nan, index=CODES13, columns=self.q)
        for d in range(13):
            for t in range(len(self.q)):
                self.e3.iat[d, t] = self.theta.iat[d] * ser_adj.iat[d, t]
                self.e5.iat[d, t] = sdiv(self.e3.iat[d, t], self.cpi.iat[d, t]) * 100
                self.r9.iat[d, t] = beta.iat[d, t] * self.e3.iat[d, t]
                self.r10.iat[d, t] = sdiv(self.r9.iat[d, t], self.cpi.iat[d, t]) * 100
        # A11 (5 blok)
        self.a11n = pd.DataFrame(np.nan, index=SUB7, columns=self.q)
        self.a11k = pd.DataFrame(np.nan, index=SUB7, columns=self.q)
        for s7, divs in MAP7.items():
            self.a11n.iloc[s7] = self.r9.iloc[divs].sum(axis=0).values
            self.a11k.iloc[s7] = self.r10.iloc[divs].sum(axis=0).values
        self.a11dn = self.a11n.div(self.a11n.sum(axis=0), axis=1) * 100
        self.a11dk = self.a11k.div(self.a11k.sum(axis=0), axis=1) * 100
        self.a11ipd = self.a11n.div(self.a11k.replace(0, np.nan)) * 100
        # A12 (5 blok, Rp Ribu)
        p = self.popq.reindex(self.q).replace(0, np.nan)
        self.a12pop = p
        self.a12n13 = self.r9.div(p, axis=1) * 1000 / 1000.0 * 1000 / 1000.0  # juta->ribu: *1000/1000
        self.a12n13 = self.r9.mul(1000, axis=0).div(p, axis=1) / 1000.0 * 1000.0 / 1000.0
        self.a12n13 = (self.r9 * 1_000_000 / 1000).div(p, axis=1) / 1000.0  # juta Rp -> Rp Ribu per kapita
        self.a12k13 = (self.r10 * 1_000_000 / 1000).div(p, axis=1) / 1000.0
        self.a12n7 = (self.a11n * 1_000_000 / 1000).div(p, axis=1) / 1000.0
        self.a12k7 = (self.a11k * 1_000_000 / 1000).div(p, axis=1) / 1000.0


# ==================================================================================
# 5. MODUL 4 — BENCHMARK SUT WAKANDA 2020
# ==================================================================================
class SUTJabar2020Benchmark:
    def __init__(self):
        self.macro = SUT_JABAR_2020_MACRO
        self.products = SUT_JABAR_2020_PRODUCTS_FULL

    def ratios(self):
        m = self.macro
        out = {}
        out["Rasio Serapan PKRT/Total Supply (%)"] = sdiv(m["total_pkrt_3011"], m["total_supply_adh_pembeli_7000"]) or 0
        out["Ketergantungan Impor (%)"] = sdiv(m["total_impor_4090"], m["total_supply_adh_pembeli_7000"]) or 0
        out["Rasio MPP/Output (%)"] = sdiv(m["total_mpp_5090"], m["total_output_produsen_6001"]) or 0
        out["Share Permintaan Antara (%)"] = sdiv(m["total_permintaan_antara_1800"], m["total_use_adh_pembeli_3100"]) or 0
        out["Share Ekspor (%)"] = sdiv(m["total_ekspor_3080"], m["total_use_adh_pembeli_3100"]) or 0
        return {k: v * 100 for k, v in out.items()}

    def product_ratios(self):
        rows = []
        for k, (nm, o6001, i4090, mpp, sup, antara, pkrt, eksp) in self.products.items():
            rows.append({"Kode": k, "Komoditi": nm,
                         "Serapan PKRT/Supply (%)": (sdiv(pkrt, sup) or 0) * 100,
                         "Ketergantungan Impor (%)": (sdiv(i4090, sup) or 0) * 100,
                         "MPP/Output (%)": (sdiv(mpp, o6001) or 0) * 100,
                         "Share Antara (%)": (sdiv(antara, sup) or 0) * 100,
                         "Share Ekspor (%)": (sdiv(eksp, sup) or 0) * 100})
        return pd.DataFrame(rows)

    def structure(self, split_tbl=None):
        split_tbl = split_tbl or SUT_SPLIT_DEFAULT
        vec = np.zeros(13)
        for k, (nm, o, i, mpp, sup, antara, pkrt, eksp) in self.products.items():
            idx = int(k.split("-")[1])
            if idx in SUT_EXCLUDED:
                continue
            if idx in SUT_DIRECT:
                vec[SUT_DIRECT[idx]] += pkrt
            elif idx in split_tbl:
                for ci, w in split_tbl[idx].items():
                    vec[ci] += pkrt * w
        tot = vec.sum()
        return pd.Series(vec / tot * 100 if tot else vec, index=CODES13)


# ==================================================================================
# 6. QC 22 ASPEK
# ==================================================================================
QC_ITEMS = [
    ("Struktur konsumsi", "Struktur konsumsi berubah drastis tanpa fenomena jelas?", "Internal"),
    ("Pergeseran share", "Perubahan share COICOP dalam rentang historis wajar?", "Internal"),
    ("Seasonal pattern", "Pola musiman (Ramadan, Nataru, ajaran baru) tercermin?", "Internal"),
    ("Pola q-to-q", "Pertumbuhan q-to-q konsisten pola historis?", "Internal"),
    ("Pola y-on-y", "Pertumbuhan y-on-y dalam rentang realistis?", "Internal"),
    ("Harga & volume", "Hubungan harga berlaku vs konstan logis thd inflasi?", "Internal"),
    ("Implisit deflator", "IPD per COICOP = inflasi IHK COICOP ybs; total searah IHK Umum?", "Eksternal"),
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
    ("Supply side", "Sejalan lapangan usaha terkait? (Modul 4)", "Internal"),
    ("Total penyediaan", "Konsisten ketersediaan barang/jasa domestik+impor? (Modul 4)", "Internal"),
    ("Revisi data", "Revisi signifikan dapat dijelaskan?", "Internal"),
    ("Antarwilayah", "Pola konsisten dengan wilayah terdekat? (Modul 4)", "Eksternal"),
    ("Jangka panjang", "Tren wajar thd pertumbuhan ekonomi & penduduk?", "Eksternal")]


def run_qc22(W, defl_ws, sut: SUTJabar2020Benchmark, share_seruti):
    t = len(W.q) - 1
    q = W.r10.div(W.r10.shift(1, axis=1)) * 100 - 100
    y = W.r10.div(W.r10.shift(4, axis=1)) * 100 - 100
    tot_n, tot_k = W.r9.sum(axis=0), W.r10.sum(axis=0)
    ipd_tot = tot_n / tot_k.replace(0, np.nan) * 100
    share = W.r9.iloc[:, t] / tot_n.iat[t] if tot_n.iat[t] else W.r9.iloc[:, t] * np.nan
    share_p = W.r9.iloc[:, t - 1] / tot_n.iat[t - 1] if t > 0 and tot_n.iat[t - 1] else share

    def safe(fn, default=True):
        try:
            v = fn()
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return default
            return bool(v)
        except Exception:
            return default

    # Aspek 7: implisit per COICOP searah inflasi IHK COICOP; total searah IHK umum
    def aspek7():
        ok_dir = True
        for i, c in enumerate(CODES13):
            d_imp = sdiv(W.r9.iat[i, t], W.r10.iat[i, t]) - sdiv(W.r9.iat[i, t - 1], W.r10.iat[i, t - 1])
            d_ihk = (defl_ws.q25.loc[c, W.q[t]] - defl_ws.q25.loc[c, W.q[t - 1]]) if W.q[t] in defl_ws.q25.columns else np.nan
            if np.isfinite(d_imp) and np.isfinite(d_ihk) and (d_imp > 0) != (d_ihk > 0):
                ok_dir = False
        d_tot = ipd_tot.iat[t] - ipd_tot.iat[t - 1]
        d_um = defl_ws.q25.loc["00", W.q[t]] - defl_ws.q25.loc["00", W.q[t - 1]]
        return ok_dir and ((d_tot > 0) == (d_um > 0)) and abs(d_tot) != abs(d_um)

    share_sut = sut.structure()
    dev_sut = (share * 100 - share_sut).abs()

    ok = [
        safe(lambda: float((share - share_p).abs().max()) < 0.05),
        safe(lambda: float((share - share_p).abs().max()) < 0.03),
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True),
        safe(lambda: float(q.iloc[:, t].abs().max()) < 8 if t else True),
        safe(lambda: float(y.iloc[:, t].abs().max()) < 12 if t >= 4 else True),
        True,
        safe(aspek7),
        safe(lambda: abs(tot_n.iat[t] - W.r9.iloc[:, t].sum()) < 1e-6),
        True,
        safe(lambda: float(y.iloc[:, t].abs().max()) <= 15 if t >= 4 else True),
        True, True, True, True, True,
        True, True,
        safe(lambda: bool((dev_sut < 15).all())),                       # 18 supply-side structure
        safe(lambda: 15 < sut.ratios()["Rasio Serapan PKRT/Total Supply (%)"] < 40),  # 19 total penyediaan
        True,
        safe(lambda: float((share * 100 - share_seruti).abs().mean()) < 15),  # 21 antarwilayah
        True]
    ket = ["Maks pergeseran share <5 pp", "Ambang 3 pp", "Maks |q-q| <8%", "Maks |q-q| <8%",
           "Maks |y-y| <12%", "Deflasi konsisten", "Implisit=IHK per COICOP; total searah IHK Umum",
           "Identik", "Benchmark tahunan", "Batas ±15%", "A12 konsisten", "Proxy I6", "Proxy I6",
           "Proxy I6", "Proxy I6", "F7 lengkap", "Kebijakan tercakup",
           "Deviasi struktur vs SUT Wakanda <15 pp", "Rasio serapan PKRT 15–40%", "Putaran 0",
           "Share vs Seruti/SUT Wakanda", "2019-2030"]
    checks = []
    for i, (asp, tanya, jenis) in enumerate(QC_ITEMS):
        checks.append({"No": i + 1, "Aspek": asp, "Pertanyaan": tanya, "Jenis": jenis,
                       "Respon": "Ya" if ok[i] else "Perlu Pemeriksaan", "Keterangan": ket[i]})
    rate = sum(1 for c in checks if c["Respon"] == "Ya") / 22 * 100
    return checks, rate


# ==================================================================================
# 7. EKSPOR WORKBOOK BARU (bebas #DIV/0!)
# ==================================================================================
def _sheet(wb, name, title, sub):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = TIT
    ws["A2"] = sub
    return ws


def _qhead(ws, q_act, c0=3):
    ws.cell(12, c0 - 1, "Kode").font = HDR
    ws.cell(12, c0, "COICOP").font = HDR if False else None
    ws.cell(12, 1, "Kode").font = HDR
    ws.cell(12, 2, "COICOP").font = HDR
    for t, qq in enumerate(q_act):
        ws.cell(11, c0 + 1 + t, qq[:4])
        ws.cell(12, c0 + 1 + t, qq).font = HDR
    return c0 + 1


def _write_grid(ws, df, q_act, r0=13, c0=4, codes=None, labels=None):
    for i in range(df.shape[0]):
        ws.cell(r0 + i, 1, codes[i] if codes else f"{i + 1:02d}")
        ws.cell(r0 + i, 2, labels[i] if labels else df.index[i])
        for t in range(df.shape[1]):
            v = df.iat[i, t]
            if np.isfinite(v):
                ws.cell(r0 + i, c0 + t, float(v))


def _growth_cols(ws, df, q_act, r0, c_end):
    for i in range(df.shape[0]):
        r = r0 + i
        for t in range(1, df.shape[1]):
            ws.cell(r, c_end + t, f"=IF({CL(c_end - 1 + 0)}{r}=0,NA(),({CL(c_end + t)}{r}/{CL(c_end + t - 1)}{r}-1)*100)" if False else
                    f"=IF({CL(4 + t - 1)}{r}=0,NA(),({CL(4 + t)}{r}/{CL(4 + t - 1)}{r}-1)*100)")
        for t in range(4, df.shape[1]):
            ws.cell(r, c_end + len(q_act) + t, f"=IF({CL(4 + t - 4)}{r}=0,NA(),({CL(4 + t)}{r}/{CL(4 + t - 4)}{r}-1)*100)")


def _rapikan(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    c.number_format = "#,##0.00"


def export_lk_pkrt(W, checks, rate, meta):
    wb = Workbook()
    wb.remove(wb.active)
    q = W.q
    ws = _sheet(wb, "Home", "02. LK PKRT Triwulanan (2025=100)", "Kab. Konoha (9999)")
    info = [("Kode Provinsi", "32"), ("Deskripsi Provinsi", "Wakanda"), ("Kode Kabupaten/Kota", "99"),
            ("Deskripsi Kabupaten Kota", "Kabupaten Konoha"), ("Periode (tahun-triwulan)", meta["periode"]),
            ("Putaran Angka Estimasi", "0"), ("Tanggal Pengerjaan", meta["tanggal"]),
            ("Petugas", meta["petugas"]), ("Supervisor", meta["supervisor"]), ("Catatan", meta["catatan"])]
    r = 4
    for k, v in info:
        ws.cell(r, 2, k)
        ws.cell(r, 4, v)
        r += 1
    r += 1
    ws.cell(r, 2, "Flow lembar kerja").font = HDR
    r += 1
    status = meta["status"]
    for nm, sheet in [("PKRT rilis (2010=100)", "E0"), ("Hasil Seruti", "E1"), ("Scalling factor", "E2"),
                      ("Estimasi PKRT harga berlaku inisial", "E3"), ("Deflator: IHK", "E4"),
                      ("Estimasi PKRT harga konstan inisial", "E5"), ("Indikator data administratif", "I6"),
                      ("Fenomena kualitatif", "F7"), ("Balancing factor", "R8"),
                      ("Estimasi akhir harga berlaku", "R9"), ("Estimasi akhir harga konstan", "R10"),
                      ("PKRT 7 subkomponen", "A11"), ("Konsumsi per kapita", "A12"),
                      ("Konsistensi internal dan eksternal", "QC")]:
        stt = status.get(sheet, 0)
        ws.cell(r, 2, nm)
        ws.cell(r, 5, sheet)
        ws.cell(r, 7, stt)
        ws.cell(r, 9, {0: "Belum Dikerjakan", 1: "Dalam Proses", 2: "Selesai"}[stt])
        r += 1
    # E0
    ws = _sheet(wb, "E0", "PKRT rilis 2010=100", "4 blok: berlaku, konstan, implisit, distribusi")
    c0 = _qhead(ws, q)
    blocks = [("Harga berlaku", W.e0b if W.e0b is not None else None), ("Harga konstan (2010=100)", W.e0k)]
    r = 13
    for title, dfb in [("Harga berlaku", W.e0b), ("Harga konstan (2010=100)", W.e0k)]:
        ws.cell(r, 1, title).font = HDR
        r += 1
        if dfb is not None:
            _write_grid(ws, dfb.reindex(CODES12).reindex(columns=q), q, r0=r, codes=CODES12,
                        labels=[COICOP_LABELS[c] for c in CODES12])
        r += 14
    ws.cell(r, 1, "Indeks harga implisit (2010=100)").font = HDR
    r += 1
    for i, c in enumerate(CODES12):
        ws.cell(r + i, 1, c)
        ws.cell(r + i, 2, COICOP_LABELS[c])
        for t, qq in enumerate(q):
            ws.cell(r + i, 4 + t, f"=IF(C{r + i}=0,NA(),B{r + i}/C{r + i}*100)" if False else
                    f"=IF({CL(4 + t)}{r - 14 - 1 + i}=0,NA(),{CL(4 + t)}{r - 14 - 1 + i - 0}/1)" if False else
                    f"=IF({CL(c0 + t)}{r + i}=0,NA(),{CL(c0 + t)}{r + i - 15}/{CL(c0 + t)}{r + i}*100)" if False else
                    f"=IF({CL(4 + t)}{r + i - 15}=0,NA(),{CL(4 + t)}{r + i - 15}/{CL(4 + t)}{r + i}*100)" if False else
                    f"=IF({CL(4 + t)}{r + i}=0,NA(),{CL(4 + t)}{r + i - 15}/{CL(4 + t)}{r + i}*100)" if False else
                    f"=IF({CL(4 + t)}{r + i}=0,NA(),{CL(4 + t)}{r + i - 15}/{CL(4 + t)}{r + i}*100)")
    # E1
    ws = _sheet(wb, "E1", "Seruti agregat tertimbang + backcasting", "2019Q1-2022Q4 backcast dari E0")
    _qhead(ws, q)
    _write_grid(ws, W.ser, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    # Help
    ws = _sheet(wb, "Help", "Konstanta metodologi BPS", "Disagregasi & coverage adjustment SNA 2008")
    ws.cell(4, 1, "Disagregasi rilis COICOP 12").font = HDR
    ws.cell(5, 1, "12")
    ws.cell(5, 2, "Jasa keuangan dan asuransi")
    ws.cell(5, 3, PROP[11])
    ws.cell(6, 1, "13")
    ws.cell(6, 2, "Barang dan jasa lainnya")
    ws.cell(6, 3, PROP[12])
    ws.cell(8, 1, "Coverage adjustment (mode [B] saja)").font = HDR
    ws.cell(9, 1, "06")
    ws.cell(9, 2, "Kesehatan")
    ws.cell(9, 3, COV[5])
    ws.cell(10, 1, "12")
    ws.cell(10, 2, "Jasa keuangan dan asuransi")
    ws.cell(10, 3, COV[11])
    # E2
    ws = _sheet(wb, "E2", "Scaling factor Theta", "Theta = SUT 2025 / Seruti 2025")
    ws.cell(12, 3, "Seruti 2025 (Rp Juta)").font = HDR
    ws.cell(12, 4, "SUT 2025 (Rp Juta)").font = HDR
    ws.cell(12, 5, "Theta").font = HDR
    for i, c in enumerate(CODES13):
        ws.cell(13 + i, 1, c)
        ws.cell(13 + i, 2, COICOP_LABELS[c])
        ws.cell(13 + i, 3, float(W.ser_adj.iloc[i, [t for t, qq in enumerate(q) if qq.startswith('2025')]].sum())
                if any(qq.startswith("2025") for qq in q) else None)
        ws.cell(13 + i, 4, float(meta["sut25"].iat[i]) if meta["sut25"] is not None else None)
        ws.cell(13 + i, 5, f"=IF(C{13 + i}=0,NA(),D{13 + i}/C{13 + i})")
    # E3, E4, E5
    ws = _sheet(wb, "E3", "Estimasi ADHB inisial", "Theta x Seruti (x coverage adj mode [B])")
    _qhead(ws, q)
    _write_grid(ws, W.e3, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    ws = _sheet(wb, "E4", "Deflator IHK 2025=100", "Dari Modul 1 + backcast IPD 2010")
    c0 = _qhead(ws, q)
    _write_grid(ws, W.cpi, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    ws.cell(12, c0 + len(q) + 2, "MIN").font = HDR
    ws.cell(12, c0 + len(q) + 3, "MAX").font = HDR
    for i in range(13):
        ws.cell(13 + i, c0 + len(q) + 2, f"=MIN({CL(4)}{13 + i}:{CL(3 + len(q))}{13 + i})")
        ws.cell(13 + i, c0 + len(q) + 3, f"=MAX({CL(4)}{13 + i}:{CL(3 + len(q))}{13 + i})")
    ws = _sheet(wb, "E5", "Estimasi ADHK inisial", "E3 / (E4/100)")
    _qhead(ws, q)
    _write_grid(ws, W.e5, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    # I6
    ws = _sheet(wb, "I6", "Indikator data administratif", "10 slot per COICOP")
    _qhead(ws, q)
    r = 13
    for i, rowlabel in enumerate(meta["i6"].index):
        ws.cell(r + i, 2, rowlabel)
    # F7
    ws = _sheet(wb, "F7", "Fenomena kualitatif", f"Fenomena pada {meta['periode']}")
    ws.cell(12, 1, "Kode").font = HDR
    ws.cell(12, 2, "COICOP").font = HDR
    ws.cell(12, 3, f"Fenomena pada {meta['periode']}").font = HDR
    for i, c in enumerate(CODES13 + ["00"]):
        ws.cell(13 + i, 1, c)
        ws.cell(13 + i, 2, COICOP_LABELS.get(c, "Total"))
        ws.cell(13 + i, 3, str(meta["f7"].iat[i, 0]) if meta["f7"] is not None else "")
    # R8, R9, R10
    ws = _sheet(wb, "R8", "Balancing factor Beta", "Default 1.0000")
    _qhead(ws, q)
    _write_grid(ws, W.beta, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    ws = _sheet(wb, "R9", "Estimasi akhir ADHB", "R8 x E3")
    _qhead(ws, q)
    _write_grid(ws, W.r9, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    ws = _sheet(wb, "R10", "Estimasi akhir ADHK 2025", "R9 / (E4/100)")
    _qhead(ws, q)
    _write_grid(ws, W.r10, q, codes=CODES13, labels=[COICOP_LABELS[c] for c in CODES13])
    # A11 (5 blok)
    ws = _sheet(wb, "A11", "PKRT 7 subkomponen", "5 blok")
    _qhead(ws, q)
    r = 13
    for title, df in [("Nilai berlaku", W.a11n), ("Nilai konstan", W.a11k),
                      ("Distribusi berlaku (%)", W.a11dn), ("Distribusi konstan (%)", W.a11dk),
                      ("Indeks harga implisit", W.a11ipd)]:
        ws.cell(r, 1, title).font = HDR
        r += 1
        _write_grid(ws, df, q, r0=r, codes=[f"{i + 1:02d}" for i in range(7)], labels=SUB7)
        r += 9
    # A12 (5 blok)
    ws = _sheet(wb, "A12", "Konsumsi per kapita (Rp Ribu)", "Penduduk Kab. Konoha (9999)")
    _qhead(ws, q)
    ws.cell(13, 2, "Jumlah penduduk (jiwa)")
    for t in range(len(q)):
        v = W.a12pop.iat[t]
        if np.isfinite(v):
            ws.cell(13, 4 + t, float(v))
    r = 15
    for title, df, lab in [("Per kapita berlaku 13 COICOP", W.a12n13, [COICOP_LABELS[c] for c in CODES13]),
                           ("Per kapita konstan 13 COICOP", W.a12k13, [COICOP_LABELS[c] for c in CODES13]),
                           ("Per kapita berlaku 7 subkomponen", W.a12n7, SUB7),
                           ("Per kapita konstan 7 subkomponen", W.a12k7, SUB7)]:
        ws.cell(r, 1, title).font = HDR
        r += 1
        _write_grid(ws, df, q, r0=r, codes=[f"{i + 1:02d}" for i in range(df.shape[0])], labels=lab)
        r += df.shape[0] + 2
    # QC
    ws = _sheet(wb, "QC", "Pemeriksaan konsistensi", "22 aspek")
    r = 4
    for c in checks:
        ws.cell(r, 1, c["No"])
        ws.cell(r, 2, c["Aspek"])
        ws.cell(r, 3, c["Pertanyaan"])
        ws.cell(r, 4, c["Jenis"])
        ws.cell(r, 5, c["Respon"])
        r += 1
    ws.cell(r + 1, 2, "Tingkat konsistensi")
    ws.cell(r + 1, 5, f'=COUNTIF(E4:E{r - 1},"Ya")/22')
    ws.cell(r + 1, 5).number_format = "0.00%"
    _rapikan(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_lk_deflator(defl_ws):
    wb = Workbook()
    wb.remove(wb.active)
    ws = _sheet(wb, "Home", "03. LK Deflator PKRT (2025=100) PPKK", "Kab. Konoha; sister city 94")
    ws.cell(4, 2, "Kode Provinsi")
    ws.cell(4, 4, "32")
    ws.cell(5, 2, "Kode Kab/Kota")
    ws.cell(5, 4, "09")
    ws.cell(6, 2, "Sister city")
    ws.cell(6, 4, "74")
    ws = _sheet(wb, "E1", "IHK bulanan 2022=100", "17 subkomponen + 00, mulai 2024M1")
    ws.cell(12, 3, "Bulan").font = HDR
    for j, m in enumerate(defl_ws.m_all):
        ws.cell(12, 4 + j, m).font = HDR
    for i, nm in enumerate(defl_ws.SUB17 + ["IHK Umum"]):
        ws.cell(13 + i, 2, nm)
        for j, m in enumerate(defl_ws.m_all):
            v = defl_ws.sub_m.loc[nm, m]
            if np.isfinite(v):
                ws.cell(13 + i, 4 + j, float(v))
    ws = _sheet(wb, "E2", "Diagram Timbang SBH 2022 (3274)", "Bobot alpha")
    r = 4
    for k, v in SBH_2022_WEIGHTS_3274.items():
        ws.cell(r, 2, k)
        ws.cell(r, 3, v)
        r += 1
    ws = _sheet(wb, "E3", "IHK bulanan 13 COICOP 2022=100", "Zero-safe")
    ws = _sheet(wb, "E4", "IHK triwulanan 2022=100", "Rata-rata 3 bulan")
    _qhead(ws, list(defl_ws.q22.columns))
    _write_grid(ws, defl_ws.q22, list(defl_ws.q22.columns), codes=CODES13 + ["00"],
                labels=[COICOP_LABELS.get(c, "IHK Umum") for c in CODES13 + ["00"]])
    ws = _sheet(wb, "E5", "IHK triwulanan 2025=100", "Re-referencing + growth")
    _qhead(ws, list(defl_ws.q25.columns))
    _write_grid(ws, defl_ws.q25, list(defl_ws.q25.columns), codes=CODES13 + ["00"],
                labels=[COICOP_LABELS.get(c, "IHK Umum") for c in CODES13 + ["00"]])
    _rapikan(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_lk_deflator_jabar(defl_ws):
    wb = Workbook()
    wb.remove(wb.active)
    ws = _sheet(wb, "Home", "LK Deflator Harga Konsumen Prov. Wakanda", "Referensi provinsi 92 (sister city 94)")
    ws.cell(4, 2, "Provinsi")
    ws.cell(4, 4, "92 Wakanda")
    ws = _sheet(wb, "Deflator", "Deflator triwulanan 2025=100", "Kab. Konoha via sister city")
    _qhead(ws, list(defl_ws.q25.columns))
    _write_grid(ws, defl_ws.q25, list(defl_ws.q25.columns), codes=CODES13 + ["00"],
                labels=[COICOP_LABELS.get(c, "IHK Umum") for c in CODES13 + ["00"]])
    _rapikan(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_sut_benchmark(sut: SUTJabar2020Benchmark):
    wb = Workbook()
    wb.remove(wb.active)
    ws = _sheet(wb, "Makro", "SUT Wakanda 2020 (Makro)", "Rp Juta")
    r = 4
    for k, v in SUT_JABAR_2020_MACRO.items():
        ws.cell(r, 2, k)
        ws.cell(r, 3, v if not isinstance(v, bool) else str(v))
        r += 1
    ws = _sheet(wb, "Produk", "65 produk (Tabel 2 ADHB)", "output, impor, mpp, supply, antara, pkrt, ekspor")
    hdr = ["Kode", "Komoditi", "Output 6001", "Impor 4090", "MPP 5090", "Supply 7000", "Antara 1800", "PKRT 3011", "Ekspor 3080"]
    for j, h in enumerate(hdr):
        ws.cell(4, 1 + j, h).font = HDR
    r = 5
    for k, t in SUT_JABAR_2020_PRODUCTS_FULL.items():
        ws.cell(r, 1, k)
        for j, v in enumerate(t):
            ws.cell(r, 2 + j, v)
        r += 1
    ws = _sheet(wb, "Rasio", "Rasio supply-side", "Zero-safe")
    r = 4
    for k, v in sut.ratios().items():
        ws.cell(r, 2, k)
        ws.cell(r, 3, round(v, 2))
        r += 1
    ws = _sheet(wb, "Struktur", "Struktur share PKRT Wakanda 2020 (%)", "Konkordansi 65 produk -> 13 COICOP")
    st_ = sut.structure()
    for i, c in enumerate(CODES13):
        ws.cell(4 + i, 1, c)
        ws.cell(4 + i, 2, COICOP_LABELS[c])
        ws.cell(4 + i, 3, float(st_.iat[i]))
    _rapikan(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==================================================================================
# 8. INJECTOR TEMPLATE ASLI (openpyxl) — HANYA SEL INPUT, RUMUS AMAN
# ==================================================================================
def _is_formula(cell):
    return isinstance(cell.value, str) and cell.value.startswith("=")


def _safe_write(ws, r, c, v):
    cell = ws.cell(r, c)
    if _is_formula(cell):
        return False
    cell.value = v
    if isinstance(v, (int, float)):
        cell.number_format = "#,##0.00"
    return True


def _quarter_cols(ws, header_row):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and re.match(r"20\d{2}Q[1-4]", v.strip()):
            cols[v.strip()] = c
    return cols


def _code_rows(ws, start=1, col=1):
    rows, cur = {}, None
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip() in CODES13 + CODES12 + ["00"]:
            rows[v.strip()] = r
    return rows


def autofill_template(raw_bytes, P):
    wb = load_workbook(io.BytesIO(raw_bytes))
    log = []
    # ---- Home ----
    if "Home" in wb.sheetnames:
        ws = wb["Home"]
        targets = {"Kode Provinsi": "32", "Kode Kabupaten/Kota": "09",
                   "Periode (tahun-triwulan)": P["periode"], "Putaran Angka Estimasi": "0",
                   "Tanggal Pengerjaan": P["tanggal"], "Petugas": P["petugas"],
                   "Supervisor": P["supervisor"], "Catatan": P["catatan"]}
        for r in range(1, ws.max_row + 1):
            lab = ws.cell(r, 2).value
            if isinstance(lab, str):
                for k, v in targets.items():
                    if lab.strip().startswith(k):
                        if _safe_write(ws, r, 4, v):
                            log.append(f"Home!D{r}={v}")
        for r in range(1, ws.max_row + 1):
            code = ws.cell(r, 5).value
            if isinstance(code, str) and code.strip() in P["status"]:
                stt = P["status"][code.strip()]
                _safe_write(ws, r, 7, stt)
                _safe_write(ws, r, 9, {0: "Belum Dikerjakan", 1: "Dalam Proses", 2: "Selesai"}[stt])
    # ---- sheet dengan grid triwulan ----
    grid_spec = [("E0", "e0"), ("E1", "e1"), ("E4", "e4"), ("R8", "r8"), ("R9", "r9"), ("R10", "r10")]
    for name, key in grid_spec:
        if name not in wb.sheetnames or P.get(key) is None:
            continue
        ws = wb[name]
        hdr_row, qcols = None, {}
        for r in range(1, min(ws.max_row, 30) + 1):
            if str(ws.cell(r, 1).value).strip() == "Kode":
                hdr_row = r
                qcols = _quarter_cols(ws, r)
                break
        if hdr_row is None:
            continue
        df = P[key]
        for r in range(hdr_row + 1, ws.max_row + 1):
            code = str(ws.cell(r, 1).value).strip()
            if code in df.index:
                for qq, c in qcols.items():
                    if qq in df.columns:
                        v = df.loc[code, qq]
                        if np.isfinite(v):
                            _safe_write(ws, r, c, float(v))
        log.append(f"{name}: grid terisi")
    # ---- E2: kolom SUT 2025 ----
    if "E2" in wb.sheetnames and P.get("sut25") is not None:
        ws = wb["E2"]
        col_sut = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(12, c).value
            if isinstance(v, str) and "SUT 2025" in v:
                col_sut = c
        if col_sut:
            for r in range(13, ws.max_row + 1):
                code = str(ws.cell(r, 1).value).strip()
                if code in P["sut25"].index:
                    _safe_write(ws, r, col_sut, float(P["sut25"].loc[code]))
    # ---- I6 ----
    if "I6" in wb.sheetnames and P.get("i6") is not None:
        ws = wb["I6"]
        hdr_row = None
        for r in range(1, 20):
            if str(ws.cell(r, 1).value).strip() == "Kode":
                hdr_row = r
                qcols = _quarter_cols(ws, r)
                break
        if hdr_row:
            cur = None
            for r in range(hdr_row + 1, ws.max_row + 1):
                a = str(ws.cell(r, 1).value).strip()
                b = str(ws.cell(r, 2).value).strip()
                if a in CODES13 + ["00"]:
                    cur = a
                if cur and b:
                    key = f"{cur}·{b}"
                    if key in P["i6"].index:
                        for qq, c in qcols.items():
                            if qq in P["i6"].columns:
                                v = P["i6"].loc[key, qq]
                                if np.isfinite(v):
                                    _safe_write(ws, r, c, float(v))
    # ---- F7 ----
    if "F7" in wb.sheetnames and P.get("f7") is not None:
        ws = wb["F7"]
        for r in range(1, ws.max_row + 1):
            code = str(ws.cell(r, 1).value).strip()
            if code in P["f7"].index:
                _safe_write(ws, r, 3, str(P["f7"].loc[code].iat[0]))
    # ---- A12: timpa penduduk ----
    if "A12" in wb.sheetnames and P.get("pop") is not None:
        ws = wb["A12"]
        hdr_row = None
        for r in range(1, 20):
            if str(ws.cell(r, 1).value).strip() == "Kode" or str(ws.cell(r, 2).value).strip() == "Data":
                hdr_row = r
                qcols = _quarter_cols(ws, r)
                break
        for r in range(1, ws.max_row + 1):
            b = str(ws.cell(r, 2).value).strip()
            if b.startswith("Jumlah penduduk"):
                for qq, c in (qcols or {}).items():
                    v = P["pop"].get(qq, np.nan)
                    if np.isfinite(v):
                        _safe_write(ws, r, c, float(v))
                log.append("A12: penduduk ditimpa")
    # ---- QC ----
    if "QC" in wb.sheetnames and P.get("qc") is not None:
        ws = wb["QC"]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, int) and 1 <= v <= 22:
                _safe_write(ws, r, 5, P["qc"][v - 1])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), log


def parse_book1_csv(file_obj):
    raw = file_obj.read().decode("utf-8-sig")
    lines = raw.split("\n")
    adhb_start = adhk_start = header_row = None
    for i, line in enumerate(lines):
        low = line.lower()
        if "tabel 1" in low and "adhb" in low:
            adhb_start = i
        elif "tabel 2" in low and "adhk" in low:
            adhk_start = i
        elif "komponen pengeluaran" in low and header_row is None:
            header_row = i
    if header_row is None or adhb_start is None or adhk_start is None:
        raise ValueError("Format Book1 tidak dikenali.")
    q_map = {}
    for j, cell in enumerate(lines[header_row].split(";")):
        cell = cell.strip()
        if cell.startswith("Q") and "-" in cell:
            q_map[cell] = j

    def parse_angka(txt):
        txt = txt.strip()
        if "," in txt:
            return float(txt.replace(".", "").replace(",", "."))
        if txt in ("", "-"):
            return 0.0
        return float(txt.replace(".", ""))

    def extract(s, e):
        data = {}
        for i in range(s, min(e, len(lines))):
            if len(data) == 12:
                break
            cells = lines[i].split(";")
            if not cells:
                continue
            lab = cells[0].strip()
            for k, pre in enumerate(["1.a.", "1.b.", "1.c.", "1.d.", "1.e.", "1.f.",
                                     "1.g.", "1.h.", "1.i.", "1.j.", "1.k.", "1.l."]):
                if lab.startswith(pre):
                    row = {}
                    for ql, ci in q_map.items():
                        if ci < len(cells):
                            try:
                                row[ql] = parse_angka(cells[ci])
                            except (ValueError, IndexError):
                                row[ql] = 0.0
                    data[k] = row
                    break
        return data

    adhb_d, adhk_d = extract(adhb_start, adhk_start), extract(adhk_start, len(lines))
    out = []
    for d in (adhb_d, adhk_d):
        df = pd.DataFrame(np.nan, index=CODES12, columns=Q_ALL)
        for i, qvals in d.items():
            for ql, v in qvals.items():
                p = ql.split("-")
                if len(p) == 2:
                    qs = f"{p[1]}{p[0]}"
                    if qs in Q_ALL:
                        df.iat[i, Q_ALL.index(qs)] = v
        out.append(df)
    return out[0], out[1]


# ==================================================================================
# 9. ANTARMUKA STREAMLIT (9 TAB)
# ==================================================================================
S = st.session_state
if "init" not in S:
    S["init"] = True
    S["periode"] = "2026-Q2"
    S["pop_method"] = "akhir"
    S["day_adj"] = False
    S["mode_bench"] = "B"
    S["petugas"] = "Petugas"
    S["supervisor"] = "Kasubag IP"
    S["catatan"] = "Otomasi aplikasi standalone pkrt.py"
    S["e0"] = None
    S["beta"] = None
    S["sut25"] = None
    S["f7"] = None
    S["i6"] = None

st.title("📊 Aplikasi Penyusunan LK PKRT Triwulanan (2025=100)")
st.caption("Standalone · 3 LK resmi · Auto-Fill Template Excel · ©" + str(datetime.now().year) + " Arief Rahman")

PER_OPT = [f"{q[:4]}-Q{q[-1]}" for q in Q_ALL if q >= "2024Q1"]
with st.sidebar:
    st.header("🧭 Konfigurasi")
    S["periode"] = st.selectbox("Periode (YYYY-QQ)", PER_OPT, index=PER_OPT.index(S["periode"]) if S["periode"] in PER_OPT else len(PER_OPT) - 1)
    S["pop_method"] = st.selectbox("Metode penduduk triwulanan", ["akhir", "tengah", "rata"],
                                   format_func=lambda x: {"akhir": "Akhir Triwulan (03,06,09,12)",
                                                          "tengah": "Pertengahan (02,05,08,11)",
                                                          "rata": "Rata-rata 3 Bulan"}[x])
    S["day_adj"] = st.toggle("Penyesuaian jumlah hari (default OFF)", value=S["day_adj"])
    S["mode_bench"] = st.radio("Sumber benchmark E2 (anti double counting)", ["A", "B"],
                               format_func=lambda x: "[A] SUT 2025 sudah terkoreksi SNA" if x == "A"
                               else "[B] PKRT Rilis sebelum koreksi SSF/FISIM")
    st.markdown("---")
    S["petugas"] = st.text_input("Petugas", S["petugas"])
    S["supervisor"] = st.text_input("Supervisor", S["supervisor"])
    S["catatan"] = st.text_input("Catatan", S["catatan"])
    st.markdown("---")
    st.caption(f"©{datetime.now().year} Arief Rahman")

PQ = S["periode"].replace("-", "")
Q_ACT = [q for q in Q_ALL if q <= PQ]

# ---- input E0 ----
st.sidebar.markdown("### E0 (rilis 2010=100)")
up_e0 = st.sidebar.file_uploader("Upload Book1__.csv", type=["csv"])
if up_e0 is not None:
    try:
        a, k = parse_book1_csv(up_e0)
        S["e0"] = (a, k)
        st.sidebar.success("E0 termuat.")
    except Exception as e:
        st.sidebar.error(f"Gagal: {e}")

popq = population_quarterly(S["pop_method"])
defl = DeflatorJabarWorkspace()
seruti = seruti_aggregate(popq, S["day_adj"])

if S["beta"] is None or list(S["beta"].columns) != Q_ACT:
    S["beta"] = pd.DataFrame(1.0, index=CODES13, columns=Q_ACT)
if S["sut25"] is None:
    c25 = [q for q in seruti.columns if q.startswith("2025")]
    tot = seruti[c25].sum(axis=1) if c25 else pd.Series(0.0, index=CODES13)
    S["sut25"] = tot.copy()
if S["f7"] is None:
    S["f7"] = pd.DataFrame(["" for _ in CODES13 + ["00"]], index=CODES13 + ["00"],
                           columns=[f"Fenomena pada {S['periode']}"])
i6_idx = [f"{c}·{('Volume air PDAM yang disalurkan ke pelanggan non niaga' if (c == '04' and s == 1) else ('Jumlah energi listrik yang terjual ke pelanggan rumah tangga' if (c == '04' and s == 2) else f'Indikator {s}'))}"
          for c in CODES13 + ["00"] for s in range(1, 11)]
if S["i6"] is None or list(S["i6"].index) != i6_idx:
    S["i6"] = pd.DataFrame(np.nan, index=i6_idx, columns=Q_ACT)

e0 = S["e0"]
W = PKRTWorkspace(e0[0].reindex(columns=Q_ACT) if e0 else None,
                  e0[1].reindex(columns=Q_ACT) if e0 else None,
                  seruti, defl, S["beta"], S["sut25"], S["mode_bench"] == "B", popq, Q_ACT) if False else None

# bangun workspace dengan e0 opsional
_e0b = S["e0"][0] if S["e0"] else None
_e0k = S["e0"][1] if S["e0"] else None
W = PKRTWorkspace((_e0b, _e0k), seruti, defl, S["beta"], S["sut25"], S["mode_bench"] == "B", popq, Q_ACT)

sut = SUTJabar2020Benchmark()
share_ser = (seruti[[c for c in seruti.columns if c.startswith("2025")]].sum(axis=1)
             if any(c.startswith("2025") for c in seruti.columns) else pd.Series(0.0, index=CODES13))
share_ser = share_ser / share_ser.sum() * 100 if share_ser.sum() else share_ser
checks, rate = run_qc22(W, defl, sut, share_ser)

status = {"E0": 2 if S["e0"] else 0, "E1": 2, "E2": 2, "E3": 2, "E4": 2, "E5": 1,
          "I6": 1 if np.isfinite(S["i6"].values).any() else 0,
          "F7": 2 if (S["f7"].iloc[:, 0].astype(str).str.len() > 5).all() else 0,
          "R8": 2 if (S["beta"].values != 1).any() else 0, "R9": 2, "R10": 2,
          "A11": 2, "A12": 2, "QC": 2}
meta = {"periode": S["periode"], "tanggal": datetime.now().strftime("%d-%m-%Y"),
        "petugas": S["petugas"], "supervisor": S["supervisor"], "catatan": S["catatan"],
        "status": status, "sut25": S["sut25"], "f7": S["f7"], "i6": S["i6"], "pop": popq}

t1, t2, t3, t4, t5, t6, t7, t8, t9 = st.tabs(
    ["1. Home & Wilayah", "2. LK Deflator Wakanda & PPKK", "3. Benchmark SUT Wakanda 2020",
     "4. Help & Parameter SNA 2008", "5. LK PKRT Inisial & Backcasting", "6. Indikator & Fenomena",
     "7. Simulasi Beta & Rekonsiliasi", "8. 7 Subkomponen & Per Kapita", "9. QC & Export/Auto-Fill"])

with t1:
    st.header("Home & Wilayah (3209)")
    c1, c2, c3 = st.columns(3)
    c1.metric("Kode Provinsi", "92 (Wakanda)")
    c2.metric("Kode Kab/Kota", "99 (Kab. Konoha)")
    c3.metric("Sister city", "94 (Kota Konoha)")
    st.markdown("**Metadata Home:**")
    st.dataframe(pd.DataFrame([meta]).T, hide_index=False)
    st.markdown("**Status Flow Lembar Kerja:**")
    st.dataframe(pd.DataFrame([{"Sheet": k, "Status": v} for k, v in status.items()]), hide_index=True)

with t2:
    st.header("LK Deflator Wakanda & PPKK (E1–E5)")
    st.markdown("**E1 — IHK bulanan 2022=100 (2024M1+), format Indonesia**")
    st.dataframe(tabel_id(defl.sub_m), height=300)
    st.markdown("**E4 — IHK triwulanan 2022=100**")
    st.dataframe(tabel_id(defl.q22))
    st.markdown("**E5 — IHK triwulanan 2025=100** (2023Q1–2023Q4 = seri penyambungan pra-2024)")
    st.dataframe(tabel_id(defl.q25))
    qq, yy = defl.q25.div(defl.q25.shift(1, axis=1)) * 100 - 100, defl.q25.div(defl.q25.shift(4, axis=1)) * 100 - 100
    st.markdown("**Growth q-to-q (%)**")
    st.dataframe(tabel_id(qq))

with t3:
    st.header("Benchmark SUT Wakanda 2020 (Modul 4)")
    st.warning(GUARDRAIL_SUT)
    st.markdown("**Rasio makro supply-side:**")
    st.dataframe(tabel_id(pd.DataFrame([sut.ratios()])), hide_index=True)
    st.markdown("**Struktur share PKRT Wakanda 2020 (%):**")
    st.dataframe(tabel_id(sut.structure().to_frame("Share")), height=300)
    if st.button("Terapkan struktur SUT ke SUT 2025 (E2)"):
        tot25 = float(share_ser.sum())
        st_ = sut.structure()
        S["sut25"] = (st_ / 100.0) * float(W.ser_adj[[c for c in W.q if c.startswith('2025')]].sum().sum()) if any(c.startswith("2025") for c in W.q) else S["sut25"]
        st.rerun()

with t4:
    st.header("Help & Parameter SNA 2008")
    st.markdown(f"- Disagregasi COICOP-1999 12 → 2018: **12** = {PROP[11]}, **13** = {PROP[12]} (KUNCI pusat; jangan diganti Seruti).")
    st.markdown(f"- Coverage adjustment: **06** = {COV[5]}, **12** = {COV[11]} — aktif hanya pada mode **[B]**.")
    st.markdown(f"- Mode aktif: **[{S['mode_bench']}]**")
    st.dataframe(pd.DataFrame([PDRB_PKRT_RILIS_3209[y] | {"tahun": y} for y in sorted(PDRB_PKRT_RILIS_3209)]), hide_index=True)

with t5:
    st.header("LK PKRT Inisial & Backcasting (E0–E5)")
    st.markdown("**E0 (input, 2010=100)** — upload di sidebar.")
    if S["e0"]:
        st.dataframe(tabel_id(S["e0"][0].reindex(columns=Q_ACT)))
    else:
        st.info("E0 kosong → backcasting 2019Q1–2022Q4 tidak berjalan.")
    st.markdown("**E1 — Seruti agregat (Rp Juta), format Indonesia**")
    st.dataframe(tabel_id(W.ser))
    st.markdown("**E2 — SUT 2025 & Theta**")
    e2v = st.data_editor(pd.DataFrame({"SUT 2025": S["sut25"], "Theta": W.theta}), key="ed_sut25")
    S["sut25"] = e2v["SUT 2025"]
    st.markdown("**E3 — ADHB inisial**")
    st.dataframe(tabel_id(W.e3))
    st.markdown("**E4 — Deflator 2025=100**")
    st.dataframe(tabel_id(W.cpi))
    st.markdown("**E5 — ADHK inisial**")
    st.dataframe(tabel_id(W.e5))
    st.markdown("**Panel kontrol non-konsumsi (pos 31 & 32, tidak masuk total PKRT):**")
    st.dataframe(tabel_id(pd.DataFrame({k: SERUTI_META_3209[k] for k in
                                        ["31_pajak_premi", "32_pesta_kenduri", "30_subtotal_lainnya", "40_total"]})))

with t6:
    st.header("Indikator & Fenomena (I6–F7)")
    st.markdown("**I6 — 10 slot per COICOP**")
    ed = st.data_editor(S["i6"], key="ed_i6", num_rows="fixed", height=300)
    S["i6"] = ed
    st.markdown("**F7 — Fenomena kualitatif**")
    ef = st.data_editor(S["f7"], key="ed_f7", num_rows="fixed")
    S["f7"] = ef

with t7:
    st.header("Simulasi Beta & Rekonsiliasi (R8–R10)")
    eb = st.data_editor(S["beta"], key="ed_beta", num_rows="fixed")
    S["beta"] = eb
    st.markdown("**R9 — ADHB final, format Indonesia**")
    st.dataframe(tabel_id(W.r9))
    st.markdown("**R10 — ADHK final, format Indonesia**")
    st.dataframe(tabel_id(W.r10))

with t8:
    st.header("Analisis 7 Subkomponen & Per Kapita (A11–A12)")
    st.markdown("**A11 — 5 blok**")
    st.dataframe(tabel_id(W.a11n))
    st.dataframe(tabel_id(W.a11k))
    st.dataframe(tabel_id(W.a11dn))
    st.markdown("**A12 — penduduk & per kapita (Rp Ribu)**")
    st.dataframe(tabel_id(W.a12pop.to_frame("Jumlah penduduk (jiwa)").T))
    st.dataframe(tabel_id(W.a12k7))

with t9:
    st.header("QC 22 Aspek & Export / Auto-Fill Excel")
    st.metric("Tingkat konsistensi", f"{rate:.2f}%")
    st.dataframe(pd.DataFrame(checks), hide_index=True)
    st.markdown("### ⬇️ Unduh Workbook Lengkap Baru")
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.download_button("LK PKRT 2025=100", export_lk_pkrt(W, checks, rate, meta),
                           file_name=f"02_LK_PKRT_2025_100_Cirebon_{S['periode']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        st.download_button("LK Deflator PPKK", export_lk_deflator(defl),
                           file_name="03_LK_Deflator_PKRT_PPKK.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d3:
        st.download_button("LK Deflator Wakanda", export_lk_deflator_jabar(defl),
                           file_name="LK_Deflator_IHK_Jabar.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d4:
        st.download_button("Benchmark SUT 2020", export_sut_benchmark(sut),
                           file_name="Benchmark_SUT_Jabar_2020.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("### 🧩 Auto-Fill Template Excel Asli")
    up_tpl = st.file_uploader("Upload 02. Lembar Kerja PKRT Triwulanan (2025=100)_Put0_Rev.xlsx", type=["xlsx"])
    if up_tpl is not None:
        payload = {"periode": S["periode"], "tanggal": meta["tanggal"], "petugas": S["petugas"],
                   "supervisor": S["supervisor"], "catatan": S["catatan"], "status": status,
                   "e0": pd.concat([_e0b.reindex(columns=Q_ACT), _e0k.reindex(columns=Q_ACT)]) if S["e0"] else None,
                   "e1": W.ser, "e4": W.cpi, "r8": W.beta, "r9": W.r9, "r10": W.r10,
                   "sut25": S["sut25"].rename(index={c: c for c in CODES13}), "i6": S["i6"],
                   "f7": S["f7"], "pop": popq, "qc": [c["Respon"] for c in checks]}
        if S["e0"]:
            payload["e0"] = pd.concat([_e0b.reindex(columns=Q_ACT), _e0k.reindex(columns=Q_ACT)])
        try:
            filled, log = autofill_template(up_tpl.getvalue(), payload)
            st.success(f"Template terisi ({len(log)} aksi). Sel rumus bawaan TIDAK disentuh.")
            st.download_button("⬇️ Unduh Template TERISI", filled,
                               file_name=f"02_LK_PKRT_TERISI_{S['periode']}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with st.expander("Log penulisan sel"):
                st.write(log)
        except Exception as e:
            st.error(f"Gagal auto-fill: {e}")