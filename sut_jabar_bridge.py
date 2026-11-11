# -*- coding: utf-8 -*-
"""
sut_jabar_bridge.py — Jembatan SUT Jabar 2020 -> 13 COICOP (metode STRUKTUR/SHARE)

Sumber : SUT_3200_2020_0.csv, Tabel 2 (Penggunaan ADHB), kolom 3011
         "Konsumsi Rumah Tangga"
Checksum: total PKRT Jabar 2020 = 1.478.218.353 (Juta Rupiah)

Prinsip: total konsumsi TETAP dari Seruti Kab. Cirebon.
         SUT Jabar hanya mengoreksi KOMPOSISI antar-divisi COICOP.
"""

import io
import re
import csv
import numpy as np
import pandas as pd
import streamlit as st

EXPECTED_TOTAL = 1_478_218_353.0
COICOP_KODE = [f"{i + 1:02d}" for i in range(13)]

COMMODITY = {
    1: "Padi", 2: "Jagung", 3: "Umbi-umbian/Palawija", 4: "Kacang-kacangan & Serelia",
    5: "Hortikultura Semusim", 6: "Perkebunan Semusim", 7: "Hortikultura Tahunan",
    8: "Karet", 9: "Kelapa Sawit", 10: "Perkebunan Tahunan Lainnya",
    11: "Ternak & Hasilnya", 12: "Jasa Pertanian & Perburuan",
    13: "Hasil Hutan & Kayu", 14: "Ikan & Biota Air", 15: "Minyak Mentah & Kondensat",
    16: "Gas Alam", 17: "Tenaga Panas Bumi", 18: "LNG", 19: "Batubara & Lignit",
    20: "Bijih Tembaga", 21: "Bijih Logam Mulia", 22: "Tambang Logam Lainnya",
    23: "Tambang & Galian Lainnya", 24: "Produk dari Batubara",
    25: "Hasil Kilang Minyak, LPG, Pelumas", 26: "Produk Makanan & Minuman",
    27: "Pengolahan Tembakau", 28: "Tekstil & Pakaian Jadi",
    29: "Kulit & Alas Kaki", 30: "Produk Kayu/Gabus/Anyaman",
    31: "Kertas & Percetakan", 32: "Kimia, Farmasi, Obat Tradisional",
    33: "Karet & Plastik", 34: "Galian bukan Logam", 35: "Logam Dasar",
    36: "Brg Logam/Komputer/Elektronik/Optik", 37: "Mesin & Perlengkapan",
    38: "Alat Angkutan", 39: "Furnitur", 40: "Pengolahan Lainnya + Reparasi",
    41: "Ketenagalistrikan", 42: "Gas Buatan/Uap/Es", 43: "Air, Sampah, Limbah",
    44: "Konstruksi", 45: "Perdagangan & Reparasi Mobil/Motor",
    46: "Perdagangan Besar & Eceran", 47: "Angkutan Rel", 48: "Angkutan Darat",
    49: "Angkutan Laut", 50: "Angkutan Sungai/Danau", 51: "Angkutan Udara",
    52: "Penunjang Angkutan, Pos, Kurir", 53: "Penyediaan Akomodasi",
    54: "Penyediaan Makan Minum", 55: "Jasa Informasi & Komunikasi",
    56: "Keuangan Perbankan", 57: "Asuransi & Dana Pensiun",
    58: "Keuangan Lainnya", 59: "Penunjang Keuangan", 60: "Real Estate",
    61: "Jasa Perusahaan", 62: "Pemerintahan Umum", 63: "Jasa Pendidikan",
    64: "Jasa Kesehatan & Sosial", 65: "Jasa Lainnya",
}

# Kelompok A — tidak masuk PKRT (16 komoditas)
EXCLUDED = {8, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 35, 46, 61, 62}

# Kelompok B — mapping langsung (36 komoditas), nilai = indeks COICOP 0..12
DIRECT = {
    1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 9: 0, 10: 0, 11: 0, 14: 0,
    13: 3,
    27: 1,
    29: 2,
    37: 4, 39: 4,
    38: 6, 45: 6, 47: 6, 48: 6, 49: 6, 50: 6, 51: 6,
    41: 3, 42: 3, 43: 3, 44: 3, 60: 3,
    53: 10, 54: 10,
    56: 11, 57: 11, 58: 11, 59: 11,
    63: 9,
    64: 5,
}

# Kelompok C — split default (13 komoditas). ASUMSI, bukan angka resmi BPS.
SPLIT_DEFAULT = {
    25: {3: 0.25, 6: 0.75},
    26: {0: 0.99, 1: 0.01},
    28: {2: 0.90, 4: 0.10},
    30: {4: 0.70, 3: 0.30},
    31: {9: 0.60, 8: 0.40},
    32: {5: 0.55, 12: 0.30, 4: 0.15},
    33: {4: 0.70, 2: 0.20, 12: 0.10},
    34: {4: 0.60, 3: 0.40},
    36: {4: 0.50, 7: 0.40, 5: 0.10},
    40: {12: 0.60, 8: 0.25, 4: 0.15},
    52: {6: 0.70, 7: 0.30},
    55: {7: 0.85, 8: 0.15},
    65: {12: 0.60, 8: 0.40},
}


# ============================ PARSER ============================
def _num(x):
    """Konversi sel SUT: '  1,234 ', '-', '(567)' -> float."""
    s = str(x).strip().strip('"').strip()
    if s in ("", "-", "- 0", "-0", "nan", "None", "#VALUE!", "#ERROR!"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").replace(" ", "")
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def parse_sut(raw_bytes):
    """Baca kolom 3011 (Konsumsi Rumah Tangga) Tabel 2. Return (dict, total_ref, col)."""
    text = None
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("File tidak bisa dibaca sebagai teks CSV.")

    rows = list(csv.reader(io.StringIO(text)))

    hdr, col = None, None
    for i, r in enumerate(rows):
        cells = [c.strip() for c in r]
        if "1800" in cells and "3011" in cells:
            hdr, col = i, cells.index("3011")
            break
    if hdr is None:
        raise ValueError(
            "Header Tabel 2 (kolom 1800 & 3011) tidak ditemukan. "
            "Pastikan file yang diunggah adalah SUT_3200_2020_0.csv, "
            "bukan SUT_2020_3200.csv (kolom U-3011 kosong)."
        )

    vals, total_ref = {}, None
    for r in rows[hdr + 1:]:
        if len(r) <= col:
            continue
        head = " ".join(r[:3])
        m = re.search(r"\bP-(\d{2})\b", head)
        if m:
            k = int(m.group(1))
            if 1 <= k <= 65 and k not in vals:
                vals[k] = _num(r[col])
            continue
        if any(c.strip() == "1900" for c in r[:4]):
            total_ref = _num(r[col])
            break

    if len(vals) < 65:
        missing = sorted(set(range(1, 66)) - set(vals))
        raise ValueError(f"Komoditas tidak terbaca: {missing}")
    return vals, total_ref, col


# ============================ AGREGASI ============================
def split_table_to_df(tbl):
    rec = []
    for k in sorted(tbl):
        for ci, w in tbl[k].items():
            rec.append({
                "Kode": f"P-{k:02d}",
                "Komoditas": COMMODITY[k],
                "COICOP": COICOP_KODE[ci],
                "Bobot": float(w),
            })
    return pd.DataFrame(rec)


def df_to_split_table(df):
    """Normalisasi bobot per komoditas. Fallback ke default bila tidak valid."""
    tbl = {}
    for _, r in df.iterrows():
        try:
            k = int(str(r["Kode"]).split("-")[1])
            ci = COICOP_KODE.index(str(r["COICOP"]).strip().zfill(2))
            w = float(r["Bobot"])
        except (ValueError, IndexError, KeyError):
            continue
        if not np.isfinite(w) or w < 0:
            continue
        tbl.setdefault(k, {})[ci] = tbl.get(k, {}).get(ci, 0.0) + w

    fixed = []
    for k in SPLIT_DEFAULT:
        d = tbl.get(k, {})
        s = sum(d.values())
        if s <= 0:
            tbl[k] = dict(SPLIT_DEFAULT[k])
            fixed.append(f"P-{k:02d}")
        else:
            tbl[k] = {ci: w / s for ci, w in d.items()}
    return tbl, fixed


def aggregate(vals, split_tbl):
    out = np.zeros(13)
    for k, v in vals.items():
        if k in EXCLUDED:
            continue
        if k in DIRECT:
            out[DIRECT[k]] += v
        elif k in split_tbl:
            for ci, w in split_tbl[k].items():
                out[ci] += v * w
    return out


# ============================ UI ============================
def render_sut_bridge():
    st.subheader("📗 SUT Jawa Barat 2020 → Struktur 13 COICOP")
    st.caption(
        "Metode SHARE: total konsumsi tetap dari Seruti Kab. Cirebon; "
        "SUT Jabar hanya mengoreksi komposisi antar-divisi."
    )
    st.warning(
        "⚠️ Bobot split Kelompok C adalah **asumsi**, bukan angka resmi BPS. "
        "Struktur tahun 2020 dibekukan untuk semua periode. "
        "Mohon direview petugas yang paham Neraca Nasional sebelum dipakai untuk rilis resmi."
    )

    up = st.file_uploader(
        "Unggah SUT_3200_2020_0.csv", type=["csv"], key="sut_file"
    )
    if up is None:
        st.info("Belum ada file. Gunakan **SUT_3200_2020_0.csv** (bukan SUT_2020_3200.csv).")
        return

    try:
        vals, total_ref, col = parse_sut(up.getvalue())
    except Exception as e:
        st.error(f"Gagal membaca SUT: {e}")
        return

    total = float(sum(vals.values()))
    ref = float(total_ref) if total_ref else EXPECTED_TOTAL

    c1, c2, c3 = st.columns(3)
    c1.metric("Total 65 komoditas", f"{total:,.0f}")
    c2.metric("Baris total (1900)", f"{ref:,.0f}")
    c3.metric("Selisih", f"{total - ref:,.0f}")

    ok = abs(total - ref) <= max(1.0, abs(ref) * 1e-6)
    if ok:
        st.success(f"✅ Checksum cocok (kolom ke-{col}). Kolom yang dibaca benar.")
    else:
        st.error(
            "❌ Checksum GAGAL — kemungkinan salah kolom atau salah file. "
            "Jangan lanjutkan; kirimkan pesan ini ke pengembang."
        )
        return

    dropped = {k: v for k, v in vals.items() if k in EXCLUDED and abs(v) > 0}
    if dropped:
        pct = sum(abs(v) for v in dropped.values()) / total * 100 if total else 0.0
        st.warning(
            f"Ada {len(dropped)} komoditas dikecualikan tetapi bernilai RT "
            f"({pct:.2f}% dari total). Share dinormalisasi ulang atas sisanya."
        )
        st.dataframe(
            pd.DataFrame([
                {"Kode": f"P-{k:02d}", "Komoditas": COMMODITY[k], "Nilai RT (Jt Rp)": v}
                for k, v in sorted(dropped.items())
            ]),
            hide_index=True,
        )

    # ---------- bobot split ----------
    st.markdown("#### Bobot split Kelompok C")
    if st.button("↩️ Reset bobot ke default", key="sut_reset"):
        st.session_state["sut_split_df"] = split_table_to_df(SPLIT_DEFAULT)
        if "sut_ed_split" in st.session_state:
            del st.session_state["sut_ed_split"]
        st.rerun()

    if "sut_split_df" not in st.session_state:
        st.session_state["sut_split_df"] = split_table_to_df(SPLIT_DEFAULT)

    edited = st.data_editor(
        st.session_state["sut_split_df"],
        key="sut_ed_split",
        num_rows="fixed",
        disabled=["Kode", "Komoditas", "COICOP"],
        column_config={
            "Bobot": st.column_config.NumberColumn(
                min_value=0.0, max_value=1.0, step=0.05, format="%.2f"
            )
        },
        hide_index=True,
    )
    st.session_state["sut_split_df"] = edited
    st.caption("Hanya kolom **Bobot** yang bisa diubah. Otomatis dinormalisasi per komoditas (jumlah = 1).")

    split_tbl, fixed = df_to_split_table(edited)
    if fixed:
        st.info(f"Bobot tidak valid/kosong pada {', '.join(fixed)} → dikembalikan ke default.")

    vec = aggregate(vals, split_tbl)
    if vec.sum() <= 0:
        st.error("Hasil agregasi nol. Periksa bobot split.")
        return
    share = vec / vec.sum()

    # ---------- bandingkan dengan Seruti ----------
    st.markdown("#### Struktur SUT vs Seruti Cirebon")
    S = st.session_state
    ser = S.get("e1_seruti")
    if ser is None:
        st.info("Muat dulu Seruti (tombol **📊 Data Riil** di sidebar) untuk melihat perbandingan & θ.")
        _show_mapping(vals, split_tbl)
        return

    q25 = [c for c in ser.columns if str(c).startswith("2025")]
    if not q25:
        st.warning("Tidak ada triwulan 2025 pada Seruti.")
        _show_mapping(vals, split_tbl)
        return

    ser25 = ser.loc[:, q25].astype(float).sum(axis=1)
    tot_ser = float(ser25.sum())
    if not np.isfinite(tot_ser) or tot_ser <= 0:
        st.warning("Total Seruti 2025 kosong/nol. Muat data riil dulu.")
        _show_mapping(vals, split_tbl)
        return

    cov = {5: 0.4980, 11: 1.1007}
    ser_adj = ser25.copy()
    for i, f in cov.items():
        ser_adj.iloc[i] = ser_adj.iloc[i] * f

    target = share * tot_ser
    theta = target / ser_adj.replace(0, np.nan)
    share_ser = ser25 / tot_ser

    cmp = pd.DataFrame({
        "COICOP": COICOP_KODE,
        "Rincian": list(ser.index),
        "Share Seruti (%)": (share_ser * 100).round(2).values,
        "Share SUT (%)": (share * 100).round(2).values,
        "Selisih (pp)": ((share - share_ser) * 100).round(2).values,
        "Target SUT (Rp Juta)": np.round(target, 1),
        "Theta": theta.round(4).values,
    })
    st.dataframe(cmp, hide_index=True)

    big = cmp.loc[cmp["Selisih (pp)"].abs() > 10, "COICOP"].tolist()
    if big:
        st.warning(
            f"⚠️ Divisi {', '.join(big)} bergeser >10 poin persen. "
            "Periksa dulu mapping/bobot split-nya sebelum menerapkan."
        )

    st.caption(
        f"Total Seruti 2025 (basis level) = **{tot_ser:,.1f}** Rp Juta. "
        "θ Kesehatan ≈ 2 karena menetralkan coverage adjustment 0,498 — itu normal."
    )

    if st.button("✅ Terapkan ke E2 (SUT 2025)", type="primary", key="sut_apply"):
        if "e2_sut" not in S:
            st.error("Sheet E2 belum ada di session. Muat data riil / simulasi dulu.")
            return
        idx = S["e2_sut"].index
        colname = S["e2_sut"].columns[0]
        S["e2_sut"] = pd.DataFrame({colname: np.round(target, 1)}, index=idx)
        for k in list(st.session_state.keys()):
            if str(k).startswith("ed_"):
                del st.session_state[k]
        tag = " + struktur SUT Jabar 2020"
        cur = str(S.get("sumber_data", ""))
        if tag not in cur:
            S["sumber_data"] = cur + tag
        st.success("E2 terisi. Buka Workbook 2 → E2/Theta untuk memeriksa.")
        st.rerun()

    _show_mapping(vals, split_tbl)


def _show_mapping(vals, split_tbl):
    with st.expander("🔍 Tabel mapping lengkap 65 komoditas"):
        rec = []
        for k in range(1, 66):
            if k in EXCLUDED:
                ket, tgt = "Dikecualikan", "—"
            elif k in DIRECT:
                ket, tgt = "Langsung", COICOP_KODE[DIRECT[k]]
            else:
                ket = "Split (asumsi)"
                tgt = " + ".join(
                    f"{COICOP_KODE[ci]}:{w:.0%}" for ci, w in split_tbl.get(k, {}).items()
                )
            rec.append({
                "Kode": f"P-{k:02d}",
                "Komoditas": COMMODITY[k],
                "Nilai RT (Jt Rp)": round(vals.get(k, 0.0), 1),
                "Jenis": ket,
                "Tujuan": tgt,
            })
        st.dataframe(pd.DataFrame(rec), hide_index=True, height=420)