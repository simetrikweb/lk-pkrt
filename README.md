# 📊 Aplikasi Penyusunan 3 Lembar Kerja PKRT & Deflator (2025=100)

Aplikasi web **standalone** untuk mengotomasi **end-to-end** seluruh alur kerja **Rebasing PDRB Pengeluaran Tahun Dasar 2025=100** sesuai standar lembar kerja.

> 🎯 **Tujuan akhir:** menghasilkan angka **PKRT 1.a s/d 1.m** (sheet **ADHB & ADHK**,
> triwulan **2019Q1–2030Q4**, satuan juta rupiah, tahun dasar **2025=100**).

---


## 🧩 Fitur Utama

1. **9 tab navigasi bertahap** sesuai alur kerja.
2. **Backcasting otomatis dengan penanda seri penyambungan 2023.
4. **Sakelar anti double-counting:**
6. **Export workbook baru bebas #DIV/0!:**
7. **QC 22 aspek** konsistensi internal & eksternal
8. **Tampilan angka kaidah Indonesia** (`5.976.606,65`) di seluruh tabel layar.
9. **Beberapa Opsi metodologi:** 

---

## 📁 Struktur Proyek

```
lk-pkrt-standalone/
├── pkrt.py     # SELURUH aplikasi (satu file, self-contained)
├── README.md               # Dokumentasi ini
├── .gitignore              # Aturan pengabaian file untuk Git
└── (opsional) data/        # File input lokal: Book1__.csv, template xlsx, dll.
```

Tidak memerlukan file modul tambahan apa pun.

---

## ⚙️ Prasyarat & Instalasi

- Python **3.9+**
- Paket:

```bash
pip install streamlit pandas numpy openpyxl
```

---

## 🚀 Menjalankan Aplikasi

```bash
streamlit run pkrt.py
```

Browser terbuka di `http://localhost:8501`.

---


## 🔐 Keamanan Data

Seluruh file data (`.csv`, `.xlsx`, `.xls`) **di-ignore oleh Git** (lihat `.gitignore`)
sehingga tidak ikut ter-upload ke GitHub. Bagikan data melalui media internal instansi.

---

## 📄 Lisensi & Hak Cipta

©2026 **Arief Rahman** — dikembangkan untuk keperluan kompilasi PDRB.
Penggunaan, modifikasi, dan distribusi internal diperbolehkan dengan mencantumkan sumber.

---

## 🙏 Ucapan Terima Kasih & Catatan

- Metodologi mengikuti lembar kerja dan kaidah SNA 2008.
- Hasil aplikasi adalah **estimasi**; verifikasi akhir tetap oleh kompilator berwenang.