# 📊 Otomasi Lembar Kerja PKRT & Deflator — Kab. Konoha (2025=100)

Aplikasi berbasis **Streamlit** untuk mengotomasi penghitungan **Pengeluaran Konsumsi Rumah Tangga (PKRT)** dan **Deflator IHK** sesuai standar dengan tahun dasar **2025=100** (rebasing). Dikembangkan khusus untuk **Kabupaten Konoha (kode XXXX)**.

> 🎯 **Tujuan utama:** Menghasilkan file Excel berformat **1.a s/d 1.m** (sheet ADHB & ADHK) untuk triwulan **2019Q1 – 2025Q4**, siap diintegrasikan dengan aplikasi lanjutan PDRB.

---

## ✨ Fitur Utama

| No | Fitur | Keterangan |
|----|-------|------------|
| 1 | **Upload Book1__.csv** | Impor rilis PDRB basis 2010=100 (12 divisi) untuk backcasting 2019–2022 |
| 2 | **Upload Template PDRB (.xlsx)** | Impor template 13 divisi (1.a–1.m) langsung ke E0 |
| 3 | **Data Riil Konoha** | Seruti per kapita, IHK Kota Konoha, & penduduk proyeksi SP2020 sudah tertanam |
| 4 | **Workbook 1 (Deflator)** | IHK bulanan → triwulanan → re-referencing 2025=100 |
| 5 | **Workbook 2 (PKRT)** | Theta, Beta, ADHB/ADHK inisial & final, agregasi 7 subkomponen, per kapita |
| 6 | **QC 22 Butir** | Pemeriksaan konsistensi internal & eksternal otomatis |
| 7 | **Export Excel Multi-sheet** | WB1 (Deflator) & WB2 (PKRT) lengkap dengan rumus |
| 8 | **Export Format 1.a–1.m** | Output akhir sheet ADHB & ADHK sesuai template PDRB |
| 9 | **SUT Wakanda Bridge** | Koreksi komposisi 13 COICOP dari SUT Wakanda 2020 (opsional) |
| 10 | **Simpan/Muat JSON** | Backup & restore seluruh pekerjaan |
| 11 | **Format Angka Indonesia** | Tampilan `5.976.606,65` sesuai kaidah |

---

## 📁 Struktur File
