# AREKO BPR - Rekapitulasi Mutasi Rekening Koran BPR

Aplikasi cerdas dan otomatis untuk menganalisis, mengekstraksi, dan merekapitulasi data rekening koran **BPR (PT BPR MITRATAMA ARTHABUANA & Bank Perkreditan Rakyat lainnya)** dari berkas PDF menjadi Formulir Validasi Mutasi Rekening resmi (Format PDF A4 siap cetak dan Microsoft Excel Spreadsheet dinamis).

---

## ✨ Fitur Utama

1. **Engine Parsing Khusus BPR**:
   - Ekstraksi tabel multi-halaman berkecepatan tinggi dengan pemilahan transaksi per bulan kalender secara otomatis.
   - Penanganan pembagian mutasi multi-bulan dalam satu berkas PDF (misal: berkas mutasi Juni-Juli atau Agustus-September).
   - Pengenalan cerdas baris *Saldo Pindahan* (Opening Balance) sehingga mutasi kredit tidak terhitung ganda (*zero duplicate turnover*).
   - Penanganan carryover saldo penutup bulan sebelumnya untuk saldo pembuka bulan baru.
   - Rekonstruksi dan validasi saldo berjalan (*running balance*) secara matematis akurat 100% pas terhadap saldo awal (*Opening Balance*) dan saldo akhir (*Closing Balance*).
   - Fallback OCR otomatis beresolusi tinggi (Tesseract OCR Engine) untuk dokumen hasil scan.

2. **Ringkasan Mutasi Komprehensif**:
   - Menghitung frekuensi Debet & Kredit per bulan.
   - Menghitung total perputaran dana mutasi Debet & Kredit per bulan.
   - Menganalisis saldo tertinggi (max), saldo terendah (min), dan saldo rata-rata (average) per bulan.
   - Menghitung baris **Rata-Rata** keseluruhan bulan yang diunggah.

3. **Ekspor Dokumen Standar Perbankan**:
   - **PDF Resmi**: Dokumen A4 siap cetak dilengkapi informasi nasabah, nomor rekening, catatan Operation Head, dan kolom tanda tangan Sales Officer & Operation Head.
   - **Excel Spreadsheet (.xlsx)**: Lembar kerja dinamis lengkap dengan tabel rekapitulasi utama dan lembar rincian transaksi per bulan.

4. **Desain Antarmuka Modern & Bersih**:
   - Tampilan tema terang (*Light Mode*) yang profesional dan responsif.
   - Tabel pratinjau interaktif dengan penanda visual khusus warna hijau lembut (`#E2EFDA`) pada baris rata-rata.

---

## 🛠️ Instalasi & Menjalankan Lokal

### 1. Kebutuhan Sistem
Pastikan sistem Anda telah memiliki **Python 3.9+** serta dependensi sistem `tesseract-ocr` dan `poppler-utils`:

- **macOS (via Homebrew)**:
  ```bash
  brew install tesseract poppler
  ```
- **Ubuntu/Debian / Streamlit Cloud**:
  Sudah terkonfigurasi di `packages.txt` (`tesseract-ocr`, `poppler-utils`).

### 2. Instal Dependensi Python
```bash
cd BPR
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi
```bash
streamlit run app.py
```

Akses aplikasi melalui browser di `http://localhost:8501`.

---

## ☁️ Panduan Deploy ke Streamlit Cloud

1. Buat repository baru di GitHub (misalnya: `AREKO-BPR`).
2. Hubungkan remote dan lakukan push:
   ```bash
   git remote add origin https://github.com/<username>/AREKO-BPR.git
   git branch -M main
   git push -u origin main
   ```
3. Buka [share.streamlit.io](https://share.streamlit.io) dan klik **New App**.
4. Pilih repositori `AREKO-BPR`, branch `main`, dan set file path ke `app.py`.
5. Streamlit Cloud akan otomatis membaca `packages.txt` dan `requirements.txt` untuk menginstal Tesseract & Poppler.
