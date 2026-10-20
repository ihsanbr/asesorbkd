# 📊 Dashboard Monitoring Asesor BKD

Dashboard interaktif untuk monitoring Asesor Beban Kerja Dosen (BKD).  
Siap dijalankan lokal atau di-hosting (Streamlit Cloud, Hugging Face, VPS, dll).

---

## 📁 Isi Folder

| File | Keterangan |
|------|------------|
| `dashboard_asesor_bkd.py` | Aplikasi utama Streamlit |
| `data_asesor.xlsx` | Sample data Asesor (20 baris) |
| `data_dosen.xlsx` | Sample data Dosen (100 baris) |
| `requirements.txt` | Daftar library Python |
| `jalankan.bat` | Script jalankan untuk Windows |
| `jalankan.sh` | Script jalankan untuk Linux/Mac |
| `README.md` | Dokumentasi ini |

---

## 🚀 Cara Menjalankan di Komputer Lokal

### Opsi 1 – Windows
1. Double-click file **jalankan.bat**
2. Tunggu install selesai, browser akan terbuka otomatis

### Opsi 2 – Linux / Mac
```bash
chmod +x jalankan.sh
./jalankan.sh
```

### Opsi 3 – Manual
```bash
pip install -r requirements.txt
streamlit run dashboard_asesor_bkd.py
```

Aplikasi akan terbuka di: **http://localhost:8501**

---

## ☁️ Cara Hosting (Gratis)

### 1. Streamlit Community Cloud (Paling Mudah)
1. Buat akun di https://share.streamlit.io
2. Upload folder ini ke GitHub repository
3. Connect repository → pilih file `dashboard_asesor_bkd.py`
4. Deploy → dapat URL publik

### 2. Hugging Face Spaces
1. Buat Space baru (SDK: Streamlit)
2. Upload semua file di folder ini
3. Spaces akan otomatis menjalankan dashboard

### 3. VPS / Server Sendiri
```bash
pip install -r requirements.txt
streamlit run dashboard_asesor_bkd.py --server.port 8501 --server.address 0.0.0.0
```

---

## 📂 Format Data Excel

### Data Asesor BKD
Kolom wajib:
- ID Peg
- NUPTK
- Nama
- Fakultas
- Prodi

### Data Dosen Masuk Periode BKD
Kolom wajib:
- ID Peg
- NUPTK
- Nama
- Fakultas
- Prodi
- Jabatan Akademik
- Status Dosen
- Status SERDOS

Sangat disarankan menambahkan kolom:
- Pendidikan Terakhir → isi S2 atau S3

---

## 📋 Fitur Dashboard

### 🏠 Dashboard Utama
- Metrik: Total Asesor, Total Dosen, Kekurangan, Eligible
- Grafik distribusi Asesor per Fakultas & Prodi
- Perbandingan Asesor vs Ideal 30%
- Pie chart status kecukupan & eligible

### 📋 Daftar Asesor BKD
- Tabel + filter + download CSV

### ⚠️ Kekurangan Asesor BKD
- Ideal = 30% dari jumlah Dosen per Prodi
- Status Cukup / Kurang + detail per Prodi

### ✅ Eligible Asesor BKD
Kriteria:
1. Status Dosen = Dosen Tetap
2. S3 + minimal Lektor  ATAU  S2 + minimal Lektor Kepala
3. Status SERDOS = Sudah

---

## 💡 Tips
- Upload data Excel asli Anda lewat sidebar aplikasi
- Jika tidak upload, sistem memakai data sample
- Matching "sudah jadi asesor" menggunakan kolom ID Peg

---

Versi 1.0 | Dashboard Monitoring Asesor BKD
