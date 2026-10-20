# 📊 Dashboard Monitoring Asesor BKD — Universitas Sebelas Maret

Dashboard interaktif monitoring Asesor BKD dengan tema warna **PEAR UNS** (#0e7da7) dan logo resmi UNS.

---

## 📁 Isi Folder

| File / Folder | Keterangan |
|---------------|------------|
| `dashboard_asesor_bkd.py` | Aplikasi utama Streamlit |
| `data_asesor.xlsx` | Sample data Asesor |
| `data_dosen.xlsx` | Sample data Dosen |
| `assets/logo_uns.png` | Logo resmi UNS |
| `requirements.txt` | Dependencies |
| `jalankan.bat` | Jalankan di Windows |
| `jalankan.sh` | Jalankan di Linux/Mac |
| `README.md` | Dokumentasi |

---

## ✨ Fitur Baru

### 📋 Daftar Asesor BKD
- **Edit** data langsung di tabel (`st.data_editor`)
- **Hapus** by ID Peg atau hapus baris di editor lalu Simpan
- **Tambah** asesor baru via form
- **Export Excel**

### ⚠️ Kekurangan Asesor BKD
- Perhitungan ideal 30% per Prodi
- **Export Excel**

### ✅ Eligible Asesor BKD
- Filter kriteria kelayakan
- **Export Excel**

### Tema
- Warna utama: **#0e7da7** (PEAR UNS)
- Logo UNS di header & sidebar

---

## 🚀 Cara Menjalankan

### Windows
Double-click `jalankan.bat`

### Linux / Mac
```bash
chmod +x jalankan.sh
./jalankan.sh
```

### Manual
```bash
pip install -r requirements.txt
streamlit run dashboard_asesor_bkd.py
```

---

## ☁️ Hosting

- **Streamlit Cloud**: upload ke GitHub → deploy di share.streamlit.io
- **Hugging Face Spaces**: buat Space (SDK Streamlit) → upload folder
- **VPS**: `streamlit run dashboard_asesor_bkd.py --server.address 0.0.0.0`

---

## 📂 Format Data Excel

**Asesor:** ID Peg, NUPTK, Nama, Fakultas, Prodi

**Dosen:** ID Peg, NUPTK, Nama, Fakultas, Prodi, Jabatan Akademik, Status Dosen, Status SERDOS  
(+ rekomendasi: Pendidikan Terakhir = S2/S3)

---

Versi 1.1 | Universitas Sebelas Maret
