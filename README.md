# PDF Verifier — Verifikasi Invoice vs Excel

Tool untuk membuka PDF berpassword dan mencocokkan data invoice dengan file Excel secara batch.

## Cara Deploy ke Railway (Gratis)

### 1. Buat akun Railway
- Buka https://railway.app
- Sign up pakai akun GitHub (gratis)

### 2. Upload kode ke GitHub
- Buat repo baru di https://github.com/new (nama bebas, misal `pdf-verifier`)
- Upload semua file dari folder ini ke repo tersebut

  **Cara cepat via GitHub web:**
  - Klik "uploading an existing file"
  - Drag semua file (main.py, requirements.txt, Procfile, railway.toml, folder static/)
  - Klik "Commit changes"

### 3. Deploy di Railway
- Buka https://railway.app/new
- Pilih **"Deploy from GitHub repo"**
- Pilih repo yang baru dibuat
- Railway otomatis detect Python dan deploy
- Tunggu 2-3 menit → dapat URL publik (misal: `https://pdf-verifier-xxx.railway.app`)

### 4. Pakai tool-nya
- Buka URL Railway kamu
- Upload PDF invoice (bisa banyak sekaligus)
- Upload file Excel (unifikasi_ah.xlsx atau sejenisnya)
- Masukkan password PDF
- Klik Verifikasi → hasil langsung muncul

## Jalankan Lokal (Opsional)

```bash
pip install -r requirements.txt
uvicorn main:app --reload
# Buka http://localhost:8000
```

## Struktur File

```
pdf-verifier/
├── main.py          ← Backend FastAPI
├── requirements.txt ← Dependensi Python
├── Procfile         ← Command start Railway
├── railway.toml     ← Config Railway
└── static/
    └── index.html   ← Frontend (UI)
```

## Cara Kerja

1. PDF diupload → dibuka dengan pikepdf (menggunakan password)
2. Teks diekstrak dengan pdfplumber
3. Nomor dokumen, DPP, PPH diparsing dari teks
4. Dicocokkan dengan baris di Excel berdasarkan nomor faktur/invoice
5. Hasil ditampilkan: Sesuai / Tidak Sesuai / Tidak Ditemukan
