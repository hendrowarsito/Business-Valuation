# 📊 Lembar Kerja Penilaian Saham — Semi Otomatis
**Versi 1.0 | SRR Appraisal | 2025**

Aplikasi Streamlit untuk mengisi lembar kerja penilaian saham secara semi-otomatis menggunakan Claude AI.

---

## Fitur Utama

- ✅ **Upload 3 dokumen**: Laporan Keuangan (PDF/XLSX), Proyeksi Keuangan (XLSX), Template (XLSX)
- 🤖 **Claude AI** mengekstrak dan memetakan akun keuangan secara otomatis
- 🛡️ **Preservasi format**: Font, style, border, merged cells template TIDAK diubah
- ∑ **Formula terjaga**: Sel subtotal/total dengan formula Excel tidak disentuh
- 📋 **Review interaktif**: Edit nilai manual sebelum generate
- ✅ **Validasi otomatis**: Cek balance sheet (Total Aset = Liabilitas + Ekuitas)
- 📄 **Review log**: Laporan lengkap akun yang berhasil/gagal dipetakan

---

## Instalasi & Menjalankan

### 1. Clone / copy folder

```bash
cd stock_valuation
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set API Key (opsional, bisa diisi di sidebar aplikasi)

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

### 4. Jalankan aplikasi

```bash
streamlit run app.py
```

Buka browser: `http://localhost:8501`

---

## Struktur Folder

```
stock_valuation/
├── app.py                   # Entry point Streamlit
├── requirements.txt
├── README.md
├── pages/
│   ├── upload.py            # Step 1: Upload dokumen
│   ├── analysis.py          # Step 2: Analisis AI
│   ├── review.py            # Step 3: Review & edit mapping
│   └── download.py          # Step 4: Generate & download
└── utils/
    ├── session.py           # Session state management
    ├── parser.py            # Baca XLSX/PDF → text/schema
    ├── mapper.py            # Claude API: ekstraksi & mapping
    └── writer.py            # openpyxl: tulis ke template
```

---

## Alur Kerja

```
Upload (3 dokumen)
    ↓
Analisis AI (Claude API)
    ├── Ekstrak laporan keuangan → dict {akun: [nilai per tahun]}
    ├── Ekstrak proyeksi keuangan → dict {akun: [nilai proyeksi]}
    ├── Baca schema template → {sel: label/formula}
    └── Account mapping → list [{akun, sel, nilai, confidence}]
        ↓
Review & Edit
    ├── Tabel interaktif (editable untuk akun confidence rendah)
    ├── Validasi balance sheet otomatis
    └── Pratinjau proyeksi & valuasi
        ↓
Generate & Download
    ├── write_to_template() → openpyxl (preservasi style)
    ├── Download .xlsx (lembar kerja terisi)
    └── Download .txt (review log)
```

---

## Aturan Penulisan Template

Aplikasi mengikuti constraint berikut (sesuai spesifikasi):

| Rule | Implementasi |
|------|-------------|
| Tidak tambah kolom/baris baru | `ws.cell()` write only, tidak ada `insert_rows/cols` |
| Font tidak diubah | openpyxl hanya set `.value`, tidak sentuh `.font` |
| Sel formula dijaga | Deteksi otomatis sel ber-formula, skip saat write |
| Subtotal/total pakai formula | Formula Excel dipertahankan (tidak di-replace nilai) |

---

## Konfigurasi

| Parameter | Default | Keterangan |
|-----------|---------|------------|
| `ANTHROPIC_API_KEY` | env var / sidebar | API key Claude |
| Mata Uang | IDR | IDR / USD / SGD |
| Satuan Angka | Juta | Juta / Miliar / Penuh |
| Bahasa Akun | Indonesia | Indonesia / Inggris / Bilingual |

---

## Pengembangan Selanjutnya (Roadmap)

- [ ] Multi-sheet template support (neraca, laba rugi, arus kas di sheet berbeda)
- [ ] Export laporan penilaian ke DOCX
- [ ] Integrasi data real-time dari IDX API
- [ ] Template schema builder (UI untuk mendefinisikan mapping tanpa AI)
- [ ] Batch processing (multiple emiten sekaligus)
- [ ] History & versioning per emiten
