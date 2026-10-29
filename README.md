# SRR Kalibata Report Maker

Aplikasi Streamlit untuk mengisi template laporan Word (`.docx`) dari data
lembar kerja Excel (`.xlsx`) secara otomatis, lengkap dengan terbilang Rupiah,
pemeriksaan kelengkapan data, dan ekspor PDF.

## Menjalankan

```bash
pip install -r requirements.txt
streamlit run CreateLP5.py
```

Aplikasi terbuka di `http://localhost:8501`. Alurnya tiga langkah:

1. **Upload** — template Word (boleh lebih dari satu) dan data Excel.
2. **Periksa** — tabel kelengkapan placeholder, cek nominal vs terbilang,
   preview isi, dan (opsional) review dengan AI.
3. **Unduh** — Word, dan PDF bila LibreOffice tersedia. Bila masih ada
   masalah, unduhan perlu dikonfirmasi dulu.

## Format Excel

Data dibaca dari dua kolom pertama sheet yang dipilih:

| A (key)        | B (nilai)        |
|----------------|------------------|
| `nama_debitur` | PT ABC           |
| `tgl_laporan`  | 28/04/2026       |
| `NP_nominal`   | 200000000        |

- Baris kop/judul di atas tabel dilewati otomatis — aplikasi mencari baris
  pertama yang key-nya dipakai di template. Bisa diatur manual di
  **Pengaturan → Lewati baris pertama**.
- Tanggal ditulis `28 April 2026`, bilangan bulat tanpa `.0`, desimal memakai
  koma (`0,25`). Persentase yang tersimpan sebagai angka (`0.25`) tampil
  `0,25`; tulis sebagai teks di Excel bila ingin `25%`.

## Placeholder di template Word

Tulis key dari kolom A sebagai `{{key}}`. Placeholder boleh berada di body,
tabel, header, footer, atau hyperlink, dan formatnya (tebal, miring, font)
tetap terjaga.

Untuk nominal, tambahkan akhiran agar terbilang dibuat otomatis
(akhiran tidak peka huruf besar/kecil):

| Placeholder                       | Hasil                   |
|-----------------------------------|-------------------------|
| `{{NP_nominal}}`                  | 200.000.000,00          |
| `{{NP_nominal_FORMAT}}`           | Rp 200.000.000,00       |
| `{{NP_nominal_TERBILANG}}`        | DUA RATUS JUTA RUPIAH   |
| `{{NP_nominal_TERBILANG_TITLE}}`  | Dua Ratus Juta Rupiah   |

Key sumber (`{{NP_nominal}}`) diformat ribuan hanya bila salah satu akhiran
dipakai. Bila Excel sudah punya baris dengan nama turunan tersebut, nilai
Excel yang dipakai.

## Pengaturan lain

- **Nama file hasil** — pola dengan key Excel, misalnya
  `LP_{{nama_debitur}}_{{tgl_laporan}}`. `{{nama_file}}` = nama template.
- **Find & Replace** — ganti kata langsung di dokumen tanpa placeholder.
  Diterapkan setelah placeholder diisi.

### Review dengan Claude AI (opsional)

Isi API key di Pengaturan, atau simpan agar tidak perlu diisi ulang di
`.streamlit/secrets.toml` (sudah di-`.gitignore`):

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

Variabel lingkungan `ANTHROPIC_API_KEY` juga dibaca. Di Streamlit Community
Cloud, isi lewat **App settings → Secrets**.

### Ekspor PDF

Butuh LibreOffice di server. `packages.txt` sudah berisi paket yang
diperlukan (dipasang otomatis oleh Streamlit Community Cloud dan devcontainer):
LibreOffice Writer serta font Liberation dan Carlito yang ukurannya setara
Arial/Times New Roman/Calibri, agar tata letak PDF sama dengan Word.
Hapus `packages.txt` bila PDF tidak diperlukan — instalasinya cukup besar.

## Struktur kode

```
CreateLP5.py               UI Streamlit (tanpa logika inti)
report_maker/
  terbilang.py             terbilang, format angka/tanggal, placeholder turunan
  docx_engine.py           pengisian placeholder & Find & Replace di .docx
  checks.py                kelengkapan data, nominal vs terbilang, deteksi baris awal
  ai.py                    review dokumen dengan Claude
  pdf.py                   konversi .docx → PDF lewat LibreOffice
assets/logo_srr.png        logo di header dan ikon tab
.streamlit/config.toml     tema warna SRR
tests/                     unit test (pytest)
```

## Tes

```bash
pip install -r requirements-dev.txt
pytest
```

## Catatan devcontainer

`.devcontainer` menjalankan Streamlit dengan `--server.enableCORS false
--server.enableXsrfProtection false`. Ini bawaan template Codespaces dari
Streamlit dan diperlukan agar upload file berjalan lewat port-forward
Codespaces; jangan pakai flag ini saat deploy di server publik.
