# SRR Kalibata Report Maker

Aplikasi Streamlit untuk mengisi template laporan Word (`.docx`) dari data
lembar kerja Excel (`.xlsx`) secara otomatis, lengkap dengan terbilang Rupiah,
pemeriksaan kelengkapan data, dan ekspor PDF.

## Menjalankan

```bash
pip install -r requirements.txt
streamlit run CreateLP5.py
```

Aplikasi terbuka di `http://localhost:8501` dengan dua halaman (menu di atas).

## Laporan otomatis (halaman utama) — tanpa placeholder

1. **Upload** template laporan (.docx — boleh banyak file: bab-bab laporan,
   surat pengantar, ringkasan, dll.), lembar kerja penilaian (.xlsx), dan
   dokumen pendukung — sertifikat tanah, IMB/PBG, proposal (.pdf / .docx).
2. Klik **Analisis dokumen**. Claude membaca template dan menemukan sendiri
   data objek yang harus diisi/diganti (pemberi tugas, alamat, sertifikat,
   NIB, pemegang hak, luas, IMB, nilai, tanggal, dll.) — baik berupa
   placeholder `{{...}}` maupun data dari laporan lama.
3. Untuk setiap data, Claude mengambil nilainya dari **setiap** sumber dan
   menyarankan sumber yang paling berwenang (data tanah → sertifikat,
   bangunan → IMB, penugasan → proposal, angka penilaian → lembar kerja).
4. Konsistensi antarsumber dicek: data yang **berbeda** (mis. luas di lembar
   kerja 312 m², di sertifikat 315 m²) tampil paling atas dengan ⚠️.
5. Periksa tabel temuan. Kolom **Gunakan** berisi sumber yang disarankan —
   ganti bila perlu, pilih *Isi manual*, atau *Jangan ubah*. Rincian setiap
   penggantian (paragraf, teks lama → teks baru) bisa dilihat sebelum unduh.
   Kolom *Muncul di* menunjukkan file mana saja yang memuat data itu — satu
   pilihan diterapkan seragam di semua bab dan surat.
6. **Preview** laporan sebelum diunduh:
   - *Tampilan halaman* — halaman laporan persis seperti hasil cetak (butuh
     LibreOffice), bisa disaring ke halaman yang berubah saja.
   - *Perubahan teks* — setiap paragraf yang berubah, teks lama dicoret dan
     teks baru disorot.
   Ubah pilihan di tabel temuan bila ada yang keliru; preview ikut diperbarui.
7. **Unduh** laporan (.docx, dan PDF bila LibreOffice tersedia). Banyak file
   diunduh sekaligus sebagai .zip, atau satu per satu. Placeholder yang
   tersisa dan nominal yang tidak sesuai terbilang diperingatkan per file.

Catatan:
- Nilai dari lembar kerja selalu dibaca dari **sel aslinya**, bukan salinan
  AI. Terbilang dihitung aplikasi dari angka terpilih, bukan ditulis AI.
- Format angka mengikuti template: `1.000.000,00`, `1.000.000`, atau `250`.
  Nomor seperti nomor sertifikat ditulis apa adanya.
- Hanya teks yang diganti yang berubah; format huruf di sekitarnya tetap.
  Di langkah Unduh, setiap file hasil **diverifikasi** terhadap templatenya:
  gaya, tabel, header/footer, gambar, dan penomoran harus identik — hanya isi
  teks yang boleh berbeda ("✅ Format template utuh").
- Angka dari Excel diberi titik ribuan (595400000 → 595.400.000) bila
  ≥ 10.000. Angka kecil (tahun, luas 416) dan data bernama nomor/kode
  (`kode_pos`, `nib`, `nomor_…`, `tahun_…`, `rt`, `rw`) ditulis apa adanya.
- Analisis berjalan di **latar belakang server** dengan tampilan kemajuan:
  aman ditinggal pindah aplikasi atau layar mati. Sesi (file, hasil, pilihan
  di tabel) dipertahankan 2 jam setelah koneksi terputus
  (`server.disconnectedSessionTTL` di `.streamlit/config.toml`).
- Hasil analisis disimpan di server selama 7 hari, dikunci dengan sidik isi
  file: meng-upload ulang file yang sama langsung memuat hasilnya **tanpa
  analisis ulang dan tanpa biaya**. Hasil hilang bila aplikasi di-reboot.
- Bawaan memakai tingkat berpikir *medium* agar cepat; aktifkan **Analisis
  lebih teliti** di Pengaturan untuk dokumen rumit.
- Satu analisis memakai `claude-opus-5-5`; perkiraan biaya kasar
  US$0,30–1 per laporan tergantung jumlah halaman sumber. Total PDF maksimal
  30 MB.
- Isi dokumen dikirim ke API Anthropic; pastikan sesuai kebijakan
  kerahasiaan data klien.

## Mode placeholder (halaman kedua)

Alur sebelumnya, untuk template berisi placeholder `{{...}}`: Upload →
Periksa → Sesuaikan dengan AI → Unduh. Penjelasan di bawah berlaku untuk
halaman ini.

## Sumber data: dua cara

### 1. Baca langsung dari lembar kerja (tanpa sheet Laporan)

Upload lembar kerja penilaian apa adanya. Untuk setiap placeholder di template,
Claude mencarikan sel yang tepat di sheet mana pun — misalnya `{{alamat_objek}}`
→ `LK Tanah!C5` — dengan memahami nama placeholder dan kalimat tempat ia dipakai.

- Klik **Petakan … data dengan AI**. Nilai selalu **dibaca dari sel**, bukan
  ditulis ulang oleh AI, jadi angka tidak mungkin dikarang.
- Hasil pemetaan tampil di tabel **Pemetaan data → sel lembar kerja**: kolom
  *Sel* bisa dikoreksi (`Nama Sheet!C5`, beberapa sel dipisah `;` dan
  digabung dengan *Pemisah*). Hapus baris agar dipetakan ulang.
- Pemetaan berkeyakinan rendah, sel kosong, atau data yang tidak ditemukan
  ditandai ⚠️ di Kelengkapan Data.
- **Simpan pemetaan (.json)** lalu upload di **Pengaturan → Pemetaan
  tersimpan** untuk laporan berikutnya dengan format lembar kerja yang sama:
  langsung terisi, tanpa AI dan tanpa biaya. Hanya placeholder baru yang
  perlu dipetakan AI.
- Angka mengikuti format sel di Excel: `#,##0` → `1.250.000.000`, persen →
  `12,5%`, tanggal → `1 Oktober 2026`.
- Nilai rumus dibaca dari hasil terakhir yang disimpan Excel — buka dan simpan
  file di Excel sebelum di-upload.
- Beri placeholder nama yang menjelaskan isinya (`{{luas_tanah}}`,
  `{{nilai_pasar}}`), agar AI memetakannya dengan tepat.

### 2. Sheet berisi key | nilai (cara lama)

Dipilih otomatis bila file punya sheet bernama `Laporan`. Data dibaca dari dua
kolom pertama sheet yang dipilih:

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

Tulis placeholder sebagai `{{nama}}`. Placeholder boleh berada di body,
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

### Sesuaikan dengan AI (tanpa placeholder)

Upload dokumen sumber — PDF (sertifikat, IMB/PBG, dsb.) dan Excel (lembar
kerja penilaian) — lalu klik **Cari ketidaksesuaian**. Claude membaca sumber
dan laporan yang sudah terisi, lalu mengusulkan penggantian kalimat yang
bertentangan dengan sumber (nomor sertifikat, luas, pemegang hak, alamat,
tanggal, nilai beserta terbilangnya, dll.).

- Setiap usulan menyebut **sumber** (file + halaman / sheet + sel) dan
  **kutipan** teksnya, agar mudah dicocokkan.
- Tidak ada yang berubah sebelum disetujui: hapus centang untuk menolak, atau
  ubah kolom *Usulan baru*. Usulan yang dicentang langsung dipakai di file
  unduhan dan dokumennya diperiksa ulang.
- Hanya teks yang diganti yang berubah — format huruf di sekitarnya tetap.
- Memakai model `claude-opus-5-5` (`report_maker/penyesuaian.py`,
  `MODEL_PENYESUAIAN`). Perkiraan biaya kasar US$0,30–0,60 per laporan;
  analisis ulang dengan sumber yang sama lebih murah karena sumber di-cache.
- Total PDF maksimal 30 MB per analisis. Nilai rumus Excel dibaca dari hasil
  terakhir yang disimpan Excel — simpan file di Excel sebelum di-upload.
- Isi dokumen sumber dikirim ke API Anthropic; pastikan sesuai kebijakan
  kerahasiaan data klien.

### Ekspor PDF

Butuh LibreOffice di server. `packages.txt` sudah berisi paket yang
diperlukan (dipasang otomatis oleh Streamlit Community Cloud dan devcontainer):
LibreOffice Writer serta font Liberation dan Carlito yang ukurannya setara
Arial/Times New Roman/Calibri, agar tata letak PDF sama dengan Word.
Hapus `packages.txt` bila PDF tidak diperlukan — instalasinya cukup besar.

## Struktur kode

```
CreateLP5.py               titik masuk: header, navigasi halaman
halaman/otomatis.py        halaman Laporan otomatis (tanpa placeholder)
halaman/placeholder.py     halaman Mode placeholder (alur lama)
report_maker/
  terbilang.py             terbilang, format angka/tanggal, placeholder turunan
  docx_engine.py           pengisian placeholder & Find & Replace di .docx
  checks.py                kelengkapan data, nominal vs terbilang, deteksi baris awal
  ai.py                    review dokumen dengan Claude
  penyesuaian.py           usulan perubahan kalimat dari dokumen sumber (PDF/Excel)
  lembar_kerja.py          baca sel lembar kerja & pemetaan placeholder → sel
  analisis.py              analisis template + sumber, cek konsistensi antarsumber
  pekerjaan.py             analisis di latar belakang + simpan hasil ke disk
  preview.py               gambar halaman dari PDF & daftar perubahan teks
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
