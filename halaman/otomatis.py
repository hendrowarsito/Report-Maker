"""
Laporan otomatis — tanpa placeholder.

1. Upload template (boleh banyak file: bab-bab laporan & surat), lembar kerja,
   dan dokumen pendukung (sertifikat, IMB, proposal)
2-4. Claude membaca template, menemukan data objek yang harus diisi, mengambil
     nilainya dari tiap sumber, dan mengecek konsistensi antarsumber
5. Penilai memeriksa tabel temuan dan memilih sumber yang dipakai
6. Unduh laporan
"""
import hashlib
import os
import zipfile
from collections import Counter
from io import BytesIO
from pathlib import Path

import anthropic
import pandas as pd
import streamlit as st

from report_maker.ai import get_client
from report_maker.analisis import Temuan, analisis_dokumen, rencana_perubahan
from report_maker.checks import check_value_consistency_table
from report_maker.docx_engine import isi_dokumen
from report_maker.pdf import docx_ke_pdf, pdf_tersedia
from report_maker.penyesuaian import daftar_paragraf, terapkan_penyesuaian, validasi_usulan

MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
ISI_MANUAL = "Isi manual"
JANGAN_UBAH = "Jangan ubah"
BATAS_PDF_MB = 30  # batas satu permintaan API 32 MB, sisakan ruang untuk teks


@st.cache_data(show_spinner=False, max_entries=20)
def _paragraf(docx_bytes: bytes) -> dict[int, str]:
    return daftar_paragraf(docx_bytes)


@st.cache_data(show_spinner=False, max_entries=20)
def _pdf(docx_bytes: bytes) -> bytes:
    return docx_ke_pdf(docx_bytes)


def _api_key() -> str:
    """API key dari secrets / variabel lingkungan, atau isian pengguna."""
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:  # secrets.toml tidak ada
        key = ""
    key = key or os.environ.get("ANTHROPIC_API_KEY", "")
    if key:
        st.caption("API key diambil dari secrets / variabel lingkungan `ANTHROPIC_API_KEY`.")
        return key
    return st.text_input(
        "API Key Anthropic",
        key="api_key",
        type="password",
        placeholder="sk-ant-...",
        help="Dapatkan API key di console.anthropic.com. "
             "Simpan di .streamlit/secrets.toml agar tidak perlu diisi ulang.",
    )


def _cocokkan_sumber(nama: str, daftar: list[str]) -> str:
    """Nama sumber dari Claude → nama file yang di-upload (abaikan huruf besar/kecil)."""
    if nama in daftar:
        return nama
    return next((d for d in daftar if d.lower() == nama.lower()), nama)


# ---------------------------------------------------------------------------
# ① Upload
# ---------------------------------------------------------------------------
st.subheader("① Upload dokumen")
c1, c2 = st.columns(2)
templates = c1.file_uploader(
    "Template laporan — bab-bab & surat (.docx)", type=["docx"], accept_multiple_files=True
)
lembar_kerja = c2.file_uploader("Lembar kerja penilaian (.xlsx)", type=["xlsx"])
pendukung = st.file_uploader(
    "Dokumen pendukung — sertifikat tanah, IMB/PBG, proposal, dll. (.pdf / .docx)",
    type=["pdf", "docx"],
    accept_multiple_files=True,
)

with st.expander("Pengaturan"):
    api_key = _api_key()
    claude_ready = bool(api_key and api_key.startswith("sk-ant-"))
    if api_key and not claude_ready:
        st.warning("Format API key tidak valid.")

sumber_files = ([lembar_kerja] if lembar_kerja else []) + list(pendukung or [])
if not (templates and sumber_files):
    st.info("Upload template laporan dan minimal satu sumber data (lembar kerja atau dokumen pendukung).")
    st.stop()

nama_template = [f.name for f in templates]
if len(set(nama_template)) < len(nama_template):
    st.error("Ada file template dengan nama sama. Ubah namanya agar bisa dibedakan.")
    st.stop()

template_bytes = {f.name: f.getvalue() for f in templates}
paragraf_per_file = {nama: _paragraf(isi) for nama, isi in template_bytes.items()}
sumber = [(f.name, f.getvalue()) for f in sumber_files]
nama_sumber = [n for n, _ in sumber]
ukuran_pdf_mb = sum(len(b) for n, b in sumber if n.lower().endswith(".pdf")) / 1_000_000

# ---------------------------------------------------------------------------
# ② Analisis
# ---------------------------------------------------------------------------
st.divider()
st.subheader("② Analisis")
st.caption(
    f"Template: {len(template_bytes)} file, "
    f"{sum(len(p) for p in paragraf_per_file.values())} paragraf · "
    f"Sumber: {', '.join(nama_sumber)}. "
    "Claude mencari data objek di template (alamat, sertifikat, luas, IMB, nilai, dll.), "
    "mengambil nilainya dari setiap sumber, dan mengecek apakah sumber-sumber itu sama."
)

kunci = hashlib.md5(
    b"".join(n.encode() + b for n, b in [*template_bytes.items(), *sumber])
).hexdigest()
kunci_analisis = f"analisis_{kunci}"

if ukuran_pdf_mb > BATAS_PDF_MB:
    st.error(
        f"Total PDF {ukuran_pdf_mb:.1f} MB melebihi batas {BATAS_PDF_MB} MB. "
        "Kurangi jumlah file atau perkecil ukuran scan."
    )
    st.stop()
if not claude_ready:
    st.warning("Isi API key di Pengaturan untuk menjalankan analisis.")
    st.stop()

if st.button(
    "Analisis ulang" if kunci_analisis in st.session_state else "Analisis dokumen",
    type="primary",
):
    try:
        with st.spinner("Claude sedang membaca template dan dokumen sumber… (bisa 1–5 menit)"):
            temuan, catatan = analisis_dokumen(get_client(api_key), paragraf_per_file, sumber)
    except (anthropic.APIError, RuntimeError, ValueError) as e:
        st.error(f"Analisis gagal: {e}")
    else:
        # Buang kemunculan yang file/teksnya tidak benar-benar ada di template
        dibuang = 0
        for t in temuan:
            sah = []
            for m in t.kemunculan:
                m.dokumen = _cocokkan_sumber(m.dokumen, nama_template)
                teks_paragraf = paragraf_per_file.get(m.dokumen, {}).get(m.paragraf, "")
                if m.teks_lama and m.teks_lama in teks_paragraf:
                    sah.append(m)
            dibuang += len(t.kemunculan) - len(sah)
            t.kemunculan = sah
        st.session_state[kunci_analisis] = (temuan, catatan, dibuang)

if kunci_analisis not in st.session_state:
    st.stop()
temuan, catatan, dibuang = st.session_state[kunci_analisis]
temuan: list[Temuan]

if catatan:
    st.info(f"**Catatan Claude:** {catatan}")
if dibuang:
    st.caption(f"{dibuang} lokasi di template diabaikan karena teksnya tidak cocok persis.")

# ---------------------------------------------------------------------------
# ③ Temuan & pilihan sumber
# ---------------------------------------------------------------------------
st.divider()
st.subheader("③ Periksa temuan & pilih sumber")

kolom_sumber = list(nama_sumber)
for t in temuan:
    for k in t.kandidat:
        k.sumber = _cocokkan_sumber(k.sumber, nama_sumber)
        if k.sumber not in kolom_sumber:
            kolom_sumber.append(k.sumber)


def _status(t: Temuan) -> str:
    if not t.kandidat:
        return "❓ Tidak ada di sumber"
    if len(t.kandidat) == 1:
        return "➖ Satu sumber"
    return "✅ Sama" if t.konsisten else "⚠️ Berbeda"


def _pilihan_awal(t: Temuan) -> str:
    sumber_ada = [k.sumber for k in t.kandidat]
    rekomendasi = _cocokkan_sumber(t.rekomendasi, nama_sumber)
    if rekomendasi in sumber_ada:
        return rekomendasi
    return sumber_ada[0] if sumber_ada else JANGAN_UBAH


# Yang berbeda antarsumber tampil paling atas
urutan = sorted(range(len(temuan)), key=lambda i: (_status(temuan[i]) != "⚠️ Berbeda", i))
jumlah_beda = sum(_status(t) == "⚠️ Berbeda" for t in temuan)
jumlah_kosong = sum(not t.kemunculan for t in temuan)

m1, m2, m3 = st.columns(3)
m1.metric("Data ditemukan", len(temuan))
m2.metric("Berbeda antarsumber", jumlah_beda)
m3.metric("Lokasi di laporan", sum(len(t.kemunculan) for t in temuan))
if jumlah_beda:
    st.warning(f"{jumlah_beda} data berbeda antarsumber — tentukan mana yang dipakai di kolom **Gunakan**.")

baris = []
for i in urutan:
    t = temuan[i]
    per_sumber = {}
    for k in t.kandidat:
        per_sumber.setdefault(k.sumber, f"{k.nilai}  ({k.lokasi})")
    baris.append({
        "Data": t.nama,
        "Status": _status(t),
        **{s: per_sumber.get(s, "") for s in kolom_sumber},
        "Gunakan": _pilihan_awal(t),
        "Nilai manual": "",
        "Di laporan saat ini": " · ".join(dict.fromkeys(m.teks_lama for m in t.kemunculan)) or "—",
        "Muncul di": ", ".join(
            f"{Path(d).stem} ({n}×)" for d, n in Counter(m.dokumen for m in t.kemunculan).items()
        ) or "—",
        "Keterangan": t.penjelasan,
    })

st.caption(
    "Periksa setiap baris. Kolom **Gunakan** sudah berisi sumber yang disarankan; ganti bila "
    f"perlu, pilih **{ISI_MANUAL}** lalu tulis di *Nilai manual*, atau **{JANGAN_UBAH}**."
)
df = st.data_editor(
    pd.DataFrame(baris),
    column_config={
        "Gunakan": st.column_config.SelectboxColumn(
            options=kolom_sumber + [ISI_MANUAL, JANGAN_UBAH], required=True
        ),
        "Nilai manual": st.column_config.TextColumn(help=f"Dipakai bila Gunakan = {ISI_MANUAL}"),
    },
    disabled=[c for c in baris[0] if c not in ("Gunakan", "Nilai manual")] if baris else True,
    hide_index=True,
    width="stretch",
    key=f"temuan_{kunci}",
)

nilai_dipakai: dict[int, str] = {}
peringatan: list[str] = []
for i, (_, r) in zip(urutan, df.iterrows()):
    t = temuan[i]
    pilihan = r["Gunakan"]
    if pilihan == JANGAN_UBAH or not t.kemunculan:
        continue
    if pilihan == ISI_MANUAL:
        manual = "" if pd.isna(r["Nilai manual"]) else str(r["Nilai manual"]).strip()
        if manual:
            nilai_dipakai[i] = manual
        else:
            peringatan.append(f"{t.nama}: {ISI_MANUAL} dipilih tetapi Nilai manual kosong — tidak diubah")
        continue
    nilai = next((k.nilai for k in t.kandidat if k.sumber == pilihan), None)
    if nilai is None:
        peringatan.append(f"{t.nama}: {pilihan} tidak memuat data ini — tidak diubah")
    else:
        nilai_dipakai[i] = nilai

perubahan_per_file, peringatan_angka = rencana_perubahan(temuan, nilai_dipakai)
perubahan_per_file = {
    nama: validasi_usulan(paragraf_per_file[nama], usulan)[0]
    for nama, usulan in perubahan_per_file.items()
}
jumlah_perubahan = sum(len(u) for u in perubahan_per_file.values())
peringatan += peringatan_angka

if jumlah_kosong:
    st.caption(f"{jumlah_kosong} data hanya untuk pengecekan (tidak ada lokasinya di template).")
for p in peringatan:
    st.warning(p)

with st.expander(f"Rincian {jumlah_perubahan} perubahan yang akan ditulis ke laporan"):
    if jumlah_perubahan:
        rincian = [
            (nama, u) for nama in nama_template for u in perubahan_per_file.get(nama, [])
        ]
        st.dataframe(
            pd.DataFrame({
                "Dokumen": [nama for nama, _ in rincian],
                "Paragraf": [u.paragraf for _, u in rincian],
                "Data": [u.alasan for _, u in rincian],
                "Teks lama": [u.teks_lama for _, u in rincian],
                "Teks baru": [u.teks_baru for _, u in rincian],
            }),
            hide_index=True,
            width="stretch",
        )
    else:
        st.write("Tidak ada perubahan — template sudah sesuai dengan pilihan di atas.")

# ---------------------------------------------------------------------------
# ④ Unduh
# ---------------------------------------------------------------------------
st.divider()
st.subheader("④ Unduh laporan")

hasil: dict[str, bytes] = {}
masalah: list[str] = []
jumlah = 0
for nama in nama_template:
    isi, n = terapkan_penyesuaian(template_bytes[nama], perubahan_per_file.get(nama, []))
    hasil[nama] = isi
    jumlah += n
    _, teks_hasil, sisa_placeholder = isi_dokumen(isi, {}, ())
    nominal = check_value_consistency_table(teks_hasil)
    salah_nominal = nominal[nominal["Status"].str.startswith("❌")]
    if sisa_placeholder:
        masalah.append(
            f"**{nama}** — {len(sisa_placeholder)} placeholder tersisa: "
            + ", ".join(f"`{{{{{p}}}}}`" for p in sisa_placeholder)
        )
    for _, r in salah_nominal.iterrows():
        masalah.append(f"**{nama}** — {r['Angka (Rp)']} ({r['Terbilang di Dokumen']}): {r['Status']}")

if masalah:
    st.warning("Periksa sebelum dikirim:\n" + "\n".join(f"- {m}" for m in masalah))
else:
    st.success(f"Laporan siap — {jumlah} penggantian di {len(hasil)} file.")

nama_hasil = {nama: f"{Path(nama).stem}_terisi" for nama in nama_template}
buat_pdf = pdf_tersedia() and st.toggle("Sertakan PDF")
pdfs: dict[str, bytes] = {}
if buat_pdf:
    try:
        with st.spinner("Membuat PDF..."):
            pdfs = {nama: _pdf(isi) for nama, isi in hasil.items()}
    except RuntimeError as e:
        st.error(str(e))

if len(hasil) == 1:
    nama, isi = next(iter(hasil.items()))
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button(
        "Unduh Word (.docx)", data=isi, file_name=f"{nama_hasil[nama]}.docx",
        mime=MIME_DOCX, type="primary", width="stretch",
    )
    if nama in pdfs:
        c2.download_button(
            "Unduh PDF", data=pdfs[nama], file_name=f"{nama_hasil[nama]}.pdf",
            mime="application/pdf", width="stretch",
        )
else:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for nama, isi in hasil.items():
            zipf.writestr(f"{nama_hasil[nama]}.docx", isi)
            if nama in pdfs:
                zipf.writestr(f"{nama_hasil[nama]}.pdf", pdfs[nama])
    st.download_button(
        f"Unduh semua ({len(hasil)} file, .zip)",
        data=zip_buffer.getvalue(),
        file_name="Laporan_terisi.zip",
        mime="application/zip",
        type="primary",
    )
    with st.expander("Unduh per file"):
        for nama, isi in hasil.items():
            st.download_button(
                f"{nama_hasil[nama]}.docx", data=isi, file_name=f"{nama_hasil[nama]}.docx",
                mime=MIME_DOCX, key=f"unduh_{nama}",
            )
