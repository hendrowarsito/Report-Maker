"""
Laporan otomatis — tanpa placeholder.

1. Upload template, lembar kerja, dan dokumen pendukung (sertifikat, IMB, proposal)
2-4. Claude membaca template, menemukan data objek yang harus diisi, mengambil
     nilainya dari tiap sumber, dan mengecek konsistensi antarsumber
5. Penilai memeriksa tabel temuan dan memilih sumber yang dipakai
6. Unduh laporan
"""
import hashlib
import os
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
template = c1.file_uploader("Template laporan (.docx)", type=["docx"])
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
if not (template and sumber_files):
    st.info("Upload template laporan dan minimal satu sumber data (lembar kerja atau dokumen pendukung).")
    st.stop()

template_bytes = template.getvalue()
sumber = [(f.name, f.getvalue()) for f in sumber_files]
nama_sumber = [n for n, _ in sumber]
ukuran_pdf_mb = sum(len(b) for n, b in sumber if n.lower().endswith(".pdf")) / 1_000_000
paragraf = _paragraf(template_bytes)

# ---------------------------------------------------------------------------
# ② Analisis
# ---------------------------------------------------------------------------
st.divider()
st.subheader("② Analisis")
st.caption(
    f"Template: {len(paragraf)} paragraf · Sumber: {', '.join(nama_sumber)}. "
    "Claude mencari data objek di template (alamat, sertifikat, luas, IMB, nilai, dll.), "
    "mengambil nilainya dari setiap sumber, dan mengecek apakah sumber-sumber itu sama."
)

kunci = hashlib.md5(
    template_bytes + b"".join(n.encode() + b for n, b in sumber)
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
        with st.spinner("Claude sedang membaca template dan dokumen sumber… (bisa 1–3 menit)"):
            temuan, catatan = analisis_dokumen(get_client(api_key), paragraf, sumber)
    except (anthropic.APIError, RuntimeError, ValueError) as e:
        st.error(f"Analisis gagal: {e}")
    else:
        # Buang kemunculan yang teksnya tidak benar-benar ada di paragraf
        dibuang = 0
        for t in temuan:
            sah = [m for m in t.kemunculan if m.teks_lama and m.teks_lama in paragraf.get(m.paragraf, "")]
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

perubahan, peringatan_angka = rencana_perubahan(temuan, nilai_dipakai)
perubahan, _ = validasi_usulan(paragraf, perubahan)
peringatan += peringatan_angka

if jumlah_kosong:
    st.caption(f"{jumlah_kosong} data hanya untuk pengecekan (tidak ada lokasinya di template).")
for p in peringatan:
    st.warning(p)

with st.expander(f"Rincian {len(perubahan)} perubahan yang akan ditulis ke laporan"):
    if perubahan:
        st.dataframe(
            pd.DataFrame({
                "Paragraf": [u.paragraf for u in perubahan],
                "Data": [u.alasan for u in perubahan],
                "Teks lama": [u.teks_lama for u in perubahan],
                "Teks baru": [u.teks_baru for u in perubahan],
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

hasil, jumlah = terapkan_penyesuaian(template_bytes, perubahan)
_, teks_hasil, sisa_placeholder = isi_dokumen(hasil, {}, ())
nominal = check_value_consistency_table(teks_hasil)
salah_nominal = nominal[nominal["Status"].str.startswith("❌")]

if sisa_placeholder:
    st.warning(
        f"{len(sisa_placeholder)} placeholder masih tersisa: "
        + ", ".join(f"`{{{{{p}}}}}`" for p in sisa_placeholder)
    )
if not salah_nominal.empty:
    st.warning("Ada nominal yang tidak sesuai terbilangnya:")
    st.dataframe(salah_nominal, hide_index=True, width="stretch")
if not sisa_placeholder and salah_nominal.empty:
    st.success(f"Laporan siap — {jumlah} penggantian diterapkan.")

nama_hasil = f"{Path(template.name).stem}_terisi"
c1, c2, _ = st.columns([1, 1, 3])
c1.download_button(
    "Unduh Word (.docx)",
    data=hasil,
    file_name=f"{nama_hasil}.docx",
    mime=MIME_DOCX,
    type="primary",
    width="stretch",
)
if pdf_tersedia():
    if c2.toggle("Buat PDF"):
        try:
            with st.spinner("Membuat PDF..."):
                pdf = _pdf(hasil)
            c2.download_button(
                "Unduh PDF", data=pdf, file_name=f"{nama_hasil}.pdf",
                mime="application/pdf", width="stretch",
            )
        except RuntimeError as e:
            st.error(str(e))
