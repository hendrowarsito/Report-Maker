"""
Analisis otomatis: template laporan + lembar kerja + dokumen pendukung.

Claude membaca template (tanpa perlu placeholder), menemukan data objek yang
harus diisi/diganti, mengambil nilainya dari setiap sumber (lembar kerja,
sertifikat, IMB, proposal), dan menilai apakah sumber-sumber itu konsisten.
Penilai memilih sumber yang dipakai; aplikasi lalu menulis ke laporan.

Angka terbilang tidak ditulis AI: aplikasi menghitungnya dari nilai terpilih.
"""
import base64
import re
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from io import BytesIO
from pathlib import Path

import anthropic
from docx import Document

from .ai import minta_json
from .docx_engine import extract_text_from_docx
from .lembar_kerja import baca_sel, sel_ke_teks
from .penyesuaian import Usulan
from .terbilang import (
    angka_ke_terbilang,
    angka_ke_terbilang_title,
    format_angka_indonesia,
    format_ribuan,
    parse_ke_int,
)

BENTUK = ["teks", "angka", "terbilang", "terbilang_judul"]


@dataclass
class Kandidat:
    sumber: str      # nama file
    lokasi: str      # "hal. 1" / "LK Tanah!C5"
    nilai: str


@dataclass
class Kemunculan:
    paragraf: int
    teks_lama: str
    bentuk: str      # salah satu BENTUK
    dokumen: str = ""  # nama file template (laporan bisa terdiri dari banyak bab/surat)


@dataclass
class Temuan:
    nama: str
    kandidat: list[Kandidat] = field(default_factory=list)
    rekomendasi: str = ""        # nama file sumber yang disarankan
    konsisten: bool = True
    penjelasan: str = ""
    kemunculan: list[Kemunculan] = field(default_factory=list)


# =============================================================================
# Dokumen sumber → blok konten untuk Claude
# =============================================================================

def blok_sumber(nama: str, isi: bytes) -> tuple[dict, dict[str, str]]:
    """
    Ubah satu file sumber menjadi blok konten Claude.
    Kembalikan (blok, sel) — sel berisi isi sel Excel ({"Sheet!C5": teks})
    agar nilai lembar kerja bisa dibaca langsung dari selnya, bukan dari AI.
    """
    ext = Path(nama).suffix.lower()
    if ext == ".pdf":
        return {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": base64.standard_b64encode(isi).decode(),
            },
            "title": nama,
        }, {}
    if ext == ".xlsx":
        sel = baca_sel(isi)
        return {"type": "text", "text": f"# Lembar kerja: {nama}\n\n{sel_ke_teks(sel)}"}, sel
    if ext == ".docx":
        teks = extract_text_from_docx(Document(BytesIO(isi)))
        return {"type": "text", "text": f"# Dokumen Word: {nama}\n\n{teks}"}, {}
    raise ValueError(f"Jenis file tidak didukung: {nama}")


# =============================================================================
# Permintaan ke Claude
# =============================================================================

_SKEMA = {
    "type": "object",
    "properties": {
        "temuan": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "nama": {"type": "string"},
                    "kandidat": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "sumber": {"type": "string"},
                                "lokasi": {"type": "string"},
                                "nilai": {"type": "string"},
                            },
                            "required": ["sumber", "lokasi", "nilai"],
                            "additionalProperties": False,
                        },
                    },
                    "rekomendasi": {"type": "string"},
                    "konsisten": {"type": "boolean"},
                    "penjelasan": {"type": "string"},
                    "kemunculan": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "dokumen": {"type": "string"},
                                "paragraf": {"type": "integer"},
                                "teks_lama": {"type": "string"},
                                "bentuk": {"type": "string", "enum": BENTUK},
                            },
                            "required": ["dokumen", "paragraf", "teks_lama", "bentuk"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": [
                    "nama", "kandidat", "rekomendasi", "konsisten", "penjelasan", "kemunculan",
                ],
                "additionalProperties": False,
            },
        },
        "catatan": {"type": "string"},
    },
    "required": ["temuan", "catatan"],
    "additionalProperties": False,
}

_SYSTEM = """\
Anda membantu penilai di Kantor Jasa Penilai Publik (KJPP) menyusun laporan \
penilaian properti. Anda menerima dokumen sumber — lembar kerja penilaian \
(Excel, per sel "Sheet!C5"), sertifikat tanah, IMB/PBG, proposal/surat \
penugasan, dsb. — dan satu atau beberapa file template (bab-bab laporan, \
surat pengantar, ringkasan, dsb.) yang paragrafnya bernomor [n] per file. \
Template bisa berisi placeholder {{...}} atau data objek dari laporan lama. \
Satu data bisa muncul di beberapa file; semuanya harus diisi sama.

Tugas:
1. Temukan setiap DATA OBJEK di template yang harus berasal dari dokumen sumber: \
mis. nama pemberi tugas, nomor & tanggal proposal/penugasan, tujuan penilaian, \
tanggal inspeksi/penilaian/laporan, alamat objek, kelurahan/kecamatan/kota, \
jenis & nomor sertifikat, NIB, nama pemegang hak, luas tanah, nomor & tanggal \
IMB/PBG, luas bangunan, peruntukan, nilai pasar/likuidasi, dsb. Jangan \
sertakan kalimat baku, metodologi, asumsi, atau opini penilai.
2. Untuk setiap data, catat nilainya di SETIAP sumber yang memuatnya \
("kandidat"): "sumber" = nama file persis, "lokasi" = halaman PDF atau sel \
Excel persis ("Sheet!C5"), "nilai" = nilainya, ditulis bergaya laporan \
Indonesia (angka 1.250.000.000; tanggal 28 April 2026). Jangan menebak; \
lewati sumber yang tidak memuat data itu.
3. "rekomendasi" = nama file sumber yang paling berwenang untuk data itu: \
data hukum tanah dari sertifikat, data bangunan dari IMB/PBG, data penugasan \
dari proposal, angka hasil penilaian dari lembar kerja. Kosongkan bila tidak \
ada kandidat.
4. "konsisten" = apakah semua kandidat menyatakan fakta yang sama (abaikan \
beda format/singkatan, mis. "Jl." vs "Jalan"). Bila tidak, jelaskan bedanya \
singkat di "penjelasan".
5. "kemunculan" = setiap tempat data itu muncul di SEMUA file template: \
"dokumen" = nama file template persis; "paragraf" = nomor [n] di file itu; \
"teks_lama" = teks persis dari paragraf yang harus diganti \
(placeholder lengkap dengan kurungnya, atau nilai lama — sependek mungkin, \
tanpa kata di sekitarnya, tidak saling tumpang tindih); "bentuk" = "teks" \
(diganti nilai apa adanya — juga untuk nomor/kode seperti nomor \
sertifikat, NIB, nomor IMB), "angka" (jumlah/besaran: luas, nilai, harga), \
"terbilang" (terbilang huruf kapital, mis. DUA RATUS JUTA RUPIAH) atau "terbilang_judul" (Dua Ratus Juta \
Rupiah). Terbilang dibuat aplikasi dari angka terpilih.
6. Sertakan juga data yang di template sudah benar, agar penilai bisa \
memeriksa konsistensi antarsumber.
Semua teks penjelasan dalam bahasa Indonesia, singkat.\
"""


def analisis_dokumen(
    client: anthropic.Anthropic,
    template: dict[str, dict[int, str]],
    sumber: list[tuple[str, bytes]],
    effort: str = "medium",
    progres: Callable[[int], None] | None = None,
) -> tuple[list[Temuan], str]:
    """
    Kirim semua template + sumber ke Claude dalam satu permintaan, agar data
    yang sama diisi seragam di semua bab dan surat.
    template: {nama file: {nomor paragraf: teks}}.
    Nilai kandidat dari lembar kerja diganti dengan isi sel aslinya bila
    lokasinya menunjuk satu sel.
    """
    konten: list[dict] = []
    sel_per_file: dict[str, dict[str, str]] = {}
    for nama, isi in sumber:
        blok, sel = blok_sumber(nama, isi)
        konten.append(blok)
        if sel:
            sel_per_file[nama] = sel
    # Sumber jarang berubah antar-percobaan: cache agar analisis ulang lebih murah
    konten[-1]["cache_control"] = {"type": "ephemeral"}

    for nama, paragraf in template.items():
        isi = "\n".join(f"[{i}] {teks}" for i, teks in paragraf.items())
        konten.append({"type": "text", "text": f"# Template: {nama} (paragraf bernomor)\n\n{isi}"})

    data = minta_json(client, _SYSTEM, konten, _SKEMA, effort=effort, progres=progres)

    temuan = []
    for t in data["temuan"]:
        kandidat = []
        for k in t["kandidat"]:
            nilai = k["nilai"]
            # Lembar kerja: percayai isi sel, bukan salinan AI. Bila angkanya sama,
            # pertahankan tulisan Claude yang sudah berformat (mis. 595.400.000).
            if (sel := sel_per_file.get(k["sumber"])) and k["lokasi"] in sel:
                if _digit(sel[k["lokasi"]]) != _digit(nilai) or not _digit(nilai):
                    nilai = sel[k["lokasi"]]
            kandidat.append(Kandidat(k["sumber"], k["lokasi"], nilai))
        temuan.append(Temuan(
            nama=t["nama"],
            kandidat=kandidat,
            rekomendasi=t["rekomendasi"],
            konsisten=t["konsisten"],
            penjelasan=t["penjelasan"],
            kemunculan=[Kemunculan(**m) for m in t["kemunculan"]],
        ))
    return temuan, data["catatan"]


# =============================================================================
# Dari pilihan penilai ke perubahan laporan
# =============================================================================

def _digit(teks: str) -> str:
    return re.sub(r"\D", "", teks)


def teks_pengganti(kemunculan: Kemunculan, nilai: str) -> str | None:
    """
    Teks yang ditulis di laporan untuk satu kemunculan. None bila bentuknya
    angka/terbilang tetapi nilainya bukan angka.
    """
    if kemunculan.bentuk == "teks":
        return nilai
    angka = parse_ke_int(nilai)
    if angka is None:
        return None
    if kemunculan.bentuk == "terbilang":
        return angka_ke_terbilang(angka)
    if kemunculan.bentuk == "terbilang_judul":
        return angka_ke_terbilang_title(angka)
    # Ikuti gaya angka di template: "1.000.000,00", "1.000.000", atau "250"
    lama = kemunculan.teks_lama.strip()
    if re.search(r",\d{2}$", lama):
        return format_angka_indonesia(angka)
    # Template bertitik ribuan, atau placeholder (tanpa contoh format): beri titik ribuan
    if re.search(r"\d\.\d{3}", lama) or "{{" in lama:
        return format_ribuan(angka)
    # Tanpa pemisah ribuan (mis. luas, nomor): angka apa adanya, nol di depan tetap
    return nilai.strip() if nilai.strip().isdigit() else str(angka)


def rencana_perubahan(
    temuan: list[Temuan], nilai_dipakai: dict[int, str]
) -> tuple[dict[str, list[Usulan]], list[str]]:
    """
    nilai_dipakai: {indeks temuan: nilai} — temuan yang tidak ada di sini
    tidak diubah. Kembalikan ({nama template: perubahan}, peringatan).
    """
    perubahan: dict[str, list[Usulan]] = {}
    peringatan: list[str] = []
    for i, nilai in nilai_dipakai.items():
        t = temuan[i]
        for m in t.kemunculan:
            baru = teks_pengganti(m, nilai)
            if baru is None:
                peringatan.append(
                    f"{t.nama}: “{nilai}” bukan angka, {m.bentuk} di {m.dokumen} "
                    f"paragraf [{m.paragraf}] dilewati"
                )
            elif baru != m.teks_lama:
                perubahan.setdefault(m.dokumen, []).append(
                    Usulan(m.paragraf, m.teks_lama, baru, t.nama, "", "")
                )
    return perubahan, peringatan


# =============================================================================
# Simpan / muat hasil analisis
# =============================================================================

def temuan_ke_dict(temuan: list[Temuan], catatan: str) -> dict:
    return {"temuan": [asdict(t) for t in temuan], "catatan": catatan}


def temuan_dari_dict(data: dict) -> tuple[list[Temuan], str]:
    temuan = [
        Temuan(
            nama=t["nama"],
            kandidat=[Kandidat(**k) for k in t["kandidat"]],
            rekomendasi=t["rekomendasi"],
            konsisten=t["konsisten"],
            penjelasan=t["penjelasan"],
            kemunculan=[Kemunculan(**m) for m in t["kemunculan"]],
        )
        for t in data["temuan"]
    ]
    return temuan, data["catatan"]
