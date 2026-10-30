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
from collections import Counter
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from io import BytesIO
from pathlib import Path

import anthropic
from docx import Document

from .ai import minta_json
from .docx_engine import extract_text_from_docx, ikuti_kapitalisasi
from .lembar_kerja import baca_sel, sel_ke_teks
from .penyesuaian import Usulan
from .terbilang import (
    angka_ke_terbilang,
    angka_ke_terbilang_title,
    format_ribuan,
    parse_angka,
)

BENTUK = ["teks", "angka", "terbilang", "terbilang_judul"]
JENIS_SUMBER = ["sertifikat", "imb_pbg", "proposal", "lembar_kerja", "lainnya"]
KATEGORI = ["sertifikat", "bangunan", "penugasan", "penilaian", "lainnya"]

# Urutan sumber per kategori data — aturan tetap, bukan pilihan AI:
# data sertifikat dari sertifikat; bila tidak ada/tidak jelas, dari lembar kerja.
URUTAN_SUMBER = {
    "sertifikat": ["sertifikat", "lembar_kerja"],
    "bangunan":   ["imb_pbg", "lembar_kerja"],
    "penugasan":  ["proposal", "lembar_kerja"],
    "penilaian":  ["lembar_kerja"],
    "lainnya":    ["lembar_kerja"],
}
LABEL_JENIS = {
    "sertifikat": "sertifikat", "imb_pbg": "IMB/PBG", "proposal": "proposal",
    "lembar_kerja": "lembar kerja", "lainnya": "dokumen lain",
}


@dataclass
class Kandidat:
    sumber: str      # nama file
    lokasi: str      # "hal. 1" / "LK Tanah!C5"
    nilai: str
    jelas: bool = True          # terbaca jelas (bukan buram/coretan/tulisan tangan ragu)
    jenis: str = "lainnya"      # jenis dokumen sumber, salah satu JENIS_SUMBER


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
    rekomendasi: str = ""        # nama file sumber yang disarankan Claude
    konsisten: bool = True
    penjelasan: str = ""
    kemunculan: list[Kemunculan] = field(default_factory=list)
    kategori: str = "lainnya"    # salah satu KATEGORI


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
                    "kategori": {"type": "string", "enum": KATEGORI},
                    "kandidat": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "sumber": {"type": "string"},
                                "lokasi": {"type": "string"},
                                "nilai": {"type": "string"},
                                "jelas": {"type": "boolean"},
                            },
                            "required": ["sumber", "lokasi", "nilai", "jelas"],
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
                    "nama", "kategori", "kandidat", "rekomendasi", "konsisten",
                    "penjelasan", "kemunculan",
                ],
                "additionalProperties": False,
            },
        },
        "sumber": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "nama": {"type": "string"},
                    "jenis": {"type": "string", "enum": JENIS_SUMBER},
                },
                "required": ["nama", "jenis"],
                "additionalProperties": False,
            },
        },
        "catatan": {"type": "string"},
    },
    "required": ["sumber", "temuan", "catatan"],
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
0. "sumber": jenis setiap dokumen sumber — sertifikat (SHM, SHGB, SHGU, \
SHMSRS, dsb.), imb_pbg, proposal (proposal/surat penugasan), lembar_kerja \
(Excel penilaian), atau lainnya.
1. Temukan setiap DATA OBJEK di template yang harus berasal dari dokumen sumber: \
mis. nama pemberi tugas, nomor & tanggal proposal/penugasan, tujuan penilaian, \
tanggal inspeksi/penilaian/laporan, alamat objek, kelurahan/kecamatan/kota, \
jenis & nomor sertifikat, NIB, nama pemegang hak, luas tanah, nomor & tanggal \
IMB/PBG, luas bangunan, peruntukan, nilai pasar/likuidasi, dsb. Jangan \
sertakan kalimat baku, metodologi, asumsi, atau opini penilai.
2. "kategori" data: sertifikat (semua data yang tercantum di sertifikat: \
jenis & nomor hak, NIB, luas tanah, pemegang hak, tempat & tanggal terbit, \
nomor & tanggal surat ukur, letak tanah desa/kecamatan/kabupaten/provinsi), \
bangunan (data IMB/PBG), penugasan (pemberi tugas, debitur, nomor & tanggal \
proposal, tujuan), penilaian (tanggal penilaian/inspeksi, nilai pasar/ \
likuidasi, kurs, hasil perhitungan), atau lainnya.
3. Untuk setiap data, catat nilainya di SETIAP sumber yang memuatnya \
("kandidat"): "sumber" = nama file persis, "lokasi" = halaman PDF atau sel \
Excel persis ("Sheet!C5"), "nilai" = nilainya, ditulis bergaya laporan \
Indonesia (angka 1.250.000.000; tanggal 28 April 2026). Bila ada beberapa \
sertifikat, catat nilai dari MASING-MASING sertifikat sebagai kandidat \
tersendiri; bila satu file memuat nilai berbeda di beberapa tempat, catat \
semuanya. "jelas" = false bila nilai di dokumen buram, tercoret/diganti, \
tulisan tangan yang meragukan, atau Anda tidak yakin membacanya. Jangan \
menebak; lewati sumber yang tidak memuat data itu. "rekomendasi" = nama file \
yang menurut Anda paling tepat (aplikasi menerapkan aturan sumber sendiri).
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
7. Jangan masukkan kemunculan yang di template sudah benar dan hanya beda \
huruf besar/kecil atau gaya penulisan (mis. "PT ABC INDONESIA" vs \
"PT ABC Indonesia") — gaya template dipertahankan.
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

    nama_sumber = [n for n, _ in sumber]
    jenis_ai = {
        _cocokkan_nama(s["nama"], nama_sumber): s["jenis"] for s in data.get("sumber", [])
    }
    jenis = {n: jenis_ai.get(n) or tebak_jenis_sumber(n) for n in nama_sumber}

    temuan = []
    for t in data["temuan"]:
        kandidat = []
        for k in t["kandidat"]:
            k["sumber"] = _cocokkan_nama(k["sumber"], nama_sumber)
            nilai = k["nilai"]
            # Lembar kerja: percayai isi sel, bukan salinan AI. Bila angkanya sama,
            # pertahankan tulisan Claude yang sudah berformat (mis. 595.400.000).
            if (sel := sel_per_file.get(k["sumber"])) and k["lokasi"] in sel:
                if _digit(sel[k["lokasi"]]) != _digit(nilai) or not _digit(nilai):
                    nilai = sel[k["lokasi"]]
            kandidat.append(Kandidat(
                k["sumber"], k["lokasi"], nilai, k.get("jelas", True),
                jenis.get(k["sumber"]) or tebak_jenis_sumber(k["sumber"]),
            ))
        temuan.append(Temuan(
            nama=t["nama"],
            kandidat=kandidat,
            rekomendasi=_cocokkan_nama(t["rekomendasi"], nama_sumber),
            kategori=t.get("kategori", "lainnya"),
            konsisten=t["konsisten"],
            penjelasan=t["penjelasan"],
            kemunculan=[Kemunculan(**m) for m in t["kemunculan"]],
        ))
    return temuan, data["catatan"]


# =============================================================================
# Jenis sumber & aturan pemilihan sumber
# =============================================================================

def _cocokkan_nama(nama: str, daftar: list[str]) -> str:
    """Nama file dari Claude → nama file yang di-upload (abaikan huruf besar/kecil)."""
    if nama in daftar:
        return nama
    return next((d for d in daftar if d.casefold() == nama.casefold()), nama)


_POLA_SERTIFIKAT = re.compile(r"\b(shm|shgb|hgb|shgu|hgu|shp|shmsrs|sertifikat|sertipikat)\b", re.I)
_POLA_IMB = re.compile(r"\b(imb|pbg|slf)\b", re.I)
_POLA_PROPOSAL = re.compile(r"\b(proposal|spn|penawaran|penugasan|spk)\b", re.I)


def tebak_jenis_sumber(nama: str) -> str:
    """Cadangan bila Claude tidak menyebut jenis: tebak dari nama file."""
    if nama.lower().endswith((".xlsx", ".xlsm")):
        return "lembar_kerja"
    dasar = re.sub(r"[_+\-.]", " ", nama)
    if _POLA_SERTIFIKAT.search(dasar):
        return "sertifikat"
    if _POLA_IMB.search(dasar):
        return "imb_pbg"
    if _POLA_PROPOSAL.search(dasar):
        return "proposal"
    return "lainnya"


def _bisa_dipakai(k: Kandidat, t: Temuan) -> bool:
    """Kandidat layak dipakai: terbaca jelas dan, untuk data angka, bukan 0."""
    if not k.jelas or not k.nilai.strip():
        return False
    berbentuk_angka = any(m.bentuk != "teks" for m in t.kemunculan)
    return not (berbentuk_angka and parse_angka(k.nilai) == 0)


def pilih_sumber(t: Temuan) -> tuple[str, str]:
    """
    Sumber yang dipakai untuk satu data, menurut aturan tetap per kategori
    (URUTAN_SUMBER). Kembalikan (nama file, dasar pilihan) — nama kosong bila
    tidak ada kandidat yang layak.
      Data sertifikat: sertifikat → bila tidak ada/tidak jelas: lembar kerja.
    """
    urutan = URUTAN_SUMBER.get(t.kategori, URUTAN_SUMBER["lainnya"])
    for posisi, jenis in enumerate(urutan):
        k = next((k for k in t.kandidat if k.jenis == jenis and _bisa_dipakai(k, t)), None)
        if k is None:
            continue
        dasar = LABEL_JENIS[jenis].capitalize()
        if posisi > 0:
            utama = LABEL_JENIS[urutan[0]]
            ada_utama = any(x.jenis == urutan[0] for x in t.kandidat)
            dasar += f" ({utama} {'tidak jelas' if ada_utama else 'tidak memuat data ini'})"
        return k.sumber, dasar
    # Di luar aturan: saran Claude, lalu kandidat layak pertama
    k = next((k for k in t.kandidat if k.sumber == t.rekomendasi and _bisa_dipakai(k, t)), None) \
        or next((k for k in t.kandidat if _bisa_dipakai(k, t)), None)
    if k is None:
        return "", "Tidak ada nilai yang layak (kosong, 0, atau tidak jelas)"
    return k.sumber, f"{LABEL_JENIS.get(k.jenis, 'dokumen lain').capitalize()} (saran)"


# =============================================================================
# Dari pilihan penilai ke perubahan laporan
# =============================================================================

def _digit(teks: str) -> str:
    return re.sub(r"\D", "", teks)


_POLA_ANGKA_TEMPLATE = re.compile(r"^(\D*?)(-?\d[\d.,]*\d|\d)(.*)$", re.S)


def teks_pengganti(kemunculan: Kemunculan, nilai: str) -> str | None:
    """
    Teks yang ditulis di laporan untuk satu kemunculan. None bila bentuknya
    angka/terbilang tetapi nilainya bukan angka atau bernilai 0.
    Angka mengikuti gaya teks lama di template: jumlah desimal, titik ribuan,
    dan awalan/satuan ("455,00 m²" + "1.974 m²" → "1.974,00 m²").
    """
    if kemunculan.bentuk == "teks":
        return ikuti_kapitalisasi(kemunculan.teks_lama, nilai)
    angka = parse_angka(nilai)
    if angka is None or angka == 0:
        return None
    if kemunculan.bentuk == "terbilang":
        return angka_ke_terbilang(round(angka))
    if kemunculan.bentuk == "terbilang_judul":
        return angka_ke_terbilang_title(round(angka))

    lama = kemunculan.teks_lama.strip()
    if "{{" in lama:
        # Placeholder tidak memberi contoh format: titik ribuan, desimal bila ada
        return format_ribuan(angka) if abs(angka) >= 1000 or not angka.is_integer() \
            else str(int(angka))
    m = _POLA_ANGKA_TEMPLATE.match(lama)
    if not m:
        return format_ribuan(angka)
    awal, contoh, akhir = m.groups()
    desimal = len(contoh.rsplit(",", 1)[1]) if "," in contoh else 0
    if desimal == 0 and not angka.is_integer():
        desimal = len(f"{angka}".split(".")[1])     # jangan buang desimal nilai
    # Diawali nol ("0999") = nomor/kode: tulis digit apa adanya
    if (re.fullmatch(r"0\d+", contoh) or re.fullmatch(r"0\d+", nilai.strip())) \
            and nilai.strip().isdigit():
        return f"{awal}{nilai.strip()}{akhir}"
    # Besaran (luas, nilai): titik ribuan bila template memakainya atau ≥ 1.000
    pakai_ribuan = bool(re.search(r"\d\.\d{3}", contoh)) or abs(angka) >= 1000
    if pakai_ribuan:
        teks = format_ribuan(angka, desimal)
    elif desimal:
        teks = f"{angka:.{desimal}f}".replace(".", ",")
    else:
        teks = str(int(angka))
    return f"{awal}{teks}{akhir}"


def rencana_perubahan(
    temuan: list[Temuan], nilai_dipakai: dict[int, str]
) -> tuple[dict[str, list[Usulan]], list[str]]:
    """
    nilai_dipakai: {indeks temuan: nilai} — temuan yang tidak ada di sini
    tidak diubah. Kembalikan ({nama template: perubahan}, peringatan) —
    satu peringatan per data, bukan per lokasi.
    """
    perubahan: dict[str, list[Usulan]] = {}
    peringatan: list[str] = []
    for i, nilai in nilai_dipakai.items():
        t = temuan[i]
        dilewati: list[str] = []
        for m in t.kemunculan:
            baru = teks_pengganti(m, nilai)
            if baru is None:
                dilewati.append(Path(m.dokumen).stem)
            elif baru.casefold() != m.teks_lama.casefold():
                perubahan.setdefault(m.dokumen, []).append(
                    Usulan(m.paragraf, m.teks_lama, baru, t.nama, "", "")
                )
        if dilewati:
            alasan = "bernilai 0 (perhitungan di lembar kerja belum selesai?)" \
                if parse_angka(nilai) == 0 else "bukan angka"
            lokasi = ", ".join(f"{d} ({n}×)" for d, n in Counter(dilewati).items())
            peringatan.append(
                f"**{t.nama}**: “{nilai}” {alasan} — {len(dilewati)} lokasi angka/terbilang "
                f"tidak diubah ({lokasi}). Pilih sumber lain atau isi manual."
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
            kandidat=[
                Kandidat(**{"jenis": tebak_jenis_sumber(k["sumber"]), **k}) for k in t["kandidat"]
            ],
            rekomendasi=t["rekomendasi"],
            kategori=t.get("kategori", "lainnya"),
            konsisten=t["konsisten"],
            penjelasan=t["penjelasan"],
            kemunculan=[Kemunculan(**m) for m in t["kemunculan"]],
        )
        for t in data["temuan"]
    ]
    return temuan, data["catatan"]
