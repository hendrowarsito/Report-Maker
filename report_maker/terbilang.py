"""Terbilang Rupiah, format angka/tanggal Indonesia, dan placeholder turunan."""
import datetime
import numbers
import re

import pandas as pd


# =============================================================================
# Konversi angka ke terbilang Rupiah (murni Python, tanpa API)
# =============================================================================

_SATUAN = [
    "", "satu", "dua", "tiga", "empat", "lima",
    "enam", "tujuh", "delapan", "sembilan",
]
_BELASAN = [
    "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas",
    "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas",
]


def _terbilang_bilangan(n: int) -> str:
    """Rekursif: ubah bilangan bulat positif ke kata Indonesia (huruf kecil)."""
    if n == 0:
        return ""

    parts: list[str] = []

    if n >= 1_000_000_000_000:
        t, n = divmod(n, 1_000_000_000_000)
        parts.append(_terbilang_bilangan(t) + " triliun")

    if n >= 1_000_000_000:
        m, n = divmod(n, 1_000_000_000)
        parts.append(_terbilang_bilangan(m) + " miliar")

    if n >= 1_000_000:
        j, n = divmod(n, 1_000_000)
        parts.append(_terbilang_bilangan(j) + " juta")

    if n >= 1_000:
        r, n = divmod(n, 1_000)
        parts.append("seribu" if r == 1 else _terbilang_bilangan(r) + " ribu")

    if n >= 100:
        h, n = divmod(n, 100)
        parts.append("seratus" if h == 1 else _SATUAN[h] + " ratus")

    if n >= 20:
        p, n = divmod(n, 10)
        parts.append(_SATUAN[p] + " puluh")
        if n:
            parts.append(_SATUAN[n])
    elif n >= 10:
        parts.append(_BELASAN[n - 10])
    elif n > 0:
        parts.append(_SATUAN[n])

    return " ".join(parts)


def angka_ke_terbilang(n: int) -> str:
    """
    Ubah bilangan bulat ke terbilang HURUF KAPITAL SEMUA.
    Contoh: 200_000_000 → 'DUA RATUS JUTA RUPIAH'
    Dipakai untuk placeholder {{KEY_TERBILANG}}.
    """
    if n == 0:
        return "NOL RUPIAH"
    prefix = "MINUS " if n < 0 else ""
    return (prefix + _terbilang_bilangan(abs(n))).upper() + " RUPIAH"


def angka_ke_terbilang_title(n: int) -> str:
    """
    Ubah bilangan bulat ke terbilang Huruf Besar Di Awal Kata (Title Case).
    Contoh: 200_000_000 → 'Dua Ratus Juta Rupiah'
    Dipakai untuk placeholder {{KEY_TERBILANG_TITLE}}.
    """
    return angka_ke_terbilang(n).title()


def format_rupiah(n: int) -> str:
    """
    Format bilangan bulat ke string Rupiah Indonesia.
    Contoh: 200_000_000 → 'Rp 200.000.000,00'
    """
    ribuan = f"{n:,}".replace(",", ".")
    return f"Rp {ribuan},00"


def parse_ke_int(nilai) -> int | None:
    """
    Coba parsing berbagai format nilai ke bilangan bulat.
    Mengembalikan None jika:
      - bukan angka
      - mengandung huruf selain prefix "Rp" (misal: tanggal, nama, %)
      - hasil < 1

    Format yang didukung: int, float, "200000000", "Rp 200.000.000,00", "200.000.000,00"
    """
    if isinstance(nilai, (int, float)):
        result = int(round(float(nilai)))
        return result if result >= 1 else None

    s = str(nilai).strip()

    # Tolak string yang mengandung huruf di luar prefix Rp, atau simbol non-numerik
    # Contoh yang DITOLAK: "28 April 2026", "Unit A", "20%", "n/a", "1.5x"
    tanpa_rp = re.sub(r"(?i)^\s*rp\.?\s*", "", s).strip()
    if re.search(r"[a-zA-Z%]", tanpa_rp):
        return None

    # Hapus prefix Rp
    s = tanpa_rp
    # Format Indonesia: titik = pemisah ribuan, koma = desimal
    if re.search(r"\d\.\d{3}", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")

    s = re.sub(r"[^\d.]", "", s)
    if not s:
        return None
    try:
        result = int(round(float(s)))
        return result if result >= 1 else None
    except ValueError:
        return None


# Angka Indonesia dengan awalan/satuan yang lazim di laporan penilaian:
# "Rp 1.250.000", "1.974,00 m²", "12,5%", "Rp 1.000.000,-"
_POLA_ANGKA_BERSATUAN = re.compile(
    r"^\s*(?:rp\.?\s*)?(-?\d[\d.,]*)\s*(?:m²|m2|m\^2|meter persegi|%|,-|m)?\s*$",
    re.IGNORECASE,
)


def parse_angka(nilai) -> float | None:
    """
    Angka dari teks bergaya Indonesia, boleh dengan "Rp" dan satuan (m², %).
    Titik = pemisah ribuan, koma = desimal. None bila bukan angka (mis. tanggal).
      "1.974,00 m²" → 1974.0   "Rp 595.400.000" → 595400000.0   "0,25" → 0.25
    """
    if isinstance(nilai, bool):
        return None
    if isinstance(nilai, (int, float)):
        return float(nilai)
    m = _POLA_ANGKA_BERSATUAN.match(str(nilai).replace("\xa0", " "))
    if not m:
        return None
    angka = m.group(1).rstrip(".,")
    if "," in angka:
        angka = angka.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", angka):
        angka = angka.replace(".", "")
    try:
        return float(angka)
    except ValueError:
        return None


def format_angka_indonesia(n: int) -> str:
    """
    Format angka ke gaya Indonesia tanpa prefix Rp.
    Contoh: 200_000_000 → '200.000.000,00'
    Digunakan ketika template sudah hardcode 'Rp ' sebelum placeholder.
    """
    return f"{n:,}".replace(",", ".") + ",00"


_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def format_tanggal(d: datetime.date) -> str:
    """Contoh: date(2026, 4, 28) → '28 April 2026'"""
    return f"{d.day} {_BULAN[d.month - 1]} {d.year}"


# Key yang nilainya nomor/kode, bukan besaran: jangan diberi titik ribuan
_POLA_KEY_IDENTITAS = re.compile(
    r"(^|_)(nib|nik|npwp|nop|kode|kodepos|pos|telp|telepon|hp|phone|fax|nomor|no|"
    r"rek|rekening|tahun|thn|year|id|rt|rw|blok|kav)(_|$|\d)",
    re.IGNORECASE,
)
NILAI_ERROR_EXCEL = {
    "#NAME?", "#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!",
    "#SPILL!", "#CALC!", "#GETTING_DATA", "#FIELD!", "#BLOCKED!", "#UNKNOWN!",
}
BATAS_RIBUAN = 10_000   # di bawah ini (mis. tahun 2026, luas 416) angka ditulis apa adanya


def format_ribuan(n: float, desimal: int | None = None) -> str:
    """1234567 → '1.234.567'; 1234.5 → '1.234,5' (desimal koma, tanpa nol berlebih)."""
    if desimal is None:
        teks = f"{n:,.6f}".rstrip("0").rstrip(".")
    else:
        teks = f"{n:,.{desimal}f}"
    return teks.replace(",", "_").replace(".", ",").replace("_", ".")


def format_nilai(value, kunci: str = "") -> str:
    """
    Ubah nilai sel Excel ke teks siap tempel di laporan.
      • kosong / NaN          → ''
      • 595400000             → '595.400.000'  (titik ribuan bila ≥ 10.000)
      • 2026, 416             → '2026', '416'  (angka kecil apa adanya)
      • 0.25                  → '0,25'         (desimal koma)
      • datetime(2026, 4, 28) → '28 April 2026'
      • teks                  → apa adanya (spasi di tepi dibuang)
    kunci: nama data (mis. "kode_pos", "nib") — bila berupa nomor/kode,
    angka ditulis tanpa titik ribuan.
    """
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, datetime.datetime):        # termasuk pd.Timestamp
        teks = format_tanggal(value)
        if (value.hour, value.minute) != (0, 0):
            teks += f" {value:%H:%M}"
        return teks
    if isinstance(value, datetime.date):
        return format_tanggal(value)
    if isinstance(value, datetime.time):
        return f"{value:%H:%M}"
    if isinstance(value, numbers.Real):
        angka = float(value)
        if angka.is_integer():
            angka = int(angka)
        identitas = bool(kunci) and _POLA_KEY_IDENTITAS.search(kunci) is not None
        if abs(angka) >= BATAS_RIBUAN and not identitas:
            return format_ribuan(angka)
        if isinstance(angka, int):
            return str(angka)
        return f"{angka:.6f}".rstrip("0").rstrip(".").replace(".", ",")
    teks = str(value).strip()
    # Error rumus Excel (#NAME?, #REF!, …) dan isian sementara ("xxx") = belum ada nilai
    if teks.upper() in NILAI_ERROR_EXCEL or re.fullmatch(r"x{3,}", teks, re.IGNORECASE):
        return ""
    return teks


# Placeholder turunan dibuat otomatis dari key Excel berdasarkan akhiran namanya.
# Urutan penting: akhiran yang lebih panjang dicek lebih dulu
# (_TERBILANG_TITLE sebelum _TERBILANG, _KALIMAT_CAP sebelum _KALIMAT).
_SUFIKS_TURUNAN = {
    "_TERBILANG_TITLE": angka_ke_terbilang_title,
    "_KALIMAT_CAP":     angka_ke_terbilang,         # {{NP_kalimat_cap}} → DUA RATUS … RUPIAH
    "_TERBILANG":       angka_ke_terbilang,
    "_KALIMAT":         angka_ke_terbilang_title,   # {{NP_kalimat}} → Dua Ratus … Rupiah
    "_FORMAT":          format_rupiah,
}
# Nama key angka sumber yang lazim untuk turunan "NP_…": NP, NP_nominal, NP_nilai, NP_angka
_AKHIRAN_SUMBER = ("", "_nominal", "_nilai", "_angka")


def kunci_sumber(dasar: str, kunci_tersedia) -> str:
    """
    Key angka untuk turunan berdasar `dasar`: "NP" → "NP_nominal" bila key
    itulah yang ada. Tidak peka huruf besar/kecil. Bila tidak ada, `dasar`.
    """
    tersedia = {str(k).casefold(): k for k in kunci_tersedia}
    for akhiran in _AKHIRAN_SUMBER:
        if (k := tersedia.get((dasar + akhiran).casefold())) is not None:
            return k
    return dasar


def pecah_turunan(ph: str):
    """'NP_nominal_TERBILANG' → ('NP_nominal', angka_ke_terbilang); bukan turunan → (None, None)."""
    for sufiks, fungsi in _SUFIKS_TURUNAN.items():
        if ph.upper().endswith(sufiks) and len(ph) > len(sufiks):
            return ph[: -len(sufiks)], fungsi
    return None, None


def generate_auto_terbilang(
    data_dict: dict[str, str], placeholders: set[str]
) -> tuple[dict[str, str], dict[str, str]]:
    """
    Buat nilai untuk placeholder turunan yang dipakai di template, berdasarkan
    konvensi nama (akhiran tidak peka huruf besar/kecil):

      {{NP_nominal_TERBILANG}} / {{NP_kalimat_cap}}  → "DUA RATUS JUTA RUPIAH"
      {{NP_nominal_TERBILANG_TITLE}} / {{NP_kalimat}} → "Dua Ratus Juta Rupiah"
      {{NP_nominal_FORMAT}}                          → "Rp 200.000.000,00"

    Angka sumbernya key dasar atau varian umumnya (NP, NP_nominal, NP_nilai,
    NP_angka). Terbilang SELALU dihitung dari angka — teks terbilang di Excel
    (sering rusak, mis. #NAME? dari rumus/makro) tidak dipakai — agar pasti
    sesuai nominalnya. Key sumber ikut diformat Indonesia: "200.000.000,00".

    Kembalikan (nilai_tambahan, peringatan) — peringatan berisi placeholder
    turunan yang tidak bisa dibuat beserta alasannya.
    """
    tambahan:   dict[str, str] = {}
    peringatan: dict[str, str] = {}

    for ph in sorted(placeholders):
        dasar, fungsi = pecah_turunan(ph)
        if dasar is None:
            continue
        sumber = kunci_sumber(dasar, data_dict)
        if sumber not in data_dict:
            if not data_dict.get(ph):
                peringatan[ph] = f"Angka sumber {{{{{dasar}}}}} / {{{{{dasar}_nominal}}}} tidak ada"
            continue
        angka = parse_angka(data_dict[sumber])
        if angka is None or angka == 0:
            alasan = "bernilai 0" if angka == 0 else "bukan angka"
            peringatan[ph] = f"Nilai {{{{{sumber}}}}} {alasan}"
            continue

        tambahan[ph]     = fungsi(round(angka))
        tambahan[sumber] = format_angka_indonesia(round(angka))

    return tambahan, peringatan
