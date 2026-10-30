"""Pemeriksaan data tanpa API: kelengkapan placeholder, nominal vs terbilang,
dan deteksi baris awal data di sheet Excel."""
import re

import pandas as pd

from .terbilang import pecah_turunan, angka_ke_terbilang, format_rupiah, parse_ke_int


# =============================================================================
# Kecocokan placeholder template vs data Excel (tanpa API)
# =============================================================================

def build_placeholder_check(
    placeholders: dict[str, list[str]],
    data_excel: dict[str, str],
    data_final: dict[str, str],
    peringatan: dict[str, str],
    asal_otomatis: dict[str, str] | None = None,
) -> pd.DataFrame:
    """
    Satu baris per placeholder di template, dengan nilai yang akan dipakai
    dan statusnya: dari Excel, dibuat otomatis, kosong, atau tidak ada.
    asal_otomatis: {placeholder turunan: key angka sumbernya} untuk keterangan.
    """
    asal_otomatis = asal_otomatis or {}
    rows = []
    for ph in sorted(placeholders, key=str.lower):
        if ph in peringatan:
            status = f"⚠️ {peringatan[ph]}"
        elif ph not in data_final:
            status = "❌ Tidak ada di Excel"
        elif not data_final[ph]:
            status = "⚠️ Nilai kosong"
        elif ph in asal_otomatis:
            status = f"✅ Otomatis dari {{{{{asal_otomatis[ph]}}}}}"
        elif ph in data_excel:
            status = "✅ Dari Excel"
        else:
            status = "✅ Otomatis"
        rows.append({
            "Placeholder": f"{{{{{ph}}}}}",
            "Nilai":       data_final.get(ph, ""),
            "Status":      status,
            "Dokumen":     ", ".join(placeholders[ph]),
        })
    return pd.DataFrame(rows, columns=["Placeholder", "Nilai", "Status", "Dokumen"])


# =============================================================================
# Konsistensi nominal angka vs terbilang (dihitung lokal — tidak butuh API)
# =============================================================================

# "Rp 200.000.000,00 (DUA RATUS JUTA RUPIAH)" — angka dan terbilang ditangkap
# berpasangan sehingga satu terbilang yang hilang tidak menggeser pasangan lain.
_POLA_NOMINAL = re.compile(
    r"Rp\.?\s*(\d[\d.]*(?:,\d{1,2})?)\s*\(([^()]*rupiah[^()]*)\)", re.IGNORECASE
)


def check_value_consistency_table(text: str) -> pd.DataFrame:
    """
    Cocokkan setiap nominal "Rp ... (... RUPIAH)" dengan terbilang yang benar
    menurut angka_ke_terbilang. Tidak peka huruf besar/kecil.
    """
    rows = []
    for angka_teks, kata in _POLA_NOMINAL.findall(text):
        angka = parse_ke_int(angka_teks)
        if angka is None:
            continue
        kata_norm = " ".join(kata.upper().split())
        benar     = angka_ke_terbilang(angka)
        rows.append({
            "Angka (Rp)":          format_rupiah(angka),
            "Terbilang di Dokumen": kata.strip(),
            "Status": "✅ Sesuai" if kata_norm == benar else f"❌ Seharusnya: {benar}",
        })
    return pd.DataFrame(rows, columns=["Angka (Rp)", "Terbilang di Dokumen", "Status"])


_POLA_KEY = re.compile(r"^[A-Za-z_][\w.-]*$")


def deteksi_baris_awal(
    kolom_key: list[str], kolom_nilai: list[str], placeholders: set[str]
) -> int:
    """
    Tebak baris pertama data (key | nilai) di sheet Excel, berbasis 0.

    Cari baris pertama yang key-nya dipakai di template, lalu mundur selama
    baris di atasnya masih berbentuk key (tanpa spasi) dan punya nilai — agar
    key yang tidak dipakai template di awal tabel tetap ikut. Kop/judul
    ("Laporan Penilaian", "LAPORAN" tanpa nilai) tidak ikut terambil.
    Kembalikan 0 jika tidak ada key template yang ditemukan.
    """
    dikenal = placeholders | {pecah_turunan(ph)[0] for ph in placeholders}
    for i, key in enumerate(kolom_key):
        if key in dikenal:
            while i > 0 and _POLA_KEY.match(kolom_key[i - 1]) and kolom_nilai[i - 1]:
                i -= 1
            return i
    return 0
