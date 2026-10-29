"""Preview laporan sebelum diunduh: gambar halaman (dari PDF) dan daftar perubahan teks."""
import html
import re
from io import BytesIO

import pypdfium2 as pdfium

from .penyesuaian import Usulan


def _rapat(teks: str) -> str:
    """Satukan spasi/baris agar teks PDF bisa dicocokkan dengan teks Word."""
    return re.sub(r"\s+", " ", teks.replace("\xa0", " ")).strip().casefold()


def teks_per_halaman(pdf: bytes) -> list[str]:
    doc = pdfium.PdfDocument(pdf)
    try:
        return [doc[i].get_textpage().get_text_range() for i in range(len(doc))]
    finally:
        doc.close()


def gambar_halaman(pdf: bytes, nomor: int, skala: float = 1.3) -> bytes:
    """PNG halaman ke-nomor (mulai 1)."""
    doc = pdfium.PdfDocument(pdf)
    try:
        gambar = doc[nomor - 1].render(scale=skala).to_pil()
    finally:
        doc.close()
    out = BytesIO()
    gambar.save(out, "PNG", optimize=True)
    return out.getvalue()


def halaman_berubah(halaman: list[str], teks_baru: list[str]) -> list[int]:
    """
    Nomor halaman (mulai 1) yang memuat salah satu teks baru. Teks panjang
    dicocokkan dari potongan awalnya karena bisa terpotong pindah halaman.
    """
    kunci = {_rapat(t)[:40] for t in teks_baru if len(_rapat(t)) >= 3}
    hasil = []
    for i, isi in enumerate(halaman, start=1):
        isi = _rapat(isi)
        if any(k in isi for k in kunci):
            hasil.append(i)
    return hasil


def perubahan_html(paragraf: dict[int, str], usulan: list[Usulan]) -> list[tuple[int, str]]:
    """
    Untuk setiap paragraf yang berubah: (nomor, HTML paragraf baru) dengan teks
    lama dicoret dan teks baru disorot.
    """
    per_paragraf: dict[int, list[Usulan]] = {}
    for u in usulan:
        per_paragraf.setdefault(u.paragraf, []).append(u)

    hasil = []
    for nomor in sorted(per_paragraf):
        teks = paragraf.get(nomor, "")
        penanda: dict[str, str] = {}
        for k, u in enumerate(per_paragraf[nomor]):
            tanda = f"\x00{k}\x00"
            if u.teks_lama in teks:
                teks = teks.replace(u.teks_lama, tanda)
                penanda[tanda] = (
                    f'<del style="color:#b42318">{html.escape(u.teks_lama)}</del> '
                    f'<mark style="background:#fff3a3">{html.escape(u.teks_baru)}</mark>'
                )
        isi = html.escape(teks)
        for tanda, ganti in penanda.items():
            isi = isi.replace(html.escape(tanda), ganti)
        hasil.append((nomor, isi))
    return hasil
