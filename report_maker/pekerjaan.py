"""
Analisis di latar belakang server.

Analisis berjalan di thread terpisah, tidak terikat pada koneksi browser:
bila HP tidur atau pengguna pindah aplikasi, analisis tetap selesai. Hasilnya
disimpan di disk dengan kunci sidik isi file, sehingga meng-upload ulang file
yang sama langsung memuat hasilnya tanpa memanggil AI lagi.
"""
import json
import tempfile
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from .analisis import Temuan, temuan_dari_dict, temuan_ke_dict

FOLDER = Path(tempfile.gettempdir()) / "report_maker_analisis"
UMUR_MAKS_DETIK = 7 * 24 * 3600   # hasil lebih tua dari 7 hari dihapus


@dataclass
class Pekerjaan:
    status: str = "berjalan"          # berjalan | selesai | gagal
    mulai: float = field(default_factory=time.time)
    karakter: int = 0                 # panjang jawaban Claude yang sudah diterima
    galat: str = ""


_pekerjaan: dict[str, Pekerjaan] = {}
_kunci = threading.Lock()


def _berkas(kunci: str) -> Path:
    return FOLDER / f"{kunci}.json"


def simpan_hasil(kunci: str, temuan: list[Temuan], catatan: str) -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    sementara = _berkas(kunci).with_suffix(".tmp")
    sementara.write_text(json.dumps(temuan_ke_dict(temuan, catatan), ensure_ascii=False))
    sementara.replace(_berkas(kunci))   # atomik: tidak pernah terbaca setengah jadi


def muat_hasil(kunci: str) -> tuple[list[Temuan], str] | None:
    try:
        return temuan_dari_dict(json.loads(_berkas(kunci).read_text()))
    except (OSError, ValueError, KeyError, TypeError):
        return None


def hapus_hasil(kunci: str) -> None:
    _berkas(kunci).unlink(missing_ok=True)


def bersihkan_lama() -> None:
    batas = time.time() - UMUR_MAKS_DETIK
    for f in FOLDER.glob("*.json"):
        try:
            if f.stat().st_mtime < batas:
                f.unlink()
        except OSError:
            pass


def status(kunci: str) -> Pekerjaan | None:
    return _pekerjaan.get(kunci)


def mulai(
    kunci: str,
    fungsi: Callable[[Callable[[int], None]], tuple[list[Temuan], str]],
) -> Pekerjaan:
    """
    Jalankan fungsi(progres) di thread latar belakang dan simpan hasilnya.
    Bila pekerjaan dengan kunci sama masih berjalan, kembalikan yang itu.
    """
    with _kunci:
        p = _pekerjaan.get(kunci)
        if p is not None and p.status == "berjalan":
            return p
        p = Pekerjaan()
        _pekerjaan[kunci] = p

    def _progres(n: int) -> None:
        p.karakter = n

    def _jalan() -> None:
        try:
            temuan, catatan = fungsi(_progres)
            simpan_hasil(kunci, temuan, catatan)
            p.status = "selesai"
        except Exception as e:  # dilaporkan ke halaman, bukan dibuang diam-diam
            p.galat = str(e) or type(e).__name__
            p.status = "gagal"

    threading.Thread(target=_jalan, daemon=True, name=f"analisis-{kunci[:8]}").start()
    return p
