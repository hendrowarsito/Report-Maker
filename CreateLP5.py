"""SRR Kalibata Report Maker — titik masuk Streamlit (header + navigasi halaman)."""
import importlib
import sys
from pathlib import Path

import streamlit as st

# Saat repo diperbarui (mis. push ke Streamlit Cloud), Streamlit menjalankan
# ulang halaman tetapi modul report_maker tetap versi lama di memori, sehingga
# impor fungsi baru gagal (ImportError). Muat ulang paket bila filenya berubah.
_FOLDER_PAKET = Path(__file__).parent / "report_maker"
# Urutan dependensi: modul yang diimpor modul lain dimuat ulang lebih dulu
_URUTAN_MODUL = [
    "terbilang", "ai", "docx_engine", "checks", "pdf",
    "lembar_kerja", "penyesuaian", "analisis", "pekerjaan",
]


def _jejak_paket() -> tuple:
    return tuple(sorted(
        (f.name, f.stat().st_mtime_ns, f.stat().st_size) for f in _FOLDER_PAKET.glob("*.py")
    ))


_paket = sys.modules.get("report_maker")
if _paket is not None and getattr(_paket, "_jejak", None) != _jejak_paket():
    for _nama in _URUTAN_MODUL:
        if (_modul := sys.modules.get(f"report_maker.{_nama}")) is not None:
            importlib.reload(_modul)

import report_maker  # noqa: E402

report_maker._jejak = _jejak_paket()

# Logo opsional: taruh file di assets/logo_srr.png (atau .jpg / .svg)
LOGO_PATH = next(
    (
        p for ext in ("png", "jpg", "jpeg", "svg")
        if (p := Path(__file__).parent / "assets" / f"logo_srr.{ext}").exists()
    ),
    None,
)

st.set_page_config(
    page_title="SRR Kalibata Report Maker",
    page_icon=str(LOGO_PATH) if LOGO_PATH else "📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# API key diisi sekali dan dipakai di semua halaman: tanpa ini Streamlit
# menghapus nilai widget saat pengguna berpindah halaman.
if "api_key" in st.session_state:
    st.session_state["api_key"] = st.session_state["api_key"]

halaman = st.navigation(
    [
        st.Page("halaman/otomatis.py", title="Laporan otomatis", default=True),
        st.Page("halaman/placeholder.py", title="Mode placeholder"),
    ],
    position="top",
)

if LOGO_PATH:
    col_logo, col_judul = st.columns([1, 8], vertical_alignment="center")
    col_logo.image(str(LOGO_PATH), width="stretch")
else:
    col_judul = st.container()
col_judul.title("SRR Kalibata Report Maker")
col_judul.caption("Susun laporan penilaian dari lembar kerja dan dokumen pendukung.")

halaman.run()
