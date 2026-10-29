"""Konversi .docx ke PDF lewat LibreOffice headless (opsional)."""
import shutil
import subprocess
import tempfile
from pathlib import Path


def _soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def pdf_tersedia() -> bool:
    return _soffice() is not None


def docx_ke_pdf(docx_bytes: bytes, timeout: int = 120) -> bytes:
    """Konversi satu dokumen; RuntimeError bila LibreOffice tidak ada atau gagal."""
    soffice = _soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice tidak terpasang.")

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "laporan.docx"
        src.write_bytes(docx_bytes)
        # Profil LibreOffice terpisah agar aman dijalankan paralel oleh beberapa sesi
        profil = (Path(tmp) / "profil").as_uri()
        hasil = subprocess.run(
            [
                soffice, f"-env:UserInstallation={profil}",
                "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src),
            ],
            capture_output=True,
            timeout=timeout,
        )
        pdf = Path(tmp) / "laporan.pdf"
        if hasil.returncode != 0 or not pdf.exists():
            pesan = hasil.stderr.decode(errors="replace").strip() or "tanpa pesan"
            raise RuntimeError(f"Konversi PDF gagal: {pesan}")
        return pdf.read_bytes()
