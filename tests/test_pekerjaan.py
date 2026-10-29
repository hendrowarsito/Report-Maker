import os
import threading
import time
from types import SimpleNamespace

import pytest

from report_maker import pekerjaan
from report_maker.ai import minta_json
from report_maker.analisis import Kandidat, Kemunculan, Temuan, temuan_dari_dict, temuan_ke_dict

TEMUAN = [Temuan(
    "Luas tanah",
    [Kandidat("LK.xlsx", "LK Tanah!C4", "312")],
    "LK.xlsx", True, "",
    [Kemunculan(2, "250", "angka", "Bab 1.docx")],
)]


@pytest.fixture(autouse=True)
def _folder_sementara(tmp_path, monkeypatch):
    monkeypatch.setattr(pekerjaan, "FOLDER", tmp_path)
    monkeypatch.setattr(pekerjaan, "_pekerjaan", {})


def _tunggu(kunci, batas=5.0):
    akhir = time.time() + batas
    while pekerjaan.status(kunci).status == "berjalan":
        assert time.time() < akhir, "pekerjaan tidak selesai"
        time.sleep(0.02)
    return pekerjaan.status(kunci)


def test_serialisasi_temuan_bolak_balik():
    assert temuan_dari_dict(temuan_ke_dict(TEMUAN, "catatan")) == (TEMUAN, "catatan")


def test_analisis_latar_belakang_progres_dan_hasil_tersimpan():
    lanjut = threading.Event()
    dipanggil = []

    def fungsi(progres):
        dipanggil.append(1)
        progres(120)
        lanjut.wait(5)
        return TEMUAN, "selesai"

    p = pekerjaan.mulai("abc", fungsi)
    time.sleep(0.05)
    assert p.status == "berjalan" and p.karakter == 120
    # Permintaan kedua untuk file yang sama tidak memulai analisis baru
    assert pekerjaan.mulai("abc", fungsi) is p
    lanjut.set()
    assert _tunggu("abc").status == "selesai"
    assert dipanggil == [1]
    assert pekerjaan.muat_hasil("abc") == (TEMUAN, "selesai")


def test_analisis_gagal_dilaporkan_dan_bisa_diulang():
    def gagal(progres):
        raise RuntimeError("API sedang sibuk")

    pekerjaan.mulai("x", gagal)
    p = _tunggu("x")
    assert (p.status, p.galat) == ("gagal", "API sedang sibuk")
    assert pekerjaan.muat_hasil("x") is None

    pekerjaan.mulai("x", lambda progres: (TEMUAN, ""))
    assert _tunggu("x").status == "selesai"


def test_hapus_dan_bersihkan_hasil_lama():
    pekerjaan.simpan_hasil("lama", TEMUAN, "")
    pekerjaan.simpan_hasil("baru", TEMUAN, "")
    tua = time.time() - pekerjaan.UMUR_MAKS_DETIK - 10
    os.utime(pekerjaan._berkas("lama"), (tua, tua))
    pekerjaan.bersihkan_lama()
    assert pekerjaan.muat_hasil("lama") is None
    assert pekerjaan.muat_hasil("baru") is not None
    pekerjaan.hapus_hasil("baru")
    assert pekerjaan.muat_hasil("baru") is None


def test_minta_json_melaporkan_progres_dan_effort():
    tercatat = {}
    respons = SimpleNamespace(
        stop_reason="end_turn", content=[SimpleNamespace(type="text", text='{"a": 1}')]
    )

    class _Stream:
        text_stream = ['{"a"', ": 1}"]
        def __enter__(s): return s
        def __exit__(s, *a): return False
        def get_final_message(s): return respons

    def stream(**kw):
        tercatat.update(kw)
        return _Stream()

    client = SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(stream=stream)))
    progres = []
    assert minta_json(client, "s", [], {}, effort="medium", progres=progres.append) == {"a": 1}
    assert progres == [4, 8]
    assert tercatat["output_config"]["effort"] == "medium"
