import datetime

import pandas as pd
import pytest

from report_maker.terbilang import (
    angka_ke_terbilang,
    angka_ke_terbilang_title,
    format_angka_indonesia,
    format_nilai,
    format_rupiah,
    generate_auto_terbilang,
    parse_ke_int,
)


@pytest.mark.parametrize("n, kata", [
    (0, "NOL RUPIAH"),
    (1, "SATU RUPIAH"),
    (11, "SEBELAS RUPIAH"),
    (19, "SEMBILAN BELAS RUPIAH"),
    (100, "SERATUS RUPIAH"),
    (1_000, "SERIBU RUPIAH"),
    (1_001_011, "SATU JUTA SERIBU SEBELAS RUPIAH"),
    (200_000_000, "DUA RATUS JUTA RUPIAH"),
    (2_500_000_000, "DUA MILIAR LIMA RATUS JUTA RUPIAH"),
    (1_000_000_000_000, "SATU TRILIUN RUPIAH"),
    (-5_000, "MINUS LIMA RIBU RUPIAH"),
])
def test_angka_ke_terbilang(n, kata):
    assert angka_ke_terbilang(n) == kata


def test_terbilang_title_dan_format():
    assert angka_ke_terbilang_title(200_000_000) == "Dua Ratus Juta Rupiah"
    assert format_rupiah(200_000_000) == "Rp 200.000.000,00"
    assert format_angka_indonesia(1_500) == "1.500,00"


@pytest.mark.parametrize("nilai, hasil", [
    (200_000_000, 200_000_000),
    (200_000_000.0, 200_000_000),
    ("200000000", 200_000_000),
    ("Rp 200.000.000,00", 200_000_000),
    ("rp. 1.500.000", 1_500_000),
    ("28 April 2026", None),
    ("20%", None),
    ("0", None),
    ("", None),
])
def test_parse_ke_int(nilai, hasil):
    assert parse_ke_int(nilai) == hasil


@pytest.mark.parametrize("nilai, teks", [
    (None, ""),
    (float("nan"), ""),
    (pd.NaT, ""),
    (200_000_000.0, "200.000.000"),
    (595_400_000, "595.400.000"),
    (416_000, "416.000"),
    (150, "150"),
    (2026, "2026"),
    (1234567.5, "1.234.567,5"),
    (0.25, "0,25"),
    (1234.5, "1234,5"),
    (datetime.datetime(2026, 4, 28), "28 April 2026"),
    (datetime.datetime(2026, 4, 28, 14, 30), "28 April 2026 14:30"),
    (pd.Timestamp("2026-12-01"), "1 Desember 2026"),
    (datetime.date(2026, 1, 5), "5 Januari 2026"),
    ("  PT ABC ", "PT ABC"),
])
def test_format_nilai(nilai, teks):
    assert format_nilai(nilai) == teks


def test_generate_auto_terbilang_dari_akhiran():
    data = {"NP": "200000000", "nama": "PT ABC"}
    tambahan, peringatan = generate_auto_terbilang(
        data, {"NP_TERBILANG", "NP_terbilang_title", "NP_FORMAT", "nama"}
    )
    assert tambahan == {
        "NP": "200.000.000,00",
        "NP_TERBILANG": "DUA RATUS JUTA RUPIAH",
        "NP_terbilang_title": "Dua Ratus Juta Rupiah",
        "NP_FORMAT": "Rp 200.000.000,00",
    }
    assert peringatan == {}


def test_generate_auto_terbilang_sumber_bukan_angka_dan_nilai_excel_menang():
    data = {"nama": "PT ABC", "X": "5", "X_TERBILANG": "manual"}
    tambahan, peringatan = generate_auto_terbilang(
        data, {"nama_TERBILANG", "X_TERBILANG", "tidak_ada_TERBILANG"}
    )
    assert tambahan == {}
    assert peringatan == {"nama_TERBILANG": "Nilai {{nama}} bukan angka"}


@pytest.mark.parametrize("kunci, nilai, teks", [
    ("np_nominal", 595_400_000, "595.400.000"),
    ("kode_pos", 10220, "10220"),
    ("nib", 10_220_304_050_607, "10220304050607"),
    ("nomor_shm", 15000, "15000"),
    ("tahun_bangun", 1998, "1998"),
    ("luas_tanah", 416, "416"),
])
def test_format_nilai_dengan_kunci(kunci, nilai, teks):
    assert format_nilai(nilai, kunci) == teks
