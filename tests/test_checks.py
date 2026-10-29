from report_maker.checks import (
    build_placeholder_check,
    check_value_consistency_table,
    deteksi_baris_awal,
)


def test_konsistensi_nominal_dipasangkan_dan_dicek_lokal():
    teks = (
        "Nilai Rp 200.000.000,00 (DUA RATUS JUTA RUPIAH). "
        "Salah Rp 5.000.000,00 (Lima Ratus Ribu Rupiah). "
        "Tanpa terbilang Rp 7.000.000,00 dan Rp 1.500.000 (Satu Juta Lima Ratus Ribu Rupiah)."
    )
    df = check_value_consistency_table(teks)
    assert df["Status"].tolist() == [
        "✅ Sesuai",
        "❌ Seharusnya: LIMA JUTA RUPIAH",
        "✅ Sesuai",
    ]


def test_konsistensi_kosong_tetap_punya_kolom():
    df = check_value_consistency_table("tidak ada nominal")
    assert df.empty and "Status" in df.columns


def test_build_placeholder_check_status():
    df = build_placeholder_check(
        {"a": ["t.docx"], "b": ["t.docx"], "c": ["t.docx"], "a_TERBILANG": ["t.docx"],
         "x_TERBILANG": ["t.docx"]},
        data_excel={"a": "5", "b": ""},
        data_final={"a": "5,00", "b": "", "a_TERBILANG": "LIMA RUPIAH"},
        peringatan={"x_TERBILANG": "Nilai {{x}} bukan angka"},
    )
    status = dict(zip(df["Placeholder"], df["Status"]))
    assert status == {
        "{{a}}": "✅ Dari Excel",
        "{{a_TERBILANG}}": "✅ Otomatis",
        "{{b}}": "⚠️ Nilai kosong",
        "{{c}}": "❌ Tidak ada di Excel",
        "{{x_TERBILANG}}": "⚠️ Nilai {{x}} bukan angka",
    }


def test_deteksi_baris_awal_lewati_kop():
    keys = ["LAPORAN PENILAIAN", "", "Kop:", "LAPORAN", "kode_internal", "NP", "nama"]
    vals = ["", "", "KJPP SRR", "", "X-01", "200", "PT ABC"]
    # kode_internal tidak dipakai template tapi ikut karena tepat di atas data
    assert deteksi_baris_awal(keys, vals, {"NP_TERBILANG", "nama"}) == 4


def test_deteksi_baris_awal_tanpa_key_dikenal():
    assert deteksi_baris_awal(["a", "b"], ["1", "2"], {"zzz"}) == 0
