from io import BytesIO

from docx import Document

from report_maker.docx_engine import (
    isi_dokumen,
    nama_file_hasil,
    placeholders_dari_bytes,
)


def _docx(*paragraf_runs, header=None) -> bytes:
    """Buat .docx; tiap paragraf = list run, agar placeholder bisa dipecah."""
    doc = Document()
    for runs in paragraf_runs:
        p = doc.add_paragraph()
        for teks in runs:
            p.add_run(teks).bold = teks.startswith("**")
    if header:
        doc.sections[0].header.paragraphs[0].text = header
    tabel = doc.add_table(rows=1, cols=1)
    tabel.cell(0, 0).text = "Sel {{nama}}"
    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def _teks(docx_bytes: bytes) -> list[str]:
    return [p.text for p in Document(BytesIO(docx_bytes)).paragraphs]


def test_placeholder_terpecah_beberapa_run_dan_header_tabel():
    src = _docx(["Tanggal {{tgl_", "inspeksi}} oleh {{nama}}."], header="Kop {{nama}}")
    assert placeholders_dari_bytes(src) == {"tgl_inspeksi", "nama"}

    hasil, teks, sisa = isi_dokumen(src, {"tgl_inspeksi": "28 April 2026", "nama": "PT ABC"})
    assert _teks(hasil)[0] == "Tanggal 28 April 2026 oleh PT ABC."
    doc = Document(BytesIO(hasil))
    assert doc.sections[0].header.paragraphs[0].text == "Kop PT ABC"
    assert doc.tables[0].cell(0, 0).text == "Sel PT ABC"
    assert "Sel PT ABC" in teks
    assert sisa == []


def test_placeholder_tanpa_data_tetap_dan_dilaporkan():
    src = _docx(["{{ada}} dan {{tidak_ada}}"])
    hasil, _, sisa = isi_dokumen(src, {"ada": "1"})
    assert _teks(hasil)[0] == "1 dan {{tidak_ada}}"
    assert sisa == ["nama", "tidak_ada"]   # {{nama}} di tabel juga belum diisi


def test_find_replace_setelah_placeholder():
    src = _docx(["Pemberi tugas: {{nama}}"])
    hasil, _, _ = isi_dokumen(src, {"nama": "PT ABC"}, (("Pemberi tugas", "Klien"),))
    assert _teks(hasil)[0] == "Klien: PT ABC"


def test_format_run_tetap_terjaga():
    src = _docx(["**{{nama}}"])
    hasil, _, _ = isi_dokumen(src, {"nama": "PT ABC"})
    run = Document(BytesIO(hasil)).paragraphs[0].runs[0]
    assert run.text == "**PT ABC" and run.bold


def test_nama_file_hasil():
    data = {"nama": "PT ABC", "tgl": "28/04/2026"}
    assert nama_file_hasil("{{nama_file}}_terisi", "LP Final.docx", data) == "LP Final_terisi"
    assert nama_file_hasil("LP_{{nama}}_{{tgl}}", "a.docx", data) == "LP_PT ABC_28-04-2026"
    assert nama_file_hasil("LP_{{tidak_ada}}", "a.docx", data) == "LP"
    assert nama_file_hasil("{{tidak_ada}}", "a.docx", data) == "a"


def test_bandingkan_format_hanya_teks_yang_boleh_berubah():
    from report_maker.docx_engine import bandingkan_format

    src = _docx(["Nilai {{nilai}}", " oleh {{nama}}"], header="Kop {{nama}}")
    hasil, _, _ = isi_dokumen(src, {"nilai": "595.400.000", "nama": "PT ABC"})
    assert bandingkan_format(src, hasil) == []

    # Perubahan format sekecil apa pun terdeteksi
    doc = Document(BytesIO(hasil))
    doc.paragraphs[0].runs[0].italic = True
    out = BytesIO()
    doc.save(out)
    assert bandingkan_format(src, out.getvalue()) == ["word/document.xml"]


def test_header_link_to_previous_tidak_dibuat_ulang_dan_kotak_teks_terisi():
    from docx.enum.section import WD_SECTION
    from docx.oxml import parse_xml

    from report_maker.docx_engine import bandingkan_format

    doc = Document()
    doc.sections[0].header.paragraphs[0].text = "Kop {{nama}}"
    doc.add_paragraph("Bab 1 {{nama}}")
    doc.add_section(WD_SECTION.NEW_PAGE)   # section 2: header "Link to Previous"
    p = doc.add_paragraph()
    # Kotak teks (VML) berisi placeholder
    p._p.append(parse_xml(
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape><v:textbox>'
        '<w:txbxContent><w:p><w:r><w:t>Kotak {{nama}}</w:t></w:r></w:p></w:txbxContent>'
        '</v:textbox></v:shape></w:pict></w:r>'
    ))
    out = BytesIO()
    doc.save(out)
    src = out.getvalue()

    hasil, teks, sisa = isi_dokumen(src, {"nama": "PT ABC"})
    assert sisa == []
    assert bandingkan_format(src, hasil) == []          # tidak ada header/footer baru
    hasil_doc = Document(BytesIO(hasil))
    assert hasil_doc.sections[1].header.is_linked_to_previous
    assert hasil_doc.sections[0].header.paragraphs[0].text == "Kop PT ABC"
    assert "Kotak PT ABC" in hasil_doc.element.xml
