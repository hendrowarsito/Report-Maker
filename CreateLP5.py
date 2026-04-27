import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import zipfile
import re
from itertools import zip_longest

st.set_page_config(page_title="Template Word Otomatis", layout="wide")
st.markdown("# SRR Kalibata Report Maker")


# ---------------------------------------------------------------------------
# Paragraph-level placeholder replacement
# ---------------------------------------------------------------------------

def _replace_cross_run_placeholder(para, placeholder: str, value: str):
    """
    Ganti placeholder yang terpecah di beberapa run berbeda.
    Formatting (bold/italic/dll) setiap run tetap dipertahankan:
    - run di awal placeholder mempertahankan formatnya sendiri
    - run di tengah yang terhapus tidak mempengaruhi run lain
    - run setelah placeholder tetap utuh dengan formatnya
    """
    runs = para.runs
    if not runs:
        return

    texts = [run.text for run in runs]
    full_text = "".join(texts)

    if placeholder not in full_text:
        return

    # Hitung posisi kumulatif setiap run
    cum_start = []
    pos = 0
    for t in texts:
        cum_start.append(pos)
        pos += len(t)

    ph_start = full_text.find(placeholder)
    ph_end = ph_start + len(placeholder)

    # Cari run yang mengandung awal dan akhir placeholder
    start_run_idx = None
    end_run_idx = None
    for i, cs in enumerate(cum_start):
        ce = cs + len(runs[i].text)
        if cs <= ph_start < ce:
            start_run_idx = i
        if cs < ph_end <= ce:
            end_run_idx = i

    # Tangani placeholder yang tepat di batas antar run
    if end_run_idx is None:
        for i, cs in enumerate(cum_start):
            if cs + len(runs[i].text) == ph_end:
                end_run_idx = i
                break

    if start_run_idx is None or end_run_idx is None or start_run_idx == end_run_idx:
        return

    # Teks sebelum placeholder di run awal
    before = texts[start_run_idx][: ph_start - cum_start[start_run_idx]]
    # Teks setelah placeholder di run akhir
    after = texts[end_run_idx][ph_end - cum_start[end_run_idx] :]

    # Terapkan: run awal mendapat teks sebelum + nilai pengganti
    runs[start_run_idx].text = before + value
    # Run akhir hanya menyimpan teks sesudah placeholder (formatting-nya tetap)
    runs[end_run_idx].text = after
    # Run di tengah dikosongkan (formatting mereka tidak relevan)
    for i in range(start_run_idx + 1, end_run_idx):
        runs[i].text = ""


def replace_placeholders_in_paragraph(para, data: dict):
    """
    Ganti semua placeholder di paragraf.

    Catatan: meng-assign run.text di python-docx HANYA mengubah konten teks
    (<w:t>) tanpa menyentuh properti run (<w:rPr>), sehingga bold/italic/
    underline/font/warna tetap terjaga secara otomatis — tidak perlu
    disimpan dan dikembalikan secara manual (yang malah bisa memicu crash
    jika run tidak punya warna eksplisit).
    """
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"

        # Tahap 1: ganti dalam satu run (kasus normal — formatting terjaga otomatis)
        for run in para.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))

        # Tahap 2: tangani placeholder yang terpecah di beberapa run
        if placeholder in "".join(r.text for r in para.runs):
            _replace_cross_run_placeholder(para, placeholder, str(value))

    return para


# ---------------------------------------------------------------------------
# Document-level replacement (body + tabel + header + footer)
# ---------------------------------------------------------------------------

def _process_paragraph_collection(paragraphs, data: dict):
    for para in paragraphs:
        replace_placeholders_in_paragraph(para, data)


def _process_table_collection(tables, data: dict):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                _process_paragraph_collection(cell.paragraphs, data)
                _process_table_collection(cell.tables, data)  # tabel tersarang


def replace_placeholders(doc: Document, data: dict) -> Document:
    # Body
    _process_paragraph_collection(doc.paragraphs, data)
    _process_table_collection(doc.tables, data)

    # Header dan footer setiap section
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            if hf is not None:
                _process_paragraph_collection(hf.paragraphs, data)
                _process_table_collection(hf.tables, data)

    return doc


# ---------------------------------------------------------------------------
# File processing helpers
# ---------------------------------------------------------------------------

def process_files(word_file, data_dict: dict):
    doc = Document(word_file)
    updated_doc = replace_placeholders(doc, data_dict)
    output = BytesIO()
    updated_doc.save(output)
    output.seek(0)
    return output, updated_doc


def extract_text_from_docx(docx_obj: Document) -> str:
    return "\n".join(para.text for para in docx_obj.paragraphs)


# ---------------------------------------------------------------------------
# Konsistensi nominal angka vs terbilang
# ---------------------------------------------------------------------------

def check_value_consistency_table(text: str) -> pd.DataFrame:
    number_pattern = r"Rp\s*([\d\.]+,\d{2})\s*\("
    verbal_block_pattern = r"\(\s*([A-Z\s]+RUPIAH)\s*\)"

    numbers = re.findall(number_pattern, text.replace(".", "").replace("Rp", "Rp"))
    number_ints = [int(float(n.replace(",", "."))) for n in numbers]

    verbal_blocks = re.findall(verbal_block_pattern, text.upper())
    normalized_verbal = [vb.strip() for vb in verbal_blocks]

    rows = []
    for angka, kata in zip_longest(number_ints, normalized_verbal, fillvalue="(tidak ditemukan)"):
        rows.append(
            {
                "Angka (Rp)": f"Rp {angka:,}".replace(",", ".")
                if isinstance(angka, int)
                else angka,
                "Penyebutan dalam Kata": kata,
                "Keterangan": "✅ Sesuai"
                if isinstance(angka, int) and kata != "(tidak ditemukan)"
                else "⚠️ Tidak cocok",
            }
        )

    return pd.DataFrame(rows)


def highlight_placeholders(text: str) -> str:
    return re.sub(r"(\{\{.*?\}\})", r"[PLACEHOLDER:\1]", text)


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("📂 Upload Dokumen")
    word_files = st.file_uploader(
        "Template Laporan (.docx)", type=["docx"], accept_multiple_files=True
    )
    excel_file = st.file_uploader("Data Lembar Kerja (.xlsx)", type=["xlsx"])

if excel_file:
    df_excel = pd.read_excel(
        excel_file, header=None, skiprows=6, sheet_name="Laporan", engine="openpyxl"
    )
    st.subheader("📊 Tabel Data pada Lembar Kerja (edit jika perlu)")
    df_editable = st.data_editor(df_excel, num_rows="dynamic", use_container_width=True)

    data_dict: dict = {}
    for _, row in df_editable.iterrows():
        key = row.iloc[0]
        value = row.iloc[1]
        if pd.notnull(key):
            data_dict[str(key)] = str(value) if pd.notnull(value) else ""

    if word_files:
        st.subheader("📄 Preview Laporan, cek sebelum dibuat filenya")
        processed_files = []

        for word_file in word_files:
            output_file, updated_doc = process_files(word_file, data_dict)
            raw_text = extract_text_from_docx(updated_doc)
            preview_text = highlight_placeholders(raw_text)
            consistency_df = check_value_consistency_table(preview_text)

            with st.expander(f"📁 {word_file.name}", expanded=True):
                st.text_area(
                    "Isi Dokumen Terisi (Placeholder ditandai):",
                    value=preview_text,
                    height=400,
                    key=word_file.name + "_preview",
                    disabled=True,
                    label_visibility="collapsed",
                )

                if not consistency_df.empty:
                    st.markdown("🔍 **Cek Konsistensi Nominal Angka vs Kata**")
                    st.dataframe(consistency_df, use_container_width=True)
                    if "⚠️ Tidak cocok" in consistency_df["Keterangan"].values:
                        st.warning("Terdapat ketidaksesuaian antara angka dan penyebutannya.")
                    else:
                        st.success("Semua penyebutan nominal sesuai dengan angkanya.")

            processed_files.append((word_file.name, output_file))

        st.subheader("⬇️ Unduh Hasil")
        if len(processed_files) == 1:
            file_name, file_obj = processed_files[0]
            st.download_button(
                label="Unduh Dokumen",
                data=file_obj,
                file_name=file_name.replace(".docx", "_terisi.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for name, file_obj in processed_files:
                    filename = name.replace(".docx", "_terisi.docx")
                    zipf.writestr(filename, file_obj.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                label="Unduh Semua Dokumen (.zip)",
                data=zip_buffer,
                file_name="Semua_Dokumen_Terisi.zip",
                mime="application/zip",
            )
