import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import zipfile
import re
from itertools import zip_longest

st.set_page_config(page_title="Template Word Otomatis", layout="wide")
st.markdown("# SRR Kalibata Report Maker")

def replace_placeholders_in_paragraph(para, data):
    for key, value in data.items():
        placeholder = f'{{{{{key}}}}}'
        replaced_in_run = False
        for run in para.runs:
            if placeholder in run.text:
                is_bold = run.bold
                is_italic = run.italic
                is_underline = run.underline
                font_name = run.font.name
                font_size = run.font.size
                font_color = run.font.color.rgb
                run.text = run.text.replace(placeholder, str(value))
                run.bold = is_bold
                run.italic = is_italic
                run.underline = is_underline
                run.font.name = font_name
                run.font.size = font_size
                run.font.color.rgb = font_color
                replaced_in_run = True
        if not replaced_in_run and placeholder in para.text:
            new_text = para.text.replace(placeholder, str(value))
            for run in para.runs:
                run.text = ""
            if para.runs:
                para.runs[0].text = new_text
            else:
                para.add_run(new_text)
    return para

def replace_placeholders(doc, data):
    for para in doc.paragraphs:
        replace_placeholders_in_paragraph(para, data)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_placeholders_in_paragraph(para, data)
    return doc

def process_files(word_file, data_dict):
    doc = Document(word_file)
    updated_doc = replace_placeholders(doc, data_dict)
    output = BytesIO()
    updated_doc.save(output)
    output.seek(0)
    return output, updated_doc

def extract_text_from_docx(docx_obj):
    return "\n".join([para.text for para in docx_obj.paragraphs])

# ✅ Fungsi terbaru dengan regex fleksibel
def check_value_consistency_table(text):
    # Regex untuk angka nominal yang diikuti tanda kurung buka
    number_pattern = r'Rp\s*([\d\.]+,\d{2})\s*\('
    
    # Regex untuk penyebutan nominal dalam huruf kapital, berakhiran "RUPIAH"
    verbal_block_pattern = r'\(\s*([A-Z\s]+RUPIAH)\s*\)'

    # Ambil daftar angka dan konversi menjadi integer
    numbers = re.findall(number_pattern, text.replace(".", "").replace("Rp", "Rp"))
    number_ints = [int(float(n.replace(",", "."))) for n in numbers]

    # Ambil daftar penyebutan verbal (huruf kapital)
    verbal_blocks = re.findall(verbal_block_pattern, text.upper())
    normalized_verbal = [vb.strip() for vb in verbal_blocks]

    # Gabungkan angka dan kata untuk ditampilkan di tabel
    rows = []
    for angka, kata in zip_longest(number_ints, normalized_verbal, fillvalue="(tidak ditemukan)"):
        rows.append({
            "Angka (Rp)": f"Rp {angka:,}".replace(",", ".") if isinstance(angka, int) else angka,
            "Penyebutan dalam Kata": kata,
            "Keterangan": "✅ Sesuai" if isinstance(angka, int) and kata != "(tidak ditemukan)" else "⚠️ Tidak cocok"
        })

    return pd.DataFrame(rows)

def highlight_placeholders(text):
    return re.sub(r"(\{\{.*?\}\})", r"[PLACEHOLDER:\1]", text)

# ========== STREAMLIT UI ==========
with st.sidebar:
    st.header("📂 Upload Dokumen")
    word_files = st.file_uploader("Template Laporan (.docx)", type=["docx"], accept_multiple_files=True)
    excel_file = st.file_uploader("Data Lembar Kerja (.xlsx)", type=["xlsx"])

if excel_file:
    df_excel = pd.read_excel(excel_file, header=None, skiprows=6, sheet_name="Laporan", engine='openpyxl')
    st.subheader("📊 Tabel Data pada Lembar Kerja (edit jika perlu)")
    df_editable = st.data_editor(df_excel, num_rows="dynamic", use_container_width=True)

    data_dict = {}
    for _, row in df_editable.iterrows():
        key = row[0]
        value = row[1]
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
                    label_visibility="collapsed"
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
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
                mime="application/zip"
            )
