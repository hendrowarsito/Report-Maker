import streamlit as st
from docx import Document
from io import BytesIO
import pandas as pd
import zipfile
import locale
import os

def format_number_indonesia(value):
    """Format number to Indonesian format (e.g., 12.000,00)."""
    try:
        locale.setlocale(locale.LC_NUMERIC, "id_ID.UTF-8")
        return locale.format_string("%.2f", value, grouping=True)
    except:
        return value

def replace_placeholders(document, replacements):
    """
    Replace placeholders in a DOCX document tanpa merusak format run.
    Placeholder harus berada dalam satu run utuh, misalnya '{placeholder}'.
    """
    # Ganti placeholder di semua paragraf (di luar tabel)
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            text = run.text
            for key, val in replacements.items():
                # Jika angka, gunakan format Indonesia
                formatted_value = format_number_indonesia(val) if isinstance(val, (int, float)) else str(val)
                placeholder = f"{{{key}}}"
                if placeholder in text:
                    text = text.replace(placeholder, formatted_value)
            run.text = text

    # Ganti placeholder di semua paragraf dalam tabel
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        text = run.text
                        for key, val in replacements.items():
                            formatted_value = format_number_indonesia(val) if isinstance(val, (int, float)) else str(val)
                            placeholder = f"{{{key}}}"
                            if placeholder in text:
                                text = text.replace(placeholder, formatted_value)
                        run.text = text

    return document

def extract_placeholders(document):
    """
    Ekstrak placeholder dari dokumen. 
    CATATAN: Fungsi ini hanya mencari '{' dan '}' pada level paragraph.text,
    sehingga placeholder yang terpecah di beberapa run mungkin tidak terdeteksi.
    Pastikan placeholder utuh dalam satu run!
    """
    placeholders = set()
    
    # Cek placeholder di paragraf (di luar tabel)
    for paragraph in document.paragraphs:
        if "{" in paragraph.text and "}" in paragraph.text:
            for part in paragraph.text.split():
                if part.startswith("{") and part.endswith("}"):
                    placeholders.add(part.strip("{}"))

    # Cek placeholder di dalam tabel
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "{" in paragraph.text and "}" in paragraph.text:
                        for part in paragraph.text.split():
                            if part.startswith("{") and part.endswith("}"):
                                placeholders.add(part.strip("{}"))

    return sorted(placeholders)

def save_docx(document):
    """Save changes to a new DOCX file."""
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer

def generate_zip(files):
    """Generate a ZIP file from a list of files."""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, file_buffer in files:
            zf.writestr(file_name, file_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.title("SRR KALIBATA REPORT MAKER")
    st.write("Upload DOCX templates and an Excel file to generate reports automatically.")

    uploaded_templates = st.file_uploader("Upload DOCX Templates", type="docx", accept_multiple_files=True)
    uploaded_excel = st.file_uploader("Upload Excel File", type="xlsx")
    
    if uploaded_templates and uploaded_excel:
        st.success(f"{len(uploaded_templates)} templates uploaded successfully!")
        data = pd.read_excel(uploaded_excel)
        st.write("Data Preview:")
        st.dataframe(data)

        templates = {}
        all_placeholders = set()
        for file in uploaded_templates:
            document = Document(file)
            placeholders = extract_placeholders(document)
            templates[file.name] = {"document": document, "placeholders": placeholders}
            all_placeholders.update(placeholders)

        unmatched_placeholders = [ph for ph in all_placeholders if ph not in data.columns]
        if unmatched_placeholders:
            st.warning(f"Unmatched placeholders: {', '.join(unmatched_placeholders)}")

        if st.button("Generate Reports"):
            st.success("Generating reports...")
            generated_files = []
            for index, row in data.iterrows():
                for template_name, template_data in templates.items():
                    # Buat salinan Document agar template asli tidak terubah
                    doc_copy = Document()
                    doc_copy.element.body.clear()
                    for element in template_data["document"].element.body:
                        doc_copy.element.body.append(element.clone())

                    # Replace placeholders di salinan
                    replace_placeholders(doc_copy, row.to_dict())

                    # Simpan salinan yang sudah diganti
                    file_name = f"{index+1}_{template_name}"
                    buffer = save_docx(doc_copy)
                    generated_files.append((file_name, buffer))

            zip_buffer = generate_zip(generated_files)
            st.download_button("Download All Reports as ZIP", data=zip_buffer, file_name="generated_reports.zip", mime="application/zip")

if __name__ == "__main__":
    main()
