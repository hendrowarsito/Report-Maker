import json
import re
from io import BytesIO
from itertools import zip_longest
import zipfile

import anthropic
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="SRR Kalibata Report Maker", layout="wide")
st.markdown("# SRR Kalibata Report Maker")


# =============================================================================
# Paragraph-level placeholder replacement
# =============================================================================

def _replace_cross_run_placeholder(para, placeholder: str, value: str):
    """
    Ganti placeholder yang terpecah di beberapa run berbeda.
    Formatting (bold/italic/dll) setiap run tetap dipertahankan.
    """
    runs = para.runs
    if not runs:
        return

    texts = [run.text for run in runs]
    full_text = "".join(texts)

    if placeholder not in full_text:
        return

    cum_start = []
    pos = 0
    for t in texts:
        cum_start.append(pos)
        pos += len(t)

    ph_start = full_text.find(placeholder)
    ph_end = ph_start + len(placeholder)

    start_run_idx = None
    end_run_idx = None
    for i, cs in enumerate(cum_start):
        ce = cs + len(runs[i].text)
        if cs <= ph_start < ce:
            start_run_idx = i
        if cs < ph_end <= ce:
            end_run_idx = i

    if end_run_idx is None:
        for i, cs in enumerate(cum_start):
            if cs + len(runs[i].text) == ph_end:
                end_run_idx = i
                break

    if start_run_idx is None or end_run_idx is None or start_run_idx == end_run_idx:
        return

    before = texts[start_run_idx][: ph_start - cum_start[start_run_idx]]
    after = texts[end_run_idx][ph_end - cum_start[end_run_idx] :]

    runs[start_run_idx].text = before + value
    runs[end_run_idx].text = after
    for i in range(start_run_idx + 1, end_run_idx):
        runs[i].text = ""


def replace_placeholders_in_paragraph(para, data: dict):
    """
    Ganti semua placeholder di paragraf.
    run.text hanya mengubah <w:t>, formatting <w:rPr> terjaga otomatis.
    """
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"

        for run in para.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))

        if placeholder in "".join(r.text for r in para.runs):
            _replace_cross_run_placeholder(para, placeholder, str(value))

    return para


# =============================================================================
# Document-level replacement (body + tabel tersarang + header + footer)
# =============================================================================

def _process_paragraph_collection(paragraphs, data: dict):
    for para in paragraphs:
        replace_placeholders_in_paragraph(para, data)


def _process_table_collection(tables, data: dict):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                _process_paragraph_collection(cell.paragraphs, data)
                _process_table_collection(cell.tables, data)


def replace_placeholders(doc: Document, data: dict) -> Document:
    _process_paragraph_collection(doc.paragraphs, data)
    _process_table_collection(doc.tables, data)

    for section in doc.sections:
        for hf in (
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer,
        ):
            if hf is not None:
                _process_paragraph_collection(hf.paragraphs, data)
                _process_table_collection(hf.tables, data)

    return doc


# =============================================================================
# Find & Replace langsung (tanpa placeholder, tanpa API)
# =============================================================================

def apply_find_replace(doc: Document, pairs: list[tuple[str, str]]) -> int:
    """
    Ganti kata secara langsung di seluruh dokumen (body, tabel, header, footer).
    Memanfaatkan logika yang sama dengan replace_placeholders sehingga
    formatting run tetap terjaga.
    Kembalikan jumlah total penggantian yang dilakukan.
    """
    total = 0
    substitutions = {f: t for f, t in pairs if f.strip()}

    def _replace_in_para(para):
        nonlocal total
        for find, replace in substitutions.items():
            for run in para.runs:
                if find in run.text:
                    count = run.text.count(find)
                    run.text = run.text.replace(find, replace)
                    total += count
            # Tangani kata yang terpecah antar run
            if find in "".join(r.text for r in para.runs):
                before_count = "".join(r.text for r in para.runs).count(find)
                _replace_cross_run_placeholder(para, find, replace)
                total += before_count

    def _replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        _replace_in_para(para)
                    _replace_in_tables(cell.tables)

    for para in doc.paragraphs:
        _replace_in_para(para)
    _replace_in_tables(doc.tables)

    for section in doc.sections:
        for hf in (
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer,
        ):
            if hf is not None:
                for para in hf.paragraphs:
                    _replace_in_para(para)
                _replace_in_tables(hf.tables)

    return total


def smart_replace_with_claude(
    doc_text: str, find: str, replace: str, client: anthropic.Anthropic
) -> str:
    """
    Minta Claude menentukan apakah setiap kemunculan kata perlu diganti
    berdasarkan konteks kalimatnya (penggantian sensitif-konteks).
    Kembalikan penjelasan keputusan Claude sebagai teks markdown.
    Penggantian aktual tetap dilakukan oleh apply_find_replace —
    fungsi ini hanya memberi analisis konteks untuk dikonfirmasi user.
    """
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        messages=[
            {
                "role": "user",
                "content": (
                    f'Dalam dokumen laporan penilaian aset/properti Indonesia berikut, '
                    f'temukan setiap kemunculan kata **"{find}"** dan analisis:\n'
                    f'- Apakah dalam konteks ini penggantian dengan **"{replace}"** tepat?\n'
                    f'- Jika ada kemunculan yang TIDAK boleh diganti, jelaskan alasannya.\n\n'
                    f'Dokumen (penggal):\n---\n{doc_text[:8000]}\n---\n\n'
                    f'Jawab dalam markdown. Mulai dengan rekomendasi: '
                    f'✅ Ganti semua / ⚠️ Ganti sebagian / ❌ Jangan ganti'
                ),
            }
        ],
    ) as stream:
        return stream.get_final_text()


# =============================================================================
# File processing helpers
# =============================================================================

def process_files(word_file, data_dict: dict, find_replace_pairs: list[tuple[str, str]] | None = None):
    doc = Document(word_file)
    updated_doc = replace_placeholders(doc, data_dict)
    if find_replace_pairs:
        apply_find_replace(updated_doc, find_replace_pairs)
    output = BytesIO()
    updated_doc.save(output)
    output.seek(0)
    return output, updated_doc


def extract_text_from_docx(docx_obj: Document) -> str:
    """
    Ekstrak semua teks: paragraf tubuh dokumen + isi sel tabel.
    Versi lama hanya mengambil doc.paragraphs sehingga tabel tidak ikut dipreview.
    """
    parts: list[str] = []

    for para in docx_obj.paragraphs:
        parts.append(para.text)

    def _collect_table(table):
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                cell_text = " | ".join(
                    para.text for para in cell.paragraphs if para.text.strip()
                )
                if cell_text:
                    row_cells.append(cell_text)
                for nested in cell.tables:
                    _collect_table(nested)
            if row_cells:
                parts.append("  ".join(row_cells))

    for table in docx_obj.tables:
        _collect_table(table)

    return "\n".join(parts)


def highlight_placeholders(text: str) -> str:
    return re.sub(r"(\{\{.*?\}\})", r"[PLACEHOLDER:\1]", text)


def scan_remaining_placeholders(text: str) -> list[str]:
    """Kembalikan daftar placeholder {{...}} yang belum terganti."""
    return sorted(set(re.findall(r"\{\{[^}]+\}\}", text)))


# =============================================================================
# Konsistensi nominal angka vs terbilang (regex — tidak butuh API)
# =============================================================================

def check_value_consistency_table(text: str) -> tuple[pd.DataFrame, list[int], list[str]]:
    """
    Kembalikan (dataframe, list angka, list kata) untuk keperluan validasi AI.

    Catatan bug yang diperbaiki vs versi lama:
    - Pasangan angka-kata sebelumnya hanya cek keberadaan, bukan kecocokan nilai.
      Sekarang disimpan terpisah agar bisa divalidasi ke Claude.
    """
    number_pattern = r"Rp\s*([\d\.]+,\d{2})\s*\("
    verbal_block_pattern = r"\(\s*([A-Z\s]+RUPIAH)\s*\)"

    raw_numbers = re.findall(number_pattern, text)
    number_ints: list[int] = []
    for n in raw_numbers:
        cleaned = n.replace(".", "").replace(",", ".")
        try:
            number_ints.append(int(float(cleaned)))
        except ValueError:
            pass

    verbal_blocks = re.findall(verbal_block_pattern, text.upper())
    normalized_verbal = [vb.strip() for vb in verbal_blocks]

    rows = []
    for angka, kata in zip_longest(number_ints, normalized_verbal, fillvalue="(tidak ditemukan)"):
        rows.append(
            {
                "Angka (Rp)": (
                    f"Rp {angka:,}".replace(",", ".")
                    if isinstance(angka, int)
                    else angka
                ),
                "Penyebutan dalam Kata": kata,
                "Keterangan (Regex)": (
                    "⚠️ Tidak cocok"
                    if angka == "(tidak ditemukan)" or kata == "(tidak ditemukan)"
                    else "❓ Belum divalidasi AI"
                ),
            }
        )

    return pd.DataFrame(rows), number_ints, normalized_verbal


# =============================================================================
# Claude API — tiga fitur utama
# =============================================================================

def _get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def validate_terbilang_ai(
    angka: int, terbilang: str, client: anthropic.Anthropic
) -> dict:
    """
    Validasi satu pasang angka-terbilang via Claude Haiku (cepat & hemat).
    Return: {"valid": bool | None, "koreksi": str | None}
    """
    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=120,
            messages=[
                {
                    "role": "user",
                    "content": (
                        f"Validasi: apakah terbilang ini sesuai angkanya?\n"
                        f"Angka: Rp {angka:,}\n"
                        f"Terbilang: {terbilang}\n\n"
                        "Jawab hanya JSON tanpa markdown:\n"
                        '{"valid": true, "koreksi": null}\n'
                        "atau jika salah:\n"
                        '{"valid": false, "koreksi": "TERBILANG YANG BENAR RUPIAH"}'
                    ),
                }
            ],
        )
        text = response.content[0].text.strip()
        text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.MULTILINE).strip()
        return json.loads(text)
    except Exception as exc:
        return {"valid": None, "koreksi": None, "error": str(exc)}


def review_document_ai(doc_text: str, client: anthropic.Anthropic) -> str:
    """
    Review QA lengkap dokumen yang sudah terisi via Claude Sonnet (streaming).
    Cek: placeholder belum terisi, inkonsistensi nilai/tanggal, bagian kosong.
    """
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        messages=[
            {
                "role": "user",
                "content": (
                    "Anda adalah auditor dokumen laporan penilaian aset/properti profesional Indonesia.\n\n"
                    "Periksa dokumen berikut:\n"
                    "1. **Placeholder belum terisi** — pola {{...}} atau [PLACEHOLDER:{{...}}]\n"
                    "2. **Inkonsistensi nilai** — nominal angka yang sama berbeda di bagian lain\n"
                    "3. **Inkonsistensi tanggal** — tanggal tidak konsisten atau tidak logis\n"
                    "4. **Bagian kosong** — paragraf tidak lengkap\n"
                    "5. **Anomali data** — nilai tidak wajar, nama tidak konsisten\n\n"
                    f"Dokumen:\n---\n{doc_text[:10000]}\n---\n\n"
                    "Laporan audit dalam markdown terstruktur. "
                    "Mulai dengan baris status: ✅ LULUS / ⚠️ PERLU REVIEW / ❌ KRITIS"
                ),
            }
        ],
    ) as stream:
        return stream.get_final_text()


def analyze_template_ai(template_text: str, client: anthropic.Anthropic) -> dict:
    """
    Analisis template Word: deteksi semua placeholder dan artinya.
    Return dict: {"placeholders": {"KEY": "deskripsi"}, "wajib": [...], "opsional": [...]}
    """
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        messages=[
            {
                "role": "user",
                "content": (
                    "Analisis template laporan penilaian ini. "
                    "Temukan semua placeholder {{NAMA_PLACEHOLDER}} dan jelaskan:\n"
                    "- Artinya dalam konteks laporan penilaian properti/aset Indonesia\n"
                    "- Apakah wajib atau opsional\n"
                    "- Format data yang diharapkan\n\n"
                    f"Template:\n---\n{template_text[:8000]}\n---\n\n"
                    "Jawab hanya JSON tanpa markdown:\n"
                    '{"placeholders": {"KEY": {"arti": "...", "format": "...", "wajib": true}}, '
                    '"ringkasan": "..."}'
                ),
            }
        ],
    )
    text = response.content[0].text.strip()
    text = re.sub(r"^```json\s*|\s*```$", "", text, flags=re.MULTILINE).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"placeholders": {}, "ringkasan": text}


# =============================================================================
# Streamlit UI
# =============================================================================

with st.sidebar:
    st.header("📂 Upload Dokumen")
    word_files = st.file_uploader(
        "Template Laporan (.docx)", type=["docx"], accept_multiple_files=True
    )
    excel_file = st.file_uploader("Data Lembar Kerja (.xlsx)", type=["xlsx"])

    st.divider()
    st.header("🤖 Claude AI")
    api_key = st.text_input(
        "API Key Anthropic",
        type="password",
        placeholder="sk-ant-...",
        help="Dapatkan API key di console.anthropic.com",
    )
    claude_ready = bool(api_key and api_key.startswith("sk-ant-"))
    if api_key and not claude_ready:
        st.warning("Format API key tidak valid.")
    elif claude_ready:
        st.success("API key siap.")

    st.divider()
    st.header("⚙️ Konfigurasi Excel")
    sheet_name = st.text_input("Nama Sheet", value="Laporan")
    skiprows = st.number_input("Lewati baris pertama (skiprows)", min_value=0, value=6, step=1)

    st.divider()
    st.header("🔄 Find & Replace")
    st.caption("Ganti kata langsung di dokumen — tidak perlu placeholder.")

    if "fr_pairs" not in st.session_state:
        st.session_state.fr_pairs = [("", "")]

    fr_pairs_ui: list[tuple[str, str]] = []
    for i, (f_val, r_val) in enumerate(st.session_state.fr_pairs):
        c1, c2 = st.columns(2)
        find_val = c1.text_input("Cari", value=f_val, key=f"fr_find_{i}", label_visibility="collapsed", placeholder="Cari…")
        repl_val = c2.text_input("Ganti", value=r_val, key=f"fr_repl_{i}", label_visibility="collapsed", placeholder="Ganti dengan…")
        fr_pairs_ui.append((find_val, repl_val))

    if st.button("＋ Tambah baris"):
        st.session_state.fr_pairs = fr_pairs_ui + [("", "")]
        st.rerun()
    else:
        st.session_state.fr_pairs = fr_pairs_ui

    active_pairs = [(f, r) for f, r in fr_pairs_ui if f.strip()]


# ---------------------------------------------------------------------------
# Baca Excel & bangun data_dict
# ---------------------------------------------------------------------------

if excel_file:
    try:
        df_excel = pd.read_excel(
            excel_file,
            header=None,
            skiprows=int(skiprows),
            sheet_name=sheet_name,
            engine="openpyxl",
        )
    except Exception as e:
        st.error(f"Gagal membaca Excel: {e}")
        st.stop()

    st.subheader("📊 Tabel Data pada Lembar Kerja (edit jika perlu)")
    df_editable = st.data_editor(df_excel, num_rows="dynamic", use_container_width=True)

    data_dict: dict = {}
    for _, row in df_editable.iterrows():
        key = row.iloc[0]
        value = row.iloc[1]
        if pd.notnull(key):
            data_dict[str(key)] = str(value) if pd.notnull(value) else ""

    # ---------------------------------------------------------------------------
    # Fitur AI: Analisis Struktur Template
    # ---------------------------------------------------------------------------
    if word_files and claude_ready:
        with st.expander("🔍 Analisis Placeholder Template dengan AI", expanded=False):
            if st.button("Jalankan Analisis Template"):
                with st.spinner("Claude sedang membaca template..."):
                    client = _get_client(api_key)
                    all_texts = []
                    for wf in word_files:
                        doc_tmp = Document(wf)
                        all_texts.append(extract_text_from_docx(doc_tmp))
                        wf.seek(0)
                    combined = "\n\n".join(all_texts)
                    result = analyze_template_ai(combined, client)

                if result.get("placeholders"):
                    rows_analysis = []
                    for ph_key, info in result["placeholders"].items():
                        rows_analysis.append(
                            {
                                "Placeholder": f"{{{{{ph_key}}}}}",
                                "Arti": info.get("arti", ""),
                                "Format": info.get("format", ""),
                                "Wajib": "✅ Ya" if info.get("wajib") else "➖ Opsional",
                                "Terisi": "✅" if ph_key in data_dict and data_dict[ph_key] else "❌",
                            }
                        )
                    st.dataframe(pd.DataFrame(rows_analysis), use_container_width=True)
                    if result.get("ringkasan"):
                        st.info(result["ringkasan"])
                else:
                    st.write(result.get("ringkasan", "Tidak ada placeholder ditemukan."))

    # ---------------------------------------------------------------------------
    # Preview & proses setiap file Word
    # ---------------------------------------------------------------------------
    if word_files:
        st.subheader("📄 Preview Laporan — cek sebelum dibuat filenya")
        processed_files = []

        for word_file in word_files:
            output_file, updated_doc = process_files(word_file, data_dict, active_pairs or None)
            raw_text = extract_text_from_docx(updated_doc)
            preview_text = highlight_placeholders(raw_text)
            consistency_df, number_ints, normalized_verbal = check_value_consistency_table(
                raw_text
            )
            remaining = scan_remaining_placeholders(raw_text)

            with st.expander(f"📁 {word_file.name}", expanded=True):

                # -- Peringatan placeholder belum terisi (tanpa AI) --
                if remaining:
                    st.error(
                        f"**{len(remaining)} placeholder belum terganti:** "
                        + ", ".join(f"`{p}`" for p in remaining)
                    )

                # -- Preview teks --
                st.text_area(
                    "Isi Dokumen Terisi (Placeholder ditandai):",
                    value=preview_text,
                    height=400,
                    key=word_file.name + "_preview",
                    disabled=True,
                    label_visibility="collapsed",
                )

                # -- Tabel konsistensi nominal --
                if not consistency_df.empty:
                    st.markdown("🔍 **Cek Konsistensi Nominal Angka vs Kata**")

                    # Tombol validasi AI (hanya tampil jika API key tersedia)
                    if claude_ready and number_ints:
                        if st.button(
                            "✅ Validasi Terbilang dengan AI",
                            key=word_file.name + "_val_btn",
                        ):
                            client = _get_client(api_key)
                            ai_results = []
                            progress = st.progress(0, text="Memvalidasi dengan Claude...")
                            for i, (angka, kata) in enumerate(
                                zip_longest(number_ints, normalized_verbal, fillvalue="(tidak ditemukan)")
                            ):
                                if isinstance(angka, int) and kata != "(tidak ditemukan)":
                                    res = validate_terbilang_ai(angka, kata, client)
                                    if res.get("valid") is True:
                                        ai_results.append("✅ Benar")
                                    elif res.get("valid") is False:
                                        koreksi = res.get("koreksi", "")
                                        ai_results.append(f"❌ Salah — Seharusnya: {koreksi}")
                                    else:
                                        ai_results.append("⚠️ Tidak dapat divalidasi")
                                else:
                                    ai_results.append("⚠️ Tidak cocok")
                                progress.progress(
                                    (i + 1) / max(len(number_ints), len(normalized_verbal), 1)
                                )
                            progress.empty()

                            consistency_df["Validasi AI"] = ai_results
                            st.session_state[word_file.name + "_consistency"] = consistency_df

                    # Tampilkan tabel (dari session state jika sudah divalidasi AI)
                    display_df = st.session_state.get(
                        word_file.name + "_consistency", consistency_df
                    )
                    st.dataframe(display_df, use_container_width=True)

                    has_error = "⚠️" in display_df.get("Keterangan (Regex)", pd.Series()).values or (
                        "Validasi AI" in display_df.columns
                        and display_df["Validasi AI"].str.contains("❌|⚠️").any()
                    )
                    if has_error:
                        st.warning("Terdapat ketidaksesuaian nominal — periksa sebelum diunduh.")
                    elif "Validasi AI" in display_df.columns:
                        st.success("Semua nominal tervalidasi AI — sesuai.")

                # -- Review dokumen lengkap dengan AI --
                if claude_ready:
                    st.markdown("---")
                    if st.button(
                        "🤖 Review Lengkap Dokumen dengan AI",
                        key=word_file.name + "_review_btn",
                    ):
                        client = _get_client(api_key)
                        with st.spinner("Claude sedang mereview dokumen..."):
                            review_result = review_document_ai(raw_text, client)
                        st.session_state[word_file.name + "_review"] = review_result

                    if word_file.name + "_review" in st.session_state:
                        st.markdown("**Hasil Review AI:**")
                        st.markdown(st.session_state[word_file.name + "_review"])

                # -- Analisis konteks Find & Replace dengan Claude --
                if claude_ready and active_pairs:
                    st.markdown("---")
                    st.markdown("**🔄 Analisis Konteks Find & Replace**")
                    for find_word, replace_word in active_pairs:
                        btn_key = f"{word_file.name}_fr_{find_word}"
                        if st.button(
                            f'Analisis: "{find_word}" → "{replace_word}"',
                            key=btn_key,
                        ):
                            client = _get_client(api_key)
                            with st.spinner(f'Claude menganalisis konteks "{find_word}"...'):
                                analysis = smart_replace_with_claude(
                                    raw_text, find_word, replace_word, client
                                )
                            st.session_state[btn_key + "_result"] = analysis

                        if btn_key + "_result" in st.session_state:
                            st.markdown(st.session_state[btn_key + "_result"])

            processed_files.append((word_file.name, output_file))

        # ---------------------------------------------------------------------------
        # Unduh hasil
        # ---------------------------------------------------------------------------
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
