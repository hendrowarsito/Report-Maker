import datetime
import numbers
import re
from io import BytesIO
import zipfile

import anthropic
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="SRR Kalibata Report Maker", layout="wide")
st.markdown("# SRR Kalibata Report Maker")


# =============================================================================
# Konversi angka ke terbilang Rupiah (murni Python, tanpa API)
# =============================================================================

_SATUAN = [
    "", "satu", "dua", "tiga", "empat", "lima",
    "enam", "tujuh", "delapan", "sembilan",
]
_BELASAN = [
    "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas",
    "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas",
]


def _terbilang_bilangan(n: int) -> str:
    """Rekursif: ubah bilangan bulat positif ke kata Indonesia (huruf kecil)."""
    if n == 0:
        return ""

    parts: list[str] = []

    if n >= 1_000_000_000_000:
        t, n = divmod(n, 1_000_000_000_000)
        parts.append(_terbilang_bilangan(t) + " triliun")

    if n >= 1_000_000_000:
        m, n = divmod(n, 1_000_000_000)
        parts.append(_terbilang_bilangan(m) + " miliar")

    if n >= 1_000_000:
        j, n = divmod(n, 1_000_000)
        parts.append(_terbilang_bilangan(j) + " juta")

    if n >= 1_000:
        r, n = divmod(n, 1_000)
        parts.append("seribu" if r == 1 else _terbilang_bilangan(r) + " ribu")

    if n >= 100:
        h, n = divmod(n, 100)
        parts.append("seratus" if h == 1 else _SATUAN[h] + " ratus")

    if n >= 20:
        p, n = divmod(n, 10)
        parts.append(_SATUAN[p] + " puluh")
        if n:
            parts.append(_SATUAN[n])
    elif n >= 10:
        parts.append(_BELASAN[n - 10])
    elif n > 0:
        parts.append(_SATUAN[n])

    return " ".join(parts)


def angka_ke_terbilang(n: int) -> str:
    """
    Ubah bilangan bulat ke terbilang HURUF KAPITAL SEMUA.
    Contoh: 200_000_000 → 'DUA RATUS JUTA RUPIAH'
    Dipakai untuk placeholder {{KEY_TERBILANG}}.
    """
    if n == 0:
        return "NOL RUPIAH"
    prefix = "MINUS " if n < 0 else ""
    return (prefix + _terbilang_bilangan(abs(n))).upper() + " RUPIAH"


def angka_ke_terbilang_title(n: int) -> str:
    """
    Ubah bilangan bulat ke terbilang Huruf Besar Di Awal Kata (Title Case).
    Contoh: 200_000_000 → 'Dua Ratus Juta Rupiah'
    Dipakai untuk placeholder {{KEY_TERBILANG_TITLE}}.
    """
    return angka_ke_terbilang(n).title()


def format_rupiah(n: int) -> str:
    """
    Format bilangan bulat ke string Rupiah Indonesia.
    Contoh: 200_000_000 → 'Rp 200.000.000,00'
    """
    ribuan = f"{n:,}".replace(",", ".")
    return f"Rp {ribuan},00"


def _parse_ke_int(nilai) -> int | None:
    """
    Coba parsing berbagai format nilai ke bilangan bulat.
    Mengembalikan None jika:
      - bukan angka
      - mengandung huruf selain prefix "Rp" (misal: tanggal, nama, %)
      - hasil < 1

    Format yang didukung: int, float, "200000000", "Rp 200.000.000,00", "200.000.000,00"
    """
    if isinstance(nilai, (int, float)):
        result = int(round(float(nilai)))
        return result if result >= 1 else None

    s = str(nilai).strip()

    # Tolak string yang mengandung huruf di luar prefix Rp, atau simbol non-numerik
    # Contoh yang DITOLAK: "28 April 2026", "Unit A", "20%", "n/a", "1.5x"
    tanpa_rp = re.sub(r"(?i)^\s*rp\.?\s*", "", s).strip()
    if re.search(r"[a-zA-Z%]", tanpa_rp):
        return None

    # Hapus prefix Rp
    s = tanpa_rp
    # Format Indonesia: titik = pemisah ribuan, koma = desimal
    if re.search(r"\d\.\d{3}", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")

    s = re.sub(r"[^\d.]", "", s)
    if not s:
        return None
    try:
        result = int(round(float(s)))
        return result if result >= 1 else None
    except ValueError:
        return None


def format_angka_indonesia(n: int) -> str:
    """
    Format angka ke gaya Indonesia tanpa prefix Rp.
    Contoh: 200_000_000 → '200.000.000,00'
    Digunakan ketika template sudah hardcode 'Rp ' sebelum placeholder.
    """
    return f"{n:,}".replace(",", ".") + ",00"


_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def format_tanggal(d: datetime.date) -> str:
    """Contoh: date(2026, 4, 28) → '28 April 2026'"""
    return f"{d.day} {_BULAN[d.month - 1]} {d.year}"


def format_nilai(value) -> str:
    """
    Ubah nilai sel Excel ke teks siap tempel di laporan.
      • kosong / NaN          → ''
      • 200000000.0           → '200000000'   (tanpa '.0')
      • 0.25                  → '0,25'        (desimal koma)
      • datetime(2026, 4, 28) → '28 April 2026'
      • teks                  → apa adanya (spasi di tepi dibuang)
    """
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, datetime.datetime):        # termasuk pd.Timestamp
        teks = format_tanggal(value)
        if (value.hour, value.minute) != (0, 0):
            teks += f" {value:%H:%M}"
        return teks
    if isinstance(value, datetime.date):
        return format_tanggal(value)
    if isinstance(value, datetime.time):
        return f"{value:%H:%M}"
    if isinstance(value, numbers.Integral):
        return str(int(value))
    if isinstance(value, numbers.Real):
        if float(value).is_integer():
            return str(int(value))
        return f"{value:.6f}".rstrip("0").rstrip(".").replace(".", ",")
    return str(value).strip()


# Placeholder turunan dibuat otomatis dari key Excel berdasarkan akhiran namanya.
# Urutan penting: _TERBILANG_TITLE harus dicek sebelum _TERBILANG.
_SUFIKS_TURUNAN = {
    "_TERBILANG_TITLE": angka_ke_terbilang_title,
    "_TERBILANG":       angka_ke_terbilang,
    "_FORMAT":          format_rupiah,
}


def _pecah_turunan(ph: str):
    """'NP_nominal_TERBILANG' → ('NP_nominal', angka_ke_terbilang); bukan turunan → (None, None)."""
    for sufiks, fungsi in _SUFIKS_TURUNAN.items():
        if ph.upper().endswith(sufiks) and len(ph) > len(sufiks):
            return ph[: -len(sufiks)], fungsi
    return None, None


def generate_auto_terbilang(
    data_dict: dict[str, str], placeholders: set[str]
) -> tuple[dict[str, str], dict[str, str]]:
    """
    Buat nilai untuk placeholder turunan yang dipakai di template, berdasarkan
    konvensi nama (akhiran tidak peka huruf besar/kecil):

      {{NP_nominal_TERBILANG}}        → "DUA RATUS JUTA RUPIAH"
      {{NP_nominal_TERBILANG_TITLE}}  → "Dua Ratus Juta Rupiah"
      {{NP_nominal_FORMAT}}           → "Rp 200.000.000,00"

    Key sumber yang punya turunan ikut diformat Indonesia: "200.000.000,00".
    Placeholder yang sudah punya nilai di Excel tidak ditimpa.

    Kembalikan (nilai_tambahan, peringatan) — peringatan berisi placeholder
    turunan yang tidak bisa dibuat beserta alasannya.
    """
    tambahan:   dict[str, str] = {}
    peringatan: dict[str, str] = {}

    for ph in sorted(placeholders):
        if ph in data_dict:
            continue
        sumber, fungsi = _pecah_turunan(ph)
        if sumber is None or sumber not in data_dict:
            continue
        angka = _parse_ke_int(data_dict[sumber])
        if angka is None:
            peringatan[ph] = f"Nilai {{{{{sumber}}}}} bukan angka"
            continue

        tambahan[ph]     = fungsi(angka)
        tambahan[sumber] = format_angka_indonesia(angka)

    return tambahan, peringatan


# =============================================================================
# Paragraph-level placeholder replacement — berbasis XML langsung
# =============================================================================
#
# MENGAPA XML, BUKAN para.runs?
# para.runs hanya mengembalikan <w:r> anak-langsung dari <w:p>.
# Placeholder bisa berada di dalam elemen pembungkus yang TIDAK terlihat oleh
# para.runs, antara lain:
#   • <w:hyperlink>   — teks yang merupakan hyperlink
#   • <w:ins>         — teks hasil tracked-change (insertion)
#   • <w:sdt>         — content control (structured document tag)
# Dengan para._element.iter(W_T) kita mendapatkan SEMUA <w:t> di seluruh
# sub-pohon XML, sehingga tidak ada placeholder yang terlewat.
# Karena kita hanya mengubah .text pada elemen <w:t> (bukan <w:rPr>),
# semua formatting (bold/italic/warna/font) tetap terjaga otomatis.
#
_W_NS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_T   = f"{{{_W_NS}}}t"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _set_t_text(t_elem, text: str):
    """Tulis teks ke elemen <w:t>, pasang xml:space="preserve" jika ada spasi di tepi."""
    t_elem.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t_elem.set(_XML_SPACE, "preserve")


def _replace_in_para_xml(para, placeholder: str, value: str) -> int:
    """
    Ganti semua kemunculan placeholder di paragraf secara langsung via XML.
    Menangani: run normal, hyperlink, tracked-change, content-control,
               dan placeholder yang terpecah di beberapa elemen <w:t>.
    Kembalikan jumlah penggantian yang dilakukan.
    """
    t_elems = list(para._element.iter(_W_T))
    if not t_elems:
        return 0

    count = 0

    # Babak 1: ganti dalam satu elemen <w:t> (kasus paling umum)
    for t in t_elems:
        txt = t.text or ""
        if placeholder in txt:
            count += txt.count(placeholder)
            _set_t_text(t, txt.replace(placeholder, value))

    # Babak 2: ganti placeholder yang terpecah di beberapa <w:t>
    # Loop sampai tidak ada lagi kemunculan (menangani beberapa kejadian sekaligus)
    for _ in range(len(t_elems)):           # batas atas iterasi yang aman
        texts    = [t.text or "" for t in t_elems]
        full     = "".join(texts)
        if placeholder not in full:
            break

        ph_start = full.index(placeholder)
        ph_end   = ph_start + len(placeholder)

        # Bangun posisi kumulatif setiap elemen
        cum, cum_pos = 0, []
        for txt in texts:
            cum_pos.append(cum)
            cum += len(txt)

        s_idx = e_idx = None
        for i, (cp, txt) in enumerate(zip(cum_pos, texts)):
            cp_end = cp + len(txt)
            if cp <= ph_start < cp_end:
                s_idx = i
            if cp < ph_end <= cp_end:
                e_idx = i

        # Tangani placeholder yang tepat berakhir di batas elemen
        if e_idx is None:
            for i, cp in enumerate(cum_pos):
                if cp + len(texts[i]) == ph_end:
                    e_idx = i
                    break

        if s_idx is None or e_idx is None or s_idx == e_idx:
            break   # tidak dapat diselesaikan, hentikan agar tidak infinite loop

        before = texts[s_idx][: ph_start - cum_pos[s_idx]]
        after  = texts[e_idx][ph_end   - cum_pos[e_idx] :]

        _set_t_text(t_elems[s_idx], before + value)
        _set_t_text(t_elems[e_idx], after)
        for i in range(s_idx + 1, e_idx):
            t_elems[i].text = ""

        count += 1

    return count


def replace_placeholders_in_paragraph(para, data: dict):
    for key, value in data.items():
        _replace_in_para_xml(para, f"{{{{{key}}}}}", str(value))
    return para


# =============================================================================
# Document-level replacement (body + tabel tersarang + header + footer)
# =============================================================================

def _iter_table_paragraphs(tables):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                yield from _iter_table_paragraphs(cell.tables)


def iter_all_paragraphs(doc: Document):
    """Semua paragraf dokumen: body, tabel tersarang, header, dan footer."""
    yield from doc.paragraphs
    yield from _iter_table_paragraphs(doc.tables)

    for section in doc.sections:
        for hf in (
            section.header, section.footer,
            section.even_page_header, section.even_page_footer,
            section.first_page_header, section.first_page_footer,
        ):
            if hf is not None:
                yield from hf.paragraphs
                yield from _iter_table_paragraphs(hf.tables)


def replace_placeholders(doc: Document, data: dict) -> Document:
    for para in iter_all_paragraphs(doc):
        replace_placeholders_in_paragraph(para, data)
    return doc


_POLA_PLACEHOLDER = re.compile(r"\{\{([^{}]+)\}\}")


def collect_placeholders(doc: Document) -> set[str]:
    """Nama semua placeholder {{...}} di dokumen, termasuk header/footer."""
    nama: set[str] = set()
    for para in iter_all_paragraphs(doc):
        teks = "".join(t.text or "" for t in para._element.iter(_W_T))
        nama.update(_POLA_PLACEHOLDER.findall(teks))
    return nama


# =============================================================================
# Find & Replace langsung (tanpa placeholder, tanpa API)
# =============================================================================

def apply_find_replace(doc: Document, pairs: list[tuple[str, str]]) -> int:
    """
    Ganti kata secara langsung di seluruh dokumen (body, tabel, header, footer).
    Menggunakan _replace_in_para_xml yang sama dengan placeholder replacement
    sehingga hyperlink, tracked-change, dan split antar <w:t> ikut tertangani.
    Kembalikan jumlah total penggantian yang dilakukan.
    """
    total = 0
    substitutions = [(f, r) for f, r in pairs if f.strip()]

    for para in iter_all_paragraphs(doc):
        for find, replace in substitutions:
            total += _replace_in_para_xml(para, find, replace)

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
    word_file.seek(0)
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


# =============================================================================
# Kecocokan placeholder template vs data Excel (tanpa API)
# =============================================================================

def build_placeholder_check(
    placeholders: dict[str, list[str]],
    data_excel: dict[str, str],
    data_final: dict[str, str],
    peringatan: dict[str, str],
) -> pd.DataFrame:
    """
    Satu baris per placeholder di template, dengan nilai yang akan dipakai
    dan statusnya: dari Excel, dibuat otomatis, kosong, atau tidak ada.
    """
    rows = []
    for ph in sorted(placeholders, key=str.lower):
        if ph in peringatan:
            status = f"⚠️ {peringatan[ph]}"
        elif ph not in data_final:
            status = "❌ Tidak ada di Excel"
        elif not data_final[ph]:
            status = "⚠️ Nilai kosong"
        elif ph in data_excel:
            status = "✅ Dari Excel"
        else:
            status = "✅ Otomatis"
        rows.append({
            "Placeholder": f"{{{{{ph}}}}}",
            "Nilai":       data_final.get(ph, ""),
            "Status":      status,
            "Dokumen":     ", ".join(placeholders[ph]),
        })
    return pd.DataFrame(rows, columns=["Placeholder", "Nilai", "Status", "Dokumen"])


# =============================================================================
# Konsistensi nominal angka vs terbilang (dihitung lokal — tidak butuh API)
# =============================================================================

# "Rp 200.000.000,00 (DUA RATUS JUTA RUPIAH)" — angka dan terbilang ditangkap
# berpasangan sehingga satu terbilang yang hilang tidak menggeser pasangan lain.
_POLA_NOMINAL = re.compile(
    r"Rp\.?\s*(\d[\d.]*(?:,\d{1,2})?)\s*\(([^()]*rupiah[^()]*)\)", re.IGNORECASE
)


def check_value_consistency_table(text: str) -> pd.DataFrame:
    """
    Cocokkan setiap nominal "Rp ... (... RUPIAH)" dengan terbilang yang benar
    menurut angka_ke_terbilang. Tidak peka huruf besar/kecil.
    """
    rows = []
    for angka_teks, kata in _POLA_NOMINAL.findall(text):
        angka = _parse_ke_int(angka_teks)
        if angka is None:
            continue
        kata_norm = " ".join(kata.upper().split())
        benar     = angka_ke_terbilang(angka)
        rows.append({
            "Angka (Rp)":          format_rupiah(angka),
            "Terbilang di Dokumen": kata.strip(),
            "Status": "✅ Sesuai" if kata_norm == benar else f"❌ Seharusnya: {benar}",
        })
    return pd.DataFrame(rows, columns=["Angka (Rp)", "Terbilang di Dokumen", "Status"])


# =============================================================================
# Claude API — tiga fitur utama
# =============================================================================

def _get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


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
    st.header("🔢 Terbilang Otomatis")
    st.caption(
        "Tambahkan akhiran pada key Excel di template Word. "
        "Contoh untuk key `NP_nominal` = 200000000:"
    )
    st.markdown(
        "- `{{NP_nominal}}` → 200.000.000,00\n"
        "- `{{NP_nominal_FORMAT}}` → Rp 200.000.000,00\n"
        "- `{{NP_nominal_TERBILANG}}` → DUA RATUS JUTA RUPIAH\n"
        "- `{{NP_nominal_TERBILANG_TITLE}}` → Dua Ratus Juta Rupiah"
    )
    st.caption("Key sumber diformat ribuan hanya bila salah satu akhiran dipakai.")

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

    # Format tanggal/angka sebelum ditampilkan, agar yang terlihat di tabel
    # sama persis dengan yang akan ditempel ke laporan.
    df_excel = df_excel.apply(lambda col: col.map(format_nilai))

    st.subheader("📊 Tabel Data pada Lembar Kerja (edit jika perlu)")
    df_editable = st.data_editor(df_excel, num_rows="dynamic", use_container_width=True)

    data_excel: dict[str, str] = {}
    for _, row in df_editable.iterrows():
        key = format_nilai(row.iloc[0])
        if key:
            data_excel[key] = format_nilai(row.iloc[1])

    # ---------------------------------------------------------------------------
    # Pindai placeholder template & buat terbilang otomatis dari akhiran nama
    # ---------------------------------------------------------------------------
    placeholders_per_file: dict[str, list[str]] = {}
    for wf in word_files or []:
        wf.seek(0)
        for ph in collect_placeholders(Document(wf)):
            placeholders_per_file.setdefault(ph, []).append(wf.name)
        wf.seek(0)

    auto_tb, peringatan_tb = generate_auto_terbilang(data_excel, set(placeholders_per_file))
    data_dict = {**data_excel, **auto_tb}

    # ---------------------------------------------------------------------------
    # Cek kecocokan placeholder template vs data
    # ---------------------------------------------------------------------------
    if word_files:
        st.subheader("✅ Cek Kelengkapan Data")
        check_df = build_placeholder_check(
            placeholders_per_file, data_excel, data_dict, peringatan_tb
        )
        if check_df.empty:
            st.info("Tidak ada placeholder {{...}} di template.")
        else:
            jumlah_ok = check_df["Status"].str.startswith("✅").sum()
            jumlah_masalah = len(check_df) - jumlah_ok
            if jumlah_masalah:
                st.warning(
                    f"{jumlah_masalah} dari {len(check_df)} placeholder perlu diperiksa."
                )
            else:
                st.success(f"Semua {len(check_df)} placeholder terisi.")
            st.dataframe(check_df, use_container_width=True, hide_index=True)

        dipakai = set(placeholders_per_file) | {
            _pecah_turunan(ph)[0] for ph in placeholders_per_file
        }
        tidak_dipakai = sorted(set(data_excel) - dipakai)
        if tidak_dipakai:
            st.caption(
                f"{len(tidak_dipakai)} key Excel tidak dipakai di template: "
                + ", ".join(f"`{k}`" for k in tidak_dipakai)
            )

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
            consistency_df = check_value_consistency_table(raw_text)
            remaining = sorted(collect_placeholders(updated_doc), key=str.lower)

            with st.expander(f"📁 {word_file.name}", expanded=True):

                # -- Peringatan placeholder belum terisi (tanpa AI) --
                if remaining:
                    st.error(
                        f"**{len(remaining)} placeholder belum terganti:** "
                        + ", ".join(f"`{{{{{p}}}}}`" for p in remaining)
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
                    st.markdown("🔍 **Cek Konsistensi Nominal Angka vs Terbilang**")
                    st.dataframe(consistency_df, use_container_width=True, hide_index=True)

                    if consistency_df["Status"].str.startswith("❌").any():
                        st.warning("Terdapat ketidaksesuaian nominal — periksa sebelum diunduh.")
                    else:
                        st.success("Semua nominal sesuai dengan terbilangnya.")

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
