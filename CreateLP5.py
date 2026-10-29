"""SRR Kalibata Report Maker — UI Streamlit. Logika inti ada di paket report_maker."""
import hashlib
import os
import zipfile
from io import BytesIO
from pathlib import Path

import anthropic
import pandas as pd
import streamlit as st

from report_maker.ai import get_client, review_document_ai
from report_maker.checks import (
    build_placeholder_check,
    check_value_consistency_table,
    deteksi_baris_awal,
)
from report_maker.docx_engine import (
    highlight_placeholders,
    konteks_placeholder,
    isi_dokumen,
    nama_file_hasil,
    placeholders_dari_bytes,
)
from report_maker.lembar_kerja import (
    KEYAKINAN,
    Pemetaan,
    baca_sel,
    kebutuhan_kunci,
    pemetaan_dari_json,
    pemetaan_ke_json,
    petakan_dengan_ai,
    terapkan_pemetaan,
)
from report_maker.pdf import docx_ke_pdf, pdf_tersedia
from report_maker.penyesuaian import (
    Usulan,
    daftar_paragraf,
    excel_ke_teks,
    terapkan_penyesuaian,
    usulkan_penyesuaian,
    validasi_usulan,
)
from report_maker.terbilang import format_nilai, generate_auto_terbilang, pecah_turunan

MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# Logo opsional: taruh file di assets/logo_srr.png (atau .jpg / .svg)
LOGO_PATH = next(
    (
        p for ext in ("png", "jpg", "jpeg", "svg")
        if (p := Path(__file__).parent / "assets" / f"logo_srr.{ext}").exists()
    ),
    None,
)

st.set_page_config(
    page_title="SRR Kalibata Report Maker",
    page_icon=str(LOGO_PATH) if LOGO_PATH else "📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

if LOGO_PATH:
    col_logo, col_judul = st.columns([1, 8], vertical_alignment="center")
    col_logo.image(str(LOGO_PATH), width="stretch")
else:
    col_judul = st.container()
col_judul.title("SRR Kalibata Report Maker")
col_judul.caption("Isi template laporan Word dari data Excel secara otomatis.")


# =============================================================================
# Cache — dokumen hanya diproses ulang bila template, data, atau F&R berubah
# =============================================================================

@st.cache_data(show_spinner=False, max_entries=20)
def _placeholders(docx_bytes: bytes) -> set[str]:
    return placeholders_dari_bytes(docx_bytes)


@st.cache_data(show_spinner=False, max_entries=50)
def _isi(docx_bytes: bytes, data: dict, pairs: tuple) -> tuple[bytes, str, list[str]]:
    return isi_dokumen(docx_bytes, data, pairs)


@st.cache_data(show_spinner=False, max_entries=50)
def _pdf(docx_bytes: bytes) -> bytes:
    return docx_ke_pdf(docx_bytes)


@st.cache_data(show_spinner=False, max_entries=20)
def _paragraf(docx_bytes: bytes) -> dict[int, str]:
    return daftar_paragraf(docx_bytes)


@st.cache_data(show_spinner=False, max_entries=20)
def _excel_teks(xlsx_bytes: bytes) -> str:
    return excel_ke_teks(xlsx_bytes)


@st.cache_data(show_spinner=False, max_entries=10)
def _sel_lembar_kerja(xlsx_bytes: bytes) -> dict[str, str]:
    return baca_sel(xlsx_bytes)


@st.cache_data(show_spinner=False, max_entries=20)
def _konteks(docx_bytes: bytes) -> dict[str, str]:
    return konteks_placeholder(docx_bytes)


def _api_key_tersimpan() -> str:
    """API key dari .streamlit/secrets.toml atau variabel lingkungan, bila ada."""
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:  # secrets.toml tidak ada
        key = ""
    return key or os.environ.get("ANTHROPIC_API_KEY", "")


# =============================================================================
# Streamlit UI — tiga langkah: Upload → Periksa → Unduh
# =============================================================================

# ---------------------------------------------------------------------------
# ① Upload
# ---------------------------------------------------------------------------
st.subheader("① Upload")
col_word, col_excel = st.columns(2)
word_files = col_word.file_uploader(
    "Template Laporan (.docx)", type=["docx"], accept_multiple_files=True
)
excel_file = col_excel.file_uploader("Lembar Kerja (.xlsx)", type=["xlsx"])

MODE_PETA = "Baca langsung dari lembar kerja"
MODE_SHEET = "Sheet berisi key | nilai"

sheet_names: list[str] = []
if excel_file:
    try:
        sheet_names = pd.ExcelFile(excel_file, engine="openpyxl").sheet_names
    except Exception as e:
        st.error(f"Gagal membaca Excel: {e}")
        st.stop()
    finally:
        excel_file.seek(0)

mode_data = st.radio(
    "Sumber data",
    [MODE_PETA, MODE_SHEET],
    index=1 if "Laporan" in sheet_names else 0,
    horizontal=True,
    help=(
        "**Baca langsung**: AI mencari sendiri sel yang cocok untuk tiap placeholder "
        "di sheet mana pun (mis. alamat di sheet LK Tanah) — tidak perlu sheet Laporan.  \n"
        "**Sheet key | nilai**: cara lama, satu sheet berisi nama key dan nilainya."
    ),
)

with st.expander("Pengaturan"):
    if mode_data == MODE_PETA:
        st.markdown("**Pemetaan tersimpan** — pakai lagi pemetaan dari laporan sebelumnya.")
        file_pemetaan = st.file_uploader(
            "Pemetaan (.json)", type=["json"], label_visibility="collapsed"
        )
        sheet_name, skiprows = None, None
    else:
        file_pemetaan = None
        st.markdown("**Excel**")
        c1, c2 = st.columns(2)
        if sheet_names:
            sheet_name = c1.selectbox(
                "Sheet",
                sheet_names,
                index=sheet_names.index("Laporan") if "Laporan" in sheet_names else 0,
            )
        else:
            sheet_name = None
            c1.selectbox("Sheet", [], placeholder="Upload Excel dulu", disabled=True)
        skiprows = c2.number_input(
            "Lewati baris pertama",
            min_value=0,
            value=None,
            step=1,
            placeholder="Otomatis",
            help="Kosongkan agar baris awal data dideteksi dari key yang dipakai template.",
        )

    st.markdown("**Nama file hasil**")
    pola_nama = st.text_input(
        "Nama file hasil",
        value="{{nama_file}}_terisi",
        label_visibility="collapsed",
        help=(
            "Boleh memakai key Excel, misalnya `LP_{{nama_debitur}}_{{tgl_laporan}}`. "
            "`{{nama_file}}` = nama template."
        ),
    )

    st.markdown("**Find & Replace** — ganti kata langsung di dokumen, tanpa placeholder.")
    df_fr = st.data_editor(
        pd.DataFrame({"Cari": pd.Series(dtype=str), "Ganti dengan": pd.Series(dtype=str)}),
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="find_replace",
    )
    active_pairs = tuple(
        (str(f), "" if pd.isna(r) else str(r))
        for f, r in zip(df_fr["Cari"], df_fr["Ganti dengan"])
        if not pd.isna(f) and str(f).strip()
    )

    st.markdown("**Claude AI** (opsional) — untuk review dokumen otomatis.")
    api_key = _api_key_tersimpan()
    if api_key:
        st.caption("API key diambil dari secrets / variabel lingkungan `ANTHROPIC_API_KEY`.")
    else:
        api_key = st.text_input(
            "API Key Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="Dapatkan API key di console.anthropic.com. "
                 "Simpan di .streamlit/secrets.toml agar tidak perlu diisi ulang.",
        )
    claude_ready = bool(api_key and api_key.startswith("sk-ant-"))
    if api_key and not claude_ready:
        st.warning("Format API key tidak valid.")

with st.expander("Panduan placeholder"):
    st.markdown(
        "Tulis placeholder sebagai `{{nama}}` di template Word. Pada mode "
        "**Baca langsung dari lembar kerja**, beri nama yang menjelaskan isinya "
        "(mis. `{{alamat_objek}}`, `{{luas_tanah}}`, `{{nilai_pasar}}`) — AI "
        "mencarikan selnya di lembar kerja. Pada mode **Sheet key | nilai**, "
        "nama = isi kolom pertama sheet. "
        "Untuk nominal, tambahkan akhiran agar terbilang dibuat otomatis. "
        "Contoh untuk key `NP_nominal` = 200000000:\n\n"
        "| Placeholder | Hasil |\n|---|---|\n"
        "| `{{NP_nominal}}` | 200.000.000,00 |\n"
        "| `{{NP_nominal_FORMAT}}` | Rp 200.000.000,00 |\n"
        "| `{{NP_nominal_TERBILANG}}` | DUA RATUS JUTA RUPIAH |\n"
        "| `{{NP_nominal_TERBILANG_TITLE}}` | Dua Ratus Juta Rupiah |"
    )
    st.caption("Key sumber diformat ribuan hanya bila salah satu akhiran dipakai.")

if not (word_files and excel_file):
    st.info("Upload template Word dan lembar kerja Excel untuk melanjutkan.")
    st.stop()


# ---------------------------------------------------------------------------
# Pindai placeholder template
# ---------------------------------------------------------------------------
template_bytes = {wf.name: wf.getvalue() for wf in word_files}
placeholders_per_file: dict[str, list[str]] = {}
for nama, isi in template_bytes.items():
    for ph in _placeholders(isi):
        placeholders_per_file.setdefault(ph, []).append(nama)

st.divider()
st.subheader("② Periksa")

# Peringatan dari sumber data (mis. sel kosong) — ditampilkan di Kelengkapan Data
peringatan_data: dict[str, str] = {}

if mode_data == MODE_PETA:
    # -----------------------------------------------------------------------
    # Baca langsung dari lembar kerja: placeholder → sel (dipetakan AI)
    # -----------------------------------------------------------------------
    excel_bytes = excel_file.getvalue()
    try:
        sel = _sel_lembar_kerja(excel_bytes)
    except Exception as e:
        st.error(f"Gagal membaca lembar kerja: {e}")
        st.stop()

    kebutuhan = kebutuhan_kunci(set(placeholders_per_file))
    peta: dict[str, Pemetaan] = st.session_state.setdefault("pemetaan", {})
    st.session_state.setdefault("pemetaan_versi", 0)

    # Pemetaan tersimpan di-upload: gabungkan sekali per file
    if file_pemetaan is not None:
        sidik = hashlib.md5(file_pemetaan.getvalue()).hexdigest()
        if st.session_state.get("pemetaan_file") != sidik:
            try:
                peta.update(pemetaan_dari_json(file_pemetaan.getvalue()))
                st.session_state["pemetaan_versi"] += 1
            except ValueError as e:
                st.error(str(e))
            st.session_state["pemetaan_file"] = sidik

    st.caption(
        f"Lembar kerja: {len(sel)} sel berisi di {len({r.rsplit('!', 1)[0] for r in sel})} sheet · "
        f"{len(kebutuhan)} data dibutuhkan template."
    )

    # Tabel pemetaan — bisa dikoreksi; baris dihapus agar dipetakan ulang AI
    kunci_tampil = [k for k in kebutuhan if k in peta]
    peta_edit: dict[str, Pemetaan] = {}
    if kunci_tampil:
        with st.expander("Pemetaan data → sel lembar kerja (bisa dikoreksi)"):
            df_peta = st.data_editor(
                pd.DataFrame({
                    "Key":       kunci_tampil,
                    "Sel":       ["; ".join(peta[k].sel) for k in kunci_tampil],
                    "Pemisah":   [peta[k].pemisah for k in kunci_tampil],
                    "Keyakinan": [peta[k].keyakinan for k in kunci_tampil],
                    "Alasan":    [peta[k].alasan for k in kunci_tampil],
                }),
                column_config={
                    "Sel": st.column_config.TextColumn(
                        help="Format `Nama Sheet!C5`; beberapa sel dipisah `;`"
                    ),
                    "Keyakinan": st.column_config.SelectboxColumn(options=KEYAKINAN),
                },
                disabled=["Key", "Alasan"],
                num_rows="delete",
                hide_index=True,
                width="stretch",
                key=f"editor_pemetaan_{st.session_state['pemetaan_versi']}",
            )
            for _, r in df_peta.iterrows():
                peta_edit[r["Key"]] = Pemetaan(
                    sel=[x.strip() for x in str(r["Sel"] or "").split(";") if x.strip()],
                    pemisah="" if pd.isna(r["Pemisah"]) else str(r["Pemisah"]),
                    keyakinan=str(r["Keyakinan"]),
                    alasan=str(r["Alasan"]),
                )
            c_simpan, c_hapus = st.columns(2)
            if c_hapus.button(
                "Kosongkan pemetaan",
                help="Untuk lembar kerja dengan format lain — semua data dipetakan ulang.",
            ):
                st.session_state["pemetaan"] = {}
                st.session_state.pop("pemetaan_catatan", None)
                st.session_state["pemetaan_versi"] += 1
                st.rerun()
            c_simpan.download_button(
                "Simpan pemetaan (.json)",
                data=pemetaan_ke_json({**peta, **peta_edit}),
                file_name="pemetaan_lembar_kerja.json",
                mime="application/json",
                help="Upload di Pengaturan untuk laporan berikutnya dengan format "
                     "lembar kerja yang sama — tanpa AI, tanpa biaya.",
            )

    belum = [k for k in kebutuhan if k not in peta_edit]
    if belum:
        st.warning(
            f"{len(belum)} data belum dipetakan ke lembar kerja: "
            + ", ".join(f"`{k}`" for k in belum)
        )
        if not claude_ready:
            st.caption(
                "Isi API key di Pengaturan agar AI bisa mencarikan selnya, "
                "atau upload pemetaan tersimpan."
            )
        elif st.button(f"Petakan {len(belum)} data dengan AI", type="primary"):
            konteks: dict[str, str] = {}
            for isi in template_bytes.values():
                for ph, k in _konteks(isi).items():
                    konteks.setdefault(ph, k)
            try:
                with st.spinner("Claude sedang membaca lembar kerja..."):
                    hasil_peta, catatan_peta = petakan_dengan_ai(
                        get_client(api_key), sel, {k: kebutuhan[k] for k in belum}, konteks
                    )
            except (anthropic.APIError, RuntimeError, ValueError) as e:
                st.error(f"Gagal memetakan: {e}")
            else:
                # Simpan koreksi di tabel + hasil AI, lalu tampilkan ulang tabelnya
                st.session_state["pemetaan"] = {**peta, **peta_edit, **hasil_peta}
                st.session_state["pemetaan_catatan"] = catatan_peta
                st.session_state["pemetaan_versi"] += 1
                st.rerun()

    if catatan_peta := st.session_state.get("pemetaan_catatan"):
        st.info(f"**Catatan Claude:** {catatan_peta}")

    data_excel, peringatan_data = terapkan_pemetaan(peta_edit, sel)

else:
    # -----------------------------------------------------------------------
    # Cara lama: satu sheet berisi key | nilai
    # -----------------------------------------------------------------------
    try:
        df_excel = pd.read_excel(
            excel_file, header=None, sheet_name=sheet_name, engine="openpyxl"
        )
    except Exception as e:
        st.error(f"Gagal membaca Excel: {e}")
        st.stop()

    if df_excel.shape[1] < 2:
        st.error(f"Sheet **{sheet_name}** harus punya minimal 2 kolom: key dan nilai.")
        st.stop()

    # Format tanggal/angka sebelum ditampilkan, agar yang terlihat di tabel
    # sama persis dengan yang akan ditempel ke laporan.
    df_excel = df_excel.apply(lambda col: col.map(format_nilai))

    if skiprows is None:
        baris_awal = deteksi_baris_awal(
            df_excel.iloc[:, 0].tolist(),
            df_excel.iloc[:, 1].tolist(),
            set(placeholders_per_file),
        )
        sumber_baris = "terdeteksi otomatis"
    else:
        baris_awal = int(skiprows)
        sumber_baris = "diatur manual"
    df_excel = df_excel.iloc[baris_awal:].reset_index(drop=True)

    st.caption(
        f"Data dibaca dari sheet **{sheet_name}** mulai baris ke-{baris_awal + 1} "
        f"({sumber_baris}). Ubah di Pengaturan bila keliru."
    )
    with st.expander("Data Excel (bisa diedit)"):
        df_editable = st.data_editor(df_excel, num_rows="dynamic", width="stretch")

    data_excel = {}
    for _, row in df_editable.iterrows():
        key = format_nilai(row.iloc[0])
        if key:
            data_excel[key] = format_nilai(row.iloc[1])

# Buat terbilang otomatis dari akhiran nama placeholder
auto_tb, peringatan_tb = generate_auto_terbilang(data_excel, set(placeholders_per_file))
data_dict = {**data_excel, **auto_tb}

# Peringatan sel/key berlaku juga untuk placeholder turunannya ({{X_TERBILANG}})
for ph in placeholders_per_file:
    sumber = pecah_turunan(ph)[0] or ph
    if sumber in peringatan_data and ph not in peringatan_tb:
        peringatan_tb[ph] = peringatan_data[sumber]

# -- Kelengkapan data: placeholder template vs Excel --
st.markdown("**Kelengkapan Data**")
check_df = build_placeholder_check(
    placeholders_per_file, data_excel, data_dict, peringatan_tb
)
if check_df.empty:
    st.info("Tidak ada placeholder {{...}} di template.")
else:
    jumlah_masalah = (~check_df["Status"].str.startswith("✅")).sum()
    if jumlah_masalah:
        st.warning(f"{jumlah_masalah} dari {len(check_df)} placeholder perlu diperiksa.")
    else:
        st.success(f"Semua {len(check_df)} placeholder terisi.")
    st.dataframe(check_df, width="stretch", hide_index=True)

dipakai = set(placeholders_per_file) | {
    pecah_turunan(ph)[0] for ph in placeholders_per_file
}
tidak_dipakai = sorted(set(data_excel) - dipakai)
if tidak_dipakai:
    st.caption(
        f"{len(tidak_dipakai)} key Excel tidak dipakai di template: "
        + ", ".join(f"`{k}`" for k in tidak_dipakai)
    )

# -- Per dokumen: placeholder tersisa, konsistensi nominal, preview, review AI --
processed_files: list[tuple[str, bytes]] = []
masalah_per_file: dict[str, list[str]] = {}

tabs = st.tabs(list(template_bytes))
for tab, (nama_template, isi_template) in zip(tabs, template_bytes.items()):
    output_bytes, raw_text, remaining = _isi(isi_template, data_dict, active_pairs)
    consistency_df = check_value_consistency_table(raw_text)
    salah_nominal = consistency_df["Status"].str.startswith("❌").sum()

    masalah = []
    if remaining:
        masalah.append(f"{len(remaining)} placeholder belum terganti")
    if salah_nominal:
        masalah.append(f"{salah_nominal} nominal tidak sesuai terbilang")
    masalah_per_file[nama_template] = masalah

    with tab:
        if remaining:
            st.error(
                f"**{len(remaining)} placeholder belum terganti:** "
                + ", ".join(f"`{{{{{p}}}}}`" for p in remaining)
            )

        if not consistency_df.empty:
            st.markdown("**Nominal vs Terbilang**")
            if salah_nominal:
                st.warning("Terdapat ketidaksesuaian nominal — periksa sebelum diunduh.")
            else:
                st.success("Semua nominal sesuai dengan terbilangnya.")
            st.dataframe(consistency_df, width="stretch", hide_index=True)

        if not masalah and consistency_df.empty:
            st.success("Tidak ada masalah ditemukan.")

        with st.expander("Preview isi dokumen"):
            st.text_area(
                "Isi Dokumen Terisi (Placeholder ditandai):",
                value=highlight_placeholders(raw_text),
                height=400,
                # Kunci ikut isi: widget berkunci tetap mengabaikan value baru
                key=f"{nama_template}_preview_{hashlib.md5(raw_text.encode()).hexdigest()}",
                disabled=True,
                label_visibility="collapsed",
            )

        # -- Satu tombol review AI (termasuk konteks Find & Replace) --
        # Hasil disimpan per isi dokumen, jadi review lama tidak tampil
        # setelah data atau Find & Replace diubah.
        review_key = "review_" + hashlib.md5(
            (raw_text + repr(active_pairs)).encode()
        ).hexdigest()
        if claude_ready:
            if st.button("Periksa dengan AI", key=nama_template + "_ai_btn"):
                with st.spinner("Claude sedang mereview dokumen..."):
                    st.session_state[review_key] = review_document_ai(
                        raw_text, get_client(api_key), list(active_pairs)
                    )
            if review_key in st.session_state:
                st.markdown(st.session_state[review_key])
        else:
            st.caption("Isi API key di Pengaturan untuk review dokumen dengan AI.")

    processed_files.append((nama_template, output_bytes))


# ---------------------------------------------------------------------------
# ③ Sesuaikan dengan AI — tanpa placeholder, dari dokumen sumber
# ---------------------------------------------------------------------------
st.divider()
st.subheader("③ Sesuaikan dengan AI")
st.caption(
    "Opsional. Claude membandingkan laporan dengan sertifikat (PDF) dan lembar kerja "
    "(Excel), lalu mengusulkan perubahan kalimat yang tidak sesuai — tanpa placeholder. "
    "Tidak ada yang diubah sebelum Anda setujui."
)

BATAS_PDF_MB = 30  # batas satu permintaan API 32 MB, sisakan ruang untuk teks

if not claude_ready:
    st.caption("Isi API key di Pengaturan untuk memakai fitur ini.")
else:
    sumber_files = st.file_uploader(
        "Dokumen sumber (PDF sertifikat, Excel lembar kerja)",
        type=["pdf", "xlsx"],
        accept_multiple_files=True,
        key="dokumen_sumber",
    )
    pdfs = [(f.name, f.getvalue()) for f in sumber_files if f.name.lower().endswith(".pdf")]
    excels = [
        (f.name, _excel_teks(f.getvalue()))
        for f in sumber_files if f.name.lower().endswith(".xlsx")
    ]
    ukuran_pdf_mb = sum(len(isi) for _, isi in pdfs) / 1_000_000
    sidik_sumber = hashlib.md5(
        b"".join(f.name.encode() + f.getvalue() for f in sumber_files)
    ).hexdigest()

    if not sumber_files:
        st.info("Upload dokumen sumber untuk mulai.")
    elif ukuran_pdf_mb > BATAS_PDF_MB:
        st.error(
            f"Total PDF {ukuran_pdf_mb:.1f} MB melebihi batas {BATAS_PDF_MB} MB. "
            "Kurangi jumlah file atau perkecil ukuran scan."
        )
    else:
        tabs_ai = st.tabs([nama for nama, _ in processed_files])
        for idx, (tab, (nama_template, output_bytes)) in enumerate(zip(tabs_ai, processed_files)):
            paragraf = _paragraf(output_bytes)
            kunci = hashlib.md5(output_bytes + sidik_sumber.encode()).hexdigest()
            kunci_usulan = f"usulan_{kunci}"

            with tab:
                st.caption(
                    f"{len(paragraf)} paragraf laporan · {len(pdfs)} PDF · "
                    f"{len(excels)} Excel"
                )
                if st.button("Cari ketidaksesuaian", key=f"cari_{kunci}", type="primary"):
                    try:
                        with st.spinner("Claude sedang membaca dokumen sumber dan laporan..."):
                            usulan, catatan = usulkan_penyesuaian(
                                get_client(api_key), paragraf, pdfs, excels
                            )
                        st.session_state[kunci_usulan] = (
                            *validasi_usulan(paragraf, usulan), catatan
                        )
                    except (anthropic.APIError, RuntimeError, ValueError) as e:
                        st.error(f"Gagal mendapatkan usulan: {e}")

                if kunci_usulan not in st.session_state:
                    continue
                valid, tidak_valid, catatan = st.session_state[kunci_usulan]

                if catatan:
                    st.info(f"**Catatan Claude:** {catatan}")
                if not valid:
                    st.success("Tidak ada ketidaksesuaian yang ditemukan.")
                    continue

                st.markdown(
                    f"**{len(valid)} usulan perubahan** — hapus centang untuk menolak, "
                    "atau ubah kolom *Usulan baru* sebelum diterapkan."
                )
                df_usulan = st.data_editor(
                    pd.DataFrame({
                        "Terapkan":       [True] * len(valid),
                        "Paragraf":       [u.paragraf for u in valid],
                        "Kalimat lama":   [u.teks_lama for u in valid],
                        "Usulan baru":    [u.teks_baru for u in valid],
                        "Alasan":         [u.alasan for u in valid],
                        "Sumber":         [u.sumber for u in valid],
                        "Kutipan sumber": [u.kutipan_sumber for u in valid],
                    }),
                    disabled=["Paragraf", "Kalimat lama", "Alasan", "Sumber", "Kutipan sumber"],
                    hide_index=True,
                    width="stretch",
                    key=f"editor_{kunci}",
                )
                if tidak_valid:
                    with st.expander(f"{len(tidak_valid)} usulan tidak dapat diterapkan"):
                        for u, alasan in tidak_valid:
                            st.markdown(f"- [{u.paragraf}] “{u.teks_lama}” → “{u.teks_baru}” — *{alasan}*")

                disetujui = [
                    Usulan(
                        paragraf=u.paragraf, teks_lama=u.teks_lama,
                        teks_baru="" if pd.isna(baru) else str(baru),
                        alasan=u.alasan, sumber=u.sumber, kutipan_sumber=u.kutipan_sumber,
                    )
                    for u, pakai, baru in zip(valid, df_usulan["Terapkan"], df_usulan["Usulan baru"])
                    if pakai
                ]
                if not disetujui:
                    continue

                # Usulan yang disetujui langsung dipakai untuk file unduhan,
                # lalu dokumen hasilnya diperiksa ulang.
                hasil_ai, jumlah = terapkan_penyesuaian(output_bytes, disetujui)
                processed_files[idx] = (nama_template, hasil_ai)
                _, teks_ai, sisa_ai = _isi(hasil_ai, {}, ())
                salah_ai = check_value_consistency_table(teks_ai)["Status"].str.startswith("❌").sum()
                masalah = []
                if sisa_ai:
                    masalah.append(f"{len(sisa_ai)} placeholder belum terganti")
                if salah_ai:
                    masalah.append(f"{salah_ai} nominal tidak sesuai terbilang")
                masalah_per_file[nama_template] = masalah

                st.success(f"{len(disetujui)} usulan disetujui ({jumlah} penggantian) — ikut di file unduhan.")
                if salah_ai:
                    st.warning("Setelah perubahan, masih ada nominal yang tidak sesuai terbilang.")


# ---------------------------------------------------------------------------
# ④ Unduh
# ---------------------------------------------------------------------------
st.divider()
st.subheader("④ Unduh")

ada_masalah = {nama: m for nama, m in masalah_per_file.items() if m}
if ada_masalah:
    st.warning(
        "Masih ada yang perlu diperiksa:\n"
        + "\n".join(f"- **{nama}**: {', '.join(m)}" for nama, m in ada_masalah.items())
    )
    boleh_unduh = st.checkbox("Saya sudah memeriksa dan tetap ingin mengunduh.")
else:
    st.success("Semua dokumen siap diunduh.")
    boleh_unduh = True

if pdf_tersedia():
    sertakan_pdf = st.toggle("Sertakan PDF", disabled=not boleh_unduh)
else:
    sertakan_pdf = False
    st.caption("Ekspor PDF butuh LibreOffice di server — lihat README.")

# Nama file hasil dari pola; tambahkan nomor bila ada yang kembar
hasil: list[tuple[str, bytes]] = []
dipakai_nama: set[str] = set()
for nama_template, output_bytes in processed_files:
    dasar = nama_file_hasil(pola_nama, nama_template, data_dict)
    nama, n = dasar, 2
    while nama in dipakai_nama:
        nama, n = f"{dasar} ({n})", n + 1
    dipakai_nama.add(nama)
    hasil.append((nama, output_bytes))

pdfs: dict[str, bytes] = {}
if sertakan_pdf and boleh_unduh:
    try:
        with st.spinner("Membuat PDF..."):
            pdfs = {nama: _pdf(isi) for nama, isi in hasil}
    except Exception as e:
        st.error(str(e))

if len(hasil) == 1:
    nama, isi = hasil[0]
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button(
        "Unduh Word (.docx)",
        data=isi,
        file_name=f"{nama}.docx",
        mime=MIME_DOCX,
        type="primary",
        disabled=not boleh_unduh,
        width="stretch",
    )
    if nama in pdfs:
        c2.download_button(
            "Unduh PDF",
            data=pdfs[nama],
            file_name=f"{nama}.pdf",
            mime="application/pdf",
            width="stretch",
        )
else:
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for nama, isi in hasil:
            zipf.writestr(f"{nama}.docx", isi)
            if nama in pdfs:
                zipf.writestr(f"{nama}.pdf", pdfs[nama])
    st.download_button(
        f"Unduh Semua Dokumen (.zip, {len(hasil)} file)",
        data=zip_buffer.getvalue(),
        file_name="Laporan_terisi.zip",
        mime="application/zip",
        type="primary",
        disabled=not boleh_unduh,
    )
