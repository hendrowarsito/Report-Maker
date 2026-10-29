"""Review dokumen dengan Claude (opsional, butuh API key)."""
import json

import anthropic

MODEL_REVIEW = "claude-sonnet-4-6"


# =============================================================================
# Claude API — review dokumen
# =============================================================================

MODEL_ANALISIS = "claude-opus-5-5"


def get_client(api_key: str) -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=api_key)


def minta_json(
    client: anthropic.Anthropic, system: str, konten: list[dict], skema: dict
) -> dict:
    """
    Satu permintaan ke Claude dengan jawaban JSON sesuai skema.
    RuntimeError bila permintaan ditolak atau jawabannya terpotong.
    """
    with client.beta.messages.stream(
        model=MODEL_ANALISIS,
        max_tokens=64000,
        system=system,
        messages=[{"role": "user", "content": konten}],
        output_config={
            "effort": "high",
            "format": {"type": "json_schema", "schema": skema},
        },
        # Bila ditolak filter keamanan, ulangi otomatis di model cadangan Anthropic
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
    ) as stream:
        respons = stream.get_final_message()

    if respons.stop_reason == "refusal":
        raise RuntimeError("Permintaan ditolak oleh Claude. Coba lagi atau periksa dokumen sumber.")
    if respons.stop_reason == "max_tokens":
        raise RuntimeError("Jawaban Claude terpotong. Coba dengan dokumen yang lebih sedikit.")

    return json.loads(next(b.text for b in respons.content if b.type == "text"))


def review_document_ai(
    doc_text: str,
    client: anthropic.Anthropic,
    find_replace_pairs: list[tuple[str, str]] | None = None,
) -> str:
    """
    Review QA lengkap dokumen yang sudah terisi via Claude Sonnet (streaming).
    Cek: placeholder belum terisi, inkonsistensi nilai/tanggal, bagian kosong,
    dan — bila ada — apakah hasil Find & Replace tepat sesuai konteks kalimat.
    """
    cek_fr = ""
    if find_replace_pairs:
        daftar = "\n".join(f'   - "{f}" → "{r}"' for f, r in find_replace_pairs)
        cek_fr = (
            "6. **Find & Replace** — penggantian kata berikut sudah diterapkan. "
            "Sebutkan kemunculan kata pengganti yang konteksnya keliru:\n"
            f"{daftar}\n"
        )

    with client.messages.stream(
        model=MODEL_REVIEW,
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
                    "5. **Anomali data** — nilai tidak wajar, nama tidak konsisten\n"
                    f"{cek_fr}\n"
                    f"Dokumen:\n---\n{doc_text[:10000]}\n---\n\n"
                    "Laporan audit dalam markdown terstruktur. "
                    "Mulai dengan baris status: ✅ LULUS / ⚠️ PERLU REVIEW / ❌ KRITIS"
                ),
            }
        ],
    ) as stream:
        return stream.get_final_text()
