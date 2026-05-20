"""
Agente de Tradução e Formatação ABNT — v2
==========================================
Correções desta versão:
  - Limpeza do texto extraído antes da tradução (quebras espúrias de linha/página)
  - Overlap usado apenas como contexto silencioso no prompt, nunca reescrito no output
  - Chunking determinístico: cada chunk tem índice de início/fim rastreado

Requisitos:
    pip install agno anthropic pymupdf python-docx tqdm python-dotenv

Uso:
    python agent.py --input arquivo.pdf --output saida.docx
    python agent.py --input arquivo.pdf --output saida.docx --format livre
    python agent.py --input arquivo.pdf --output saida.docx --chunk-size 2000
"""

import os
import re
import time
import json
import argparse
from pathlib import Path
from dataclasses import dataclass, field
from datetime import date

import fitz
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tqdm import tqdm
from dotenv import load_dotenv

from agno.agent import Agent
from agno.models.anthropic import Claude

load_dotenv()

# ─────────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────────

CHUNK_SIZE   = 3000
OVERLAP      = 400   # usado apenas como contexto, não copiado no output
MAX_RETRIES  = 3
RETRY_DELAY  = 5
MODEL_ID     = "claude-opus-4-5"


# ─────────────────────────────────────────────
# ESTRUTURAS
# ─────────────────────────────────────────────

@dataclass
class Chunk:
    index: int
    text: str           # texto limpo, SEM overlap — é o que será traduzido
    context: str = ""   # overlap do chunk anterior — apenas para contexto no prompt
    translated: str = ""

@dataclass
class DocumentMeta:
    title: str = "Documento Traduzido"
    language_detected: str = "inglês"
    area: str = "geral"
    total_pages: int = 0
    chunks: list[Chunk] = field(default_factory=list)


# ─────────────────────────────────────────────
# EXTRAÇÃO
# ─────────────────────────────────────────────

def extract_from_pdf(path: str) -> tuple[str, int]:
    doc = fitz.open(path)
    pages = []
    for page in doc:
        pages.append(page.get_text("text"))
    return "\n\n".join(pages), len(doc)

def extract_from_docx(path: str) -> tuple[str, int]:
    doc = Document(path)
    paras = [p.text for p in doc.paragraphs if p.text.strip()]
    text = "\n\n".join(paras)
    return text, max(1, len(text.split()) // 250)

def extract_text(path: str) -> tuple[str, int]:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return extract_from_pdf(path)
    elif ext in (".docx", ".doc"):
        return extract_from_docx(path)
    raise ValueError(f"Formato não suportado: {ext}")


# ─────────────────────────────────────────────
# LIMPEZA DO TEXTO EXTRAÍDO
# ─────────────────────────────────────────────

def clean_extracted_text(text: str) -> str:
    """
    Corrige problemas comuns de extração de PDF/DOCX:

    1. Remove hifenização no fim de linha (word-wrap do PDF)
       ex: "tra-\nbalho" → "trabalho"
    2. Une quebras de linha dentro do mesmo parágrafo.
       Heurística: se a linha não termina com pontuação final nem está
       em branco, e a próxima linha começa com letra minúscula ou
       continuação natural, é quebra espúria.
    3. Normaliza múltiplas linhas em branco (máx. 2 → separador de parágrafo)
    4. Remove espaços e caracteres invisíveis extras.
    """

    # 1. Reúne hifenização de fim de linha: "pala-\nvra" → "palavra"
    text = re.sub(r'-\n(?=[a-záéíóúàâêôãõçA-ZÁÉÍÓÚÀÂÊÔÃÕÇ])', '', text)

    # 2. Une linhas dentro do mesmo parágrafo.
    #    Lógica: divide em parágrafos (2+ newlines), depois dentro de cada
    #    parágrafo une linhas que são claramente continuação.
    paragraphs = re.split(r'\n{2,}', text)
    cleaned_paragraphs = []

    for para in paragraphs:
        lines = para.splitlines()
        unified_lines = []
        buffer = ""

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            if not buffer:
                buffer = line
                continue

            prev_ends_sentence = re.search(r'[.!?:;]\s*$', buffer)
            next_starts_upper  = re.match(r'^[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ0-9]', line)
            looks_like_heading = len(line) < 80 and line.isupper()
            looks_like_list    = re.match(r'^[\-\*•◦▪▸]\s', line) or re.match(r'^\d+[.)]\s', line)

            # Se a linha anterior termina frase OU a próxima parece novo
            # elemento estrutural, preserva como parágrafo separado
            if prev_ends_sentence or looks_like_heading or looks_like_list:
                unified_lines.append(buffer)
                buffer = line
            else:
                # Quebra espúria: une com espaço
                buffer = buffer + " " + line

        if buffer:
            unified_lines.append(buffer)

        cleaned_paragraphs.append("\n".join(unified_lines))

    # 3. Reconstrói com separadores de parágrafo duplos
    text = "\n\n".join(p for p in cleaned_paragraphs if p.strip())

    # 4. Remove espaços múltiplos inline
    text = re.sub(r' {2,}', ' ', text)

    # 5. Remove caracteres de controle (exceto newline e tab)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)

    return text.strip()


# ─────────────────────────────────────────────
# CHUNKING — overlap como contexto, nunca no output
# ─────────────────────────────────────────────

def split_into_chunks(text: str, chunk_size: int = CHUNK_SIZE, overlap: int = OVERLAP) -> list[Chunk]:
    """
    Divide o texto em chunks respeitando parágrafos.
    O overlap do chunk anterior é armazenado em chunk.context
    e enviado ao modelo como contexto, mas o modelo é instruído
    a traduzir APENAS chunk.text — eliminando duplicações no output.
    """
    paragraphs = re.split(r'\n{2,}', text)
    paragraphs = [p.strip() for p in paragraphs if p.strip()]

    chunks: list[Chunk] = []
    current_paras: list[str] = []
    current_len = 0
    chunk_index = 0

    for para in paragraphs:
        # Se adicionar este parágrafo ultrapassa o limite, fecha o chunk
        if current_len + len(para) > chunk_size and current_paras:
            chunk_text = "\n\n".join(current_paras)

            # Contexto = últimos N chars do chunk anterior (se houver)
            context = ""
            if chunks:
                prev_text = chunks[-1].text
                context = prev_text[-overlap:] if len(prev_text) > overlap else prev_text

            chunks.append(Chunk(index=chunk_index, text=chunk_text, context=context))
            chunk_index += 1
            current_paras = [para]
            current_len = len(para)
        else:
            current_paras.append(para)
            current_len += len(para)

    # Último chunk
    if current_paras:
        chunk_text = "\n\n".join(current_paras)
        context = ""
        if chunks:
            prev_text = chunks[-1].text
            context = prev_text[-overlap:] if len(prev_text) > overlap else prev_text
        chunks.append(Chunk(index=chunk_index, text=chunk_text, context=context))

    return chunks


# ─────────────────────────────────────────────
# AGENTES AGNO
# ─────────────────────────────────────────────

def build_translation_agent() -> Agent:
    return Agent(
        model=Claude(id=MODEL_ID),
        description=(
            "Você é um tradutor acadêmico especializado em textos científicos e técnicos. "
            "Traduz com fidelidade semântica, fluência em português brasileiro e terminologia precisa."
        ),
        instructions=[
            "Traduza APENAS o bloco marcado com [TRADUZIR] para o português do Brasil.",
            "O bloco [CONTEXTO ANTERIOR] é fornecido apenas para você entender o fio do texto — NÃO o traduza nem o repita.",
            "Preserve toda a estrutura de parágrafos do bloco [TRADUZIR].",
            "Mantenha termos técnicos corretos; coloque o original entre parênteses quando relevante.",
            "Preserve referências bibliográficas, fórmulas e código exatamente como estão.",
            "Retorne SOMENTE o texto traduzido, sem prefácio, comentário ou marcação extra.",
        ],
        markdown=False,
    )

def build_review_agent() -> Agent:
    return Agent(
        model=Claude(id=MODEL_ID),
        description="Revisor acadêmico brasileiro. Garante fluência, coerência e registro formal.",
        instructions=[
            "Revise o trecho para fluência e coerência em português do Brasil.",
            "Corrija apenas erros de fluência, repetição ou inconsistência de terminologia.",
            "NÃO altere o conteúdo semântico nem a estrutura de parágrafos.",
            "Retorne APENAS o texto revisado, sem comentários.",
        ],
        markdown=False,
    )

def detect_meta(agent: Agent, sample: str) -> dict:
    prompt = (
        "Analise o trecho abaixo e responda SOMENTE em JSON com as chaves:\n"
        "\"language\", \"title\", \"area\"\n\n"
        f"Trecho:\n{sample[:2000]}"
    )
    for _ in range(MAX_RETRIES):
        try:
            resp = agent.run(prompt)
            raw = re.sub(r"```json|```", "", resp.content).strip()
            return json.loads(raw)
        except Exception:
            time.sleep(RETRY_DELAY)
    return {"language": "inglês", "title": "Documento Traduzido", "area": "geral"}


# ─────────────────────────────────────────────
# TRADUÇÃO COM CONTEXTO SEPARADO
# ─────────────────────────────────────────────

def translate_chunk(agent: Agent, chunk: Chunk, doc_context: str) -> str:
    """
    Envia o chunk ao agente com contexto separado.
    O modelo é instruído explicitamente a traduzir SOMENTE [TRADUZIR].
    """
    context_block = ""
    if chunk.context.strip():
        context_block = f"[CONTEXTO ANTERIOR — não traduza]\n{chunk.context}\n\n"

    prompt = (
        f"Documento: {doc_context}\n\n"
        f"{context_block}"
        f"[TRADUZIR]\n{chunk.text}"
    )

    for attempt in range(MAX_RETRIES):
        try:
            resp = agent.run(prompt)
            return resp.content.strip()
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                print(f"  ⚠️  Tentativa {attempt+1} falhou: {e}. Aguardando {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY)
            else:
                print(f"  ❌ Chunk {chunk.index} falhou — mantendo original.")
                return chunk.text
    return chunk.text


# ─────────────────────────────────────────────
# FORMATAÇÃO ABNT — DOCX
# ─────────────────────────────────────────────

def apply_abnt_page_style(doc: Document):
    section = doc.sections[0]
    section.top_margin    = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(3)
    section.right_margin  = Cm(2)
    section.page_height   = Cm(29.7)
    section.page_width    = Cm(21.0)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    pf = style.paragraph_format
    pf.alignment          = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing       = Pt(18)
    pf.space_before       = Pt(0)
    pf.space_after        = Pt(0)
    pf.first_line_indent  = Cm(1.25)

def add_title_page(doc: Document, meta: DocumentMeta):
    for text, bold, size in [
        ("DOCUMENTO TRADUZIDO", True, 14),
        ("", False, 12),
        (meta.title.upper(), True, 14),
        ("", False, 12),
        ("", False, 12),
        (f"Traduzido do {meta.language_detected} para o Português do Brasil", False, 12),
        (f"Área: {meta.area}", False, 12),
        (f"Páginas estimadas: {meta.total_pages}", False, 12),
        ("", False, 12),
        (str(date.today().year), False, 12),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)
            run.font.name = "Times New Roman"
    doc.add_page_break()

def is_heading(text: str) -> bool:
    return (
        len(text) < 120 and (
            text.isupper() or
            re.match(r'^\d+[\.\d]*\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÇ]', text) or
            re.match(r'^(abstract|introduction|conclusion|references|resumo|introdução|conclusão|referências)\s*$', text, re.I)
        )
    )

def write_docx(meta: DocumentMeta, output_path: str, format_type: str = "abnt"):
    doc = Document()
    if format_type == "abnt":
        apply_abnt_page_style(doc)
        add_title_page(doc, meta)

    for chunk in meta.chunks:
        body = chunk.translated or chunk.text
        for para_text in body.split("\n\n"):
            para_text = para_text.strip()
            if not para_text:
                continue
            if is_heading(para_text):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(para_text)
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after  = Pt(6)
            else:
                p = doc.add_paragraph(para_text)
                if format_type == "abnt":
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.save(output_path)
    print(f"\n✅ Documento salvo: {output_path}")


# ─────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────

def run_pipeline(input_path: str, output_path: str, format_type: str, force_review: bool):
    print("\n" + "="*60)
    print("  AGENTE DE TRADUÇÃO E FORMATAÇÃO — v2")
    print("="*60)

    # 1. Extração
    print(f"\n📄 Extraindo texto de: {input_path}")
    raw_text, total_pages = extract_text(input_path)
    print(f"   ✔ {total_pages} páginas | {len(raw_text):,} caracteres brutos")

    # 2. Limpeza
    print("\n🧹 Limpando texto extraído (quebras espúrias, hifenização, espaços)...")
    clean_text = clean_extracted_text(raw_text)
    removed = len(raw_text) - len(clean_text)
    print(f"   ✔ Limpeza concluída | {abs(removed):,} chars ajustados")

    # 3. Chunking
    print(f"\n✂️  Dividindo em chunks ({CHUNK_SIZE} chars, overlap {OVERLAP} como contexto)...")
    chunks = split_into_chunks(clean_text)
    print(f"   ✔ {len(chunks)} chunks gerados")

    # 4. Agentes
    translator = build_translation_agent()
    reviewer   = build_review_agent()

    # 5. Metadados
    print("\n🔍 Detectando idioma e metadados...")
    meta_info = detect_meta(translator, clean_text[:3000])
    meta = DocumentMeta(
        title=meta_info.get("title", "Documento Traduzido"),
        language_detected=meta_info.get("language", "inglês"),
        area=meta_info.get("area", "geral"),
        total_pages=total_pages,
        chunks=chunks,
    )
    print(f"   ✔ Idioma: {meta.language_detected} | Área: {meta.area}")
    print(f"   ✔ Título: {meta.title}")

    doc_context = f"Título: {meta.title}. Área: {meta.area}. Idioma: {meta.language_detected}."

    # 6. Tradução
    print(f"\n🌐 Traduzindo {len(chunks)} chunks...\n")
    for chunk in tqdm(chunks, desc="Traduzindo", unit="chunk"):
        chunk.translated = translate_chunk(translator, chunk, doc_context)
        time.sleep(0.5)

    # 7. Revisão
    do_review = force_review or len(chunks) <= 20
    if do_review:
        label = "forçada" if force_review else "automática (doc ≤ 20 chunks)"
        print(f"\n🔎 Revisão {label}...\n")
        for chunk in tqdm(chunks, desc="Revisando", unit="chunk"):
            try:
                resp = reviewer.run(chunk.translated)
                chunk.translated = resp.content.strip()
            except Exception as e:
                print(f"  ⚠️  Revisão do chunk {chunk.index} pulada: {e}")
            time.sleep(0.5)
    else:
        print(f"\n⚡ {len(chunks)} chunks — revisão ignorada. Use --review para forçar.")

    # 8. Documento final
    print(f"\n📝 Gerando documento ({format_type.upper()})...")
    write_docx(meta, output_path, format_type)

    print("\n" + "="*60)
    print("  CONCLUÍDO!")
    print(f"  Arquivo: {output_path}")
    print(f"  Chunks: {len(chunks)} | Páginas: {total_pages}")
    print("="*60 + "\n")


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

def main():
    global CHUNK_SIZE
    parser = argparse.ArgumentParser(description="Agente de tradução e formatação ABNT — v2")
    parser.add_argument("--input",  "-i", required=True, help="Arquivo de entrada (PDF ou DOCX)")
    parser.add_argument("--output", "-o", required=True, help="Arquivo de saída (.docx)")
    parser.add_argument("--format", "-f", default="abnt", choices=["abnt", "livre"])
    parser.add_argument("--chunk-size", type=int, default=CHUNK_SIZE)
    parser.add_argument("--review", action="store_true", help="Forçar revisão completa")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"❌ Arquivo não encontrado: {args.input}")
        return

    
    CHUNK_SIZE = args.chunk_size

    run_pipeline(args.input, args.output, args.format, args.review)

if __name__ == "__main__":
    main()