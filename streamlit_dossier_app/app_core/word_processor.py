"""
Module dédié au traitement des fichiers Word (.docx, .doc) pour l'analyse de synthèse.
Gère l'extraction de texte et la conversion vers PDF de manière robuste.
"""

import os
import csv
import json
import math
import tempfile
import subprocess
import re
from pathlib import Path
from typing import Optional, Tuple, Any
from collections import OrderedDict

import requests
from openai import OpenAI
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def _extract_docx_text(docx_path: Path) -> str:
    """
    Extrait le texte complet d'un document Word (.docx).

    Version améliorée qui:
    - Extrait les paragraphes ET les tableaux dans l'ordre réel
    - Préserve les en-têtes et pieds de page
    - Gère les erreurs proprement

    Args:
        docx_path: Chemin vers le fichGet-ChildItem -Path "AGENT ANTIGRAVITY" -Recurse -Filter "__pycache__" -Directory | Remove-Item -Recurse -Force
er .docx

    Returns:
        str: Texte complet extrait (vide si erreur)

    Example:
        >>> text = _extract_docx_text(Path("document.docx"))
        >>> print(f"Extrait: {len(text)} caractères")
    """
    if not docx_path.exists():
        print(f"[EXTRACTION-DOCX] Fichier introuvable: {docx_path}")
        return ""

    try:
        doc = Document(str(docx_path))
        extracted_parts = []

        # 1. En-têtes et pieds de page
        for i, section in enumerate(doc.sections, 1):
            if section.header:
                header_text = "\n".join([
                    p.text.strip()
                    for p in section.header.paragraphs
                    if p.text.strip()
                ])
                if header_text:
                    extracted_parts.append(f"[EN-TÊTE SECTION {i}]\n{header_text}\n")

            if section.footer:
                footer_text = "\n".join([
                    p.text.strip()
                    for p in section.footer.paragraphs
                    if p.text.strip()
                ])
                if footer_text:
                    extracted_parts.append(f"[PIED DE PAGE SECTION {i}]\n{footer_text}\n")

        # 2. Corps du document (paragraphes + tableaux dans l'ordre)
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraphe
                para = Paragraph(element, doc)
                text = para.text.strip()
                if text:
                    extracted_parts.append(text)

            elif element.tag.endswith('tbl'):  # Tableau
                table = Table(element, doc)
                rows_text = []

                for row in table.rows:
                    cells_text = []
                    for cell in row.cells:
                        # Combiner tous les paragraphes de la cellule
                        cell_content = " ".join([
                            p.text.strip()
                            for p in cell.paragraphs
                            if p.text.strip()
                        ])
                        cells_text.append(cell_content)

                    row_text = " | ".join(cells_text)
                    if row_text.strip():
                        rows_text.append(row_text)

                if rows_text:
                    table_text = "\n".join(rows_text)
                    extracted_parts.append(f"\n[TABLEAU]\n{table_text}\n[/TABLEAU]\n")

        # 3. Assemblage
        full_text = "\n".join(extracted_parts)

        # 4. Nettoyage
        full_text = re.sub(r'\n{3,}', '\n\n', full_text)  # Max 2 sauts de ligne
        full_text = re.sub(r' {2,}', ' ', full_text)      # Pas d'espaces multiples
        full_text = full_text.strip()

        # 5. Validation
        if not full_text:
            print(f"[EXTRACTION-DOCX] ATTENTION: Document vide: {docx_path.name}")
            return ""

        print(f"[EXTRACTION-DOCX] {docx_path.name}: {len(full_text):,} caractères extraits")
        return full_text

    except Exception as e:
        print(f"[EXTRACTION-DOCX] ERREUR lors de l'extraction de {docx_path.name}: {e}")
        print("[EXTRACTION-DOCX] Tentative de récupération via le parseur XML direct...")
        try:
            import zipfile
            import xml.etree.ElementTree as ET
            extracted_parts = []
            with zipfile.ZipFile(docx_path) as z:
                # Trier les noms pour avoir document.xml en premier, puis les headers/footers
                names = sorted(z.namelist())
                for name in names:
                    if name == 'word/document.xml' or name.startswith('word/header') or name.startswith('word/footer'):
                        try:
                            xml_content = z.read(name)
                            root = ET.fromstring(xml_content)
                            texts = []
                            for elem in root.iter():
                                if elem.tag.endswith('t') and elem.text:
                                    texts.append(elem.text)
                            if texts:
                                extracted_parts.append(" ".join(texts))
                        except Exception as xml_err:
                            print(f"[EXTRACTION-XML-FALLBACK] Erreur sur {name}: {xml_err}")
            
            full_text = "\n\n".join(extracted_parts)
            if full_text.strip():
                print(f"[EXTRACTION-DOCX] Récupération XML réussie ! {len(full_text):,} caractères récupérés.")
                return full_text
        except Exception as fallback_err:
            print(f"[EXTRACTION-DOCX] Échec du fallback XML : {fallback_err}")
            
        import traceback
        traceback.print_exc()
        return f"[Erreur extraction Word: {e}]"


def extract_docx_text(docx_path: str | Path) -> str:
    """
    Extrait le texte d'un fichier .docx (wrapper public pour compatibilité).
    """
    return _extract_docx_text(Path(docx_path))


def convert_doc_to_pdf_using_word(doc_path: str | Path, output_pdf_path: str | Path) -> bool:
    """
    Convertit un fichier .doc ou .docx en PDF en utilisant Microsoft Word.
    
    Args:
        doc_path: Chemin vers le fichier Word
        output_pdf_path: Chemin de sortie pour le PDF
        
    Returns:
        True si la conversion réussit, False sinon
    """
    try:
        # Vérifier si Word est disponible
        import win32com.client
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        try:
            doc = word.Documents.Open(str(doc_path))
            doc.SaveAs(str(output_pdf_path), FileFormat=17)  # 17 = PDF format
            doc.Close()
            return True
            
        except Exception as e:
            print(f"Erreur Word COM: {e}")
            return False
            
        finally:
            word.Quit()
            
    except ImportError:
        print("win32com.client non disponible - conversion Word désactivée")
        return False
    except Exception as e:
        print(f"Erreur générale Word COM: {e}")
        return False


def convert_doc_to_pdf_fallback(doc_path: str | Path, output_pdf_path: str | Path) -> bool:
    """
    Fallback pour la conversion Word -> PDF utilisant des outils externes.
    """
    try:
        # Essayer LibreOffice si disponible
        result = subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf', 
            '--outdir', str(Path(output_pdf_path).parent),
            str(doc_path)
        ], capture_output=True, timeout=30)
        
        if result.returncode == 0:
            return True
            
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    
    return False


def process_word_document(word_path: str | Path, temp_dir: str | Path) -> Tuple[str, Optional[str]]:
    """
    Traite un document Word et retourne le texte extrait + chemin PDF généré.
    
    Args:
        word_path: Chemin vers le fichier Word (.docx ou .doc)
        temp_dir: Répertoire temporaire pour les fichiers intermédiaires
        
    Returns:
        Tuple (texte_extrait, chemin_pdf) - chemin_pdf peut être None si échec
    """
    word_path = Path(word_path)
    temp_dir = Path(temp_dir)
    
    extracted_text = ""
    pdf_path = None
    
    try:
        # Extraction du texte selon le format
        if word_path.suffix.lower() == '.docx':
            # Extraction directe .docx -> texte
            extracted_text = extract_docx_text(word_path)
            
            # Générer un PDF de compatibilité
            pdf_path = temp_dir / f"{word_path.stem}_compat.pdf"
            
            # Essayer conversion Word native d'abord
            if not convert_doc_to_pdf_using_word(word_path, pdf_path):
                # Fallback: générer un PDF basique depuis le texte
                from reportlab.lib.pagesizes import A4
                from reportlab.pdfgen import canvas
                
                c = canvas.Canvas(str(pdf_path), pagesize=A4)
                c.setFont("Helvetica", 10)
                
                y = 800
                for line in extracted_text.split('\n'):
                    if y < 50:
                        c.showPage()
                        y = 800
                        c.setFont("Helvetica", 10)
                    
                    c.drawString(50, y, line[:100])
                    y -= 15
                
                c.save()
        
        elif word_path.suffix.lower() == '.doc':
            # Pour .doc, conversion directe en PDF
            pdf_path = temp_dir / f"{word_path.stem}.pdf"
            
            if convert_doc_to_pdf_using_word(word_path, pdf_path):
                # Extraire le texte du PDF généré
                from app_core.pdf_enrichment import extract_text_from_pdf
                extracted_text = extract_text_from_pdf(pdf_path)
            else:
                raise RuntimeError("Impossible de convertir le fichier .doc")
        
        else:
            raise ValueError(f"Format Word non supporté: {word_path.suffix}")
        
        return extracted_text, str(pdf_path) if pdf_path and pdf_path.exists() else None
        
    except Exception as e:
        print(f"Erreur lors du traitement Word: {e}")
        # Fallback: essayer d'extraire le texte brut
        try:
            if not extracted_text:
                # Lecture brute du fichier comme texte
                with open(word_path, 'r', encoding='utf-8', errors='ignore') as f:
                    extracted_text = f.read()
        except:
            extracted_text = f"Erreur: impossible de lire le fichier Word {word_path.name}"
        
        return extracted_text, None


def is_word_document(file_path: str | Path) -> bool:
    """Vérifie si un fichier est un document Word."""
    suffixes = {'.docx', '.doc'}
    return Path(file_path).suffix.lower() in suffixes


def get_word_extensions() -> list:
    """Retourne les extensions supportées pour les documents Word."""
    return ['.docx', '.doc']


def _resolve_nvidia_api_key(explicit_api_key: str | None = None) -> str:
    if explicit_api_key and explicit_api_key.strip():
        return explicit_api_key.strip()

    for env_key in ("NVIDIA_API_KEY", "NVAPI_KEY", "NIM_API_KEY"):
        value = os.environ.get(env_key, "").strip()
        if value:
            return value

    repo_root = Path(__file__).resolve().parents[2]
    key_file = repo_root / "nvidia.txt"
    if key_file.exists():
        text = key_file.read_text(encoding="utf-8", errors="replace")
        match = re.search(r"(nvapi-[A-Za-z0-9_-]+)", text)
        if match:
            return match.group(1)

    raise RuntimeError(
        "Clé NVIDIA introuvable. Renseigne NVIDIA_API_KEY (ou NVAPI_KEY/NIM_API_KEY), "
        "ou mets une clé nvapi valide dans nvidia.txt."
    )


def _is_heading_style(style_name: str, text: str) -> bool:
    style = (style_name or "").strip().lower()
    raw = (text or "").strip()
    if not raw:
        return False
    if "heading" in style or "titre" in style:
        return True
    if re.match(r"^(lot|article)\s+\d+", raw, flags=re.IGNORECASE):
        return True
    if re.match(r"^lot\s*[-:]", raw, flags=re.IGNORECASE):
        return True
    return False


def _docx_table_to_markdown(table: Any) -> str:
    rows: list[list[str]] = []
    for row in table.rows:
        cells = [re.sub(r"\s+", " ", cell.text or "").strip() for cell in row.cells]
        if any(cells):
            rows.append(cells)
    if not rows:
        return ""

    max_cols = max(len(r) for r in rows)
    normalized = [r + [""] * (max_cols - len(r)) for r in rows]
    header = normalized[0]
    sep = ["---"] * max_cols
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(sep) + " |",
    ]
    for row in normalized[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def _extract_docx_sections_with_markdown_tables(docx_path: Path) -> list[dict[str, str]]:
    try:
        import docx
    except ImportError as exc:
        raise RuntimeError("python-docx est requis pour le parsing Word structuré.") from exc

    doc = docx.Document(docx_path)
    sections: list[dict[str, str]] = []
    current_title = "Préambule"
    current_parts: list[str] = []
    table_index = 0

    def flush_section() -> None:
        nonlocal current_parts
        content = "\n\n".join(part for part in current_parts if part.strip()).strip()
        if content:
            sections.append({"title": current_title, "content": content})
        current_parts = []

    for para in doc.paragraphs:
        text = re.sub(r"\s+", " ", para.text or "").strip()
        if not text:
            continue
        style_name = para.style.name if para.style is not None else ""
        if _is_heading_style(style_name, text):
            flush_section()
            current_title = text
            continue
        current_parts.append(text)

    for table in doc.tables:
        table_index += 1
        markdown_table = _docx_table_to_markdown(table)
        if markdown_table:
            current_parts.append(f"[Tableau {table_index}]\n{markdown_table}")

    flush_section()
    if not sections:
        sections = [{"title": docx_path.stem, "content": extract_docx_text(docx_path)}]
    return sections


def _chunk_sections_by_char_count(
    sections: list[dict[str, str]],
    chunk_size: int = 1000,
    chunk_overlap: int = 120,
) -> list[dict[str, str]]:
    chunks: list[dict[str, str]] = []
    for section in sections:
        title = section.get("title", "").strip() or "Section"
        content = section.get("content", "").strip()
        if not content:
            continue
        start = 0
        while start < len(content):
            end = min(start + chunk_size, len(content))
            body = content[start:end].strip()
            if body:
                chunks.append(
                    {
                        "section_title": title,
                        "text": f"Section: {title}\n{body}",
                    }
                )
            if end >= len(content):
                break
            start = max(0, end - chunk_overlap)
    return chunks


def _embed_passages(api_key: str, passages: list[str]) -> list[list[float]]:
    if not passages:
        return []
    url = "https://integrate.api.nvidia.com/v1/embeddings"
    payload = {
        "model": "nvidia/llama-3_2-nemoretriever-300m-embed-v1",
        "input": passages,
        "encoding_format": "float",
        "input_type": "passage",
    }
    response = requests.post(
        url,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=payload,
        timeout=90,
    )
    response.raise_for_status()
    data = response.json().get("data", [])
    ordered = sorted(data, key=lambda item: item.get("index", 0))
    return [item.get("embedding", []) for item in ordered]


def _embed_query(api_key: str, query: str) -> list[float]:
    vectors = _embed_passages(api_key, [query])
    return vectors[0] if vectors else []


def _cosine_similarity(vec_a: list[float], vec_b: list[float]) -> float:
    if not vec_a or not vec_b or len(vec_a) != len(vec_b):
        return -1.0
    dot = sum(a * b for a, b in zip(vec_a, vec_b))
    norm_a = math.sqrt(sum(a * a for a in vec_a))
    norm_b = math.sqrt(sum(b * b for b in vec_b))
    if norm_a == 0 or norm_b == 0:
        return -1.0
    return dot / (norm_a * norm_b)


def _retrieve_relevant_chunks(
    api_key: str,
    chunks: list[dict[str, str]],
    top_k: int = 8,
) -> list[dict[str, str]]:
    if not chunks:
        return []
    chunk_vectors = _embed_passages(api_key, [chunk["text"] for chunk in chunks])
    queries = [
        "Identifier les intervenants du projet et leurs rôles.",
        "Lister les lots techniques de type Lot 01 - Gros Œuvre, Lot 03 - Fluides, etc.",
        "Extraire toutes les puissances électriques mentionnées (kVA, kW, W).",
        "Identifier les pénalités de retard et leurs conditions d'application.",
    ]

    ranked: list[tuple[float, int]] = []
    for query in queries:
        query_vector = _embed_query(api_key, query)
        for idx, chunk_vector in enumerate(chunk_vectors):
            score = _cosine_similarity(query_vector, chunk_vector)
            ranked.append((score, idx))

    seen: set[int] = set()
    selected: list[dict[str, str]] = []
    for _, idx in sorted(ranked, key=lambda item: item[0], reverse=True):
        if idx in seen:
            continue
        seen.add(idx)
        selected.append(chunks[idx])
        if len(selected) >= top_k:
            break
    return selected


def _run_mistral_clause_analysis(api_key: str, context: str, source_name: str) -> str:
    client = OpenAI(base_url="https://integrate.api.nvidia.com/v1", api_key=api_key)
    system_prompt = (
        "Tu es un ingénieur en bureau d'études (BET). "
        "Ta réponse doit se baser EXCLUSIVEMENT sur le texte fourni. "
        "Si une information (ex: puissance kVA) n'est pas écrite, réponds 'Non mentionné'."
    )
    user_prompt = (
        f"Document source: {source_name}\n\n"
        "Analyse le contexte ci-dessous et prépare une extraction technique:\n"
        "- intervenants\n"
        "- lots techniques\n"
        "- puissances électriques mentionnées\n"
        "- pénalités de retard\n\n"
        "Contexte:\n"
        f"{context}"
    )
    completion = client.chat.completions.create(
        model="mistralai/mistral-large-3-675b-instruct-2512",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.15,
        top_p=1.0,
        max_tokens=2048,
        stream=False,
    )
    if not completion.choices:
        return ""
    return completion.choices[0].message.content or ""


def _run_qwen_structuring(api_key: str, analysis_text: str, source_name: str) -> dict[str, Any]:
    client = OpenAI(base_url="https://integrate.api.nvidia.com/v1", api_key=api_key)
    schema_prompt = (
        "Convertis l'analyse en JSON strict (sans markdown) avec ce schéma:\n"
        "{\n"
        '  "document": "<nom_document>",\n'
        '  "intervenants": [{"nom": "...", "role": "...", "source": "..."}],\n'
        '  "lots_techniques": [{"lot": "...", "description": "...", "source": "..."}],\n'
        '  "puissances_electriques": [{"valeur": "...", "unite": "...", "contexte": "...", "source": "..."}],\n'
        '  "penalites_retard": [{"clause": "...", "valeur": "...", "conditions": "...", "source": "..."}]\n'
        "}\n"
        "Si non trouvé pour une section: renvoyer un tableau vide.\n"
        "Ne renvoie rien d'autre que le JSON."
    )
    completion = client.chat.completions.create(
        model="qwen/qwen3-coder-480b-a35b-instruct",
        messages=[
            {"role": "system", "content": "Tu structures des extractions techniques en JSON exploitable."},
            {"role": "user", "content": f"{schema_prompt}\n\nDocument: {source_name}\n\nAnalyse:\n{analysis_text}"},
        ],
        temperature=0.2,
        top_p=0.8,
        max_tokens=4096,
        stream=False,
    )
    raw = completion.choices[0].message.content if completion.choices else "{}"
    raw = (raw or "{}").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"JSON Qwen invalide: {exc}") from exc
    if not isinstance(parsed, dict):
        raise RuntimeError("Le JSON structuré renvoyé par Qwen doit être un objet.")
    parsed.setdefault("document", source_name)
    for key in ("intervenants", "lots_techniques", "puissances_electriques", "penalites_retard"):
        value = parsed.get(key)
        parsed[key] = value if isinstance(value, list) else []
    return parsed


def save_to_local(data: Any, filename: str | Path) -> Path:
    output_path = Path(filename)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()
    if suffix == ".json":
        output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return output_path
    if suffix == ".csv":
        rows: list[dict[str, Any]] = []
        if isinstance(data, dict):
            base_doc = data.get("document", "")
            for section in ("intervenants", "lots_techniques", "puissances_electriques", "penalites_retard"):
                values = data.get(section, [])
                if isinstance(values, list):
                    for item in values:
                        if isinstance(item, dict):
                            row = OrderedDict()
                            row["document"] = base_doc
                            row["categorie"] = section
                            for key, value in item.items():
                                row[str(key)] = value
                            rows.append(row)
        if not rows:
            rows = [OrderedDict({"document": "", "categorie": "", "valeur": ""})]
        headers = sorted({key for row in rows for key in row.keys()})
        with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
        return output_path
    raise ValueError("Extension non supportée pour save_to_local. Utilise .json ou .csv.")


def run_word_nim_extraction(
    word_path: str | Path,
    output_dir: str | Path,
    api_key: str | None = None,
) -> dict[str, str]:
    source_path = Path(word_path)
    if source_path.suffix.lower() != ".docx":
        raise RuntimeError("Le pipeline NIM Word structuré est actuellement disponible pour les fichiers .docx.")

    resolved_api_key = _resolve_nvidia_api_key(api_key)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    sections = _extract_docx_sections_with_markdown_tables(source_path)
    chunks = _chunk_sections_by_char_count(sections, chunk_size=1000, chunk_overlap=120)
    selected_chunks = _retrieve_relevant_chunks(resolved_api_key, chunks, top_k=8)
    context = "\n\n---\n\n".join(chunk["text"] for chunk in selected_chunks)
    if not context.strip():
        context = "\n\n".join(chunk["text"] for chunk in chunks[:8])

    analysis = _run_mistral_clause_analysis(resolved_api_key, context, source_path.name)
    structured = _run_qwen_structuring(resolved_api_key, analysis, source_path.name)

    stem = source_path.stem
    json_path = save_to_local(structured, out_dir / f"{stem}_extraction_nim.json")
    csv_path = save_to_local(structured, out_dir / f"{stem}_extraction_nim.csv")
    context_path = save_to_local({"document": source_path.name, "context": context}, out_dir / f"{stem}_rag_context.json")

    return {
        "json_path": str(json_path),
        "csv_path": str(csv_path),
        "context_path": str(context_path),
    }
