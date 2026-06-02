"""
Synthèse de documents Word (.docx) via Gemini API.
Même workflow que le traitement PDF mais adapté pour Word.
Provider: Gemini (Google Generative AI)
Version optimisée pour la performance
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
from pathlib import Path
from typing import Any

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from app_core.pipeline import create_run_dir, SynthesisOutputs


# Regex précompilées pour la performance
_HEADING_PATTERN = re.compile(r"^(lot|article|chapitre|section|partie)\s+\d", re.IGNORECASE)
_SECTION_PATTERN = re.compile(r"lot\s*\d|chapitre\s*\d|article\s*\d|section\s*\d")
_WHITESPACE_PATTERN = re.compile(r"\s+")
_MULTI_NEWLINE_PATTERN = re.compile(r"\n{3,}")
_MULTI_SPACE_PATTERN = re.compile(r" {2,}")

# Cache pour les documents extraits (évite de réextraire le même document)
_DOCUMENT_CACHE: dict[str, tuple[list[dict[str, Any]], str]] = {}
_CACHE_MAX_SIZE = 10  # Maximum 10 documents en cache


def _is_heading(docx_path: Path, para: Paragraph) -> bool:
    """Détecte si un paragraphe est un titre de section Word (version optimisée)."""
    # Vérification rapide du texte
    text = para.text
    if not text or not text.strip():
        return False
    
    text_stripped = text.strip()
    
    # Vérification du style (une seule fois)
    if para.style:
        style_name = para.style.name
        if style_name:
            style_lower = style_name.strip().lower()
            if "heading" in style_lower or "titre" in style_lower:
                return True
    
    # Vérification du gras (optimisée)
    if para.runs:
        # Vérification rapide: si le premier run est en gras et qu'il y a du texte
        first_run = para.runs[0]
        if first_run.text.strip() and first_run.bold:
            # Vérification rapide de la longueur
            if len(text_stripped) < 120:
                return True
    
    # Vérification par regex précompilée
    return bool(_HEADING_PATTERN.match(text_stripped))


def _table_to_markdown(table: Table) -> str:
    """Convertit un tableau Word en Markdown (version optimisée)."""
    rows_data = []
    has_content = False
    
    for row in table.rows:
        row_cells = []
        for cell in row.cells:
            # Extraction optimisée du texte de la cellule
            cell_text_parts = []
            for para in cell.paragraphs:
                text = para.text
                if text:
                    stripped = text.strip()
                    if stripped:
                        cell_text_parts.append(stripped)
            
            if cell_text_parts:
                # Joindre les parties et normaliser les espaces avec regex précompilée
                cell_text = " ".join(cell_text_parts)
                cell_text = _WHITESPACE_PATTERN.sub(" ", cell_text).strip()
                row_cells.append(cell_text)
                has_content = True
            else:
                row_cells.append("")
        
        if has_content or any(row_cells):
            rows_data.append(row_cells)
    
    if not rows_data:
        return ""
    
    # Trouver le nombre maximum de colonnes
    max_cols = 0
    for row in rows_data:
        if len(row) > max_cols:
            max_cols = len(row)
    
    # Construire le markdown de manière optimisée
    lines = []
    
    # Première ligne (en-tête)
    if rows_data:
        first_row = rows_data[0]
        padded_row = first_row + [""] * (max_cols - len(first_row))
        lines.append("| " + " | ".join(padded_row) + " |")
    
    # Séparateur
    if max_cols > 0:
        lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    
    # Autres lignes
    for row in rows_data[1:]:
        padded_row = row + [""] * (max_cols - len(row))
        lines.append("| " + " | ".join(padded_row) + " |")
    
    return "\n".join(lines)


def _extract_docx_sections(docx_path: Path) -> list[dict[str, Any]]:
    """
    Extrait le document Word en sections structurées avec leurs titres.
    Chaque élément: {'title': str, 'content': str, 'tables': list[str]}
    Version ultra-optimisée pour la performance avec cache.
    """
    if not docx_path.exists():
        return []

    # Vérifier le cache d'abord
    cache_key = str(docx_path.absolute())
    if cache_key in _DOCUMENT_CACHE:
        sections, _ = _DOCUMENT_CACHE[cache_key]
        return sections

    doc = Document(str(docx_path))
    sections: list[dict[str, Any]] = []
    current_title = "__preambule__"
    current_parts: list[str] = []
    current_tables: list[str] = []
    table_count = 0
    body_elements = doc.element.body  # Cache local pour performance
    
    # Optimisation: utiliser des références locales pour éviter les accès répétés
    sections_append = sections.append
    current_parts_append = current_parts.append
    current_tables_append = current_tables.append

    def flush() -> None:
        nonlocal current_title, current_parts, current_tables
        # Construction ultra-optimisée du contenu
        if current_parts or current_tables:
            # Filtrer les parties vides efficacement avec list comprehension
            filtered_parts = [part for part in current_parts if part.strip()]
            
            content = "\n".join(filtered_parts) if filtered_parts else ""
            
            sections_append({
                "title": current_title,
                "content": content,
                "tables": list(current_tables) if current_tables else [],
            })
        current_parts.clear()
        current_tables.clear()

    for element in body_elements:
        tag = element.tag
        
        if tag.endswith("p"):
            para = Paragraph(element, doc)
            text = para.text
            if text:
                stripped = text.strip()
                if stripped:
                    if _is_heading(docx_path, para):
                        flush()
                        current_title = stripped
                        continue
                    current_parts_append(stripped)
                else:
                    current_parts_append("")
        
        elif tag.endswith("tbl"):
            table = Table(element, doc)
            md_table = _table_to_markdown(table)
            if md_table:
                table_count += 1
                current_tables_append(md_table)
                current_parts_append(f"\n[TABLEAU-{table_count}]\n{md_table}\n[/TABLEAU-{table_count}]\n")

    flush()
    
    # Mettre en cache (avec texte vide pour l'instant, sera rempli par _extract_docx_text)
    if len(_DOCUMENT_CACHE) >= _CACHE_MAX_SIZE:
        # Supprimer le plus ancien élément (FIFO)
        oldest_key = next(iter(_DOCUMENT_CACHE))
        del _DOCUMENT_CACHE[oldest_key]
    
    _DOCUMENT_CACHE[cache_key] = (sections, "")
    
    return sections


def _extract_docx_text(docx_path: Path) -> str:
    """
    Extrait le texte complet d'un .docx (paragraphes + tableaux Markdown dans l'ordre).
    Version ultra-optimisée pour la performance avec cache.
    """
    # Vérifier le cache d'abord
    cache_key = str(docx_path.absolute())
    if cache_key in _DOCUMENT_CACHE:
        sections, cached_text = _DOCUMENT_CACHE[cache_key]
        if cached_text:  # Si le texte est déjà en cache, le retourner
            return cached_text
        # Sinon, continuer pour générer le texte
    
    sections = _extract_docx_sections(docx_path)
    if not sections:
        return ""
    
    parts = []
    # Pré-allocation approximative pour réduire les réallocations
    parts_append = parts.append
    
    for sec in sections:
        title = sec["title"]
        if title != "__preambule__":
            parts_append(f"\n## {title}\n")
        
        tables = sec.get("tables")
        if tables:
            # Utiliser join pour les tables au lieu d'append multiples
            tables_text = "\n".join(tables)
            parts_append(f"\n{tables_text}\n")
        
        content = sec["content"]
        if content.strip():
            parts_append(content)
    
    text = "\n".join(parts)
    # Utilisation de regex précompilées avec une seule passe
    text = _MULTI_NEWLINE_PATTERN.sub("\n\n", text)
    text = _MULTI_SPACE_PATTERN.sub(" ", text)
    text = text.strip()
    
    # Mettre à jour le cache avec le texte généré
    if cache_key in _DOCUMENT_CACHE:
        _DOCUMENT_CACHE[cache_key] = (sections, text)
    
    return text


def _find_key_section_boundaries(sections: list[dict[str, Any]]) -> tuple[int, int]:
    """
    Trouve les indices des sections clés du document (lots techniques, corps d'état).
    Retourne (idx_premiere_section_lot, idx_derniere_section) ou (-1, -1) si non trouvées.
    Version ultra-optimisée avec regex précompilée et recherche rapide.
    """
    if not sections:
        return -1, -1
    
    premiere_section_lot = -1
    derniere_section = len(sections) - 1
    
    # Recherche optimisée: parcourir seulement jusqu'à trouver la première section
    for i, sec in enumerate(sections):
        title = sec["title"]
        # Vérification rapide: si le titre contient "lot" ou "chapitre"
        if "lot" in title.lower() or "chapitre" in title.lower():
            if _SECTION_PATTERN.search(title.lower()):
                premiere_section_lot = i
                break
    
    return premiere_section_lot, derniere_section


def _limit_text_smart(
    text: str,
    max_chars: int,
    sections: list[dict[str, Any]],
) -> str:
    """
    Stratégie de troncature intelligente pour documents Word volumineux.
    - Prend le DÉBUT (avec objet/projet toujours au début)
    - Force la couverture des sections FLUIDES si présentes
    - Prend la FIN (conclusions, clauses finales)
    Version ultra-optimisée pour la performance.
    """
    text_len = len(text)
    if text_len <= max_chars:
        return text

    chunk_size = max_chars // 3
    if chunk_size <= 0:
        chunk_size = 1

    # Extraction ultra-optimisée du début et de la fin
    debut = text[:chunk_size]
    fin = text[-chunk_size:] if chunk_size > 0 else ""

    premiere_section_lot, derniere_section = _find_key_section_boundaries(sections)

    idx_start_middle = chunk_size
    idx_end_middle = text_len - chunk_size

    if premiere_section_lot != -1:
        # Recherche ultra-optimisée: utiliser find avec slice limitée
        # On limite la recherche à la partie pertinente du texte
        search_start = max(0, chunk_size - 1000)
        search_end = min(text_len, chunk_size + 10000)
        search_text = text[search_start:search_end].lower()
        lot_pos = search_text.find("lot")
        if lot_pos != -1:
            # Ajuster la position pour le texte complet
            lot_pos += search_start
            if lot_pos > chunk_size and lot_pos < idx_end_middle:
                idx_start_middle = max(idx_start_middle, lot_pos - 2000)

    middle = text[idx_start_middle:idx_end_middle]

    # Construction ultra-optimisée du résultat
    result_parts = []
    result_parts_append = result_parts.append
    
    result_parts_append(debut)
    
    middle_stripped = middle.strip()
    if middle_stripped:
        start_page = idx_start_middle // 2000
        end_page = idx_end_middle // 2000
        result_parts_append(f"\n\n[...SECTION CRITIQUE (pages {start_page} a {end_page})...]\n\n")
        result_parts_append(middle_stripped)
    
    fin_stripped = fin.strip()
    if fin_stripped:
        result_parts_append(f"\n\n[...FIN DU DOCUMENT...]\n\n{fin_stripped}")

    return "".join(result_parts)


def _call_gemini(api_key: str, model_name: str, prompt: str) -> str:
    """Appelle l'API Gemini avec urllib (version optimisée)."""
    url = f"https://generativelanguage.googleapis.com/v1beta/{model_name}:generateContent"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}]
    }
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url=f"{url}?key={api_key}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=180) as response:  # Augmenté de 90 à 180 secondes
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace") if hasattr(exc, "read") else str(exc)
        raise RuntimeError(message) from exc

    parsed = json.loads(raw)
    candidates = parsed.get("candidates") or []
    if not candidates:
        raise RuntimeError("Aucune réponse Gemini.")
    content = (candidates[0] or {}).get("content") or {}
    parts = content.get("parts") or []
    text_parts = [part.get("text", "") for part in parts if isinstance(part, dict)]
    return "\n".join(text_parts).strip()


def _call_gemini_with_retries(api_key: str, model_name: str, prompt: str, max_retries: int = 2) -> str:
    """Appelle Gemini avec retry sur erreur 503 (version optimisée avec moins de retries)."""
    for attempt in range(max_retries):
        try:
            return _call_gemini(api_key, model_name, prompt)
        except RuntimeError as exc:
            message = str(exc)
            is_overload = "503" in message or "unavailable" in message.lower() or "overload" in message.lower()
            if is_overload and attempt < max_retries - 1:
                wait_time = (attempt + 1) * 1  # Temps d'attente réduit
                time.sleep(wait_time)
                continue
            raise
    return "Échec après plusieurs tentatives."


# Cache pour les modèles Gemini (évite de récupérer la liste à chaque appel)
_GEMINI_MODEL_CACHE = {}

def _get_gemini_model(api_key: str) -> str:
    """Récupère le meilleur modèle Gemini disponible (version optimisée avec cache)."""
    # Vérifier le cache d'abord
    cache_key = api_key[:20] if api_key else "default"
    if cache_key in _GEMINI_MODEL_CACHE:
        return _GEMINI_MODEL_CACHE[cache_key]
    
    url = "https://generativelanguage.googleapis.com/v1beta/models"
    request = urllib.request.Request(
        url=f"{url}?key={api_key}",
        headers={"Content-Type": "application/json"},
        method="GET",
    )
    
    try:
        with urllib.request.urlopen(request, timeout=30) as response:  # Timeout réduit
            raw = response.read().decode("utf-8", errors="replace")
    except Exception:
        # Retourner le modèle par défaut et le mettre en cache
        default_model = "models/gemini-1.5-flash"
        _GEMINI_MODEL_CACHE[cache_key] = default_model
        return default_model

    try:
        parsed = json.loads(raw)
        models = parsed.get("models") or []
        
        # Filtrage optimisé
        candidates = []
        for m in models:
            if isinstance(m, dict):
                name = m.get("name")
                if isinstance(name, str) and name.startswith("models/"):
                    candidates.append(name)
        
        preferred_order = (
            "models/gemini-2.5-flash",
            "models/gemini-2.0-flash",
            "models/gemini-2.0-flash-lite",
            "models/gemini-1.5-flash",
            "models/gemini-1.5-flash-latest",
            "models/gemini-1.5-pro",
            "models/gemini-1.5-pro-latest",
        )
        
        for pref in preferred_order:
            if pref in candidates:
                _GEMINI_MODEL_CACHE[cache_key] = pref
                return pref
        
        # Si aucun modèle préféré n'est trouvé, prendre le premier disponible
        if candidates:
            selected = candidates[0]
        else:
            selected = "models/gemini-1.5-flash"
        
        _GEMINI_MODEL_CACHE[cache_key] = selected
        return selected
        
    except Exception:
        default_model = "models/gemini-1.5-flash"
        _GEMINI_MODEL_CACHE[cache_key] = default_model
        return default_model


def _analyze_identity(api_key: str, model: str, doc_text: str, doc_name: str, sections: list[dict[str, Any]]) -> str:
    print(f"[ANALYSE-IDENTITE] {doc_name} ({len(doc_text):,} caracteres, {len(sections)} sections)")
    limited = _limit_text_smart(doc_text, 60000, sections)  # Réduit de 80k à 60k
    prompt = f"""Expert DCE. Extrait STRICTEMENT du texte. Pas d'invention.

DOCUMENT: {doc_name}

{limited}

JSON UNIQUEMENT:

{{
  "informations_generales": {{
    "objet_consultation": "Texte exact",
    "situation_geographique": "Texte exact",
    "consistance_projet": "Texte exact",
    "surface_terrain": "Texte exact ou 'Non mentionne'",
    "superficies_globales_shob": "Texte exact ou 'Non mentionne'"
  }},
  "intervenants_principaux": {{
    "maitrise_ouvrage": "Texte exact",
    "maitrise_oeuvre_architectes": "Texte exact",
    "bet_technique_fluides": "Texte exact",
    "bet_structure": "Texte exact",
    "bureau_controle": "Texte exact",
    "opc": "Texte exact ou 'Non mentionne'"
  }},
  "calendrier_validite": {{
    "duree_previsionnelle_travaux": "Texte exact ou 'A definir'",
    "date_limite_remise_offres": "Texte exact ou 'Non mentionne'",
    "delai_validite_offres": "Texte exact ou 'Non mentionne'"
  }},
  "conditions_contractuelles_majeures": {{
    "nature_forme_marche": "Texte exact ou 'Non mentionne'",
    "retenues_garantie_cautions": "Texte exact ou 'Non mentionne'",
    "modalites_delais_paiement": "Texte exact ou 'Non mentionne'",
    "clauses_penalites": "Texte exact ou 'Non mentionne'"
  }}
}}

JSON seulement."""

    try:
        result = _call_gemini_with_retries(api_key, model, prompt)
        print(f"  Identite OK: {len(result):,} chars")
        return result
    except Exception as e:
        print(f"  Erreur identite: {e}")
        return f"Erreur API: {e}"


def _analyze_synthese(api_key: str, model: str, doc_text: str, doc_name: str, sections: list[dict[str, Any]]) -> str:
    print(f"[ANALYSE-SYNTHESE] {doc_name} ({len(doc_text):,} caracteres, {len(sections)} sections)")
    limited = _limit_text_smart(doc_text, 80000, sections)  # Réduit de 100k à 80k
    prompt = f"""Expert CVC. Extrait strictement du texte. Pas d'invention.

DOCUMENT: {doc_name}

{limited}

Extrait pour chaque section:

1. DISPOSITIONS GENERALES
- Objet prestations
- Documents fournir
- Coordination

2. PLOMBERIE SANITAIRE / ECS
- Systemes ECS
- Materiaux canalisation
- Evacuation eaux
- Normes

3. GENIE CLIMATIQUE
- Systeme climatisation
- Puissances frigorifiques
- Reseaux aerauliques
- Traitement air
- Regulation

4. PROTECTION INCENDIE
- Equipements protection
- Systemes desenfumage
- Normes securite

5. EQUIPEMENTS POMPAGE
- Stations relevage
- Traitement eau
- Filtration

6. EXECUTION QUALITE
- Modes operatoires
- Normes mise oeuvre
- Jonctions assemblages

7. CONTROLES ESSAIS
- Epreuves etancheite
- Documents reception

Texte exact seulement. Si absent: "Non mentionne"."""

    try:
        result = _call_gemini_with_retries(api_key, model, prompt)
        print(f"  Synthese OK: {len(result):,} chars")
        return result
    except Exception as e:
        print(f"  Erreur synthese: {e}")
        return f"Erreur API: {e}"


def _process_document_batch(
    api_key: str,
    model: str,
    doc_data: dict[str, dict[str, Any]],
    batch_size: int = 3,
) -> tuple[dict[str, str], dict[str, str]]:
    """
    Traite un lot de documents en parallèle pour optimiser les appels API.
    Version optimisée pour la performance.
    """
    identity_results: dict[str, str] = {}
    synthese_results: dict[str, str] = {}
    
    # Diviser les documents en lots pour éviter la surcharge
    doc_items = list(doc_data.items())
    num_batches = (len(doc_items) + batch_size - 1) // batch_size
    
    print(f"  Traitement par lots: {len(doc_items)} documents en {num_batches} lot(s)")
    
    for batch_idx in range(num_batches):
        start_idx = batch_idx * batch_size
        end_idx = min(start_idx + batch_size, len(doc_items))
        batch_items = doc_items[start_idx:end_idx]
        
        print(f"  Lot {batch_idx + 1}/{num_batches}: {len(batch_items)} document(s)")
        
        # Traiter chaque lot avec un nombre limité de workers
        batch_workers = min(4, len(batch_items) * 2)
        
        with ThreadPoolExecutor(max_workers=batch_workers) as executor:
            id_futures = {}
            syn_futures = {}
            
            for name, d in batch_items:
                if "text" in d and "sections" in d:
                    id_futures[executor.submit(_analyze_identity, api_key, model, d["text"], name, d["sections"])] = name
                    syn_futures[executor.submit(_analyze_synthese, api_key, model, d["text"], name, d["sections"])] = name
            
            # Collecter les résultats du lot actuel
            for fut in as_completed(id_futures):
                name = id_futures[fut]
                try:
                    identity_results[name] = fut.result(timeout=150)  # Timeout réduit à 2.5 minutes
                except Exception as e:
                    identity_results[name] = f"Erreur lors de l'analyse d'identité: {e}"
                    print(f"    Erreur identité pour {name}: {e}")
            
            for fut in as_completed(syn_futures):
                name = syn_futures[fut]
                try:
                    synthese_results[name] = fut.result(timeout=150)  # Timeout réduit à 2.5 minutes
                except Exception as e:
                    synthese_results[name] = f"Erreur lors de l'analyse de synthèse: {e}"
                    print(f"    Erreur synthèse pour {name}: {e}")
    
    return identity_results, synthese_results


def run_synthesis_word(
    docx_paths: list[Path],
    workspace_root: Path,
    api_key: str,
) -> SynthesisOutputs:
    """
    Genere identite + synthese a partir de documents Word (.docx) via Gemini.
    Utilise une extraction structuree par sections + tables Markdown.
    Version ultra-optimisée pour la performance.
    """
    print("=" * 80)
    print("SYNTHESE DCE - DOCUMENTS WORD (Gemini) - Version ultra-optimisée")
    print("=" * 80)

    if not api_key:
        raise ValueError("Une cle API Gemini est requise pour traiter les documents Word.")

    model = _get_gemini_model(api_key)
    print(f"  Modele Gemini utilise: {model}")

    run_dir = create_run_dir(workspace_root, "synthese-word")
    output_dir = run_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # Optimisation: ajuster le nombre de workers en fonction du nombre de documents
    num_docs = len(docx_paths)
    extract_workers = min(4, max(2, num_docs))  # Entre 2 et 4 workers
    
    print(f"\nETAPE 1/3: Extraction de {num_docs} document(s) avec {extract_workers} workers")
    doc_data: dict[str, dict[str, Any]] = {}
    
    with ThreadPoolExecutor(max_workers=extract_workers) as executor:
        fut_to_path = {executor.submit(_extract_docx_sections, p): p for p in docx_paths}
        for fut in as_completed(fut_to_path):
            p = fut_to_path[fut]
            try:
                sections = fut.result()
                full_text = _extract_docx_text(p)
                doc_data[p.name] = {
                    "sections": sections,
                    "text": full_text,
                }
                print(f"  {p.name}: {len(full_text):,} caracteres, {len(sections)} sections")
            except Exception as e:
                doc_data[p.name] = {"sections": [], "text": f"[Erreur: {e}]"}
                print(f"  Erreur {p.name}: {e}")

    print(f"\nETAPE 2/3: Analyse de {len(doc_data)} document(s) avec traitement par lots")
    identity_results: dict[str, str] = {}
    synthese_results: dict[str, str] = {}
    
    # Utiliser le traitement par lots optimisé
    identity_results, synthese_results = _process_document_batch(
        api_key=api_key,
        model=model,
        doc_data=doc_data,
        batch_size=3,  # Traiter 3 documents par lot maximum
    )

    print(f"\nETAPE 3/3: Generation documents finaux")

    # Optimisation: construction ultra-efficace des textes combinés avec limitation de taille
    id_parts = []
    syn_parts = []
    
    id_parts_append = id_parts.append
    syn_parts_append = syn_parts.append
    
    # Limiter la taille totale pour éviter les prompts trop longs
    total_id_chars = 0
    max_total_chars = 120000  # Limite de 120k caractères pour les prompts combinés
    
    for name, result in identity_results.items():
        if total_id_chars + len(result) > max_total_chars:
            # Tronquer le résultat si nécessaire
            remaining = max_total_chars - total_id_chars
            if remaining > 1000:  # Garder au moins 1000 caractères
                truncated = result[:remaining] + "\n[...tronque...]"
                id_parts_append(f"### {name}\n{truncated}")
                total_id_chars += len(truncated)
            break
        else:
            id_parts_append(f"### {name}\n{result}")
            total_id_chars += len(result)
    
    total_syn_chars = 0
    for name, result in synthese_results.items():
        if total_syn_chars + len(result) > max_total_chars:
            remaining = max_total_chars - total_syn_chars
            if remaining > 1000:
                truncated = result[:remaining] + "\n[...tronque...]"
                syn_parts_append(f"### {name}\n{truncated}")
                total_syn_chars += len(truncated)
            break
        else:
            syn_parts_append(f"### {name}\n{result}")
            total_syn_chars += len(result)
    
    combined_id = "\n\n".join(id_parts)
    combined_syn = "\n\n".join(syn_parts)
    
    print(f"  Identite combinee: {len(combined_id):,} chars")
    print(f"  Synthese combinee: {len(combined_syn):,} chars")

    id_prompt = f"""Expert DCE. Genere document "IDENTITE DU PROJET" en Markdown.

ANALYSES:
{combined_id}

Document Markdown:

# IDENTITE DU PROJET
## 1. Informations Generales
## 2. Intervenants Principaux
## 3. Calendrier et Validite
## 4. Conditions Contractuelles
## 5. Points Cles

RÈGLE ABSOLUE ET IMPÉRATIVE : Si une information, un intervenant ou un élément n'est pas clairement spécifié, IGNORE ET SUPPRIME TOTALEMENT L'ÉLÉMENT. Ne mentionne jamais 'Non spécifié', 'Non mentionné', 'N/A' ou 'Inconnu'. NE RÉDIGE JAMAIS de section ou de partie intitulée 'Points de Vigilance' ou 'Points à confirmer'. Texte exact des analyses seulement."""

    syn_prompt = f"""Expert CVC. Genere document "SYNTHESE TECHNIQUE DETAILLEE" en Markdown.

ANALYSES:
{combined_syn}

Document Markdown:

# SYNTHESE TECHNIQUE DETAILLEE
## 1. Dispositions Generales
## 2. Plomberie Sanitaire / ECS
## 3. Genie Climatique
## 4. Protection Incendie
## 5. Equipements de Pompage
## 6. Execution et Qualite
## 7. Controles et Receptions

RÈGLE ABSOLUE ET IMPÉRATIVE : Si un lot ou un équipement n'est pas spécifié, IGNORE ET SUPPRIME TOTALEMENT LA SECTION. Ne mentionne jamais 'Non spécifié', 'Non mentionné', 'N/A' ou 'Inconnu'. NE RÉDIGE JAMAIS de section ou de partie intitulée 'Points de Vigilance'. Texte exact des analyses seulement."""

    identite = _call_gemini_with_retries(api_key, model, id_prompt)
    print(f"  Identite finale: {len(identite):,} chars")

    synthese = _call_gemini_with_retries(api_key, model, syn_prompt)
    print(f"  Synthese finale: {len(synthese):,} chars")

    print("\n" + "=" * 80)
    print("TERMINE")
    print("=" * 80)

    identite_path = output_dir / "identite-projet.md"
    synthese_path = output_dir / "synthese-detaillee.md"
    identite_path.write_text(identite, encoding="utf-8")
    synthese_path.write_text(synthese, encoding="utf-8")

    return SynthesisOutputs(
        output_dir=output_dir,
        identite_path=identite_path,
        synthese_path=synthese_path,
        stdout="",
        stderr="",
    )
