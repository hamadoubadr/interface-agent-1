from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .models import ExtractedItem, ExtractionResult, SheetCandidate, SheetExtractionSummary, TableDetectionResult
from .normalizer import (
    CHAPTER_HINTS,
    COMMENT_KEYWORDS,
    LOT_KEYWORDS,
    SUBTOTAL_KEYWORDS,
    TOTAL_KEYWORDS,
    normalize_for_match,
    normalize_text,
    normalize_unit,
    parse_number,
)
from .reader import read_workbook
from .sheet_detector import detect_relevant_sheets
from .table_detector import detect_main_table

LOGGER = logging.getLogger(__name__)


@dataclass(slots=True)
class PendingItemParts:
    base_code: str | None = None
    base_designation: str | None = None
    detail_code: str | None = None
    detail_parts: list[str] = field(default_factory=list)
    source_rows: list[tuple[int, list[Any]]] = field(default_factory=list)

    def clear_details(self) -> None:
        self.detail_code = None
        self.detail_parts = []
        self.source_rows = []

    def reset(self) -> None:
        self.base_code = None
        self.base_designation = None
        self.clear_details()


def _looks_like_hierarchical_code(code: str) -> bool:
    return bool(re.fullmatch(r"[A-Z0-9]+(?:[.\-/][A-Z0-9]+)+", code, re.IGNORECASE))


def _is_upperish(text: str) -> bool:
    letters = [ch for ch in text if ch.isalpha()]
    if not letters:
        return False
    upper = sum(ch.isupper() for ch in letters)
    return upper / len(letters) >= 0.75


def _looks_like_section_label(text: str) -> bool:
    label = normalize_text(text)
    if not label or len(label) > 90:
        return False

    if re.match(r"^[A-Z0-9]{1,4}\s*[/:-]\s*[A-Za-z]", label):
        return True

    words = [word for word in re.split(r"\s+", label) if word]
    if not words:
        return False

    uppercase_like_tokens = 0
    for word in words:
        letters = [ch for ch in word if ch.isalpha()]
        if not letters:
            continue
        upper_ratio = sum(ch.isupper() for ch in letters) / len(letters)
        if upper_ratio >= 0.6:
            uppercase_like_tokens += 1

    return len(words) <= 8 and uppercase_like_tokens >= max(1, len(words) // 2)


def _is_explicit_major_heading(text: str) -> bool:
    label = normalize_text(text)
    lowered = label.lower()
    if re.match(r"^[A-Z0-9]{1,4}\s*[/:-]\s*[A-Za-z]", label):
        return True
    return any(keyword in lowered for keyword in LOT_KEYWORDS + CHAPTER_HINTS)


def _clean_heading_designation(designation: str, line_type: str) -> str:
    label = normalize_text(designation)
    if not label:
        return label

    if line_type == "chapter":
        label = re.sub(r"^([A-Za-z0-9]{1,4})\s*/\s*", r"\1/ ", label)
        return label.strip()

    if line_type == "subtotal":
        match = re.search(r"SOUS\s+TOTAL\s+([A-Z0-9]+)", label, flags=re.IGNORECASE)
        if match:
            return f"Sous-total {match.group(1).upper()}"
        return "Sous-total"

    if line_type == "total":
        normalized = normalize_for_match(label)
        compact = re.sub(r"[^a-z0-9]+", "", normalized)
        if "ttc" in compact:
            return "Total lot general TTC"
        if "ht" in compact:
            return "Total lot general HT"
        return "Total"

    return label


def _get_value(row: list[Any], column_map: dict[str, int], key: str) -> Any:
    index = column_map.get(key)
    if index is None or index >= len(row):
        return None
    return row[index]


def _raw_row_dict(row: list[Any], headers: dict[int, str]) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for index, value in enumerate(row):
        header = headers.get(index)
        key = header if header else f"col_{index + 1}"
        payload[key] = value
    return payload


def _raw_rows_dict(rows: list[tuple[int, list[Any]]], headers: dict[int, str]) -> dict[str, Any]:
    if len(rows) == 1:
        return _raw_row_dict(rows[0][1], headers)

    payload: dict[str, Any] = {"merged_rows": [row_index for row_index, _ in rows]}
    for row_index, row_values in rows:
        payload[f"row_{row_index}"] = _raw_row_dict(row_values, headers)
    return payload


def _collect_text_candidates(row: list[Any], column_map: dict[str, int]) -> list[str]:
    excluded_indexes = {index for index in column_map.values()}
    candidates: list[str] = []
    for index, value in enumerate(row):
        if index in excluded_indexes:
            continue
        text = normalize_text(value)
        if text:
            candidates.append(text)
    return candidates


def _resolve_designation(row: list[Any], table: TableDetectionResult) -> str:
    value = _get_value(row, table.column_map, "designation")
    designation = normalize_text(value)
    if designation:
        return designation

    text_candidates = _collect_text_candidates(row, table.column_map)
    text_candidates = [text for text in text_candidates if len(text) > 2]
    return " | ".join(text_candidates[:3]).strip()


def _is_measure_row(unite: str | None, quantite: float | None, prix_unitaire: float | None, montant: float | None) -> bool:
    return any(value is not None for value in (unite, quantite, prix_unitaire, montant))


def _is_short_detail_code(code_article: str | None) -> bool:
    if not code_article:
        return False
    compact = code_article.strip()
    return bool(re.fullmatch(r"[A-Za-z0-9]{1,3}", compact))


def _is_parent_article_code(code_article: str | None) -> bool:
    if not code_article:
        return False
    compact = code_article.strip()
    return bool(re.fullmatch(r"[A-Za-z]{0,4}-\d{1,4}(?:[.\-/]\d{1,4})*", compact))


def _combine_code(base_code: str | None, detail_code: str | None, current_code: str | None) -> str | None:
    if current_code and base_code and current_code != base_code:
        return current_code
    if detail_code and base_code:
        return f"{base_code}.{detail_code}"
    return current_code or base_code or detail_code


def _combine_designation(
    base_designation: str | None,
    detail_parts: list[str],
    current_designation: str | None,
) -> str:
    parts: list[str] = []
    for value in [base_designation, *detail_parts, current_designation]:
        text = normalize_text(value)
        if not text:
            continue
        if parts and text == parts[-1]:
            continue
        parts.append(text)
    return " | ".join(parts).strip()


def _is_generic_measure_phrase(text: str) -> bool:
    normalized = normalize_for_match(text)
    normalized = normalized.replace(":", "").strip()
    generic_prefixes = (
        "l ensemble",
        "ensemble",
        "l unite",
        "unite",
        "le metre lineaire",
        "metre lineaire",
        "ml",
        "u",
        "e",
    )
    return normalized in generic_prefixes


def _clean_designation_parts(
    designation: str,
    unite: str | None,
    quantite: float | None,
) -> str:
    if not designation:
        return designation

    parts = [normalize_text(part) for part in designation.split("|")]
    parts = [part for part in parts if part]
    has_measure = unite is not None or quantite is not None
    if has_measure:
        parts = [part for part in parts if not _is_generic_measure_phrase(part)]

    cleaned: list[str] = []
    for part in parts:
        if cleaned and part == cleaned[-1]:
            continue
        cleaned.append(part)
    return " | ".join(cleaned).strip()


def _pick_heading_seed(designation: str | None, code_article: str | None, raw_text: str) -> str:
    candidates = [normalize_text(value) for value in (designation, code_article, raw_text)]
    candidates = [value for value in candidates if value]
    if not candidates:
        return ""
    return max(candidates, key=len)


def classify_line(
    designation: str,
    code_article: str | None,
    unite: str | None,
    quantite: float | None,
    prix_unitaire: float | None,
    montant: float | None,
    raw_text: str,
) -> str:
    joined = normalize_text(f"{designation} {code_article or ''} {raw_text}").strip()
    lowered = joined.lower()
    has_numeric = any(value is not None for value in (quantite, prix_unitaire, montant))
    has_code = bool(code_article)
    has_unit = bool(unite)

    if not joined:
        return "unknown"
    if any(keyword in lowered for keyword in SUBTOTAL_KEYWORDS):
        return "subtotal"
    if any(keyword in lowered for keyword in TOTAL_KEYWORDS):
        return "total"
    if any(keyword in lowered for keyword in COMMENT_KEYWORDS) and not has_numeric and not has_code:
        return "comment"
    if designation and (has_numeric or has_unit):
        return "item"
    if designation and has_code and not has_numeric and not has_unit:
        if _looks_like_hierarchical_code(code_article) or _is_parent_article_code(code_article) or _is_short_detail_code(code_article):
            return "subchapter"
        return "chapter"
    if designation:
        looks_like_chapter = (
            _is_upperish(designation)
            or _looks_like_section_label(designation)
            or any(keyword in lowered for keyword in LOT_KEYWORDS + CHAPTER_HINTS)
            or bool(re.fullmatch(r"[A-Z0-9 ./_\-]{6,}", designation.strip()))
        )
        if code_article and _looks_like_hierarchical_code(code_article):
            return "subchapter"
        if looks_like_chapter:
            return "chapter"
        return "comment"
    return "unknown"


def _update_context(
    line_type: str,
    designation: str,
    code_article: str | None,
    context: dict[str, str | None],
) -> None:
    label = designation.strip()
    lowered = label.lower()

    if line_type == "chapter":
        if any(keyword in lowered for keyword in LOT_KEYWORDS):
            context["lot"] = label
        else:
            # Chapter lines update the context for the following item rows.
            context["chapter"] = label
            context["subchapter"] = None
    elif line_type == "subchapter":
        if any(keyword in lowered for keyword in LOT_KEYWORDS):
            context["lot"] = label
        elif code_article and _looks_like_hierarchical_code(code_article):
            context["subchapter"] = label
        else:
            context["chapter"] = context.get("chapter") or label
            context["subchapter"] = label


def _compute_row_confidence(
    line_type: str,
    designation: str,
    code_article: str | None,
    unite: str | None,
    quantite: float | None,
    prix_unitaire: float | None,
    montant: float | None,
    table_score: float,
) -> float:
    confidence = 0.15 + min(table_score, 1.0) * 0.35
    if designation:
        confidence += 0.2
    if code_article:
        confidence += 0.08
    if unite:
        confidence += 0.08
    if quantite is not None:
        confidence += 0.1
    if prix_unitaire is not None:
        confidence += 0.07
    if montant is not None:
        confidence += 0.07
    if line_type == "item":
        confidence += 0.1
    elif line_type in {"chapter", "subchapter"}:
        confidence += 0.03
    elif line_type in {"unknown", "comment"}:
        confidence -= 0.08
    return max(min(round(confidence, 3), 0.99), 0.0)


def extract_table_items(
    table: TableDetectionResult,
    relevant_score: float,
) -> tuple[list[ExtractedItem], SheetExtractionSummary]:
    items: list[ExtractedItem] = []
    warnings = list(table.warnings)
    context: dict[str, str | None] = {"lot": None, "chapter": None, "subchapter": None}
    pending = PendingItemParts()

    for table_row in table.table_rows:
        row = table_row.values
        code_article = normalize_text(_get_value(row, table.column_map, "code_article")) or None
        designation = _resolve_designation(row, table)
        unite = normalize_unit(_get_value(row, table.column_map, "unite"))
        quantite = parse_number(_get_value(row, table.column_map, "quantite"))
        prix_unitaire = parse_number(_get_value(row, table.column_map, "prix_unitaire"))
        montant = parse_number(_get_value(row, table.column_map, "montant"))
        raw_text = " ".join(_collect_text_candidates(row, table.column_map))
        measure_row = _is_measure_row(unite, quantite, prix_unitaire, montant)

        line_type = classify_line(
            designation=designation,
            code_article=code_article,
            unite=unite,
            quantite=quantite,
            prix_unitaire=prix_unitaire,
            montant=montant,
            raw_text=raw_text,
        )

        if line_type == "unknown":
            continue

        if (
            pending.base_code
            and line_type == "chapter"
            and not measure_row
            and not code_article
            and designation
            and not _is_explicit_major_heading(designation)
        ):
            pending.detail_parts.append(designation)
            pending.source_rows.append((table_row.row_index, row))
            continue

        if line_type in {"total", "subtotal", "chapter"}:
            pending.reset()

        if (
            line_type == "subchapter"
            and _is_parent_article_code(code_article)
            and not measure_row
            and designation
        ):
            pending.base_code = code_article
            pending.base_designation = designation
            pending.clear_details()
            pending.source_rows = [(table_row.row_index, row)]

        elif pending.base_code and line_type == "subchapter" and not measure_row:
            if code_article and code_article != pending.base_code and _is_short_detail_code(code_article):
                pending.detail_code = code_article
            if designation and designation != pending.base_designation:
                pending.detail_parts = [designation]
            base_rows = pending.source_rows[:1] if pending.source_rows else []
            pending.source_rows = [*base_rows, (table_row.row_index, row)]
            continue

        elif pending.base_code and not measure_row and line_type in {"comment", "chapter", "subchapter"} and designation:
            if designation != pending.base_designation:
                pending.detail_parts.append(designation)
            pending.source_rows.append((table_row.row_index, row))
            continue

        merged_rows = [(table_row.row_index, row)]
        effective_code = code_article
        effective_designation = designation
        effective_unite = unite
        effective_quantite = quantite
        effective_prix_unitaire = prix_unitaire
        effective_montant = montant

        if pending.base_code and measure_row and line_type == "item":
            effective_code = _combine_code(pending.base_code, pending.detail_code, code_article)
            effective_designation = _combine_designation(
                pending.base_designation,
                pending.detail_parts,
                designation,
            )
            merged_rows = [*pending.source_rows, (table_row.row_index, row)] if pending.source_rows else merged_rows
            pending.clear_details()

        if line_type in {"chapter", "subtotal", "total"}:
            effective_designation = _pick_heading_seed(effective_designation, effective_code, raw_text)
            effective_unite = None
            effective_quantite = None
            effective_prix_unitaire = None
            if line_type in {"subtotal", "total"}:
                effective_code = None

        effective_designation = _clean_designation_parts(
            effective_designation,
            effective_unite,
            effective_quantite,
        )
        effective_designation = _clean_heading_designation(effective_designation, line_type)

        if line_type in {"chapter", "subchapter"}:
            _update_context(line_type, effective_designation, effective_code, context)

        lot = normalize_text(_get_value(row, table.column_map, "lot")) or context.get("lot")
        chapitre = normalize_text(_get_value(row, table.column_map, "chapitre")) or context.get("chapter")
        sous_chapitre = (
            normalize_text(_get_value(row, table.column_map, "sous_chapitre")) or context.get("subchapter")
        )
        if line_type == "subtotal":
            sous_chapitre = None
        elif line_type == "total":
            chapitre = None
            sous_chapitre = None

        item = ExtractedItem(
            sheet_name=table.sheet_name,
            row_index=table_row.row_index,
            line_type=line_type,
            lot=lot or None,
            chapter=chapitre or None,
            subchapter=sous_chapitre or None,
            code_article=effective_code,
            designation=effective_designation or None,
            unite=effective_unite,
            quantite=effective_quantite,
            prix_unitaire=effective_prix_unitaire,
            montant=effective_montant,
            raw_row=_raw_rows_dict(merged_rows, table.normalized_headers),
            confidence=_compute_row_confidence(
                line_type=line_type,
                designation=effective_designation,
                code_article=effective_code,
                unite=effective_unite,
                quantite=effective_quantite,
                prix_unitaire=effective_prix_unitaire,
                montant=effective_montant,
                table_score=max(table.score, relevant_score),
            ),
        )
        items.append(item)

    extraction_density = len(items) / max(len(table.table_rows), 1)
    sheet_confidence = min(
        round(relevant_score * 0.35 + table.score * 0.45 + min(extraction_density, 1.0) * 0.20, 3),
        0.99,
    )

    summary = SheetExtractionSummary(
        sheet_name=table.sheet_name,
        relevant_score=relevant_score,
        table_score=table.score,
        sheet_confidence=sheet_confidence,
        rows_read=len(table.table_rows),
        rows_extracted=len(items),
        mapping_columns=sorted(table.column_map.keys()),
        warnings=warnings,
    )
    return items, summary


def _candidate_lookup(candidates: list[SheetCandidate]) -> dict[str, SheetCandidate]:
    return {candidate.sheet_name: candidate for candidate in candidates}


def extract_workbook(
    input_path: str | Path,
    forced_sheets: list[str] | None = None,
    all_sheets: bool = False,
    debug: bool = False,
) -> ExtractionResult:
    workbook_path = Path(input_path)
    sheets = read_workbook(workbook_path)
    candidates, selected_sheets, selection_warnings = detect_relevant_sheets(
        sheets=sheets,
        forced_sheets=forced_sheets,
        all_sheets=all_sheets,
    )
    lookup = _candidate_lookup(candidates)

    items: list[ExtractedItem] = []
    warnings: list[str] = list(selection_warnings)
    summaries: list[SheetExtractionSummary] = []
    total_rows_read = 0

    for sheet in sheets:
        if sheet.name not in selected_sheets:
            continue
        try:
            table = detect_main_table(sheet)
            relevant_score = lookup.get(sheet.name, SheetCandidate(sheet.name, 0.0)).score
            sheet_items, summary = extract_table_items(table, relevant_score)
            items.extend(sheet_items)
            summaries.append(summary)
            total_rows_read += summary.rows_read
            warnings.extend(summary.warnings)
            if debug:
                LOGGER.info(
                    "Feuille %s: %s lignes lues, %s lignes extraites.",
                    sheet.name,
                    summary.rows_read,
                    summary.rows_extracted,
                )
        except Exception as exc:  # pragma: no cover - defensive branch
            message = f"Echec sur la feuille {sheet.name}: {exc}"
            warnings.append(message)
            LOGGER.exception(message)

    stats = {
        "total_rows_read": total_rows_read,
        "total_items_extracted": sum(1 for item in items if item.line_type == "item"),
        "total_lines_extracted": len(items),
        "selected_sheet_count": len(selected_sheets),
    }

    return ExtractionResult(
        source_file=workbook_path,
        sheets_analyzed=[sheet.name for sheet in sheets],
        selected_sheets=selected_sheets,
        items=items,
        warnings=sorted(set(warnings)),
        stats=stats,
        sheet_summaries=summaries,
    )
