from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Iterable
from typing import Any

STANDARD_COLUMNS = (
    "code_article",
    "designation",
    "unite",
    "quantite",
    "prix_unitaire",
    "montant",
    "lot",
    "chapitre",
    "sous_chapitre",
)

HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "code_article": (
        "article",
        "code",
        "numero",
        "num",
        "n",
        "repere",
        "reference",
        "ref",
        "poste",
        "item",
        "n prix",
        "n° prix",
        "n°",
    ),
    "designation": (
        "designation",
        "designation des ouvrages",
        "description",
        "ouvrage",
        "prestation",
        "travaux",
        "libelle",
        "intitule",
        "designation prestation",
        "descriptif",
    ),
    "unite": ("unite", "u", "um", "u.m", "u m", "unit", "uds"),
    "quantite": ("quantite", "qte", "qt", "q", "qty", "quant"),
    "prix_unitaire": (
        "prix unitaire",
        "pu",
        "p u",
        "p.u",
        "prix unit",
        "prix u",
        "pu ht",
        "prix unitaire ht",
    ),
    "montant": (
        "montant",
        "total",
        "pt",
        "p.t",
        "p t",
        "prix total",
        "montant ht",
        "montant ttc",
        "total ht",
        "total ttc",
        "cout total",
        "prix global",
    ),
    "lot": ("lot", "macro lot", "ensemble"),
    "chapitre": ("chapitre", "chap", "section"),
    "sous_chapitre": ("sous chapitre", "sous-chapitre", "ss chapitre", "s section"),
}

TOTAL_KEYWORDS = ("total general", "total", "totaux", "montant total", "grand total")
SUBTOTAL_KEYWORDS = ("sous total", "sous-total", "subtotal")
COMMENT_KEYWORDS = ("note", "observation", "remarque", "commentaire", "memo")
LOT_KEYWORDS = ("lot", "macro lot", "ouvrage", "zone")
CHAPTER_HINTS = ("chapitre", "section", "poste", "article", "phase")
UNIT_HINTS = {"u", "ml", "m", "m2", "m3", "kg", "kgf", "ff", "ens", "forfait", "paire", "set"}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value)
    text = text.replace("\u00a0", " ").replace("\u202f", " ").replace("\r", "\n")
    text = text.replace("\t", " ")
    text = re.sub(r"[ ]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def normalize_for_match(value: Any) -> str:
    text = normalize_text(value).lower()
    text = "".join(
        ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch)
    )
    text = text.replace("\u00b0", " ")
    text = re.sub(r"[^a-z0-9.\n ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def is_empty(value: Any) -> bool:
    return normalize_text(value) == ""


def count_non_empty(values: Iterable[Any]) -> int:
    return sum(1 for value in values if not is_empty(value))


def is_mostly_numeric(text: str) -> bool:
    if not text:
        return False
    digits = sum(ch.isdigit() for ch in text)
    letters = sum(ch.isalpha() for ch in text)
    return digits > 0 and digits >= letters * 2


def parse_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)

    text = normalize_text(value)
    if not text:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]

    text = text.replace(" ", "").replace("\u00a0", "").replace("\u202f", "")
    text = re.sub(r"[^0-9,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    if text.count(",") > 0 and text.count(".") > 0:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") > 0:
        if text.count(",") == 1 and len(text.split(",")[-1]) in {1, 2, 3}:
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(".") > 1:
        parts = text.split(".")
        decimal = parts[-1]
        if len(decimal) in {1, 2, 3}:
            text = "".join(parts[:-1]) + "." + decimal
        else:
            text = "".join(parts)

    try:
        number = float(text)
    except ValueError:
        return None

    return -number if negative else number


def normalize_unit(value: Any) -> str | None:
    text = normalize_text(value)
    if not text:
        return None
    compact = normalize_for_match(text).replace(" ", "")
    if compact in UNIT_HINTS:
        return compact.upper() if compact != "forfait" else "Forfait"
    return text


def match_header_label(label: str) -> tuple[str | None, float]:
    normalized = normalize_for_match(label)
    if not normalized:
        return None, 0.0

    best_column: str | None = None
    best_score = 0.0
    tokens = set(normalized.split())

    for standard, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            alias_norm = normalize_for_match(alias)
            alias_tokens = set(alias_norm.split())
            if normalized == alias_norm:
                score = 1.0
            elif len(alias_norm.replace(" ", "")) <= 2:
                score = 0.92 if alias_norm in tokens else 0.0
            elif alias_norm in normalized:
                score = 0.92
            elif tokens and alias_tokens:
                overlap = len(tokens & alias_tokens) / len(alias_tokens)
                score = 0.55 + 0.35 * overlap if overlap >= 0.5 else 0.0
            else:
                score = 0.0

            if score > best_score:
                best_column = standard
                best_score = score

    return best_column, round(best_score, 3)


def map_headers(headers: dict[int, str]) -> tuple[dict[str, int], dict[int, str], float, list[str]]:
    scored_matches: list[tuple[float, str, int, str]] = []
    warnings: list[str] = []

    for index, label in headers.items():
        standard, score = match_header_label(label)
        if standard and score > 0:
            scored_matches.append((score, standard, index, label))

    scored_matches.sort(reverse=True)

    column_map: dict[str, int] = {}
    normalized_headers: dict[int, str] = {}
    used_columns: set[int] = set()

    for score, standard, index, label in scored_matches:
        if standard in column_map or index in used_columns:
            continue
        column_map[standard] = index
        normalized_headers[index] = label
        used_columns.add(index)

    mapping_score = 0.0
    if column_map:
        mapping_score += min(len(column_map) / 6.0, 1.0) * 0.7
        mapping_score += 0.15 if "designation" in column_map else 0.0
        mapping_score += 0.1 if "quantite" in column_map else 0.0
        mapping_score += 0.1 if "unite" in column_map else 0.0
        mapping_score += 0.1 if "montant" in column_map or "prix_unitaire" in column_map else 0.0
    else:
        warnings.append("Aucune colonne standard reconnue.")

    if "designation" not in column_map:
        warnings.append("Colonne désignation non détectée.")

    return column_map, normalized_headers, min(round(mapping_score, 3), 1.0), warnings
