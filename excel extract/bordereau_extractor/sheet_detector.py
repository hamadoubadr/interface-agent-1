from __future__ import annotations

from .models import SheetCandidate, SheetData
from .normalizer import HEADER_ALIASES, count_non_empty, normalize_for_match

POSITIVE_KEYWORDS = (
    "bordereau",
    "bpu",
    "dpgf",
    "devis",
    "quantite",
    "prix",
    "ouvrage",
    "prestation",
    "lot",
    "article",
)

NEGATIVE_KEYWORDS = (
    "page de garde",
    "garde",
    "synthese",
    "recapitulatif",
    "mode emploi",
    "sommaire",
    "dashboard",
    "temp",
    "tmp",
    "notes",
)


def score_sheet(sheet: SheetData) -> SheetCandidate:
    flattened = []
    for row in sheet.matrix[:30]:
        flattened.extend(row[:20])

    text = " ".join(normalize_for_match(value) for value in flattened if value is not None)
    sheet_name = normalize_for_match(sheet.name)
    reasons: list[str] = []
    score = 0.0

    for keyword in POSITIVE_KEYWORDS:
        if keyword in text or keyword in sheet_name:
            score += 0.08
            reasons.append(f"Mot-clé détecté: {keyword}")

    for keyword in NEGATIVE_KEYWORDS:
        if keyword in text or keyword in sheet_name:
            score -= 0.12
            reasons.append(f"Signal negatif: {keyword}")

    header_hits = 0
    for row in sheet.matrix[:20]:
        normalized_row = [normalize_for_match(cell) for cell in row[:20]]
        row_text = " ".join(cell for cell in normalized_row if cell)
        for aliases in HEADER_ALIASES.values():
            if any(normalize_for_match(alias) in row_text for alias in aliases):
                header_hits += 1
                break
    if header_hits:
        score += min(header_hits / 8.0, 1.0) * 0.4
        reasons.append(f"Indices de colonnes détectés: {header_hits}")

    dense_rows = sum(1 for row in sheet.matrix if count_non_empty(row) >= 3)
    if dense_rows:
        density = dense_rows / max(len(sheet.matrix), 1)
        score += min(density, 0.4)
        reasons.append(f"Densite utile: {density:.2f}")

    if sheet.merged_cells_expanded:
        score += 0.05
        reasons.append("Cellules fusionnees gerees")

    return SheetCandidate(sheet_name=sheet.name, score=max(round(score, 3), 0.0), reasons=reasons)


def detect_relevant_sheets(
    sheets: list[SheetData],
    forced_sheets: list[str] | None = None,
    all_sheets: bool = False,
) -> tuple[list[SheetCandidate], list[str], list[str]]:
    candidates = [score_sheet(sheet) for sheet in sheets]
    candidate_by_name = {candidate.sheet_name.lower(): candidate for candidate in candidates}
    analyzed_names = [sheet.name for sheet in sheets]

    if forced_sheets:
        selected: list[str] = []
        warnings: list[str] = []
        for requested in forced_sheets:
            candidate = candidate_by_name.get(requested.lower())
            if candidate:
                selected.append(candidate.sheet_name)
            else:
                warnings.append(f"Feuille demandee introuvable: {requested}")
        return candidates, selected, warnings

    if all_sheets:
        return candidates, analyzed_names, []

    sorted_candidates = sorted(candidates, key=lambda item: item.score, reverse=True)
    selected = [candidate.sheet_name for candidate in sorted_candidates if candidate.score >= 0.35]
    if not selected and sorted_candidates:
        selected = [sorted_candidates[0].sheet_name]
    return candidates, selected[:5], []
