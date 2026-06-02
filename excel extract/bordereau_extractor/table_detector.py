from __future__ import annotations

from typing import Any

from .models import SheetData, TableDetectionResult, TableRow
from .normalizer import count_non_empty, is_mostly_numeric, map_headers, normalize_text


def _row_cells(matrix: list[list[Any]], row_index: int) -> list[Any]:
    if 0 <= row_index < len(matrix):
        return matrix[row_index]
    return []


def _combine_header_rows(matrix: list[list[Any]], start: int, depth: int) -> dict[int, str]:
    rows = [_row_cells(matrix, start + offset) for offset in range(depth)]
    max_cols = max((len(row) for row in rows), default=0)
    combined: dict[int, str] = {}

    for col_index in range(max_cols):
        parts: list[str] = []
        for row in rows:
            if col_index >= len(row):
                continue
            cell = normalize_text(row[col_index])
            if not cell:
                continue
            # Long text outside the first column is usually body content, not a header label.
            if len(cell) > 80 and col_index > 0:
                continue
            if cell not in parts:
                parts.append(cell)
        combined[col_index] = " ".join(parts).strip()
    return combined


def _score_header_candidate(
    matrix: list[list[Any]],
    row_index: int,
    depth: int,
) -> tuple[float, dict[str, int], dict[int, str], list[str]]:
    combined = _combine_header_rows(matrix, row_index, depth)
    header_map, normalized_headers, mapping_score, warnings = map_headers(combined)
    row = _row_cells(matrix, row_index)
    header_texts = [normalize_text(cell) for cell in row if normalize_text(cell)]
    numeric_penalty = (
        0.2
        if header_texts and sum(is_mostly_numeric(text) for text in header_texts) > len(header_texts) / 2
        else 0.0
    )

    score = mapping_score
    if "designation" in header_map:
        score += 0.2
    if "quantite" in header_map:
        score += 0.1
    if "unite" in header_map:
        score += 0.08
    if "prix_unitaire" in header_map or "montant" in header_map:
        score += 0.08
    score -= numeric_penalty

    return round(max(score, 0.0), 3), header_map, normalized_headers, warnings


def detect_main_table(sheet: SheetData) -> TableDetectionResult:
    best_score = -1.0
    best_header_row: int | None = None
    best_depth = 1
    best_map: dict[str, int] = {}
    best_headers: dict[int, str] = {}
    warnings: list[str] = list(sheet.warnings)

    scan_limit = min(len(sheet.matrix), 40)
    for row_index in range(scan_limit):
        if count_non_empty(sheet.matrix[row_index]) < 2:
            continue
        for depth in (1, 2):
            # Many bordereaux split headers over two stacked rows: "Prix" / "unitaire".
            score, column_map, headers, local_warnings = _score_header_candidate(sheet.matrix, row_index, depth)
            if score > best_score:
                best_score = score
                best_header_row = row_index
                best_depth = depth
                best_map = column_map
                best_headers = headers
                warnings = list(sheet.warnings) + local_warnings

    if best_header_row is None:
        warnings.append("Aucune ligne d'en-tête plausible détectée.")
        return TableDetectionResult(
            sheet_name=sheet.name,
            header_row_index=None,
            header_depth=0,
            column_map={},
            normalized_headers={},
            table_rows=[],
            score=0.0,
            warnings=warnings,
        )

    data_start = best_header_row + best_depth
    table_rows: list[TableRow] = []
    empty_streak = 0
    for matrix_index in range(data_start, len(sheet.matrix)):
        values = sheet.matrix[matrix_index]
        if count_non_empty(values) == 0:
            empty_streak += 1
            if empty_streak >= 5 and table_rows:
                break
            continue
        empty_streak = 0
        table_rows.append(TableRow(row_index=matrix_index + 1, values=values))

    if not table_rows:
        warnings.append("Zone tabulaire vide après l'en-tête détecté.")

    return TableDetectionResult(
        sheet_name=sheet.name,
        header_row_index=best_header_row + 1,
        header_depth=best_depth,
        column_map=best_map,
        normalized_headers=best_headers,
        table_rows=table_rows,
        score=max(best_score, 0.0),
        warnings=warnings,
    )
