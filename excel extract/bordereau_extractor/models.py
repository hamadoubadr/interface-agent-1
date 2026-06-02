from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class SheetData:
    name: str
    matrix: list[list[Any]]
    source: str
    warnings: list[str] = field(default_factory=list)
    merged_cells_expanded: bool = False
    formula_density: float = 0.0
    non_empty_rows: int = 0
    non_empty_cols: int = 0


@dataclass(slots=True)
class SheetCandidate:
    sheet_name: str
    score: float
    reasons: list[str] = field(default_factory=list)


@dataclass(slots=True)
class TableRow:
    row_index: int
    values: list[Any]


@dataclass(slots=True)
class TableDetectionResult:
    sheet_name: str
    header_row_index: int | None
    header_depth: int
    column_map: dict[str, int]
    normalized_headers: dict[int, str]
    table_rows: list[TableRow]
    score: float
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ExtractedItem:
    sheet_name: str
    row_index: int
    line_type: str
    lot: str | None = None
    chapter: str | None = None
    subchapter: str | None = None
    code_article: str | None = None
    designation: str | None = None
    unite: str | None = None
    quantite: float | None = None
    prix_unitaire: float | None = None
    montant: float | None = None
    raw_row: dict[str, Any] = field(default_factory=dict)
    confidence: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class SheetExtractionSummary:
    sheet_name: str
    relevant_score: float
    table_score: float
    sheet_confidence: float
    rows_read: int
    rows_extracted: int
    mapping_columns: list[str]
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class ExtractionResult:
    source_file: Path
    sheets_analyzed: list[str]
    selected_sheets: list[str]
    items: list[ExtractedItem]
    warnings: list[str]
    stats: dict[str, Any]
    sheet_summaries: list[SheetExtractionSummary] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_file": str(self.source_file),
            "sheets_analyzed": self.sheets_analyzed,
            "selected_sheets": self.selected_sheets,
            "items": [item.to_dict() for item in self.items],
            "warnings": self.warnings,
            "stats": self.stats,
            "sheet_summaries": [summary.to_dict() for summary in self.sheet_summaries],
        }
