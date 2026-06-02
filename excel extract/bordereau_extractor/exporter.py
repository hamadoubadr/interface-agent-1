from __future__ import annotations

import csv
import json
from pathlib import Path

from openpyxl import Workbook

from .models import ExtractionResult


def ensure_output_dir(path: str | Path) -> Path:
    output_dir = Path(path)
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def export_json(result: ExtractionResult, output_dir: str | Path) -> Path:
    destination = ensure_output_dir(output_dir) / "bordereau_extraction.json"
    destination.write_text(
        json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return destination


def export_csv(result: ExtractionResult, output_dir: str | Path) -> Path:
    destination = ensure_output_dir(output_dir) / "bordereau_items.csv"
    fieldnames = [
        "sheet_name",
        "row_index",
        "line_type",
        "lot",
        "chapter",
        "subchapter",
        "code_article",
        "designation",
        "unite",
        "quantite",
        "prix_unitaire",
        "montant",
        "confidence",
    ]
    with destination.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for item in result.items:
            writer.writerow({field: item.to_dict().get(field) for field in fieldnames})
    return destination


def export_debug_workbook(result: ExtractionResult, output_dir: str | Path) -> Path:
    destination = ensure_output_dir(output_dir) / "bordereau_debug.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "items"

    headers = [
        "sheet_name",
        "row_index",
        "line_type",
        "lot",
        "chapter",
        "subchapter",
        "code_article",
        "designation",
        "unite",
        "quantite",
        "prix_unitaire",
        "montant",
        "confidence",
        "raw_row_json",
    ]
    worksheet.append(headers)
    for item in result.items:
        payload = item.to_dict()
        worksheet.append(
            [
                payload["sheet_name"],
                payload["row_index"],
                payload["line_type"],
                payload["lot"],
                payload["chapter"],
                payload["subchapter"],
                payload["code_article"],
                payload["designation"],
                payload["unite"],
                payload["quantite"],
                payload["prix_unitaire"],
                payload["montant"],
                payload["confidence"],
                json.dumps(payload["raw_row"], ensure_ascii=False),
            ]
        )

    meta = workbook.create_sheet(title="meta")
    meta.append(["source_file", str(result.source_file)])
    meta.append(["selected_sheets", ", ".join(result.selected_sheets)])
    meta.append(["total_rows_read", result.stats.get("total_rows_read")])
    meta.append(["total_items_extracted", result.stats.get("total_items_extracted")])
    for warning in result.warnings:
        meta.append(["warning", warning])

    workbook.save(destination)
    return destination
