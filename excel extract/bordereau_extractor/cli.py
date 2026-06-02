from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from .exporter import export_csv, export_debug_workbook, export_json
from .extractor import extract_workbook
from .reader import read_workbook


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="extract_bordereau.py",
        description="Extraction robuste de lignes de bordereaux / BPU / DPGF de construction.",
    )
    parser.add_argument("input_file", help="Chemin du fichier .xlsx ou .xlsm")
    parser.add_argument("--out", default="output", help="Dossier de sortie")
    parser.add_argument("--sheet", action="append", help="Nom d'une feuille a forcer")
    parser.add_argument("--all-sheets", action="store_true", help="Analyser toutes les feuilles")
    parser.add_argument("--debug", action="store_true", help="Active les logs detailes et exporte un xlsx de debug")
    parser.add_argument("--json-only", action="store_true", help="N'exporte que le JSON")
    parser.add_argument("--csv-only", action="store_true", help="N'exporte que le CSV")
    parser.add_argument("--list-sheets", action="store_true", help="Affiche uniquement les noms des feuilles")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.json_only and args.csv_only:
        parser.error("Choisir soit --json-only soit --csv-only, pas les deux.")

    logging.basicConfig(
        level=logging.INFO if args.debug else logging.WARNING,
        format="%(levelname)s %(message)s",
    )

    input_path = Path(args.input_file)
    if not input_path.exists():
        parser.error(f"Fichier introuvable: {input_path}")
    if input_path.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
        parser.error("Le fichier doit etre au format .xls, .xlsx ou .xlsm")

    if args.list_sheets:
        sheets = read_workbook(input_path)
        print(json.dumps([sheet.name for sheet in sheets], ensure_ascii=False, indent=2))
        return 0

    result = extract_workbook(
        input_path=input_path,
        forced_sheets=args.sheet,
        all_sheets=args.all_sheets,
        debug=args.debug,
    )

    output_dir = Path(args.out)
    generated: list[Path] = []
    if not args.csv_only:
        generated.append(export_json(result, output_dir))
    if not args.json_only:
        generated.append(export_csv(result, output_dir))
    if args.debug:
        generated.append(export_debug_workbook(result, output_dir))

    summary = {
        "source_file": str(result.source_file),
        "selected_sheets": result.selected_sheets,
        "stats": result.stats,
        "warnings_count": len(result.warnings),
        "generated_files": [str(path) for path in generated],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0
