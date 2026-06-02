import json

from bordereau_extractor.exporter import export_json
from bordereau_extractor.models import ExtractedItem, ExtractionResult


def test_export_json(tmp_path):
    result = ExtractionResult(
        source_file=tmp_path / "input.xlsx",
        sheets_analyzed=["BPU"],
        selected_sheets=["BPU"],
        items=[
            ExtractedItem(
                sheet_name="BPU",
                row_index=4,
                line_type="item",
                designation="Tube PPR",
                quantite=10.0,
                confidence=0.91,
            )
        ],
        warnings=[],
        stats={"total_rows_read": 1, "total_items_extracted": 1},
    )

    output_path = export_json(result, tmp_path)
    payload = json.loads(output_path.read_text(encoding="utf-8"))
    assert payload["source_file"].endswith("input.xlsx")
    assert payload["items"][0]["designation"] == "Tube PPR"
