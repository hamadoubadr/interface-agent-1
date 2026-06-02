from bordereau_extractor.models import SheetData
from bordereau_extractor.table_detector import detect_main_table


def test_detect_main_table_with_offset_header():
    sheet = SheetData(
        name="BPU",
        source="pandas",
        matrix=[
            ["Projet", "Test", None],
            [None, None, None],
            ["Article", "Désignation", "Qté", "P.U"],
            ["1.1", "Tube cuivre", "12", "120,00"],
            ["1.2", "Vanne", "4", "85,00"],
        ],
    )

    result = detect_main_table(sheet)
    assert result.header_row_index == 3
    assert result.column_map["designation"] == 1
    assert len(result.table_rows) == 2
