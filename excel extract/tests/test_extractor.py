from bordereau_extractor.extractor import classify_line, extract_table_items
from bordereau_extractor.models import TableDetectionResult, TableRow


def test_classify_item_line():
    line_type = classify_line(
        designation="Fourniture et pose de tube PPR",
        code_article="1.2.3",
        unite="ML",
        quantite=12.0,
        prix_unitaire=None,
        montant=None,
        raw_text="",
    )
    assert line_type == "item"


def test_classify_subtotal_line():
    line_type = classify_line(
        designation="Sous-total plomberie",
        code_article=None,
        unite=None,
        quantite=None,
        prix_unitaire=None,
        montant=1500.0,
        raw_text="",
    )
    assert line_type == "subtotal"


def test_extract_multiline_item_pattern():
    table = TableDetectionResult(
        sheet_name="Feuil1",
        header_row_index=3,
        header_depth=1,
        column_map={
            "code_article": 0,
            "designation": 1,
            "unite": 2,
            "quantite": 3,
            "prix_unitaire": 4,
            "montant": 5,
        },
        normalized_headers={
            0: "N°",
            1: "DESIGNATION DES PRESTATIONS",
            2: "UNITE",
            3: "Q",
            4: "P.U (H.T)",
            5: "P.T (H.T)",
        },
        table_rows=[
            TableRow(5, ["", "A/Evacuation EP .EV. EU.", "", "", "", ""]),
            TableRow(7, ["A-01", "Tuyau P.V.C  évacuation", "", "", "", ""]),
            TableRow(8, ["a", "diam.40", "", "", "", ""]),
            TableRow(9, ["", "le mètre linéaire", "ML", "71,00", "", ""]),
        ],
        score=0.9,
        warnings=[],
    )

    items, summary = extract_table_items(table, relevant_score=0.8)
    extracted_items = [item for item in items if item.line_type == "item"]
    assert summary.rows_extracted >= 2
    assert len(extracted_items) == 1
    assert extracted_items[0].code_article == "A-01.a"
    assert "Tuyau P.V.C" in extracted_items[0].designation
    assert "diam.40" in extracted_items[0].designation
    assert "mètre linéaire" not in extracted_items[0].designation
    assert extracted_items[0].quantite == 71.0


def test_classify_section_label_without_code():
    line_type = classify_line(
        designation="A/Evacuation EP .EV. EU.",
        code_article=None,
        unite=None,
        quantite=None,
        prix_unitaire=None,
        montant=None,
        raw_text="",
    )
    assert line_type == "chapter"


def test_extract_subtotal_keeps_clean_label():
    table = TableDetectionResult(
        sheet_name="Feuil1",
        header_row_index=3,
        header_depth=1,
        column_map={
            "code_article": 0,
            "designation": 1,
            "unite": 2,
            "quantite": 3,
            "prix_unitaire": 4,
            "montant": 5,
        },
        normalized_headers={
            0: "N°",
            1: "DESIGNATION DES PRESTATIONS",
            2: "UNITE",
            3: "Q",
            4: "P.U (H.T)",
            5: "P.T (H.T)",
        },
        table_rows=[
            TableRow(40, ["", "B/ Alimentation eau froide et chaude", "", "", "", ""]),
            TableRow(42, ["B-01", "Equipement compteur d'eau", "", "", "", ""]),
            TableRow(44, ["", "l'ensemble", "E", "1,00", "", ""]),
            TableRow(93, ["SOUS TOTAL B /………………………………………..", "", "", "", "", "0,00"]),
        ],
        score=0.9,
        warnings=[],
    )

    items, _ = extract_table_items(table, relevant_score=0.8)
    subtotal = [item for item in items if item.line_type == "subtotal"][0]
    assert subtotal.designation == "Sous-total B"
    assert subtotal.chapter == "B/ Alimentation eau froide et chaude"
    assert subtotal.subchapter is None
