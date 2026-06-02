from bordereau_extractor.normalizer import map_headers, match_header_label, parse_number


def test_match_header_label_designation():
    label, score = match_header_label("Désignation des ouvrages")
    assert label == "designation"
    assert score > 0.8


def test_map_headers_basic():
    headers = {
        0: "Article",
        1: "Désignation",
        2: "U",
        3: "Qté",
        4: "P.U HT",
        5: "Montant",
    }
    column_map, _, mapping_score, warnings = map_headers(headers)
    assert column_map["code_article"] == 0
    assert column_map["designation"] == 1
    assert column_map["unite"] == 2
    assert column_map["quantite"] == 3
    assert mapping_score > 0.7
    assert warnings == []


def test_parse_number_fr():
    assert parse_number("1 234,50") == 1234.5
    assert parse_number("11.460,00") == 11460.0
    assert parse_number("(950,25)") == -950.25
    assert parse_number("N/A") is None
