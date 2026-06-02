# Excel Extract

Module Python dedie a l'extraction de bordereaux / BPU / DPGF / devis quantitatifs de construction, sans LLM ni service externe.

## Architecture proposee

```text
excel extract/
|-- extract_bordereau.py
|-- pyproject.toml
|-- requirements.txt
|-- README.md
|-- examples/
|   `-- sample_output.json
|-- bordereau_extractor/
|   |-- __init__.py
|   |-- __main__.py
|   |-- cli.py
|   |-- exporter.py
|   |-- extractor.py
|   |-- models.py
|   |-- normalizer.py
|   |-- reader.py
|   |-- sheet_detector.py
|   `-- table_detector.py
`-- tests/
    |-- test_exporter.py
    |-- test_extractor.py
    |-- test_normalizer.py
    `-- test_table_detector.py
```

## Heuristiques metier

- Prioriser les feuilles qui ressemblent a un bordereau via le nom de feuille, les mots-cles et la densite tabulaire.
- Detecter la vraie ligne d'en-tete meme si elle commence apres plusieurs lignes de garde.
- Combiner jusqu'a deux lignes d'en-tete pour recuperer des colonnes eclatees.
- Reconnaitre les variantes de colonnes metier : article, designation, unite, quantite, PU, montant, lot, chapitre.
- Ne pas confondre `total` / `sous-total` avec des lignes produits.
- Conserver le contexte `lot > chapitre > sous-chapitre` pour les lignes produits suivantes.
- Accepter les feuilles purement quantitatives : une ligne peut etre un `item` meme sans prix si la designation, l'unite ou la quantite sont presentes.
- Tolerer les cellules fusionnees via un fallback `openpyxl` qui propage la valeur de la cellule source sur toute la zone fusionnee.
- Garder la ligne brute dans `raw_row` pour audit et debogage.

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Utilisation

```bash
python extract_bordereau.py input.xlsx --out output_dir
python extract_bordereau.py input.xls --out output_dir
```

Options utiles :

- `--sheet "NomFeuille"` : force une feuille precise. Peut etre repete.
- `--all-sheets` : analyse toutes les feuilles.
- `--debug` : logs detailles + export d'un classeur de debug.
- `--json-only` : ne sort que le JSON.
- `--csv-only` : ne sort que le CSV.
- `--list-sheets` : affiche uniquement les noms de feuilles detectes.

## Fichiers generes

- `bordereau_extraction.json`
- `bordereau_items.csv`
- `bordereau_debug.xlsx` avec `--debug`

## Tests

```bash
pytest
```

## Limites actuelles

- Les fichiers `.xls`, `.xlsx` et `.xlsm` sont acceptes.
- La reconstruction de colonnes tres eclatees sur plus de 2 lignes d'en-tete reste volontairement conservative.
- Certains bordereaux atypiques peuvent necessiter de forcer la feuille avec `--sheet`.
