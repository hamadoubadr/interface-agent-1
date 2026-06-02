import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "streamlit_dossier_app"))

from app_core.pdf_enrichment import extract_products_from_descriptif, enrich_products_dataframe

PDF_PATH = Path(r"C:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia\test_results_alysse_gemini\input\ALYSSE-DESCRIPTIF_FLUIDES - IND C.pdf")
WORKSPACE_ROOT = Path(r"C:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia\test_enrich_brand")

print(f"Test enrichment marques depuis: {PDF_PATH.name}")
print("=" * 80)

products_df = extract_products_from_descriptif(
    [PDF_PATH],
    WORKSPACE_ROOT,
    use_nvidia=True
)

print(f"\nProduits extraits: {len(products_df)}")

enriched_df = enrich_products_dataframe(
    products_df,
    [PDF_PATH],
    use_nvidia=True
)

print(f"\nProduits après enrichment: {len(enriched_df)}")
print(f"Colonnes: {list(enriched_df.columns)}")
print("\n" + "=" * 80)

if not enriched_df.empty:
    print("\nProduits avec marques (brand):\n")
    for idx, row in enriched_df.head(50).iterrows():
        brand = row.get('brand', 'N/A')
        code = row.get('code_article', 'N/A')
        parent = row.get('parent_code', '')
        sub = row.get('sub_code', '')
        name = row.get('product_name', 'N/A')

        if parent:
            code_display = f"{parent}-{sub}"
        else:
            code_display = code if code else ""

        print(f"{code_display:12} | {name[:45]:45} | Marque: {brand if brand else '-'}")
else:
    print("Aucun produit après enrichment!")

output_dir = WORKSPACE_ROOT / "test_output"
output_dir.mkdir(parents=True, exist_ok=True)
csv_path = output_dir / "produits_enrichis_test.csv"
enriched_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
print(f"\nRésultats sauvegardés dans: {csv_path}")