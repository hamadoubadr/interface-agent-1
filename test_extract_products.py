import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "streamlit_dossier_app"))

from app_core.pdf_enrichment import extract_products_from_descriptif

PDF_PATH = Path(r"C:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia\test_results_alysse_gemini\input\ALYSSE-DESCRIPTIF_FLUIDES - IND C.pdf")
WORKSPACE_ROOT = Path(r"C:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia\test_descriptif_products")

print(f"Test extraction depuis: {PDF_PATH.name}")
print("=" * 60)

products_df = extract_products_from_descriptif(
    [PDF_PATH],
    WORKSPACE_ROOT,
    use_nvidia=True
)

print(f"\nTotal produits extraits: {len(products_df)}")
print(f"Colonnes: {list(products_df.columns)}")
print("\n" + "=" * 60)

if not products_df.empty:
    print("\nProduits extraits:\n")
    for idx, row in products_df.iterrows():
        print(f"{idx+1}. [{row.get('chapter', 'N/A')}] [{row.get('subchapter', 'N/A')}]")
        print(f"   Nom: {row.get('product_name', 'N/A')}")
        print(f"   Code: {row.get('code_article', 'N/A')}")
        print(f"   Unité: {row.get('unite', 'N/A')} | Qté: {row.get('quantite', 'N/A')}")
        print()
else:
    print("Aucun produit extrait!")

output_dir = WORKSPACE_ROOT / "test_output"
output_dir.mkdir(parents=True, exist_ok=True)
csv_path = output_dir / "produits_extraits_test.csv"
products_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
print(f"\nRésultats sauvegardés dans: {csv_path}")