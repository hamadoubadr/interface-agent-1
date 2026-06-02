import os
import sys
import time
from pathlib import Path

# Résolution dynamique des chemins
WORKSPACE_ROOT = Path(r"c:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia")
sys.path.insert(0, str(WORKSPACE_ROOT / "streamlit_dossier_app"))

from app_core.pipeline import run_dossier_synthesis

def test_synthesis():
    print("=" * 80)
    print("TEST DE SYNTHÈSE NVIDIA MULTI-AGENTS - FERME EXPÉRIMENTALE")
    print("=" * 80)
    
    docx_file = Path(r"C:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia\EXPERIMENTAL\2026-02-16_Ferme Exprimentale_DCE_Lot unique.docx")
    
    if not docx_file.exists():
        print(f"ERREUR : Le fichier Word n'existe pas à : {docx_file}")
        return
        
    print(f"\n[1/3] Analyse du fichier : {docx_file.name}...")
    
    start_time = time.time()
    try:
        outputs = run_dossier_synthesis(
            pdf_paths=[docx_file],
            workspace_root=WORKSPACE_ROOT,
            provider="nvidia",
            api_key=None
        )
        duration = time.time() - start_time
        print(f"\n[2/3] Synthèse réussie avec succès en {duration:.2f} secondes !")
        
        # Chemins des rapports
        print(f"\nRapports générés dans : {outputs.output_dir}")
        print(f"Fiche d'identité : {outputs.identite_path}")
        print(f"Synthèse technique : {outputs.synthese_path}")
        
        # Lecture et affichage des résultats
        print("\n" + "=" * 80)
        print("CONTENU DU RAPPORT - FICHE D'IDENTITÉ :")
        print("" * 80)
        with open(outputs.identite_path, "r", encoding="utf-8") as f:
            print(f.read())
            
        print("\n" + "=" * 80)
        print("CONTENU DU RAPPORT - SYNTHÈSE TECHNIQUE APPROFONDIE :")
        print("=" * 80)
        with open(outputs.synthese_path, "r", encoding="utf-8") as f:
            print(f.read())
            
    except Exception as e:
        print(f"\nERREUR lors de la synthèse : {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_synthesis()
