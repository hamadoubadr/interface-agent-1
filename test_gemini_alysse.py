import os
import sys
import subprocess
from pathlib import Path

# Configuration des chemins
current_dir = Path(os.getcwd())
# On utilise la racine du projet pour accéder au binaire C#
project_root = current_dir.parent 
sys.path.append(str(current_dir / "streamlit_dossier_app"))

def run_gemini_synthesis_test():
    print("="*80)
    print("TEST SYNTHÈSE GEMINI - PROJET ALYSSE")
    print("="*80)

    # 1. Définition des chemins
    alysse_dir = Path(r"C:\Users\Admin\Documents\stage pfe\agent antigravity a marche - Copie\AGENT ANTIGRAVITY\ALYSSE")
    pdf_files = [
        alysse_dir / "ALYSSE-DESCRIPTIF_FLUIDES - IND C.pdf",
        alysse_dir / "Projet ALYSSE - Consultation restreinte Fluide.pdf"
    ]
    
    # Clé Gemini fournie
    gemini_key = "AIzaSyAe9JO12Jd9EOHsBvqjxlNoULefwRM4pGw"
    
    # 2. Création du dossier de sortie
    output_root = current_dir / "test_results_alysse_gemini"
    output_root.mkdir(parents=True, exist_ok=True)
    
    # On prépare le dossier input pour le binaire C#
    input_dir = output_root / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    import shutil
    for f in pdf_files:
        if f.exists():
            shutil.copy2(f, input_dir / f.name)
        else:
            print(f"ERREUR : Fichier introuvable : {f}")
            return

    output_dir = output_root / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    # 3. Lancement du binaire C# (Moteur Gemini original)
    print("\n[1/2] Lancement du moteur de synthèse Gemini (C#)...")
    
    # Chemin vers le binaire
    exe_path = project_root / "AGENT ANTIGRAVITY" / "ConstructionDossierAgent.exe"
    
    if not exe_path.exists():
        print(f"ERREUR : Binaire C# introuvable à {exe_path}")
        return

    command = [
        str(exe_path),
        "--input", str(input_dir),
        "--output", str(output_dir),
        "--provider", "gemini",
        "--api-key", gemini_key
    ]

    try:
        # On définit la variable d'environnement pour être sûr
        env = os.environ.copy()
        env["GEMINI_API_KEY"] = gemini_key
        
        result = subprocess.run(
            command,
            cwd=str(exe_path.parent),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env
        )
        
        if result.returncode != 0:
            print(f"  ERREUR du moteur C# : {result.stderr or result.stdout}")
            return
        
        print("  - Synthèse terminée avec succès.")
        
    except Exception as e:
        print(f"  ERREUR lors de l'exécution : {e}")
        return

    # 4. Vérification des résultats
    print("\n[2/2] Vérification des rapports générés...")
    
    # Le binaire crée un sous-dossier avec un timestamp
    # On cherche le dossier le plus récent dans output
    subdirs = sorted(output_dir.glob("*/run-*"), key=os.path.getmtime, reverse=True)
    if not subdirs:
        # Essayer un autre pattern si nécessaire
        subdirs = sorted(output_dir.glob("run-*"), key=os.path.getmtime, reverse=True)
        if not subdirs:
            # Fallback : chercher directement les fichiers .md dans l'arborescence
            md_files = list(output_dir.rglob("*.md"))
            if not md_files:
                print("  ERREUR : Aucun fichier Markdown généré.")
                return
            final_dir = md_files[0].parent
        else:
            final_dir = subdirs[0]
    else:
        final_dir = subdirs[0]

    identite_path = final_dir / "identite-projet.md"
    synthese_path = final_dir / "synthese-detaillee.md"

    print("\n" + "="*80)
    print("TEST TERMINÉ")
    print(f"Rapports générés dans : {final_dir}")
    if identite_path.exists(): print(f"1. {identite_path.name} ({identite_path.stat().st_size} octets)")
    if synthese_path.exists(): print(f"2. {synthese_path.name} ({synthese_path.stat().st_size} octets)")
    print("="*80)

if __name__ == "__main__":
    run_gemini_synthesis_test()
