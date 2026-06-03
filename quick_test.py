#!/usr/bin/env python3
"""
Test rapide pour vérifier que l'application fonctionne.
"""

import sys
from pathlib import Path

def test_basic_imports():
    """Test des imports de base."""
    print("Test des imports de base...")
    
    try:
        import streamlit
        print("  OK - streamlit")
    except ImportError:
        print("  ERREUR - streamlit non installé")
        return False
    
    try:
        import pandas
        print("  OK - pandas")
    except ImportError:
        print("  ERREUR - pandas non installé")
        return False
    
    try:
        import pdfplumber
        print("  OK - pdfplumber")
    except ImportError:
        print("  ERREUR - pdfplumber non installé")
        return False
    
    return True

def test_app_imports():
    """Test des imports de l'application."""
    print("\nTest des imports de l'application...")
    
    # Ajouter le chemin de l'application
    app_path = Path("streamlit_dossier_app")
    if not app_path.exists():
        print("  ERREUR - Dossier streamlit_dossier_app manquant")
        return False
    
    sys.path.insert(0, str(app_path.absolute()))
    
    try:
        from app_core.pdf_enrichment import extract_products_from_descriptif, enrich_products_dataframe
        print("  OK - pdf_enrichment")
    except ImportError as e:
        print(f"  ERREUR - pdf_enrichment: {e}")
        return False
    
    try:
        from app_core.pipeline import run_dossier_synthesis
        print("  OK - pipeline")
    except ImportError as e:
        print(f"  ERREUR - pipeline: {e}")
        return False
    
    return True

def test_file_structure():
    """Test de la structure des fichiers."""
    print("\nTest de la structure des fichiers...")
    
    required_files = [
        "requirements.txt",
        "streamlit_dossier_app/app.py",
        ".streamlit/config.toml",
        "Procfile",
        "README.md"
    ]
    
    all_ok = True
    for file_path in required_files:
        if Path(file_path).exists():
            print(f"  OK - {file_path}")
        else:
            print(f"  ERREUR - {file_path} manquant")
            all_ok = False
    
    return all_ok

def main():
    print("=" * 60)
    print("TEST RAPIDE - Agent DCE")
    print("=" * 60)
    
    tests = [
        ("Imports de base", test_basic_imports),
        ("Imports application", test_app_imports),
        ("Structure fichiers", test_file_structure),
    ]
    
    all_passed = True
    for test_name, test_func in tests:
        print(f"\n[TEST] {test_name}")
        if test_func():
            print("  -> PASSÉ")
        else:
            print("  -> ÉCHEC")
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("SUCCES - L'application est prête pour le déploiement !")
        print("\nETAPES pour déployer sur Streamlit Cloud :")
        print("1. Créez un dépôt GitHub (ex: agent-dce)")
        print("2. Poussez votre code :")
        print("   git init")
        print("   git add .")
        print("   git commit -m 'Initial commit'")
        print("   git remote add origin https://github.com/VOTRE-USERNAME/agent-dce.git")
        print("   git push -u origin main")
        print("3. Allez sur https://share.streamlit.io")
        print("4. Connectez votre compte GitHub")
        print("5. Sélectionnez votre dépôt et déployez")
        print("\nLIEN : Votre application sera sur https://VOTRE-APP.streamlit.app")
    else:
        print("ATTENTION - Des problèmes doivent être corrigés avant le déploiement.")
    
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())