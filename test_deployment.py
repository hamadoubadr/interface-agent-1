#!/usr/bin/env python3
"""
Script de test pour vérifier que l'application est prête pour le déploiement.
"""

import sys
import os
from pathlib import Path

def check_requirements():
    """Vérifie que les fichiers requis existent."""
    print("VERIFICATION des fichiers requis...")
    
    required_files = [
        "requirements.txt",
        "streamlit_dossier_app/app.py",
        "streamlit_dossier_app/app_core/pdf_enrichment.py",
        "streamlit_dossier_app/app_core/pipeline.py",
        ".streamlit/config.toml"
    ]
    
    all_ok = True
    for file_path in required_files:
        if Path(file_path).exists():
            print(f"  [OK] {file_path}")
        else:
            print(f"  [ERREUR] {file_path} - MANQUANT")
            all_ok = False
    
    return all_ok

def check_imports():
    """Vérifie que les imports principaux fonctionnent."""
    print("\nVERIFICATION des imports...")
    
    try:
        # Test des imports de base
        import streamlit
        import pandas
        import numpy
        print("  [OK] Imports de base")
        
        # Test des imports de l'application
        sys.path.insert(0, str(Path("streamlit_dossier_app").absolute()))
        from app_core.pdf_enrichment import extract_products_from_descriptif, enrich_products_dataframe
        print("  [OK] Imports de l'application")
        
        return True
    except ImportError as e:
        print(f"  [ERREUR] Erreur d'import: {e}")
        return False
    except Exception as e:
        print(f"  [ATTENTION] Autre erreur: {e}")
        return False

def check_dependencies():
    """Vérifie les dépendances critiques."""
    print("\nVERIFICATION des dépendances...")
    
    dependencies = [
        ("streamlit", ">=1.32.0"),
        ("pandas", ">=2.2.0"),
        ("pdfplumber", ">=0.10.0"),
        ("openai", ">=1.12.0"),
        ("sentence-transformers", ">=2.2.0"),
    ]
    
    all_ok = True
    for dep, version in dependencies:
        try:
            module = __import__(dep)
            print(f"  [OK] {dep}{version}")
        except ImportError:
            print(f"  [ERREUR] {dep}{version} - NON INSTALLÉ")
            all_ok = False
    
    return all_ok

def check_app_structure():
    """Vérifie la structure de l'application."""
    print("\nVERIFICATION de la structure de l'application...")
    
    app_dir = Path("streamlit_dossier_app")
    if not app_dir.exists():
        print("  [ERREUR] Dossier streamlit_dossier_app manquant")
        return False
    
    required_dirs = [
        app_dir / "app_core",
        app_dir / ".workspace"
    ]
    
    all_ok = True
    for dir_path in required_dirs:
        if dir_path.exists():
            print(f"  [OK] {dir_path.relative_to('.')}")
        else:
            print(f"  [ATTENTION] {dir_path.relative_to('.')} - CRÉATION")
            dir_path.mkdir(exist_ok=True)
    
    return all_ok

def main():
    print("=" * 60)
    print("TEST de déploiement - Agent DCE")
    print("=" * 60)
    
    tests = [
        ("Fichiers requis", check_requirements),
        ("Structure application", check_app_structure),
        ("Imports", check_imports),
        ("Dépendances", check_dependencies),
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n[TEST] {test_name}")
        result = test_func()
        results.append((test_name, result))
    
    print("\n" + "=" * 60)
    print("RESUME des tests")
    print("=" * 60)
    
    all_passed = True
    for test_name, passed in results:
        status = "[OK] PASSÉ" if passed else "[ERREUR] ÉCHEC"
        print(f"{status} - {test_name}")
        if not passed:
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("[SUCCES] Tous les tests sont passés ! L'application est prête pour le déploiement.")
        print("\n[ETAPES] Prochaines étapes :")
        print("1. Poussez le code sur GitHub")
        print("2. Allez sur share.streamlit.io")
        print("3. Connectez votre compte GitHub")
        print("4. Déployez l'application")
        print("\n[LIEN] Votre application sera accessible sur : https://VOTRE-APP.streamlit.app")
    else:
        print("[ATTENTION] Certains tests ont échoué. Veuillez corriger les problèmes avant le déploiement.")
    
    return 0 if all_passed else 1

if __name__ == "__main__":
    sys.exit(main())