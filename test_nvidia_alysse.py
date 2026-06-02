"""

Script générique d'analyse de PDF via le pipeline Multi-Agents NVIDIA.

Usage :

  python test_nvidia_alysse.py [dossier_pdfs] [dossier_sortie]



Si aucun argument n'est fourni, les valeurs par défaut pointent sur le

dossier ALYSSE du projet courant.



Le script :

  1. Découvre automatiquement tous les PDF présents dans le dossier source.

  2. Génère une fiche d'identité du projet (maître d'ouvrage, architecte,

     intervenants, localisation, délai, etc.) — universelle.

  3. Génère une synthèse technique descriptive des lots / prestations — universelle.

  4. Sauvegarde les deux rapports en Markdown dans le dossier de sortie.

"""



import argparse

import os

import sys

from pathlib import Path



# ---------------------------------------------------------------------------

# Résolution dynamique du chemin vers app_core

# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent

sys.path.insert(0, str(SCRIPT_DIR / "streamlit_dossier_app"))



from app_core.multi_agent_pipeline import MultiAgentChatbot  # noqa: E402





# ---------------------------------------------------------------------------

# Questions génériques (ni ALYSSE-spécifiques, ni fluides-spécifiques)

# ---------------------------------------------------------------------------

QUESTION_IDENTITE = (

    "Fais une fiche d'identité complète du projet en listant, si disponibles : "

    "Maître d'ouvrage, Architecte / Maître d'œuvre, Bureaux d'études techniques (BET), "

    "Bureau de contrôle, situation géographique / adresse, objet / nature du marché, "

    "délai des travaux, et tout autre intervenant mentionné (coordinateur, économiste, etc.)."

)



QUESTION_SYNTHESE = (

    "Fais une synthèse technique détaillée de toutes les prestations et lots décrits "

    "dans les documents : liste chaque lot ou domaine technique (exemples : gros œuvre, "

    "charpente, façades, plomberie, électricité, CVC, etc.), décris les matériaux, "

    "équipements et normes utilisés, et signale les points de vigilance importants."

)





# ---------------------------------------------------------------------------

# Valeurs par défaut

# ---------------------------------------------------------------------------

DEFAULT_PDF_DIR = (

    SCRIPT_DIR.parent

    / "AGENT ANTIGRAVITY"

    / "ALYSSE"

)





def parse_args() -> argparse.Namespace:

    parser = argparse.ArgumentParser(

        description="Analyse générique de PDF avec le pipeline Multi-Agents NVIDIA."

    )

    parser.add_argument(

        "pdf_dir",

        nargs="?",

        default=str(DEFAULT_PDF_DIR),

        help=(

            "Dossier contenant les fichiers PDF à analyser "

            f"(défaut : {DEFAULT_PDF_DIR})"

        ),

    )

    parser.add_argument(

        "output_dir",

        nargs="?",

        default=None,

        help=(

            "Dossier de sortie pour les rapports Markdown "

            "(défaut : test_results_<nom_du_dossier>_nvidia/ dans le répertoire du script)"

        ),

    )

    return parser.parse_args()





def discover_pdfs(directory: Path) -> list[Path]:

    """Trouve récursivement tous les PDF dans *directory*."""

    return sorted(directory.glob("**/*.pdf"))





def slugify(name: str) -> str:

    """Transforme un nom de dossier en slug pour les noms de fichiers."""

    return name.lower().replace(" ", "_").replace("-", "_")





def run(pdf_dir: Path, output_dir: Path) -> None:

    print("=" * 80)

    print(f"ANALYSE PDF - {pdf_dir.name.upper()}")

    print("=" * 80)



    # 1. Découverte des PDF

    pdf_files = discover_pdfs(pdf_dir)

    if not pdf_files:

        print(f"ERREUR : aucun fichier PDF trouvé dans {pdf_dir}")

        sys.exit(1)



    print(f"\n{len(pdf_files)} PDF(s) détecté(s) :")

    for f in pdf_files:

        print(f"  • {f.name}")



    # 2. Initialisation du pipeline Multi-Agents NVIDIA

    print("\n[1/3] Initialisation du Chatbot Multi-Agents NVIDIA...")

    try:

        bot = MultiAgentChatbot(use_nvidia_embeddings=True)

        print("  - Embeddings  : Llama-3.2 NemoRetriever (NVIDIA)")

        print("  - Extracteur  : Qwen3 Coder (NVIDIA)")

        print(f"  - Synthétiseur: {bot.synthesizer.model}")

    except Exception as exc:

        print(f"  ERREUR d'initialisation : {exc}")

        sys.exit(1)



    # 3. Indexation (RAG)

    print(f"\n[2/3] Indexation de {len(pdf_files)} PDF(s)...")

    try:

        bot.load_documents(pdf_files)

        print("  - Indexation terminée avec succès.")

    except Exception as exc:

        print(f"  ERREUR lors de l'indexation : {exc}")

        sys.exit(1)



    # 4. Génération des rapports

    print("\n[3/3] Génération des rapports via Multi-Agents...")



    questions = {

        "identite": QUESTION_IDENTITE,

        "synthese": QUESTION_SYNTHESE,

    }



    results: dict[str, str] = {}

    for key, question in questions.items():

        print(f"  - Analyse : {key}...")

        try:

            response = bot.ask(question)

            results[key] = response["answer"]

            print(f"    OK ({len(results[key])} caractères)")

        except Exception as exc:

            results[key] = f"Erreur lors de l'analyse : {exc}"

            print(f"    ERREUR : {exc}")



    # 5. Sauvegarde

    output_dir.mkdir(parents=True, exist_ok=True)

    slug = slugify(pdf_dir.name)



    identite_path = output_dir / f"identite-{slug}-nvidia.md"

    synthese_path = output_dir / f"synthese-{slug}-nvidia.md"



    project_title = pdf_dir.name.upper()



    identite_path.write_text(

        f"# IDENTITÉ DU PROJET - {project_title} (NVIDIA)\n\n{results['identite']}",

        encoding="utf-8",

    )

    synthese_path.write_text(

        f"# SYNTHÈSE TECHNIQUE - {project_title} (NVIDIA)\n\n{results['synthese']}",

        encoding="utf-8",

    )



    print("\n" + "=" * 80)

    print("ANALYSE TERMINÉE")

    print(f"Rapports générés dans : {output_dir}")

    print(f"  1. {identite_path.name}")

    print(f"  2. {synthese_path.name}")

    print("=" * 80)





def main() -> None:

    args = parse_args()



    pdf_dir = Path(args.pdf_dir).resolve()

    if not pdf_dir.is_dir():

        print(f"ERREUR : le dossier '{pdf_dir}' n'existe pas.")

        sys.exit(1)



    if args.output_dir:

        output_dir = Path(args.output_dir).resolve()

    else:

        slug = slugify(pdf_dir.name)

        output_dir = SCRIPT_DIR / f"test_results_{slug}_nvidia"



    run(pdf_dir, output_dir)





if __name__ == "__main__":

    main()

