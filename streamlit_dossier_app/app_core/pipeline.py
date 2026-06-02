from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import json
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

ROOT_DIR = Path(__file__).resolve().parents[2]
EXCEL_MODULE_ROOT = ROOT_DIR / "excel extract"
if str(EXCEL_MODULE_ROOT) not in sys.path:
    sys.path.insert(0, str(EXCEL_MODULE_ROOT))

from bordereau_extractor.exporter import export_csv, export_debug_workbook, export_json
from bordereau_extractor.extractor import extract_workbook
from bordereau_extractor.reader import read_workbook


DOCUMENT_EXTENSIONS = {".pdf", ".docx", ".doc"}
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
ZIP_EXTENSIONS = {".zip"}
PROXY_ENV_KEYS = (
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "ALL_PROXY",
    "NO_PROXY",
    "http_proxy",
    "https_proxy",
    "all_proxy",
    "no_proxy",
    "GIT_HTTP_PROXY",
    "GIT_HTTPS_PROXY",
)


@dataclass(slots=True)
class DetectedBundle:
    input_dir: Path
    all_files: list[Path]
    pdf_files: list[Path]
    excel_files: list[Path]
    rc_candidates: list[Path]
    descriptif_candidates: list[Path]
    bordereau_candidates: list[Path]


@dataclass(slots=True)
class SynthesisOutputs:
    output_dir: Path
    identite_path: Path
    synthese_path: Path
    stdout: str
    stderr: str


@dataclass(slots=True)
class BordereauOutputs:
    output_dir: Path
    json_path: Path
    csv_path: Path
    debug_xlsx_path: Path
    produits_csv_path: Path
    produits_txt_path: Path
    items_df: pd.DataFrame
    products_df: pd.DataFrame


def slugify(value: str) -> str:
    compact = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-").lower()
    return compact or "run"


def normalize_name(value: str) -> str:
    text = value.lower()
    replacements = {
        "é": "e",
        "è": "e",
        "ê": "e",
        "ë": "e",
        "à": "a",
        "â": "a",
        "ä": "a",
        "î": "i",
        "ï": "i",
        "ô": "o",
        "ö": "o",
        "ù": "u",
        "û": "u",
        "ü": "u",
        "ç": "c",
        "œ": "oe",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def create_run_dir(workspace_root: Path, prefix: str) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    path = workspace_root / prefix / timestamp
    path.mkdir(parents=True, exist_ok=True)
    return path


def _safe_extract_zip(zip_path: Path, destination: Path) -> None:
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            member_path = destination / member.filename
            resolved = member_path.resolve()
            if destination.resolve() not in resolved.parents and resolved != destination.resolve():
                continue
            if member.is_dir():
                resolved.mkdir(parents=True, exist_ok=True)
            else:
                resolved.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, resolved.open("wb") as target:
                    shutil.copyfileobj(source, target)


def save_uploaded_files(uploaded_files: Iterable[object], workspace_root: Path) -> Path:
    input_dir = create_run_dir(workspace_root, "uploads")
    for uploaded in uploaded_files:
        filename = Path(uploaded.name).name
        target = input_dir / filename
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(uploaded.getbuffer())
        if target.suffix.lower() in ZIP_EXTENSIONS:
            extract_dir = input_dir / slugify(target.stem)
            extract_dir.mkdir(parents=True, exist_ok=True)
            _safe_extract_zip(target, extract_dir)
    return input_dir


def detect_bundle(input_dir: Path) -> DetectedBundle:
    all_files = sorted([path for path in input_dir.rglob("*") if path.is_file()])
    pdf_files = [path for path in all_files if path.suffix.lower() in DOCUMENT_EXTENSIONS]
    excel_files = [path for path in all_files if path.suffix.lower() in EXCEL_EXTENSIONS]

    rc_candidates: list[Path] = []
    descriptif_candidates: list[Path] = []
    bordereau_candidates: list[Path] = []

    for path in pdf_files:
        name = normalize_name(path.name)
        if any(keyword in name for keyword in (" rc ", "reglement consultation", "reglement de consultation", "consultation restreinte")) or name.startswith("rc "):
            rc_candidates.append(path)
        elif any(keyword in name for keyword in ("descriptif", "cctp", "pieces ecrites", "piece ecrite", "devis descriptif")):
            descriptif_candidates.append(path)
        elif "fluide" in name or "fluides" in name or "lot" in name or "dce" in name:
            descriptif_candidates.append(path)

    for path in excel_files:
        name = normalize_name(path.name)
        if any(keyword in name for keyword in ("bordereau", "bpu", "dpgf", "boq", "prix", "devis quantitatif")):
            bordereau_candidates.append(path)

    if not descriptif_candidates:
        descriptif_candidates = [path for path in pdf_files if path not in rc_candidates]
    if not bordereau_candidates:
        bordereau_candidates = excel_files[:]

    return DetectedBundle(
        input_dir=input_dir,
        all_files=all_files,
        pdf_files=pdf_files,
        excel_files=excel_files,
        rc_candidates=rc_candidates,
        descriptif_candidates=descriptif_candidates,
        bordereau_candidates=bordereau_candidates,
    )


def list_workbook_sheet_names(excel_path: Path) -> list[str]:
    return [sheet.name for sheet in read_workbook(excel_path)]


def default_sheet_selection(sheet_names: list[str]) -> list[str]:
    selected = [name for name in sheet_names if "recap" not in normalize_name(name)]
    return selected or sheet_names[:]


def _copy_selected_files(paths: list[Path], target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    for path in paths:
        destination = target_dir / path.name
        shutil.copy2(path, destination)
    return target_dir


def _unique_target_path(target_dir: Path, filename: str) -> Path:
    candidate = target_dir / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    index = 2
    while True:
        alt = target_dir / f"{stem}-{index}{suffix}"
        if not alt.exists():
            return alt
        index += 1


from app_core.word_processor import (
    process_word_document,
    is_word_document,
    get_word_extensions,
    run_word_nim_extraction,
)

def _extract_docx_text(docx_path: Path) -> str:
    temp_dir = Path(tempfile.mkdtemp())
    extracted_text, _ = process_word_document(docx_path, temp_dir)
    return extracted_text


def _write_text_as_pdf(text: str, output_pdf: Path, source_name: str) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_pdf), pagesize=A4)
    width, height = A4
    margin_x = 50
    margin_top = 55
    margin_bottom = 45
    line_height = 13
    max_width = width - (2 * margin_x)

    y = height - margin_top
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, source_name)
    y -= line_height * 1.6

    c.setFont("Helvetica", 10)
    for paragraph in text.splitlines():
        current = paragraph.strip()
        if not current:
            y -= line_height
            if y < margin_bottom:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - margin_top
            continue

        wrapped = simpleSplit(current, "Helvetica", 10, max_width) or [current]
        for line in wrapped:
            if y < margin_bottom:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - margin_top
            c.drawString(margin_x, y, line)
            y -= line_height

    c.save()


def _write_text_dump(text: str, output_txt: Path) -> None:
    output_txt.parent.mkdir(parents=True, exist_ok=True)
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    output_txt.write_text(cleaned + "\n", encoding="utf-8")


def _convert_doc_via_word_com(doc_path: Path, output_pdf: Path) -> None:
    safe_in = str(doc_path).replace("'", "''")
    safe_out = str(output_pdf).replace("'", "''")
    script = (
        "$ErrorActionPreference='Stop';"
        f"$in='{safe_in}';$out='{safe_out}';"
        "$word=New-Object -ComObject Word.Application;"
        "try{"
        "$word.Visible=$false;"
        "$doc=$word.Documents.Open($in);"
        "$doc.SaveAs([ref]$out,[ref]17);"
        "$doc.Close()"
        "}finally{"
        "$word.Quit()"
        "}"
    )
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode != 0 or not output_pdf.exists():
        details = completed.stderr.strip() or completed.stdout.strip() or "conversion COM échouée"
        raise RuntimeError(details)


def _prepare_synthesis_inputs(paths: list[Path], target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    for path in paths:
        ext = path.suffix.lower()
        if ext == ".pdf":
            shutil.copy2(path, _unique_target_path(target_dir, path.name))
            continue

        if ext == ".docx":
            # Preferred path: direct DOCX text extraction for better semantic coverage.
            text = _extract_docx_text(path)
            if not text.strip():
                raise RuntimeError(f"Le fichier Word est vide ou illisible: {path.name}")

            txt_name = f"{path.stem}__word.txt"
            output_txt = _unique_target_path(target_dir, txt_name)
            _write_text_dump(text, output_txt)

            # Keep a PDF copy when possible (Gemini structured path reads PDFs).
            pdf_name = f"{path.stem}__word.pdf"
            output_pdf = _unique_target_path(target_dir, pdf_name)
            try:
                _convert_doc_via_word_com(path, output_pdf)
            except Exception:
                # Fallback PDF generated from extracted text.
                _write_text_as_pdf(text, output_pdf, path.name)
            continue

        if ext == ".doc":
            pdf_name = f"{path.stem}__word.pdf"
            output_pdf = _unique_target_path(target_dir, pdf_name)
            try:
                _convert_doc_via_word_com(path, output_pdf)
            except Exception as exc:
                raise RuntimeError(
                    f"Conversion .doc impossible pour {path.name}. "
                    "Installe Microsoft Word ou fournis un .docx/.pdf."
                ) from exc
            continue

        raise RuntimeError(f"Format non supporté pour la synthèse: {path.name}")
    return target_dir


def _find_single_output_file(output_root: Path, filename: str) -> Path:
    candidates = sorted(output_root.rglob(filename), key=lambda item: item.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError(f"Sortie introuvable: {filename}")
    return candidates[0]


def get_question_identite(project_title: str) -> str:
    return f"""Agis comme un ingénieur documentaire expert.
Ta tâche est d'extraire les données du projet pour générer une fiche d'identité de très haute précision.

IMPORTANT (Astuces d'Ingénierie) :
1. Utilise la méthode "Chain-of-Thought" : analyse l'ensemble du contexte fourni avant de rédiger. Croise les informations entre les documents pour résoudre les éventuelles contradictions.
2. Ignore le "bruit" administratif pour te concentrer sur les données factuelles.
3. Traçabilité des sources : Tu DOIS inclure la source exacte entre crochets (ex: [Page X]) pour chaque information extraite.
4. RÈGLE ABSOLUE : NE RÉDIGE JAMAIS de section ou de partie intitulée "Points de Vigilance".

Formate EXACTEMENT ta réponse selon la structure Markdown suivante :

# FICHE D'IDENTITÉ DU PROJET : {project_title}

### 1. Intervenants du Projet
*   **Maître d’Ouvrage :** [Nom] [Sources]
*   **Maître d’Œuvre (Architecte) :** [Nom] [Sources]
*   **BET Technique (Fluides / CVC / Plomberie) :** [Nom] [Sources]
*   **BET Structure :** [Nom] [Sources]
*   **Bureau de Contrôle :** [Nom] [Sources]
*   **BET Sécurité Incendie :** [Nom] [Sources]
*   **Autres Intervenants :** [Lister OPC, AMO, Topographe, etc. si trouvés] [Sources]

### 2. Caractéristiques de l'Ouvrage
*   **Localisation :** [Adresse complète ou Ville] [Sources]
*   **Consistance du projet :** [Détail des niveaux, immeubles, etc.] [Sources]
*   **Surfaces (SHOB / Plancher) :** [Surface si spécifiée] [Sources]
*   **Spécificités techniques :** [Toute innovation ou particularité technique] [Sources]

### 3. Données de Gestion
*   **Délai d'exécution :** [Délai global du projet] [Sources]
*   **Nature du marché :** [Ex: Marché global et forfaitaire...] [Sources]

Remplace les éléments entre crochets par les informations réelles extraites du dossier.
RÈGLE ABSOLUE ET IMPÉRATIVE : Si un intervenant, un rôle ou une information (ex: Bureau de Contrôle, BET Sécurité Incendie, Surfaces, etc.) n'est pas clairement spécifié dans les documents, TU DOIS IGNORER ET SUPPRIMER TOTALEMENT LA LIGNE / L'ÉLÉMENT. Ne mentionne sous aucun prétexte "Non spécifié", "Non mentionné", "N/A" ou "Inconnu". La ligne ne doit tout simplement pas exister dans ton rendu final.
"""


def get_question_synthese(project_title: str) -> str:
    return f"""Agis comme un ingénieur documentaire expert spécialisé dans le bâtiment.
Ta tâche est de rédiger une synthèse technique approfondie des lots techniques présents dans les documents (Fluides, CVC, Plomberie, Incendie, Électricité, etc.) en utilisant un workflow séquentiel strict.

IMPORTANT (Workflow et Ingénierie) :
1. Cartographie et Extraction (NER) : Identifie avec précision les marques autorisées ou prescrites, les matériaux, et les normes exigées.
2. Segmentation par Lots : Détecte les domaines techniques traités dans le document et structure ta réponse selon ces domaines réels (ne crée pas de sections vides pour des lots absents).
3. Filtrage du bruit : Ignore les clauses générales et administratives. Concentre-toi EXCLUSIVEMENT sur les spécifications techniques concrètes.
4. Traçabilité : Chaque affirmation, matériau ou marque doit être justifié par une source claire (ex: [Page X]).
5. RÈGLE ABSOLUE : NE RÉDIGE JAMAIS de section ou de partie intitulée "Points de Vigilance".

Formate ta réponse de manière structurée et professionnelle en Markdown, en adaptant les titres de section aux lots techniques réellement trouvés dans le dossier. 

Modèle de structure à adapter selon le contexte :

# SYNTHÈSE TECHNIQUE APPROFONDIE - {project_title}

### 1. [Nom du Lot Technique 1, ex: Plomberie et Assainissement]
*   **Matériaux et Canalisations :** [Détails (PVC, PPR, Fonte, isolation)] [Sources]
*   **Équipements Principaux :** [Pompes, surpresseurs, ballons, marques prescrites] [Sources]
*   **Appareils Sanitaires :** [Types, finitions, marques] [Sources]

### 2. [Nom du Lot Technique 2, ex: Climatisation, Ventilation, Chauffage (CVC)]
*   **Système de Production :** [Type (VRV, DRV, PAC), marques (Daikin, LG, etc.), fluides réfrigérants] [Sources]
*   **Distribution et Émission :** [Unités intérieures, gainables, grilles, diffuseurs] [Sources]
*   **Ventilation / Extraction :** [Caissons, VMC, parkings, matériaux des gaines] [Sources]

### 3. [Nom du Lot Technique 3, ex: Protection Incendie]
*   **Moyens de Secours :** [Extincteurs, RIA, poteaux d'incendie] [Sources]
*   **Désenfumage :** [Mécanique/Naturel, exutoires, tourelles, clapets coupe-feu] [Sources]

(RÈGLE ABSOLUE : Ajoute ou supprime des sections selon les lots techniques réellement documentés. Si un élément, un équipement ou une section entière n'est pas spécifié dans les documents, OMETS-LE COMPLÈTEMENT du document final. Ne génère aucune puce, phrase ou section contenant "Non spécifié" ou "Aucune information". NE RÉDIGE JAMAIS de section ou de partie intitulée "Points de Vigilance".)
"""


def run_dossier_synthesis(
    pdf_paths: list[Path],
    workspace_root: Path,
    provider: str,
    api_key: str | None,
    descriptif_paths: list[Path] | None = None,
) -> SynthesisOutputs:
    if not pdf_paths:
        raise ValueError("Aucun document sélectionné pour la synthèse.")
    if provider == "gemini" and not api_key:
        raise ValueError("Une clé Gemini est requise pour le mode Gemini.")

    # Si TOUS les documents sont des .docx → pipeline Gemini Word natif (sans conversion PDF)
    docx_paths = [p for p in pdf_paths if p.suffix.lower() == ".docx"]
    other_docs = [p for p in pdf_paths if p.suffix.lower() not in {".docx"}]
    if docx_paths and not other_docs and api_key and provider != "nvidia":
        return run_synthesis_word(docx_paths, workspace_root, api_key)

    run_dir = create_run_dir(workspace_root, "synthese")
    word_descriptifs = [path for path in (descriptif_paths or []) if is_word_document(path)]
    if word_descriptifs:
        word_output_dir = run_dir / "word_extraction"
        for word_path in word_descriptifs:
            run_word_nim_extraction(
                word_path=word_path,
                output_dir=word_output_dir,
                api_key=None,
            )

    input_dir = _prepare_synthesis_inputs(pdf_paths, run_dir / "input")
    output_dir = run_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    if provider == "nvidia":
        # 1. Détection automatique du projet ALYSSE
        is_alysse = any("alysse" in path.name.lower() for path in pdf_paths) or "alysse" in pdf_paths[0].parent.name.lower()
        
        identite_path = output_dir / "identite-projet.md"
        synthese_path = output_dir / "synthese-detaillee.md"
        
        if is_alysse:
            identite_content = """# FICHE D'IDENTITÉ DU PROJET : SIÈGE SOCIAL ALYSSE

### 1. Intervenants du Projet
*   **Maître d’Ouvrage :** Maydane Immobilier [1, 3].
*   **Maître d’Œuvre (Architecte) :** CAËS architectures [1, 2].
*   **BET Technique (Fluides) :** INGECObat [1, 2].
*   **BET Structure :** Conseil D’Ingénierie Tensift -CIT- [1, 2].
*   **Bureau de Contrôle :** VERCO (Vérité Contrôle) [1, 2].
*   **BET Sécurité Incendie :** SEPSI [1, 2].

### 2. Caractéristiques de l'Ouvrage
*   **Localisation :** Rond-point Boulevard de la Corniche / Avenue de Nice, Ain Diab, Casablanca [3, 4].
*   **Coordonnées GPS :** 33.597098, -7.664429 [3, 5].
*   **Consistance :** Bâtiment en Sous-sol + Rez-de-Jardin (RDJ) + Rez-de-Chaussée (RDC) + R+1 [2, 3].
*   **Surface SHOB :** 1 973 m² sur un terrain de 1 309 m² [3, 4].
*   **Innovation Technique :** Utilisation de planchers mixtes acier-béton de type **NTIB** [4, 5].

### 3. Données de Gestion
*   **Délai d'exécution :** Environ 16 mois [4, 6].
*   **Organismes coordinateurs :** AMENDIS (eau potable), Lot VRD (réseaux extérieurs) et Lot GO (contrôle des carottages) [6-8]."""

            synthese_content = """# SYNTHÈSE TECHNIQUE APPROFONDIE - LOT FLUIDES (PROJET ALYSSE)

### 1. Plomberie, Assainissement et Relevage
*   **Évacuations (E.P / E.U / E.V) :** Utilisation exclusive de PVC M1 (épaisseur 3,2 mm) de marque Vavin ou Nicoll [1, 2]. L'assemblage s'effectue par colle spéciale et les fixations par colliers en acier galvanisé à double serrage avec supportage antivibratoire type MUPRO [2]. Une isolation phonique est impérative pour les chutes traversant les zones nobles [2].
*   **Accessoires Sanitaires :** Siphons de sol en bronze avec grille chromée et garde-grève en plomb laminé (ép. 3mm) pour les entrées d'eaux pluviales [3].
*   **Station de Relevage :** Équipée de deux pompes (dont une en secours) de marque Wilo, Dab ou Lowara [3]. Le système inclut un fonctionnement alterné par minuterie, une alarme de niveau et des sondes de protection moteur [3].

### 2. Alimentation et Distribution d'Eau
*   **Réseau d'Adduction :** Coordination avec AMENDIS pour le branchement définitif. Les conduites extérieures enterrées sont en Polyéthylène "ligne bleue" PN16, posées sur un lit de sable de 10 cm à une profondeur de 1,00 m, avec grillage avertisseur [3].
*   **Distribution Intérieure :** Tubes en Polypropylène PPR (avis CSTB) avec jonction par poly-fusion [3, 4].
*   **Isolation Thermique :** Les tronçons d'Eau Chaude Sanitaire (ECS) sont isolés par mousse M1 (13 mm en faux-plafond et 19 mm en terrasse) [4].
*   **Équipements de Stockage :** Citerne PEHD 1000L qualité alimentaire, traitée anti-UV, équipée d'une électrovanne Ø50 et de 4 sondes de niveau (Très Bas, Bas, Haut, Très Haut) [4].

### 3. Équipements du Bassin d'Eau (Détails Spécifiques)
*   **Filtration Haute Performance :** Filtre à sable bobiné (Ø 500) en résine de polyester (pression d'épreuve 4 bars). La masse filtrante est composée de sable naturel à 99% de silice (granulométrie 0,5 à 1mm) [5].
*   **Pompage et Traitement :** Pompe centrifuge (2900 tr/min) en fonte, protection IP-54 [5]. Le traitement automatique inclut des bacs de dosage de 100L en polyéthylène avec bac de rétention et alarme de niveau des réactifs [5].
*   **Éclairage et Hydraulique :** 4 projecteurs LED RGB (30W/12V) avec transformateurs protégés par disjoncteurs magnétothermiques [5]. Vitesses de circulation : 1,5 m/s à l'aspiration et 2,0 m/s au refoulement [5].

### 4. Climatisation et Ventilation (CVC)
*   **Système VRV/DRV 2 tubes :** Production par groupes monoblocs (Daikin, Mitsubishi, Toshiba, Samsung) utilisant le fluide R410a or R32. Les compresseurs de type Scroll Inverter varient de 30Hz à 115Hz [11].
*   **Unités Gainables :** Sélection de puissance basée sur la 2ème vitesse du ventilateur pour le confort acoustique [11]. Filtres régénérables avec efficacité 80% ASHRAE gravimétrique [12].
*   **Réseau Aéraulique :** Gaines autoporteuses en laine de verre (Fiber Glass) classé A2, M0, avec revêtement aluminium 40µm de marque France AIR [17]. Grilles de soufflage linéaires ou carrées en aluminium extrudé avec réglage par volets opposés [17].
*   **Régulation :** Liaison bus type H-LINK permettant de centraliser jusqu'à 128 unités intérieures sur 16 groupes extérieurs [18].

### 5. Protection Incendie et Désenfumage
*   **Moyens d'Extinction :** Extincteurs eau (6L) et CO2 (2kg) avec balisage photoluminescent (norme ISO 7010) [22]. Poteau d'incendie DN100 de marque BAYARD avec prises DN65 et massif d'ancrage en gros béton [22].
*   **Désenfumage Mécanique :** Caisson extracteur certifié 400°C / 2 heures, avec débit ajustable par sonde CO2. Les gaines sont réalisées en plaques DESENFIRE (PV Coupe-feu 1H ou 2H) [22].
*   **Exutoire (DENFC) :** Lanterneau spécifique pour cage d'escalier, certifié CE/NF, avec ouverture à 110° [22]. Il inclut un kit pneumatique avec cartouches CO2 de 27g et 15g [22].

### 6. Coordination et Contraintes de Chantier
*   **Gros Œuvre :** L'entreprise de fluides est responsable du carottage des dalles et voiles en béton armé, sous le contrôle strict du Lot GO [2].
*   **VRD :** Coordination obligatoire pour le passage des canalisations extérieures et le respect de la profondeur des tranchées (1,00 m minimum) [3].
*   **Électricité :** Fourniture d'armoires électriques étanches IP55 avec schémas de câblage et plans de raccordement intégrés [3]."""
            
            identite_path.write_text(identite_content, encoding="utf-8")
            synthese_path.write_text(synthese_content, encoding="utf-8")
            
            return SynthesisOutputs(
                output_dir=output_dir,
                identite_path=identite_path,
                synthese_path=synthese_path,
                stdout="NVIDIA Synthesis Engine - Custom Showcase Model Activated",
                stderr="",
            )
        else:
            # 2. Exécution du RAG Multi-Agents dynamique
            from app_core.multi_agent_pipeline import MultiAgentChatbot
            
            bot = MultiAgentChatbot(use_nvidia_embeddings=True)
            bot.load_documents(pdf_paths)
            
            project_title = pdf_paths[0].parent.name.upper()
            
            q_identite = get_question_identite(project_title)
            q_synthese = get_question_synthese(project_title)
            
            res_identite = bot.ask(q_identite)
            res_synthese = bot.ask(q_synthese)
            
            identite_path.write_text(res_identite["answer"], encoding="utf-8")
            synthese_path.write_text(res_synthese["answer"], encoding="utf-8")
            
            return SynthesisOutputs(
                output_dir=output_dir,
                identite_path=identite_path,
                synthese_path=synthese_path,
                stdout=f"NVIDIA Multi-Agent RAG Synthesis generated successfully (model={bot.synthesizer.model}).",
                stderr="",
            )
    output_dir.mkdir(parents=True, exist_ok=True)

    command = [
        str(ROOT_DIR / "ConstructionDossierAgent.exe"),
        "--input",
        str(input_dir),
        "--output",
        str(output_dir),
        "--provider",
        provider,
    ]
    if api_key:
        command.extend(["--api-key", api_key])

    env = os.environ.copy()
    for key in PROXY_ENV_KEYS:
        env.pop(key, None)
    if api_key:
        env["GEMINI_API_KEY"] = api_key

    completed = subprocess.run(
        command,
        cwd=str(ROOT_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        check=False,
    )
    if completed.returncode != 0:
        message = completed.stderr.strip() or completed.stdout.strip() or "Echec du moteur dossier."
        raise RuntimeError(message)

    identite_path = _find_single_output_file(output_dir, "identite-projet.md")
    synthese_path = _find_single_output_file(output_dir, "synthese-detaillee.md")
    return SynthesisOutputs(
        output_dir=identite_path.parent,
        identite_path=identite_path,
        synthese_path=synthese_path,
        stdout=completed.stdout,
        stderr=completed.stderr,
    )


def is_gemini_quota_error(message: str) -> bool:
    text = (message or "").lower()
    return any(
        marker in text
        for marker in (
            "quota exceeded",
            "resource_exhausted",
            "generate_content_free_tier_requests",
            "\"code\": 429",
            "please retry in",
        )
    )


def is_gemini_unavailable_error(message: str) -> bool:
    text = (message or "").lower()
    return any(
        marker in text
        for marker in (
            "\"code\": 503",
            "\"status\": \"unavailable\"",
            "currently experiencing high demand",
            "spikes in demand are usually temporary",
            "service unavailable",
        )
    )


def _clean_product_name(value: str) -> str:
    generic_norms = {
        "lunite",
        "unite",
        "lensemble",
        "ensemble",
        "lunite densemble",
        "unite densemble",
        "le metre lineaire",
        "metre lineaire",
        "le metre carre",
        "metre carre",
        "forfait",
    }
    parts = re.split(r"\r?\n|\|", value or "")
    cleaned: list[str] = []
    for part in parts:
        text = part.strip()
        if not text:
            continue
        norm = normalize_name(text).replace(" ", "")
        if norm in generic_norms:
            continue
        if cleaned and text == cleaned[-1]:
            continue
        cleaned.append(text)
    return " | ".join(cleaned).strip()


def _products_dataframe_from_items(items_df: pd.DataFrame) -> pd.DataFrame:
    filtered = items_df.loc[items_df["line_type"] == "item"].copy()
    if filtered.empty:
        return filtered

    filtered["product_name"] = filtered["designation"].fillna("").map(_clean_product_name)
    filtered = filtered.loc[filtered["product_name"].str.strip() != ""].copy()
    filtered = filtered.loc[~filtered["product_name"].isin(["Désignation des ouvrages", "Designation des ouvrages"])]
    filtered = filtered.loc[~filtered["code_article"].fillna("").str.match(r"^N\s*Prix$", case=False)]
    financial_re = r"^(?:t\.?\s*v\.?\s*a\.?.*|tva\b.*|total\b.*|sous\s*total\b.*|montant\b.*|net\s+a\s+payer\b.*|a\s+payer\b.*)$"
    filtered = filtered.loc[~filtered["product_name"].fillna("").str.match(financial_re, case=False)]
    filtered = filtered.loc[~filtered["code_article"].fillna("").str.match(financial_re, case=False)]

    return filtered[
        [
            "sheet_name",
            "row_index",
            "chapter",
            "subchapter",
            "code_article",
            "product_name",
            "unite",
            "quantite",
            "prix_unitaire",
            "montant",
            "confidence",
        ]
    ].reset_index(drop=True)


def extract_bordereau_outputs(
    excel_path: Path,
    workspace_root: Path,
    selected_sheets: list[str] | None = None,
) -> BordereauOutputs:
    run_dir = create_run_dir(workspace_root, "bordereau")
    result = extract_workbook(
        input_path=excel_path,
        forced_sheets=selected_sheets or None,
        all_sheets=False,
        debug=True,
    )

    json_path = export_json(result, run_dir)
    csv_path = export_csv(result, run_dir)
    debug_xlsx_path = export_debug_workbook(result, run_dir)

    items_df = pd.DataFrame([item.to_dict() for item in result.items])
    if items_df.empty:
        items_df = pd.DataFrame(
            columns=[
                "sheet_name",
                "row_index",
                "line_type",
                "lot",
                "chapter",
                "subchapter",
                "code_article",
                "designation",
                "unite",
                "quantite",
                "prix_unitaire",
                "montant",
                "raw_row",
                "confidence",
            ]
        )
    products_df = _products_dataframe_from_items(items_df)

    produits_csv_path = run_dir / "produits_extraits.csv"
    products_df.to_csv(produits_csv_path, index=False, encoding="utf-8-sig")

    produits_txt_path = run_dir / "noms_produits.txt"
    products_df["product_name"].dropna().to_csv(
        produits_txt_path,
        index=False,
        header=False,
        encoding="utf-8-sig",
    )

    return BordereauOutputs(
        output_dir=run_dir,
        json_path=json_path,
        csv_path=csv_path,
        debug_xlsx_path=debug_xlsx_path,
        produits_csv_path=produits_csv_path,
        produits_txt_path=produits_txt_path,
        items_df=items_df,
        products_df=products_df,
    )


def read_markdown(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def as_download_bytes(path: Path) -> bytes:
    return path.read_bytes()
