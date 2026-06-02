"""
Pipeline Multi-Agents pour le chatbot DCE
Combine intelligemment plusieurs modèles NVIDIA pour maximiser l'efficacité:
- NVIDIA Embeddings (llama-3_2-nemoretriever-300m-embed-v1): Indexation et recherche
- Qwen3 Coder (qwen3-coder-480b-a35b-instruct): Extraction de données/tableaux
- Mistral Nemotron (mistralai/mistral-nemotron): Synthèse et rédaction
"""

import hashlib
import os
import requests
import re
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Tuple, Optional
from pathlib import Path
from sentence_transformers import SentenceTransformer
from langchain_chroma import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import PyMuPDFLoader, Docx2txtLoader
from langchain_core.documents import Document
from openai import OpenAI


# Configuration NVIDIA
NVIDIA_API_KEY = "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"
NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"
NVIDIA_CLIENT = OpenAI(base_url=NVIDIA_BASE_URL, api_key=NVIDIA_API_KEY)


class NVIDIAEmbeddings:
    """Embeddings NVIDIA utilisant llama-3_2-nemoretriever-300m-embed-v1"""

    def __init__(self, api_key: str = NVIDIA_API_KEY):
        self.api_key = api_key
        self.model = "nvidia/llama-3_2-nemoretriever-300m-embed-v1"
        self.invoke_url = f"{NVIDIA_BASE_URL}/embeddings"

    def embed_documents(self, texts: List[str]) -> List[List[float]]:
        """Embed une liste de documents en un seul appel API"""
        payload = {
            "model": self.model,
            "input": texts,
            "encoding_format": "float",
            "input_type": "passage"
        }
        response = requests.post(self.invoke_url, headers=self._headers(), json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        return [item["embedding"] for item in sorted(result["data"], key=lambda x: x["index"])]

    def embed_query(self, text: str) -> List[float]:
        """Embed une requête unique"""
        payload = {
            "model": self.model,
            "input": [text],
            "encoding_format": "float",
            "input_type": "query"
        }
        response = requests.post(self.invoke_url, headers=self._headers(), json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        if "data" in result and len(result["data"]) > 0:
            return result["data"][0]["embedding"]
        return []

    def _headers(self):
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }


class LocalSentenceTransformerEmbeddings:
    """Fallback local embeddings si l'API NVIDIA échoue"""

    def __init__(self, model_name: str = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"):
        self.model = SentenceTransformer(model_name)

    def embed_documents(self, texts: List[str]) -> List[List[float]]:
        vectors = self.model.encode(list(texts), convert_to_numpy=True, normalize_embeddings=True)
        return vectors.tolist()

    def embed_query(self, text: str) -> List[float]:
        vector = self.model.encode(text, convert_to_numpy=True, normalize_embeddings=True)
        return vector.tolist()


class QwenDataExtractor:
    """Agent d'extraction de données utilisant Qwen3 Coder via OpenAI client"""

    def __init__(self, api_key: str = NVIDIA_API_KEY):
        self.api_key = api_key
        self.model = "qwen/qwen3-coder-480b-a35b-instruct"
        self.client = OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

    def extract_table_data(self, context: str, question: str) -> str:
        """Extrait des données de tableaux depuis le contexte"""
        system_prompt = """Tu es un expert en extraction de données techniques de tableaux de bordereau.
Ta tâche est d'extraire précisément les informations demandées depuis le contexte fourni.
Focus sur les données chiffrées: quantités, prix unitaires, montants totaux, références.
Ne fais pas de calculs complexes, juste extrait les données brutes."""

        completion = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"CONTEXTE:\n{context}\n\nQUESTION: {question}"}
            ],
            temperature=0.3,
            top_p=0.8,
            max_tokens=2048,
            stream=False
        )

        if completion.choices and len(completion.choices) > 0:
            return completion.choices[0].message.content or ""
        return ""


class MistralSynthesizer:
    """Agent de synthèse avec fallback de modèles NVIDIA"""

    def __init__(self, api_key: str = NVIDIA_API_KEY):
        self.api_key = api_key
        self.client = OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)
        env_models = os.getenv("NVIDIA_SYNTH_MODELS", "").strip()
        if env_models:
            self.models = [item.strip() for item in env_models.split(",") if item.strip()]
        else:
            self.models = [
                "mistralai/mistral-nemotron",
                "meta/llama-3.1-70b-instruct",
                "microsoft/phi-3.5-mini-instruct",
            ]
        self.model = self.models[0]

    def synthesize_answer(self, question: str, context: str, extracted_data: str = "") -> str:
        """Synthétise une réponse claire à partir du contexte et des données extraites"""
        system_prompt = """Tu es un assistant expert en construction et dossiers techniques DCE.
Tu rédiges des réponses claires, précises et professionnelles en français technique.
Tu cites toujours tes sources (nom du document et page).
Si des données chiffrées sont fournies, intègre-les naturellement dans ta réponse.
Si l'information n'est pas dans les documents, dis-le clairement.
RÈGLE ABSOLUE : NE RÉDIGE JAMAIS de section ou de partie intitulée "Points de Vigilance" ni dans l'identité ni dans la synthèse technique."""

        user_content = f"QUESTION: {question}\n\n"

        question_lower = question.lower()
        if any(keyword in question_lower for keyword in ["identité", "identite", "architecte", "maître d'ouvrage", "maitre d'ouvrage", "bet"]):
            user_content += (
                "FORMAT ATTENDU:\n"
                "- Maître d'ouvrage\n"
                "- Architecte\n"
                "- BET / bureaux d'études\n"
                "- Bureau de contrôle\n"
                "- Localisation / situation géographique\n"
                "- Nature du marché / objet\n"
                "- Délai des travaux\n"
                "- Autres intervenants utiles\n\n"
                "CONSIGNE:\n"
                "Privilégie les pages de garde, règlements de consultation et tableaux d'identification. "
                "Si le contexte contient un bloc 'OCR STRUCTURE PAGE DE GARDE', considère ce bloc comme prioritaire pour associer chaque rôle à la bonne entité. "
                "Recopie exactement les noms d'entreprises lorsqu'ils apparaissent. "
                "S'il existe plusieurs entités proches, indique-les sans fusion abusive.\n"
                "RÈGLE ABSOLUE ET IMPÉRATIVE : Si un intervenant, un rôle ou une information (ex: Bureau de contrôle, BET, Surfaces, etc.) n'est pas clairement spécifié dans les documents, TU DOIS IGNORER ET SUPPRIMER TOTALEMENT L'ÉLÉMENT. Ne mentionne sous aucun prétexte 'Non spécifié', 'Non mentionné', 'N/A' ou 'Inconnu'. La ligne ne doit tout simplement pas exister dans ton rendu final.\n\n"
            )
        elif any(keyword in question_lower for keyword in ["technique", "plomberie", "ecs", "climatisation", "ventilation", "incendie"]):
            user_content += (
                "FORMAT ATTENDU:\n"
                "- Vue d'ensemble du lot fluides\n"
                "- Plomberie / évacuation\n"
                "- ECS\n"
                "- Climatisation / ventilation\n"
                "- Protection incendie / désenfumage\n"
                "- Matériaux, marques et normes\n\n"
            )

        if extracted_data:
            user_content += f"DONNÉES EXTRAITES:\n{extracted_data}\n\n"

        user_content += f"CONTEXTE DES DOCUMENTS:\n{context}"

        last_error = None
        for model_name in self.models:
            try:
                completion = self.client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_content}
                    ],
                    temperature=0.15,
                    top_p=1.00,
                    max_tokens=2048,
                    stream=False
                )
                self.model = model_name
                if completion.choices and len(completion.choices) > 0:
                    return completion.choices[0].message.content or ""
                return ""
            except Exception as exc:
                last_error = exc
                continue

        if last_error:
            raise last_error
        return ""


class MultiAgentChatbot:
    """Orchestrateur multi-agents pour le chatbot DCE"""

    def __init__(self, use_nvidia_embeddings: bool = True):
        # Choix du modèle d'embeddings
        if use_nvidia_embeddings:
            try:
                self.embeddings = NVIDIAEmbeddings()
                # Test rapide
                self.embeddings.embed_query("test")
            except Exception:
                print("NVIDIA embeddings indisponible, fallback vers local")
                self.embeddings = LocalSentenceTransformerEmbeddings()
        else:
            self.embeddings = LocalSentenceTransformerEmbeddings()

        # Agents spécialisés
        self.data_extractor = QwenDataExtractor()
        self.synthesizer = MistralSynthesizer()

        self.vectorstore = None
        self.raw_documents: List[Document] = []
        self.source_paths: List[Path] = []
        self.ocr_front_documents: List[Document] = []

    def load_documents(self, pdf_paths: List[Path]):
        """Charge et indexe les PDF/DOCX pour le RAG avec cache vectoriel"""
        self.source_paths = list(pdf_paths)
        # Calcule une signature des documents
        sig = hashlib.md5(str(sorted([str(p) for p in pdf_paths])).encode()).hexdigest()[:8]
        # Utiliser un dossier local au lieu de /tmp pour éviter les conflits Windows/Linux et les locks
        persist_dir = str(Path(os.getcwd()) / f".chroma_cache_dce_{sig}")

        documents = []
        for doc_path in pdf_paths:
            suffix = doc_path.suffix.lower()
            try:
                if suffix == ".docx":
                    # Utiliser Docx2txtLoader pour Word
                    loader = Docx2txtLoader(str(doc_path))
                    documents.extend(loader.load())
                else:
                    # PDF par défaut
                    loader = PyMuPDFLoader(str(doc_path))
                    documents.extend(loader.load())
            except Exception as e:
                # Fallback pour PDF
                if suffix != ".docx":
                    try:
                        loader = PyMuPDFLoader(str(doc_path))
                        documents.extend(loader.load())
                    except Exception as e2:
                        print(f"Erreur lors du chargement de {doc_path}: {e2}")
                        continue
                else:
                    print(f"Erreur lors du chargement de {doc_path}: {e}")
                    continue

        self.raw_documents = documents[:]
        self.ocr_front_documents = []

        if Path(persist_dir).exists():
            # Réutilise l'index existant, mais garde aussi les pages brutes pour le mode hybride.
            self.vectorstore = Chroma(
                collection_name="dce_multi_agent",
                embedding_function=self.embeddings,
                persist_directory=persist_dir
            )
            return

        # Découpage intelligent (Chunking)
        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=1000,
            chunk_overlap=200,
            separators=["\n\n", "\n", ".", " "]
        )
        chunks = text_splitter.split_documents(documents)

        # Création de la base vectorielle avec persistance
        self.vectorstore = Chroma.from_documents(
            documents=chunks,
            embedding=self.embeddings,
            collection_name="dce_multi_agent",
            persist_directory=persist_dir
        )

    def _is_data_question(self, question: str) -> bool:
        """Détecte si la question concerne des données chiffrées/tableaux"""
        data_keywords = [
            "prix", "coût", "montant", "quantité", "total", "budget",
            "combien", "combien coûte", "tarif", "devise", "dh", "euro",
            "tableau", "bordereau", "article", "référence"
        ]
        question_lower = question.lower()
        return any(keyword in question_lower for keyword in data_keywords)

    def _is_identity_question(self, question: str) -> bool:
        identity_keywords = [
            "identité", "identite", "maître d'ouvrage", "maitre d'ouvrage",
            "architecte", "bet", "bureau de contrôle", "bureau de controle",
            "situation", "localisation", "nature du marché", "nature du marche",
            "intervenants", "maîtrise d'oeuvre", "maitrise d'oeuvre"
        ]
        question_lower = question.lower()
        return any(keyword in question_lower for keyword in identity_keywords)

    def _keyword_score(self, text: str, keywords: List[str]) -> int:
        lowered = text.lower()
        score = 0
        for keyword in keywords:
            score += len(re.findall(re.escape(keyword), lowered))
        return score

    def _dedupe_documents(self, documents: List[Document]) -> List[Document]:
        seen = set()
        deduped = []
        for doc in documents:
            source = doc.metadata.get("source", "N/A")
            page = doc.metadata.get("page", "N/A")
            key = (source, page, doc.page_content[:300])
            if key in seen:
                continue
            seen.add(key)
            deduped.append(doc)
        return deduped

    def _get_front_pages(self, max_page: int = 2) -> List[Document]:
        front_pages = []
        for doc in self.raw_documents:
            page = doc.metadata.get("page")
            if isinstance(page, int) and page <= max_page:
                front_pages.append(doc)
        return front_pages

    def _get_ocr_front_pages(self, max_page: int = 1) -> List[Document]:
        if self.ocr_front_documents:
            return [doc for doc in self.ocr_front_documents if doc.metadata.get("page", 999) <= max_page]

        try:
            import fitz
            import numpy as np
            import easyocr
        except Exception:
            return []

        try:
            reader = easyocr.Reader(["fr", "en"], gpu=False)
        except Exception:
            return []

        ocr_docs: List[Document] = []
        for source_path in self.source_paths:
            if source_path.suffix.lower() != ".pdf":
                continue
            try:
                pdf = fitz.open(str(source_path))
            except Exception:
                continue

            for page_index in range(min(max_page + 1, len(pdf))):
                try:
                    page = pdf[page_index]
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                    lines = reader.readtext(img, detail=0, paragraph=False)
                    text = "\n".join(line.strip() for line in lines if isinstance(line, str) and line.strip())
                    if not text.strip():
                        continue
                    ocr_docs.append(
                        Document(
                            page_content=text,
                            metadata={
                                "source": str(source_path),
                                "page": page_index,
                                "extraction": "ocr_front_page",
                            },
                        )
                    )
                except Exception:
                    continue

        self.ocr_front_documents = ocr_docs
        return [doc for doc in self.ocr_front_documents if doc.metadata.get("page", 999) <= max_page]

    def _get_structured_identity_ocr_documents(self) -> List[Document]:
        try:
            import fitz
            import numpy as np
            import easyocr
        except Exception:
            return []

        try:
            reader = easyocr.Reader(["fr", "en"], gpu=False)
        except Exception:
            return []

        def normalize(text: str) -> str:
            return (
                text.lower()
                .replace("é", "e")
                .replace("è", "e")
                .replace("ê", "e")
                .replace("à", "a")
                .replace("ù", "u")
                .replace("ô", "o")
                .replace("'", " ")
            )

        def clean_text(text: str) -> str:
            cleaned = re.sub(r"\s+", " ", text).strip(" -|,:;/")
            replacements = {
                "cals architectures": "CAES architectures",
                "cais architectures": "CAES architectures",
                "cals architccturcs": "CAES architectures",
                "cais architccturcs": "CAES architectures",
                "cals architcctures": "CAES architectures",
                "cals": "CAES",
                "cais": "CAES",
                "sseps1": "SEPSI",
                "ssepsi": "SEPSI",
                "seps1": "SEPSI",
                "ingecobat": "INGECObat",
                "maydane immobilier": "Maydane Immobilier",
                "alysse realty": "ALYSSE REALTY",
            }
            norm = normalize(cleaned)
            for key, value in replacements.items():
                if norm == key:
                    return value
            if "conseil" in norm and ("tensift" in norm or "cit" in norm):
                return "Conseil D'Ingenierie Tensift (CIT)"
            if "ingeco" in norm:
                return "INGECObat"
            if "sepsi" in norm or "sseps1" in norm or "ssepsi" in norm:
                return "SEPSI"
            if "archit" in norm and ("caes" in norm or "cals" in norm or "cais" in norm):
                return "CAES architectures"
            return cleaned

        def is_address_like(text: str) -> bool:
            norm = normalize(text)
            address_markers = [
                "casablanca", "casa", "rabat", "marrakech", "avenue", "av ", "rue",
                "bd", "boulevard", "etage", "appt", "extension", "tel", "mob",
                "galerie", "quartier", "agdal", "main street", "origin office",
                "papillons", "yacoub", "mansour", "michlifen"
            ]
            if any(marker in norm for marker in address_markers):
                return True
            return bool(re.search(r"\d", text))

        def is_noise(text: str) -> bool:
            norm = normalize(text)
            noise_markers = [
                "royaume du maroc", "wilaya", "prefecture", "lot fluides",
                "dce", "version du", "projet de construction", "bureaux alysse",
                "bureau d etude", "bureaux d etudes techniques"
            ]
            return any(marker in norm for marker in noise_markers)

        def pick_title_entity(items: List[Dict]) -> str:
            candidates = []
            for item in items:
                if item["y1"] < 330 or item["y1"] > 650:
                    continue
                mid_x = (item["x1"] + item["x2"]) / 2
                if abs(mid_x - center_x) > 220:
                    continue
                if is_noise(item["text"]) or is_address_like(item["text"]):
                    continue
                norm = normalize(item["text"])
                if any(token in norm for token in ["arrondissement", "prefecture", "royaume", "wilaya"]):
                    continue
                if len(norm.split()) < 2:
                    continue
                candidates.append(clean_text(item["text"]))
            return candidates[0] if candidates else ""

        def classify_heading(norm: str) -> Optional[str]:
            if "maitre" in norm and "ouvrage" in norm:
                return "Maitre d'ouvrage"
            if "maitre" in norm and ("oeuvre" in norm or "qeuvre" in norm):
                return "Maitre d'oeuvre"
            if ("bet" in norm or "bel" in norm) and "tech" in norm:
                return "BET technique"
            if ("bet" in norm or "bel" in norm) and "struct" in norm:
                return "BET structure"
            if "bureau" in norm and "controle" in norm:
                return "Bureau de controle"
            if ("bet" in norm or "ssep" in norm or "sepsi" in norm) and ("securite" in norm or "incendie" in norm):
                return "BET securite incendie"
            return None

        def extract_heading_values(items: List[Dict], headings: List[Dict]) -> Dict[str, List[str]]:
            headings.sort(key=lambda h: (h["col"], h["y1"]))
            by_col = {"left": [], "right": []}
            for heading in headings:
                by_col[heading["col"]].append(heading)

            role_values: Dict[str, List[str]] = {}
            for col, col_headings in by_col.items():
                for idx, heading in enumerate(col_headings):
                    next_y = col_headings[idx + 1]["y1"] if idx + 1 < len(col_headings) else float("inf")
                    next_y = min(next_y, heading["y1"] + 260)
                    values = []
                    for item in items:
                        if item["col"] != col:
                            continue
                        if item["y1"] <= heading["y2"] or item["y1"] >= next_y:
                            continue
                        text_norm = item["norm"]
                        if classify_heading(text_norm):
                            continue
                        if len(item["text"]) <= 2 or is_noise(item["text"]):
                            continue
                        values.append(clean_text(item["text"]))
                    role_values[heading["label"]] = values
            return role_values

        def first_non_address(values: List[str]) -> str:
            text_bits = [value for value in values if not is_address_like(value)]
            if not text_bits:
                return ""
            merged = " ".join(text_bits[:3]).strip()
            return clean_text(merged)

        def choose_role_value(label: str, values: List[str], title_entity: str) -> str:
            joined_norm = " ".join(normalize(value) for value in values)
            if label == "Maitre d'ouvrage":
                if title_entity:
                    return title_entity
                merged = first_non_address(values)
                return merged or ""
            if label == "Maitre d'oeuvre":
                if "archit" in joined_norm and ("caes" in joined_norm or "cals" in joined_norm or "cais" in joined_norm):
                    return "CAES architectures"
                return first_non_address(values)
            if label == "BET technique":
                if "ingeco" in joined_norm:
                    return "INGECObat"
                return first_non_address(values)
            if label == "BET structure":
                if "conseil" in joined_norm and ("tensift" in joined_norm or "cit" in joined_norm):
                    return "Conseil D'Ingenierie Tensift (CIT)"
                return first_non_address(values)
            if label == "BET securite incendie":
                if "sepsi" in joined_norm or "sseps1" in joined_norm or "ssepsi" in joined_norm:
                    return "SEPSI"
                return first_non_address(values)
            if label == "Bureau de controle":
                if "verite" in joined_norm and "controle" in joined_norm:
                    return "VERITE CONTROLE"
                return first_non_address(values)
            return first_non_address(values)

        structured_docs: List[Document] = []

        for source_path in self.source_paths:
            if source_path.suffix.lower() != ".pdf":
                continue
            try:
                pdf = fitz.open(str(source_path))
                if len(pdf) == 0:
                    continue
                page = pdf[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                raw_items = reader.readtext(img, detail=1, paragraph=False)
            except Exception:
                continue

            page_width = pix.width
            center_x = page_width / 2
            items = []
            for box, text, *_ in raw_items:
                if not isinstance(text, str) or not text.strip():
                    continue
                xs = [pt[0] for pt in box]
                ys = [pt[1] for pt in box]
                items.append(
                    {
                        "text": text.strip(),
                        "norm": normalize(text.strip()),
                        "x1": min(xs),
                        "x2": max(xs),
                        "y1": min(ys),
                        "y2": max(ys),
                        "col": "left" if ((min(xs) + max(xs)) / 2) < center_x else "right",
                    }
                )

            headings = []
            for item in items:
                label = classify_heading(item["norm"])
                if label:
                    headings.append({**item, "label": label})

            if not headings:
                continue

            title_entity = pick_title_entity(items)
            role_values = extract_heading_values(items, headings)
            lines = []

            if title_entity:
                lines.append(f"Entite projet principale: {title_entity}")

            ordered_roles = [
                "Maitre d'ouvrage",
                "Maitre d'oeuvre",
                "BET technique",
                "BET structure",
                "Bureau de controle",
                "BET securite incendie",
            ]
            for role in ordered_roles:
                selected = choose_role_value(role, role_values.get(role, []), title_entity)
                if selected:
                    lines.append(f"{role}: {selected}")

            moa_contact = first_non_address(role_values.get("Maitre d'ouvrage", []))
            if moa_contact and title_entity and normalize(moa_contact) != normalize(title_entity):
                lines.append(f"Contact consultation / depot: {moa_contact}")

            if lines:
                structured_docs.append(
                    Document(
                        page_content="OCR STRUCTURE PAGE DE GARDE\n" + "\n".join(lines),
                        metadata={
                            "source": str(source_path),
                            "page": 0,
                            "extraction": "ocr_structured_identity",
                        },
                    )
                )

        return structured_docs

    def _get_keyword_documents(self, keywords: List[str], limit: int = 8) -> List[Document]:
        scored_docs = []
        for doc in self.raw_documents:
            score = self._keyword_score(doc.page_content, keywords)
            if score > 0:
                page = doc.metadata.get("page")
                scored_docs.append((score, 999 if page is None else page, doc))
        scored_docs.sort(key=lambda item: (-item[0], item[1]))
        return [item[2] for item in scored_docs[:limit]]

    def _retrieve_semantic_documents(self, query: str, k: int = 8, fetch_k: int = 24) -> List[Document]:
        retriever = self.vectorstore.as_retriever(
            search_type="mmr",
            search_kwargs={"k": k, "fetch_k": fetch_k}
        )
        return retriever.invoke(query)

    def _build_hybrid_context_documents(self, question: str) -> List[Document]:
        semantic_docs = self._retrieve_semantic_documents(question, k=10, fetch_k=28)

        if self._is_identity_question(question):
            identity_keywords = [
                "maître d'ouvrage", "maitre d'ouvrage", "architecte", "bet",
                "bureau de contrôle", "bureau de controle", "adresse",
                "localisation", "situation", "marché", "consultation",
                "maîtrise d'oeuvre", "maitrise d'oeuvre", "verite controle",
                "ingecobat", "caes", "cais", "alysse realty", "sepsi", "cit"
            ]
            combined = (
                semantic_docs
                + self._get_front_pages(max_page=3)
                + self._get_ocr_front_pages(max_page=1)
                + self._get_structured_identity_ocr_documents()
                + self._get_keyword_documents(identity_keywords, limit=10)
            )
            return self._dedupe_documents(combined)[:18]

        technical_keywords = [
            "plomberie", "ecs", "climatisation", "ventilation", "désenfumage",
            "protection incendie", "station de relevage", "vrv", "vmc", "pompe"
        ]
        combined = semantic_docs + self._get_keyword_documents(technical_keywords, limit=6)
        return self._dedupe_documents(combined)[:14]

    def ask(self, question: str, chat_history: List[Tuple[str, str]] = None) -> Dict:
        """Pose une question au chatbot multi-agents"""
        if not self.vectorstore:
            return {"answer": "Erreur: Le chatbot n'est pas initialisé.", "sources": []}

        if chat_history is None:
            chat_history = []

        # Étape 1: Construit un contexte hybride:
        # - RAG sémantique NVIDIA pour la précision
        # - pages de garde + passages administratifs pour récupérer les infos globales que Gemini capte mieux
        relevant_docs = self._build_hybrid_context_documents(question)

        # Construction du contexte
        context_parts = []
        for doc in relevant_docs:
            source = doc.metadata.get('source', 'N/A')
            page = doc.metadata.get('page', 'N/A')
            context_parts.append(f"[Source: {source} - Page {page}]\n{doc.page_content}")

        context = "\n\n---\n\n".join(context_parts)

        is_data_q = self._is_data_question(question)
        extracted_data = ""

        if is_data_q:
            # Lancer Qwen et Mistral en parallèle
            with ThreadPoolExecutor(max_workers=2) as executor:
                fut_extract = executor.submit(self.data_extractor.extract_table_data, context, question)
                fut_synth = executor.submit(self.synthesizer.synthesize_answer, question, context, "")
                extracted_data = fut_extract.result()
                baseline_answer = fut_synth.result()
                # Re-synthèse rapide avec les données extraites si utile
                if extracted_data and extracted_data.strip():
                    answer = self.synthesizer.synthesize_answer(question, context, extracted_data)
                else:
                    answer = baseline_answer
        else:
            # Pas de Qwen du tout → direct Mistral
            answer = self.synthesizer.synthesize_answer(question, context, "")

        return {
            "answer": answer,
            "sources": relevant_docs,
            "extracted_data": extracted_data if is_data_q else ""
        }
