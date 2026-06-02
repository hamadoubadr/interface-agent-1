"""
Pipeline Multi-Agents pour le chatbot DCE
Combine intelligemment plusieurs modèles NVIDIA pour maximiser l'efficacité:
- NVIDIA Embeddings (llama-3_2-nemoretriever-300m-embed-v1): Indexation et recherche
- Qwen3 Coder (qwen3-coder-480b-a35b-instruct): Extraction de données/tableaux
- Mistral Nemotron (mistralai/mistral-nemotron): Synthèse et rédaction
"""

import hashlib
import requests
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Tuple, Optional
from pathlib import Path
from sentence_transformers import SentenceTransformer
from langchain_chroma import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import PyPDFLoader, Docx2txtLoader
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
    """Agent de synthèse utilisant Mistral Nemotron via OpenAI client"""

    def __init__(self, api_key: str = NVIDIA_API_KEY):
        self.api_key = api_key
        self.model = "mistralai/mistral-nemotron"
        self.client = OpenAI(base_url=NVIDIA_BASE_URL, api_key=api_key)

    def synthesize_answer(self, question: str, context: str, extracted_data: str = "") -> str:
        """Synthétise une réponse claire à partir du contexte et des données extraites"""
        system_prompt = """Tu es un assistant expert en construction et dossiers techniques DCE.
Tu rédiges des réponses claires, précises et professionnelles en français technique.
Tu cites toujours tes sources (nom du document et page).
Si des données chiffrées sont fournies, intègre-les naturellement dans ta réponse.
Si l'information n'est pas dans les documents, dis-le clairement."""

        user_content = f"QUESTION: {question}\n\n"

        if extracted_data:
            user_content += f"DONNÉES EXTRAITES:\n{extracted_data}\n\n"

        user_content += f"CONTEXTE DES DOCUMENTS:\n{context}"

        completion = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            temperature=0.15,
            top_p=1.00,
            max_tokens=2048,
            stream=False
        )

        if completion.choices and len(completion.choices) > 0:
            return completion.choices[0].message.content or ""
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

    def load_documents(self, pdf_paths: List[Path]):
        """Charge et indexe les PDF/DOCX pour le RAG avec cache vectoriel"""
        import os
        # Calcule une signature des documents
        sig = hashlib.md5(str(sorted([str(p) for p in pdf_paths])).encode()).hexdigest()[:8]
        persist_dir = str(Path(os.getcwd()) / f".chroma_cache_dce_{sig}")

        if Path(persist_dir).exists():
            # Réutilise l'index existant
            self.vectorstore = Chroma(
                collection_name="dce_multi_agent",
                embedding_function=self.embeddings,
                persist_directory=persist_dir
            )
            return

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
                    loader = PyPDFLoader(str(doc_path), extraction_mode="layout")
                    documents.extend(loader.load())
            except Exception as e:
                # Fallback pour PDF
                if suffix != ".docx":
                    try:
                        loader = PyPDFLoader(str(doc_path), extraction_mode="plain_text")
                        documents.extend(loader.load())
                    except Exception as e2:
                        print(f"Erreur lors du chargement de {doc_path}: {e2}")
                        continue
                else:
                    print(f"Erreur lors du chargement de {doc_path}: {e}")
                    continue

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

    def ask(self, question: str, chat_history: List[Tuple[str, str]] = None) -> Dict:
        """Pose une question au chatbot multi-agents"""
        if not self.vectorstore:
            return {"answer": "Erreur: Le chatbot n'est pas initialisé.", "sources": []}

        if chat_history is None:
            chat_history = []

        # Étape 1: Récupération des documents pertinents
        retriever = self.vectorstore.as_retriever(search_kwargs={"k": 4})
        relevant_docs = retriever.invoke(question)

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
