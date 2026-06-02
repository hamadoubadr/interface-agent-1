from __future__ import annotations

import os
import json
import re
import time
import urllib.error
import urllib.request
from pathlib import Path
from io import BytesIO

import requests
import streamlit as st
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_chroma import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import PyPDFLoader
from langchain.chains import ConversationalRetrievalChain
from langchain_core.messages import HumanMessage, AIMessage
from sentence_transformers import SentenceTransformer
import markdown
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

from app_core.pdf_enrichment import build_pdf_corpus, enrich_products_dataframe, export_enriched_excel, extract_products_from_descriptif
from app_core.pipeline import (
    as_download_bytes,
    create_run_dir,
    default_sheet_selection,
    detect_bundle,
    extract_bordereau_outputs,
    is_gemini_quota_error,
    is_gemini_unavailable_error,
    list_workbook_sheet_names,
    read_markdown,
    run_dossier_synthesis,
    save_uploaded_files,
)
from app_core.pdf_generator import markdown_to_pdf_bytes
from app_core.RAG_chatbot import MultiAgentChatbot


APP_ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = APP_ROOT / ".workspace"
WORKSPACE_ROOT.mkdir(parents=True, exist_ok=True)


class LocalSentenceTransformerEmbeddings:
    def __init__(self, model_name: str):
        self.model = SentenceTransformer(model_name)

    def embed_documents(self, texts: list[str]) -> list[list[float]]:
        vectors = self.model.encode(list(texts), convert_to_numpy=True, normalize_embeddings=True)
        return vectors.tolist()

    def embed_query(self, text: str) -> list[float]:
        vector = self.model.encode(text, convert_to_numpy=True, normalize_embeddings=True)
        return vector.tolist()


def _upload_signature(uploaded_files: list[object]) -> tuple[tuple[str, int], ...]:
    return tuple((file.name, file.size) for file in uploaded_files)


def _path_by_name(paths: list[Path]) -> dict[str, Path]:
    return {path.name: path for path in paths}


def _render_download_button(label: str, path: Path, mime: str) -> None:
    st.download_button(
        label=label,
        data=as_download_bytes(path),
        file_name=path.name,
        mime=mime,
        use_container_width=True,
    )


def _markdown_to_pdf_bytes(markdown_text: str, title: str = "Document") -> bytes:
    """Convert Markdown text to PDF bytes using the professional GIDNAI template"""
    if title == "Identité du Projet":
        return markdown_to_pdf_bytes(markdown_text, title, "identite")
    else:
        return markdown_to_pdf_bytes(markdown_text, title, "synthese")


def _bundle_from_state() -> tuple[Path | None, object | None]:
    input_dir = st.session_state.get("input_dir")
    if not input_dir:
        return None, None
    bundle = detect_bundle(Path(input_dir))
    return Path(input_dir), bundle


def _list_gemini_models(api_key: str) -> list[dict[str, object]]:
    url = "https://generativelanguage.googleapis.com/v1beta/models"
    request = urllib.request.Request(
        url=f"{url}?key={api_key}",
        headers={"Content-Type": "application/json"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace") if hasattr(exc, "read") else str(exc)
        raise RuntimeError(message) from exc
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc
    parsed = json.loads(raw)
    models = parsed.get("models") or []
    return [model for model in models if isinstance(model, dict)]


def _pick_gemini_model(models: list[dict[str, object]]) -> str:
    candidates: list[str] = []
    for model in models:
        name = str(model.get("name") or "")
        methods = model.get("supportedGenerationMethods") or []
        if "generateContent" not in methods:
            continue
        if not name.startswith("models/"):
            continue
        candidates.append(name)

    preferred_order = (
        "models/gemini-2.5-flash",
        "models/gemini-2.0-flash",
        "models/gemini-2.0-flash-lite",
        "models/gemini-1.5-flash",
        "models/gemini-1.5-flash-latest",
        "models/gemini-1.5-pro",
        "models/gemini-1.5-pro-latest",
    )
    for preferred in preferred_order:
        if preferred in candidates:
            return preferred

    flash_candidates = [name for name in candidates if "flash" in name]
    if flash_candidates:
        return sorted(flash_candidates)[0]
    if candidates:
        return sorted(candidates)[0]
    raise RuntimeError("Aucun modele Gemini compatible avec generateContent n'a ete trouve.")


def _get_cached_gemini_model(api_key: str) -> str:
    cache = st.session_state.get("gemini_models_cache") or {}
    if not isinstance(cache, dict):
        cache = {}
    cached = cache.get("selected_model")
    if isinstance(cached, str) and cached.startswith("models/"):
        return cached
    models = _list_gemini_models(api_key)
    selected = _pick_gemini_model(models)
    cache["selected_model"] = selected
    st.session_state["gemini_models_cache"] = cache
    return selected


def _call_gemini(api_key: str, model_name: str, prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/{model_name}:generateContent"
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [{"text": prompt}],
            }
        ]
    }
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url=f"{url}?key={api_key}",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        message = exc.read().decode("utf-8", errors="replace") if hasattr(exc, "read") else str(exc)
        raise RuntimeError(message) from exc
    except Exception as exc:
        raise RuntimeError(str(exc)) from exc

    parsed = json.loads(raw)
    candidates = parsed.get("candidates") or []
    if not candidates:
        raise RuntimeError("Aucune reponse Gemini.")
    content = (candidates[0] or {}).get("content") or {}
    parts = content.get("parts") or []
    text_parts = [part.get("text", "") for part in parts if isinstance(part, dict)]
    return "\n".join(text_parts).strip()


def _call_gemini_with_retries(api_key: str, model_name: str, prompt: str, max_retries: int = 3) -> str:
    for attempt in range(max_retries):
        try:
            return _call_gemini(api_key, model_name, prompt)
        except RuntimeError as exc:
            message = str(exc)
            is_overload = "503" in message or "unavailable" in message.lower() or "overload" in message.lower()
            if is_overload and attempt < max_retries - 1:
                wait_time = (attempt + 1) * 2
                time.sleep(wait_time)
                continue
            raise
    return "Échec après plusieurs tentatives."


class NVIDIAChatbot:
    def __init__(self, api_key: str = None):
        self.api_key = api_key or "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"
        self.invoke_url = "https://integrate.api.nvidia.com/v1/chat/completions"
        self.model = "mistralai/mistral-large-3-675b-instruct-2512"
        
        # Modèle d'embeddings (gratuit, local)
        self.embeddings = LocalSentenceTransformerEmbeddings(
            model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
        )
        
        self.vectorstore = None
        
    def load_documents(self, pdf_paths: list[Path]):
        """Charge et indexe les PDF pour le RAG"""
        documents = []
        for pdf_path in pdf_paths:
            loader = PyPDFLoader(str(pdf_path))
            documents.extend(loader.load())
        
        # Découpage intelligent (Chunking)
        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=1000,
            chunk_overlap=200,
            separators=["\n\n", "\n", ".", " "]
        )
        chunks = text_splitter.split_documents(documents)
        
        # Création de la base vectorielle locale (Chroma)
        self.vectorstore = Chroma.from_documents(
            documents=chunks,
            embedding=self.embeddings,
            collection_name=f"dce_{int(time.time())}"
        )
    
    def _call_nvidia(self, messages: list[dict]) -> str:
        """Appelle l'API NVIDIA avec le modèle Mistral Large 3"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        
        payload = {
            "model": self.model,
            "messages": messages,
            "max_tokens": 2048,
            "temperature": 0.15,
            "top_p": 1.00,
            "frequency_penalty": 0.00,
            "presence_penalty": 0.00,
            "stream": False
        }
        
        response = requests.post(self.invoke_url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        
        result = response.json()
        choices = result.get("choices", [])
        if choices and len(choices) > 0:
            return choices[0].get("message", {}).get("content", "")
        return ""
    
    def ask(self, question: str, chat_history: list[tuple[str, str]]):
        """Pose une question au chatbot RAG avec Mistral Large 3"""
        if not self.vectorstore:
            return {"answer": "Erreur: Le chatbot n'est pas initialise.", "sources": []}
        
        # Récupération des documents pertinents
        retriever = self.vectorstore.as_retriever(search_kwargs={"k": 6})
        relevant_docs = retriever.invoke(question)
        
        # Construction du contexte à partir des documents récupérés
        context_parts = []
        for doc in relevant_docs:
            source = doc.metadata.get('source', 'N/A')
            page = doc.metadata.get('page', 'N/A')
            context_parts.append(f"[Source: {source} - Page {page}]\n{doc.page_content}")
        
        context = "\n\n---\n\n".join(context_parts)
        
        # Construction des messages pour l'API NVIDIA
        messages = []
        
        # Message système avec le contexte
        system_prompt = f"""Tu es un assistant expert en construction et dossiers techniques DCE (Dossier de Consultation des Entreprises).
Tu réponds aux questions en te basant UNIQUEMENT sur les documents fournis ci-dessous.

CONTEXTE DES DOCUMENTS:
{context}

Instructions:
- Réponds de manière concise et précise
- Cite les sources (nom du document et page) quand tu réponds
- Si l'information n'est pas dans les documents, dis-le clairement
- Réponds en français"""
        
        messages.append({"role": "system", "content": system_prompt})
        
        # Historique de conversation
        for human, ai in chat_history:
            messages.append({"role": "user", "content": human})
            messages.append({"role": "assistant", "content": ai})
        
        # Question actuelle
        messages.append({"role": "user", "content": question})
        
        # Appel à l'API
        answer = self._call_nvidia(messages)
        
        return {
            "answer": answer,
            "sources": relevant_docs
        }


class ChatbotDCE(MultiAgentChatbot):
    """Alias pour compatibilité avec le pipeline multi-agents"""
    pass


def get_chatbot():
    if "chatbot" not in st.session_state:
        st.session_state.chatbot = MultiAgentChatbot()
    return st.session_state.chatbot

def main() -> None:
    st.set_page_config(page_title="Dossier Construction AI", layout="wide")
    st.title("Agent IA 1 – Lecture & Synthèse de Dossiers TCE")
    st.caption("Analyse IA de dossiers TCE")

    api_key = os.environ.get("GEMINI_API_KEY", "")

    uploaded_files = st.file_uploader(
        "",
        type=["pdf", "docx", "doc", "xls", "xlsx", "xlsm", "zip"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        signature = _upload_signature(uploaded_files)
        if st.session_state.get("upload_signature") != signature:
            input_dir = save_uploaded_files(uploaded_files, WORKSPACE_ROOT)
            st.session_state["upload_signature"] = signature
            st.session_state["input_dir"] = str(input_dir)

    input_dir, bundle = _bundle_from_state()
    if not input_dir or not bundle:
        st.info("Téléversez d'abord les fichiers du dossier pour activer les options.")
        return

    st.success(f"Dossier chargé dans {input_dir}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Documents détectés", len(bundle.pdf_files))
    col2.metric("Excels détectés", len(bundle.excel_files))
    col3.metric("Fichiers totaux", len(bundle.all_files))

    pdf_map = _path_by_name(bundle.pdf_files)
    excel_map = _path_by_name(bundle.excel_files)

    st.subheader("Detection des pieces")
    detected_rc = [path.name for path in bundle.rc_candidates]
    detected_descriptif = [path.name for path in bundle.descriptif_candidates]
    detected_bordereau = bundle.bordereau_candidates[0].name if bundle.bordereau_candidates else None

    selected_rc_names = st.multiselect(
        "RC",
        options=list(pdf_map.keys()),
        default=detected_rc,
    )
    selected_descriptif_names = st.multiselect(
        "Descriptif / CCTP / pieces ecrites",
        options=list(pdf_map.keys()),
        default=detected_descriptif,
    )
    selected_bordereau_name = st.selectbox(
        "Bordereau Excel",
        options=list(excel_map.keys()),
        index=list(excel_map.keys()).index(detected_bordereau) if detected_bordereau in excel_map else 0 if excel_map else None,
    ) if excel_map else None

    selected_sheet_names: list[str] = []
    if selected_bordereau_name:
        sheet_names = list_workbook_sheet_names(excel_map[selected_bordereau_name])
        defaults = default_sheet_selection(sheet_names)
        selected_sheet_names = st.multiselect(
            "Feuilles bordereau a inclure",
            options=sheet_names,
            default=defaults,
            help="Laisser RECAP de cote si tu veux seulement les vraies feuilles produits.",
        )

    st.divider()

    synth_col, enrich_col, chat_col = st.columns(3)

    with synth_col:
        st.subheader("Génération de synthèses")
        if st.button("Start", key="synth_start", use_container_width=True):
            selected_pdf_paths = [pdf_map[name] for name in dict.fromkeys(selected_rc_names + selected_descriptif_names) if name in pdf_map]
            word_docx_paths = [p for p in selected_pdf_paths if p.suffix.lower() in {".docx", ".doc"}]
            pdf_only_paths = [p for p in selected_pdf_paths if p.suffix.lower() == ".pdf"]
            if not selected_pdf_paths:
                st.error("Sélectionnez au moins un RC ou un descriptif (PDF/Word).")
            else:
                with st.spinner("Génération des synthèses en cours..."):
                    try:
                        outputs = run_dossier_synthesis(
                            pdf_paths=selected_pdf_paths,
                            workspace_root=WORKSPACE_ROOT,
                            provider="nvidia",
                            api_key=api_key,
                            descriptif_paths=[pdf_map[name] for name in selected_descriptif_names if name in pdf_map],
                        )
                        st.session_state["synthesis_outputs"] = {
                            "identite_path": str(outputs.identite_path),
                            "synthese_path": str(outputs.synthese_path),
                            "output_dir": str(outputs.output_dir),
                            "stdout": outputs.stdout,
                            "stderr": outputs.stderr,
                            "provider_used": "nvidia",
                        }
                    except Exception as exc:
                        st.error(str(exc))

    with enrich_col:
        st.subheader("Génération de Bordereau")
        enrich_source = st.radio(
            "Source des produits",
            options=["Depuis bordereau Excel", "Depuis simple descriptif (IA)"],
            index=0,
            horizontal=True,
            key="enrich_source_choice"
        )
        enrich_engine = st.radio(
            "Moteur d'enrichissement",
            options=["Python Local (Regex)", "NVIDIA Mistral (Recommandé)"],
            index=1,
            horizontal=True,
            key="enrich_engine_choice"
        )
        if st.button("Start", key="enrich_start", use_container_width=True):
            descriptif_paths = [pdf_map[name] for name in selected_descriptif_names if name in pdf_map]
            descriptif_pdf_paths = [path for path in descriptif_paths if path.suffix.lower() == ".pdf"]
            if not descriptif_pdf_paths:
                st.error("Sélectionnez au moins un descriptif PDF.")
            elif enrich_source == "Depuis bordereau Excel":
                if not selected_bordereau_name:
                    st.error("Aucun bordereau Excel détecté.")
                else:
                    with st.spinner("Extraction puis enrichissement en cours..."):
                        try:
                            bord_outputs = extract_bordereau_outputs(
                                excel_path=excel_map[selected_bordereau_name],
                                workspace_root=WORKSPACE_ROOT,
                                selected_sheets=selected_sheet_names,
                            )
                            use_nvidia_engine = (enrich_engine == "NVIDIA Mistral (Recommandé)")
                            enriched_df = enrich_products_dataframe(
                                bord_outputs.products_df, 
                                descriptif_pdf_paths,
                                use_nvidia=use_nvidia_engine
                            )
                            enrich_dir = create_run_dir(WORKSPACE_ROOT, "enrichment")
                            excel_path = export_enriched_excel(enriched_df, enrich_dir / "produits_enrichis.xlsx")
                            csv_path = enrich_dir / "produits_enrichis.csv"
                            enriched_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                            txt_path = enrich_dir / "produits_enrichis.txt"
                            enriched_df["product_name"].dropna().to_csv(txt_path, index=False, header=False, encoding="utf-8-sig")

                            st.session_state["enrichment_outputs"] = {
                                "excel_path": str(excel_path),
                                "csv_path": str(csv_path),
                                "txt_path": str(txt_path),
                                "preview": enriched_df.head(50).to_dict(orient="records"),
                            }
                        except Exception as exc:
                            st.error(str(exc))
            else:
                with st.spinner("Extraction des produits depuis le descriptif via IA..."):
                    try:
                        use_nvidia_engine = (enrich_engine == "NVIDIA Mistral (Recommandé)")
                        products_df = extract_products_from_descriptif(
                            descriptif_pdf_paths,
                            WORKSPACE_ROOT,
                            use_nvidia=use_nvidia_engine
                        )
                        if products_df.empty:
                            st.warning("Aucun produit n'a pu être extrait du descriptif.")
                        else:
                            enriched_df = enrich_products_dataframe(
                                products_df, 
                                descriptif_pdf_paths,
                                use_nvidia=use_nvidia_engine
                            )
                            enrich_dir = create_run_dir(WORKSPACE_ROOT, "enrichment")
                            excel_path = export_enriched_excel(enriched_df, enrich_dir / "produits_enrichis.xlsx")
                            csv_path = enrich_dir / "produits_enrichis.csv"
                            enriched_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                            txt_path = enrich_dir / "produits_enrichis.txt"
                            enriched_df["product_name"].dropna().to_csv(txt_path, index=False, header=False, encoding="utf-8-sig")

                            st.session_state["enrichment_outputs"] = {
                                "excel_path": str(excel_path),
                                "csv_path": str(csv_path),
                                "txt_path": str(txt_path),
                                "preview": enriched_df.head(50).to_dict(orient="records"),
                            }
                    except Exception as exc:
                        st.error(str(exc))

    with chat_col:
        st.subheader("CHATBOT")
        selected_chat_pdfs = [pdf_map[name] for name in dict.fromkeys(selected_rc_names + selected_descriptif_names) if name in pdf_map]
        
        if not selected_chat_pdfs:
            st.info("Selectionne des documents pour activer le chatbot.")
        else:
            if "chatbot" not in st.session_state:
                if st.button(" Initialiser le Chatbot", use_container_width=True):
                    with st.spinner("Indexation RAG en cours..."):
                        bot = get_chatbot() # Utilise le pipeline multi-agents NVIDIA mis en cache
                        bot.load_documents(selected_chat_pdfs)
                        st.session_state.chatbot = bot
                        st.session_state.chat_history = []
                        st.success("Chatbot prêt (Pipeline Multi-Agents) !")
                        st.rerun()
            else:
                # Interface de chat
                if "chat_history" not in st.session_state:
                    st.session_state.chat_history = []

                # Affichage historique
                for role, content in st.session_state.chat_history:
                    with st.chat_message(role):
                        st.markdown(content)

                if prompt := st.chat_input("Posez votre question technique..."):
                    st.session_state.chat_history.append(("user", prompt))
                    with st.chat_message("user"):
                        st.markdown(prompt)

                    with st.chat_message("assistant"):
                        with st.spinner("Analyse du dossier..."):
                            # On passe l'historique formatté pour LangChain
                            history_for_bot = []
                            for i in range(0, len(st.session_state.chat_history) - 1, 2):
                                try:
                                    u = st.session_state.chat_history[i][1]
                                    a = st.session_state.chat_history[i+1][1]
                                    history_for_bot.append((u, a))
                                except IndexError: break

                            response = st.session_state.chatbot.ask(prompt, history_for_bot)
                            answer = response["answer"]
                            st.markdown(answer)
                            
                            # Affichage des sources
                            with st.expander("📄 Sources (Traçabilité)"):
                                for i, doc in enumerate(response["sources"], 1):
                                    st.write(f"**Source {i}:** {doc.metadata.get('source', 'N/A')} - Page {doc.metadata.get('page', 'N/A')}")
                                    st.caption(doc.page_content[:300] + "...")
                            
                            st.session_state.chat_history.append(("assistant", answer))

    if "synthesis_outputs" in st.session_state:
        outputs = st.session_state["synthesis_outputs"]
        st.divider()
        st.subheader("Résultats du rapport de synthèse")
        provider_used = outputs.get("provider_used")
        if provider_used == "nvidia":
            st.success("Synthèse générée avec succès via le pipeline Multi-Agents NVIDIA (Llama 3.2 + Qwen 3 Coder + Mistral Nemotron) !")
        elif provider_used == "local":
            st.info("Synthèse générée avec le moteur local. Gemini était indisponible ou son quota était saturé.")
        identite_path = Path(outputs["identite_path"])
        synthese_path = Path(outputs["synthese_path"])
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Identite du projet**")
            st.code(read_markdown(identite_path), language="markdown")
            _render_download_button("Télécharger identite-projet.md", identite_path, "text/markdown")
            identite_pdf = _markdown_to_pdf_bytes(read_markdown(identite_path), "Identité du Projet")
            st.download_button(
                label="📄 Télécharger PDF",
                data=identite_pdf,
                file_name="identite-projet.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with col_b:
            st.markdown("**Synthèse détaillée**")
            st.code(read_markdown(synthese_path), language="markdown")
            _render_download_button("Télécharger synthèse-détaillée.md", synthese_path, "text/markdown")
            synthese_pdf = _markdown_to_pdf_bytes(read_markdown(synthese_path), "Synthèse Détaillée")
            st.download_button(
                label="📄 Télécharger PDF",
                data=synthese_pdf,
                file_name="synthese-detaillee.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    if "enrichment_outputs" in st.session_state:
        outputs = st.session_state["enrichment_outputs"]
        st.divider()
        st.subheader("Résultats d'enrichissement produits")
        preview = outputs["preview"]
        if preview:
            st.dataframe(preview, use_container_width=True)
        _render_download_button("Télécharger produits_enrichis.xlsx", Path(outputs["excel_path"]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        _render_download_button("Télécharger produits_enrichis.csv", Path(outputs["csv_path"]), "text/csv")
        _render_download_button("Télécharger noms_produits.txt", Path(outputs["txt_path"]), "text/plain")


if __name__ == "__main__":
    main()
