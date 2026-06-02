# Agent DCE - Extraction et Enrichissement de Bordereaux

Agent IA pour extraire et enrichir automatiquement les bordereaux de prix à partir de descriptifs techniques (CCTP).

## 🚀 Fonctionnalités

- **Extraction depuis descriptif** : Identification automatique des produits techniques depuis les PDF descriptifs
- **Enrichissement des marques** : Recherche intelligente des marques et spécifications
- **Génération Excel** : Production de bordereaux Excel structurés
- **Chatbot RAG** : Assistant conversationnel pour questions techniques
- **Multi-modèles** : Utilisation de Qwen3 Coder, Mistral Large, et embeddings NVIDIA

## 📋 Prérequis

- Python 3.8 ou supérieur
- 4GB RAM minimum (8GB recommandé)
- Connexion internet pour les modèles cloud (optionnel pour mode local)

## 🔧 Installation locale

### Option 1 : Installation rapide
```bash
# Rendre le script exécutable (Linux/Mac)
chmod +x setup.sh

# Exécuter l'installation
./setup.sh
```

### Option 2 : Installation manuelle
```bash
# Créer un environnement virtuel
python -m venv venv

# Activer l'environnement
# Linux/Mac:
source venv/bin/activate
# Windows:
venv\Scripts\activate

# Installer les dépendances
pip install -r requirements.txt

# Télécharger les modèles d'embeddings
python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')"
```

## 🎯 Utilisation

### Démarrer l'application
```bash
# Activer l'environnement virtuel
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate     # Windows

# Lancer Streamlit
streamlit run streamlit_dossier_app/app.py
```

L'application sera accessible sur : [http://localhost:8501](http://localhost:8501)

### Workflow principal
1. **Téléverser des fichiers** : PDF descriptifs et/ou Excel de bordereau
2. **Sélectionner la source** : 
   - "Depuis bordereau Excel" : Extraction depuis un fichier Excel existant
   - "Depuis simple descriptif (IA)" : Identification automatique des produits
3. **Configurer l'enrichissement** :
   - Activer/désactiver l'extraction des marques
   - Choisir le modèle (NVIDIA API ou local)
4. **Générer le bordereau** : Télécharger le fichier Excel enrichi

## 🌐 Déploiement web

### Option 1 : Streamlit Cloud (Recommandé)
1. Poussez votre code sur GitHub
2. Allez sur [share.streamlit.io](https://share.streamlit.io)
3. Connectez votre compte GitHub
4. Sélectionnez le dépôt et le fichier `streamlit_dossier_app/app.py`
5. Configurez les variables d'environnement si nécessaire

### Option 2 : Heroku
```bash
# Se connecter à Heroku
heroku login

# Créer une application
heroku create votre-nom-app

# Déployer
git push heroku main

# Ouvrir l'application
heroku open
```

### Option 3 : Railway
1. Importez votre dépôt GitHub sur Railway
2. Railway détectera automatiquement l'application Streamlit
3. Configurez les variables d'environnement
4. Déployez en un clic

## 🔑 Configuration API

### Clés API requises (optionnel)
- **NVIDIA API Key** : Pour utiliser les modèles Qwen3 Coder et Mistral Large
- **Google Gemini API Key** : Alternative pour l'extraction

### Variables d'environnement
```bash
# NVIDIA API
export NVIDIA_API_KEY="nvapi-votre-cle-ici"

# Google Gemini
export GOOGLE_API_KEY="votre-cle-gemini-ici"

# Mode local (sans API)
export USE_LOCAL_MODELS="true"
```

## 📁 Structure du projet
```
agent-dce/
├── streamlit_dossier_app/
│   ├── app.py              # Application principale Streamlit
│   ├── app_core/           # Modules de traitement
│   │   ├── pdf_enrichment.py    # Extraction et enrichissement
│   │   ├── pipeline.py          # Pipeline de traitement
│   │   ├── RAG_chatbot.py       # Chatbot avec RAG
│   │   └── word_processor.py    # Traitement de texte
│   └── .workspace/         # Dossier de travail temporaire
├── requirements.txt        # Dépendances Python
├── Procfile               # Configuration déploiement
├── runtime.txt            # Version Python
├── setup.sh               # Script d'installation
└── README.md              # Documentation
```

## 🛠️ Développement

### Exécuter les tests
```bash
pytest test_extract_products.py
pytest test_extract_with_brands.py
```

### Formater le code
```bash
black .
ruff check --fix .
```

### Mode debug
```bash
streamlit run streamlit_dossier_app/app.py --logger.level=debug
```

## 🤝 Contribution

1. Fork le projet
2. Créez une branche (`git checkout -b feature/amazing-feature`)
3. Committez vos changements (`git commit -m 'Add amazing feature'`)
4. Poussez la branche (`git push origin feature/amazing-feature`)
5. Ouvrez une Pull Request

## 📄 Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.

## 🙏 Remerciements

- **NVIDIA** pour les modèles NIM et l'API
- **Mistral AI** pour les modèles Mistral Large
- **Qwen** pour Qwen3 Coder
- **Streamlit** pour la plateforme d'applications web