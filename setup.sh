#!/bin/bash

echo "🔧 Installation de l'agent DCE Streamlit..."

# Vérifier Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 n'est pas installé. Veuillez installer Python 3.8 ou supérieur."
    exit 1
fi

# Créer un environnement virtuel
echo "📦 Création de l'environnement virtuel..."
python3 -m venv venv

# Activer l'environnement virtuel
if [[ "$OSTYPE" == "linux-gnu"* ]] || [[ "$OSTYPE" == "darwin"* ]]; then
    source venv/bin/activate
elif [[ "$OSTYPE" == "msys" ]] || [[ "$OSTYPE" == "win32" ]]; then
    source venv/Scripts/activate
fi

# Mettre à jour pip
echo "⬆️ Mise à jour de pip..."
pip install --upgrade pip

# Installer les dépendances
echo "📚 Installation des dépendances..."
pip install -r requirements.txt

# Télécharger les modèles SentenceTransformer
echo "🤖 Téléchargement des modèles d'embeddings..."
python -c "from sentence_transformers import SentenceTransformer; SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')"

echo "✅ Installation terminée !"
echo ""
echo "Pour démarrer l'application :"
echo "1. Activez l'environnement virtuel :"
echo "   - Linux/Mac: source venv/bin/activate"
echo "   - Windows: venv\\Scripts\\activate"
echo "2. Lancez Streamlit :"
echo "   streamlit run streamlit_dossier_app/app.py"
echo ""
echo "L'application sera accessible sur : http://localhost:8501"