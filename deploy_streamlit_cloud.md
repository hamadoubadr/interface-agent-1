# Guide de déploiement sur Streamlit Cloud

Ce guide vous explique comment déployer votre agent DCE sur Streamlit Cloud pour le partager avec d'autres personnes.

## 📋 Prérequis

1. **Compte GitHub** : [github.com](https://github.com)
2. **Compte Streamlit Cloud** : [share.streamlit.io](https://share.streamlit.io)
3. **Code source** : Votre projet doit être sur GitHub

## 🚀 Étapes de déploiement

### Étape 1 : Préparer votre dépôt GitHub

1. **Initialiser un dépôt Git** (si ce n'est pas déjà fait) :
   ```bash
   cd "c:\Users\badr\Documents\stage pfe new\agent antigravity a marche - Copie\agent 1 nvidia"
   git init
   git add .
   git commit -m "Initial commit: Agent DCE"
   ```

2. **Créer un dépôt sur GitHub** :
   - Allez sur [github.com](https://github.com)
   - Cliquez sur "+" → "New repository"
   - Nom : `agent-dce` (ou autre nom)
   - Description : "Agent IA pour extraction et enrichissement de bordereaux DCE"
   - Public (recommandé pour Streamlit Cloud gratuit)
   - Ne pas initialiser avec README

3. **Pousser votre code sur GitHub** :
   ```bash
   git remote add origin https://github.com/VOTRE-USERNAME/agent-dce.git
   git branch -M main
   git push -u origin main
   ```

### Étape 2 : Configurer Streamlit Cloud

1. **Se connecter à Streamlit Cloud** :
   - Allez sur [share.streamlit.io](https://share.streamlit.io)
   - Connectez-vous avec votre compte GitHub

2. **Créer une nouvelle application** :
   - Cliquez sur "New app"
   - Sélectionnez votre dépôt : `VOTRE-USERNAME/agent-dce`
   - Branche : `main`
   - Fichier principal : `streamlit_dossier_app/app.py`
   - Cliquez sur "Deploy"

### Étape 3 : Configurer les variables d'environnement (Optionnel)

Si vous utilisez des clés API (NVIDIA, Gemini), configurez-les :

1. Dans Streamlit Cloud, allez dans "App settings"
2. Cliquez sur "Secrets"
3. Ajoutez vos clés :
   ```toml
   # .streamlit/secrets.toml
   NVIDIA_API_KEY = "nvapi-votre-cle-ici"
   GOOGLE_API_KEY = "votre-cle-gemini-ici"
   ```

### Étape 4 : Vérifier le déploiement

1. **Attendez 2-3 minutes** pour l'installation des dépendances
2. **Vérifiez les logs** :
   - Cliquez sur "Manage app" → "Logs"
   - Vérifiez qu'il n'y a pas d'erreurs
3. **Testez l'application** :
   - Ouvrez le lien fourni par Streamlit Cloud
   - Testez l'upload de fichiers PDF
   - Testez l'extraction et l'enrichissement

## 🔧 Configuration avancée

### Fichier `requirements.txt` optimisé pour Streamlit Cloud

Streamlit Cloud a des limitations de mémoire (1GB). Voici une version optimisée :

```txt
# requirements-streamlit.txt
streamlit==1.32.0
pandas==2.2.0
numpy==1.24.0
pdfplumber==0.10.0
pypdf==4.0.0
openai==1.12.0
sentence-transformers==2.2.0
requests==2.31.0
openpyxl==3.1.2
```

### Configuration mémoire

Pour optimiser l'utilisation mémoire sur Streamlit Cloud :

1. **Limiter la taille des fichiers uploadés** :
   ```python
   # Dans app.py
   st.set_page_config(
       page_title="Agent DCE",
       page_icon="🤖",
       layout="wide",
       initial_sidebar_state="expanded"
   )
   
   # Limiter à 50MB par fichier
   MAX_UPLOAD_SIZE = 50 * 1024 * 1024
   ```

2. **Utiliser le cache Streamlit** :
   ```python
   @st.cache_data(ttl=3600)
   def process_pdf(file_path):
       # Traitement lourd
       return result
   ```

## 🚨 Dépannage

### Problèmes courants et solutions

| Problème | Solution |
|----------|----------|
| **"ModuleNotFoundError"** | Vérifiez que toutes les dépendances sont dans `requirements.txt` |
| **Mémoire insuffisante** | Utilisez la version optimisée des requirements |
| **Timeout lors du déploiement** | Réduisez la taille des fichiers de test |
| **Erreur d'import** | Vérifiez les chemins relatifs dans le code |
| **Clés API non reconnues** | Vérifiez le format dans les secrets |

### Vérifier les logs

Les logs Streamlit Cloud sont accessibles via :
1. "Manage app" → "Logs"
2. Vérifiez les erreurs d'installation
3. Vérifiez les erreurs d'exécution

## 🔗 Partager votre application

Une fois déployée, vous pouvez :
1. **Partager le lien** : `https://VOTRE-APP.streamlit.app`
2. **Intégrer dans un site web** : Utiliser un iframe
3. **Partager sur les réseaux** : Le lien est public

## 📊 Monitoring

Streamlit Cloud fournit :
- **Statistiques d'utilisation** : Nombre de vues, temps d'exécution
- **Logs en temps réel** : Pour le débogage
- **État de l'application** : En ligne, erreur, déployé

## 🔄 Mises à jour

Pour mettre à jour votre application :
1. **Poussez les changements sur GitHub** :
   ```bash
   git add .
   git commit -m "Description des changements"
   git push origin main
   ```
2. **Streamlit Cloud redéploie automatiquement**
3. **Vérifiez les logs** après 2-3 minutes

## 🆘 Support

- **Documentation Streamlit** : [docs.streamlit.io](https://docs.streamlit.io)
- **Forum Streamlit** : [discuss.streamlit.io](https://discuss.streamlit.io)
- **GitHub Issues** : Pour les bugs spécifiques à votre code

---

**Félicitations !** 🎉 Votre agent DCE est maintenant accessible en ligne. Partagez le lien avec la personne qui doit le tester.