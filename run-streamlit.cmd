@echo off
setlocal
py -3.13 -m streamlit run "%~dp0streamlit_dossier_app\app.py" --server.port 8501 --server.headless true
