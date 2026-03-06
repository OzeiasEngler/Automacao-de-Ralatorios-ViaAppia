@echo off
REM Sobe a API ARTESP em teste local (porta 8000)
cd /d "%~dp0"
python -m uvicorn render_api.app:app --host 127.0.0.1 --port 8000
pause
