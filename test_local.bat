@echo off
REM Testes de integração da API (na raiz do repositório)
cd /d "%~dp0"
python -m pytest
if errorlevel 1 pause
