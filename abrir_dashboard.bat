@echo off
cd /d "%~dp0"

echo Iniciando el dashboard...
start "" http://localhost:8080/
python -m http.server 8080
