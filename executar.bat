@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo Ambiente ainda nao foi instalado nesta maquina.
    echo Rode primeiro o arquivo instalar.bat, depois volte a usar este.
    pause
    exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" gui_automacao.py
