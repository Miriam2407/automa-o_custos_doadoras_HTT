@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================
echo   Instalador - Automacao de Doadoras
echo ================================================
echo.

where python >nul 2>nul
if %errorlevel%==0 (
    echo Python ja esta instalado nesta maquina.
    goto :instalar_dependencias
)

echo Python nao foi encontrado nesta maquina.
echo Baixando o instalador oficial do Python 3.12...
echo (isso precisa de internet - pode levar alguns minutos)
echo.

set "PYTHON_URL=https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe"
set "INSTALADOR=%TEMP%\python_installer.exe"

powershell -NoProfile -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALADOR%'"

if not exist "%INSTALADOR%" (
    echo.
    echo ERRO: nao foi possivel baixar o instalador do Python.
    echo Verifique sua conexao com a internet, ou instale o Python manualmente
    echo em https://www.python.org/downloads/ marcando a opcao "Add python.exe to PATH".
    pause
    exit /b 1
)

echo Instalando Python silenciosamente (aguarde)...
"%INSTALADOR%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

echo Aguardando a instalacao finalizar...
timeout /t 20 /nobreak >nul

del "%INSTALADOR%" >nul 2>nul

REM o PATH so atualiza numa janela nova do terminal - reabre este script
REM automaticamente numa nova janela para garantir que "python" ja funciona
echo.
echo Reabrindo em uma nova janela para concluir a instalacao...
start "" cmd /c "cd /d "%~dp0" && instalar.bat"
exit /b 0

:instalar_dependencias
echo.
echo Criando ambiente virtual (.venv)...
python -m venv .venv

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo ERRO: nao consegui criar o ambiente virtual. Confirme que o Python
    echo foi instalado corretamente e tente rodar instalar.bat de novo.
    pause
    exit /b 1
)

echo Instalando as bibliotecas necessarias...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip >nul
pip install -r requirements.txt

echo.
echo ================================================
echo   Instalacao concluida!
echo   Use o arquivo executar.bat para abrir o programa.
echo ================================================
pause
