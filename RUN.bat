@echo off
rem Atalho para rodar a extracao sem precisar digitar o comando no terminal.
rem Basta dar duplo-clique neste arquivo.

setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo Python nao encontrado. Instale o Python 3.13 e rode novamente:
    echo   py -3.13 -m pip install -r requirements.txt
    echo   py -3.13 -m playwright install chromium
    echo.
    pause
    exit /b 1
)

py -3.13 download_nfce.py baixar --perfil perfil

echo.
echo Processo encerrado.
pause >nul
