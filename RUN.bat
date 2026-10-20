@echo off
rem Atalho para rodar a extracao sem precisar digitar o comando no terminal.
rem Basta dar duplo-clique neste arquivo. Funciona em qualquer maquina Windows:
rem   1. acha o Python (instala pelo winget se nao houver);
rem   2. cria um ambiente virtual proprio em .venv;
rem   3. instala as dependencias (requirements.txt) e o Chromium do Playwright;
rem   4. roda o robo.
rem Argumentos extras sao repassados ao robo. Ex.: RUN.bat --limite 10

setlocal
cd /d "%~dp0"
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

set "VENV=%~dp0.venv"
set "VPY=%VENV%\Scripts\python.exe"
set "MARCA=%VENV%\requirements.instalado"

rem --- 1. Ambiente virtual ja existe e funciona? --------------------------
rem (Uma .venv copiada de outra maquina aponta para um Python que nao existe
rem aqui; nesse caso ela e recriada.)
if exist "%VPY%" (
    "%VPY%" -c "import sys" >nul 2>nul
    if errorlevel 1 (
        echo Ambiente virtual invalido nesta maquina. Recriando...
        rmdir /s /q "%VENV%"
    )
)
if exist "%VPY%" goto :dependencias

rem --- 2. Localiza um Python 3.9+ ----------------------------------------
call :achar_python
if not defined PYBASE (
    echo Python nao encontrado. Tentando instalar o Python 3.13 pelo winget...
    where winget >nul 2>nul
    if errorlevel 1 goto :sem_python
    winget install -e --id Python.Python.3.13 --scope user --accept-package-agreements --accept-source-agreements
    call :achar_python
)
if not defined PYBASE goto :sem_python

echo Criando ambiente virtual em .venv ...
%PYBASE% -m venv "%VENV%"
if errorlevel 1 (
    echo Falha ao criar o ambiente virtual.
    goto :fim_erro
)

rem --- 3. Dependencias ----------------------------------------------------
:dependencias
rem Reinstala quando requirements.txt mudou ou quando falta algum modulo.
set "INSTALAR=0"
if not exist "%MARCA%" set "INSTALAR=1"
if exist "%MARCA%" (
    fc /b "requirements.txt" "%MARCA%" >nul 2>nul
    if errorlevel 1 set "INSTALAR=1"
)
"%VPY%" -c "import playwright, openpyxl, dotenv" >nul 2>nul
if errorlevel 1 set "INSTALAR=1"

if "%INSTALAR%"=="1" (
    echo Instalando dependencias do projeto...
    "%VPY%" -m pip install --disable-pip-version-check -q --upgrade pip
    "%VPY%" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 (
        echo Falha ao instalar as dependencias. Verifique a internet/proxy e rode de novo.
        goto :fim_erro
    )
    copy /y "requirements.txt" "%MARCA%" >nul
)

rem O Chromium fica no perfil do usuario (LOCALAPPDATA\ms-playwright), entao
rem pode faltar mesmo com a .venv pronta. Se ja estiver instalado, e instantaneo.
rem NODE_USE_SYSTEM_CA faz o download confiar nos certificados do Windows
rem (proxy corporativo com inspecao de HTTPS).
set "NODE_USE_SYSTEM_CA=1"
"%VPY%" -m playwright install chromium
if errorlevel 1 (
    echo.
    echo AVISO: nao consegui baixar o Chromium do Playwright.
    echo O robo vai tentar usar o Microsoft Edge instalado na maquina.
    echo.
)

rem --- 4. Roda o robo -----------------------------------------------------
echo.
"%VPY%" download_nfce.py baixar --perfil perfil %*

echo.
echo Processo encerrado.
pause >nul
exit /b 0

rem ------------------------------------------------------------------------
:achar_python
set "PYBASE="
for %%V in (3.13 3.12 3.11 3.10 3.9) do (
    if not defined PYBASE (
        py -%%V -c "import sys" >nul 2>nul && set "PYBASE=py -%%V"
    )
)
if not defined PYBASE (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>nul && set "PYBASE=python"
)
rem Logo apos instalar pelo winget o PATH desta janela ainda nao foi atualizado.
if not defined PYBASE if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PYBASE="%LOCALAPPDATA%\Programs\Python\Python313\python.exe""
)
exit /b 0

:sem_python
echo.
echo Nao foi possivel encontrar nem instalar o Python.
echo Instale o Python 3.13 em https://www.python.org/downloads/
echo (marque "Add python.exe to PATH") e rode este arquivo novamente.

:fim_erro
echo.
pause
exit /b 1
