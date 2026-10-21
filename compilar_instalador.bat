@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

echo ======================================================================
echo       Compilando Instalador Unico do Robo DET (Inno Setup)
echo ======================================================================
echo.
echo [INFO] Este processo compila o Robo DET no modo pasta permanente [--onedir]
echo [INFO] e gera um instalador executavel unico [dist\Instalador_RoboDET.exe].
echo [INFO] A instalacao e feita uma unica vez e o app abrira instantaneamente!
echo.

:: 1. Detecta ambiente virtual (.venv) ou Python padrao
set "PYTHON_CMD=python"
if exist ".\.venv\Scripts\activate.bat" (
    echo [INFO] Ativando ambiente virtual .venv...
    call .\.venv\Scripts\activate.bat
    set "PYTHON_CMD=python"
) else if exist "..\.venv\Scripts\activate.bat" (
    echo [INFO] Ativando ambiente virtual .venv externo...
    call ..\.venv\Scripts\activate.bat
    set "PYTHON_CMD=python"
)

:: Verifica se o Python esta funcional
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado no sistema!
    echo Certifique-se de que o Python esta instalado e adicionado ao PATH.
    pause
    exit /b 1
)

:: 2. Verifica se o PyInstaller esta instalado
%PYTHON_CMD% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [INFO] PyInstaller nao detectado. Instalando agora via pip...
    %PYTHON_CMD% -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar o PyInstaller!
        pause
        exit /b 1
    )
)

:: 3. Detecta o compilador do Inno Setup (ISCC.exe)
set "ISCC="
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC for /f "delims=" %%I in ('where iscc 2^>nul') do set "ISCC=%%I"

if not defined ISCC (
    echo [AVISO] Compilador Inno Setup ISCC.exe nao foi encontrado!
    echo [INFO] Para gerar o instalador executavel unico Instalador_RoboDET.exe,
    echo        instale o Inno Setup executando no terminal:
    echo        winget install JRSoftware.InnoSetup
    echo.
)

:: 4. Limpa builds anteriores
echo [INFO] Limpando pastas de build anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist\RoboDET" rmdir /s /q "dist\RoboDET"
if exist "dist\Instalador_RoboDET.exe" del /f /q "dist\Instalador_RoboDET.exe"

echo.
echo [1/2] Empacotando aplicacao com PyInstaller [--onedir]...
echo Isso levara de 1 a 2 minutos...
echo.

:: 5. Executa o PyInstaller
%PYTHON_CMD% -m PyInstaller --noconfirm --clean --noconsole --onedir --name "RoboDET" ^
    --add-data "assets;assets" ^
    --collect-all customtkinter ^
    --collect-all playwright ^
    --collect-all openpyxl ^
    --collect-all pyautogui ^
    --collect-all pywinauto ^
    --exclude-module torch ^
    --exclude-module torchvision ^
    --exclude-module torchaudio ^
    --exclude-module scipy ^
    --exclude-module pandas ^
    --exclude-module langchain ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERRO] Ocorreu uma falha durante a compilacao do PyInstaller!
    pause
    exit /b 1
)

:: 6. Garante que a pasta assets esteja disponivel na raiz de dist\RoboDET tambem
if exist "assets" (
    if not exist "dist\RoboDET\assets" (
        mkdir "dist\RoboDET\assets" >nul 2>&1
    )
    xcopy /e /y /i "assets" "dist\RoboDET\assets" >nul 2>&1
)

:: 7. Gera o Instalador com Inno Setup (se disponivel)
if defined ISCC (
    echo.
    echo [2/2] Gerando instalador unico .exe com Inno Setup...
    "%ISCC%" inno_setup.iss
    if errorlevel 1 (
        echo.
        echo [ERRO] Falha ao compilar com Inno Setup!
        pause
        exit /b 1
    )
)

echo.
echo ======================================================================
echo               Compilacao Concluida com Sucesso!
echo ======================================================================
echo.
if exist "dist\Instalador_RoboDET.exe" (
    echo  [PRODUTO FINAL PRONTO PARA DISTRIBUICAO]
    echo  - Instalador gerado: dist\Instalador_RoboDET.exe
    echo.
    echo  Como enviar para os usuarios:
    echo    1. Envie apenas o arquivo "dist\Instalador_RoboDET.exe" para o outro computador.
    echo    2. O usuario da 2 cliques e instala em segundos - sem pedir senha de admin.
    echo    3. Atalhos sao criados na Area de Trabalho e Menu Iniciar.
    echo    4. O Robo DET abrira instantaneamente a partir de entao!
) else (
    echo  - Pasta do aplicativo compilada: dist\RoboDET
    echo  - Execute dist\RoboDET\RoboDET.exe diretamente ou instale o Inno Setup para gerar o instalador unico.
)
echo.
pause
