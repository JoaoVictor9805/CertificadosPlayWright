@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ======================================================================
echo           Compilando Executavel Unico do Robo Societario
echo ======================================================================
echo.
echo [INFO] Este processo compila o Robo Societario no modo pasta permanente [--onedir]
echo [INFO] com console habilitado para exibicao em tempo real de logs e erros.
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

:: 3. Limpa builds anteriores
echo [INFO] Limpando pastas de build anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist\RoboSocietario" rmdir /s /q "dist\RoboSocietario"
if exist "dist\RoboDET" rmdir /s /q "dist\RoboDET"

echo.
echo [INFO] Compilando aplicacao com PyInstaller [--onedir --console]...
echo Isso levara cerca de 1 a 2 minutos...
echo.

%PYTHON_CMD% -m PyInstaller --noconfirm --clean --console --onedir --name "RoboSocietario" ^
    --add-data "assets;assets" ^
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

:: 4. Garante que os assets estejam disponiveis na pasta de distribuicao
if exist "assets" (
    echo [INFO] Copiando pasta assets para dist\RoboSocietario\assets...
    if not exist "dist\RoboSocietario\assets" mkdir "dist\RoboSocietario\assets" >nul 2>&1
    xcopy /e /y /i "assets" "dist\RoboSocietario\assets" >nul 2>&1
)

:: 5. Copia a planilha CNPJs.xlsx e utilitarios para dentro de dist\RoboSocietario
if exist "CNPJs.xlsx" (
    echo [INFO] Incluindo planilha base CNPJs.xlsx em dist\RoboSocietario...
    copy /y "CNPJs.xlsx" "dist\RoboSocietario\CNPJs.xlsx" >nul 2>&1
)
if exist "Atualizar_Certificados.bat" (
    echo [INFO] Incluindo utilitario Atualizar_Certificados.bat em dist\RoboSocietario...
    copy /y "Atualizar_Certificados.bat" "dist\RoboSocietario\Atualizar_Certificados.bat" >nul 2>&1
)
if exist "Manual_do_Usuario_RoboSocietario.docx" (
    echo [INFO] Incluindo Manual do Usuario .docx em dist\RoboSocietario...
    copy /y "Manual_do_Usuario_RoboSocietario.docx" "dist\RoboSocietario\Manual_do_Usuario_RoboSocietario.docx" >nul 2>&1
)

:: 6. Gera arquivo ZIP automaticamente para distribuicao imediata
if exist "dist\RoboSocietario.zip" del /f /q "dist\RoboSocietario.zip"
echo [INFO] Compactando pasta para dist\RoboSocietario.zip...
tar -a -cf "dist\RoboSocietario.zip" -C "dist" "RoboSocietario"

echo.
echo ======================================================================
echo               Compilacao Concluida com Sucesso!
echo ======================================================================
echo.
echo  Pasta gerada:  dist\RoboSocietario
if exist "dist\RoboSocietario.zip" (
    echo  Arquivo ZIP:   dist\RoboSocietario.zip
    echo.
    echo  [PRONTO PARA ENVIO]
    echo  Basta enviar o arquivo "dist\RoboSocietario.zip" para o usuario.
    echo  Ao descompactar, tudo fica na mesma pasta:
    echo    - Executavel: RoboSocietario.exe (da 2 cliques direto para rodar)
    echo    - Utilitario: Atualizar_Certificados.bat (atualizacao rapida com 2 cliques)
    echo    - Planilha base: CNPJs.xlsx (com abas CNPJs_consultados e Certificados_Disponiveis)
    echo    - Manual: Manual_do_Usuario_RoboSocietario.docx (passo a passo para o usuario)
    echo    - Evidencias: pasta CapturasDeTela-Downloads (com prints e anexos baixados)
    echo    - Relatorio geral: pasta Relatorios
) else (
    echo.
    echo  Como distribuir para os usuarios:
    echo    1. Copie ou compacte (ZIP) a pasta "dist\RoboSocietario".
    echo    2. O usuario clica duas vezes direto em "RoboSocietario.exe" para rodar!
)
echo.
pause
