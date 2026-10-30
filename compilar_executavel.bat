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

:: 4. Garante que os assets fiquem exclusivamente em _internal e nunca na raiz
if exist "dist\RoboSocietario\assets" (
    rmdir /s /q "dist\RoboSocietario\assets" >nul 2>&1
)
if exist "assets" (
    if not exist "dist\RoboSocietario\_internal\assets" mkdir "dist\RoboSocietario\_internal\assets" >nul 2>&1
    xcopy /e /y /i "assets" "dist\RoboSocietario\_internal\assets" >nul 2>&1
)

:: 5. Copia as planilhas base exclusivamente para PlanilhasBase, utilitários e manual
if exist "dist\RoboSocietario\CNPJs.xlsx" (
    del /f /q "dist\RoboSocietario\CNPJs.xlsx" >nul 2>&1
)
if not exist "dist\RoboSocietario\PlanilhasBase" mkdir "dist\RoboSocietario\PlanilhasBase" >nul 2>&1
if exist "PlanilhasBase\CNPJs.xlsx" (
    echo [INFO] Incluindo planilha base em dist\RoboSocietario\PlanilhasBase\CNPJs.xlsx...
    copy /y "PlanilhasBase\CNPJs.xlsx" "dist\RoboSocietario\PlanilhasBase\CNPJs.xlsx" >nul 2>&1
) else if exist "CNPJs.xlsx" (
    echo [INFO] Incluindo planilha base em dist\RoboSocietario\PlanilhasBase\CNPJs.xlsx...
    copy /y "CNPJs.xlsx" "dist\RoboSocietario\PlanilhasBase\CNPJs.xlsx" >nul 2>&1
)
if exist "Atualizar_Certificados.bat" (
    echo [INFO] Incluindo utilitario Atualizar_Certificados.bat em dist\RoboSocietario...
    copy /y "Atualizar_Certificados.bat" "dist\RoboSocietario\Atualizar_Certificados.bat" >nul 2>&1
)
if exist "Manual_do_Usuario_RoboSocietario.pdf" (
    echo [INFO] Incluindo Manual do Usuario .pdf em dist\RoboSocietario...
    copy /y "Manual_do_Usuario_RoboSocietario.pdf" "dist\RoboSocietario\Manual_do_Usuario_RoboSocietario.pdf" >nul 2>&1
)

:: 6. Oculta a pasta _internal na distribuicao
if exist "dist\RoboSocietario\_internal" (
    echo [INFO] Ocultando pasta _internal...
    attrib +h "dist\RoboSocietario\_internal" >nul 2>&1
)

:: 7. Gera arquivo ZIP automaticamente para distribuicao imediata preservando atributos
if exist "dist\RoboSocietario.zip" del /f /q "dist\RoboSocietario.zip"
echo [INFO] Compactando pasta para dist\RoboSocietario.zip...
%PYTHON_CMD% -c "import os, zipfile, ctypes; dist_dir = 'dist'; src_dir = os.path.join(dist_dir, 'RoboSocietario'); z_path = os.path.join(dist_dir, 'RoboSocietario.zip'); zf = zipfile.ZipFile(z_path, 'w', zipfile.ZIP_DEFLATED); [zf.writestr(zipfile.ZipInfo(os.path.relpath(os.path.join(r, d), dist_dir).replace('\\\\', '/') + '/', ((ctypes.windll.kernel32.GetFileAttributesW(os.path.join(r, d)) & 0xFF) | 0x10) if ctypes.windll.kernel32.GetFileAttributesW(os.path.join(r, d)) != -1 else 0x10), '') for r, dirs, _ in os.walk(src_dir) for d in dirs]; [zf.writestr(zipfile.ZipInfo(os.path.relpath(os.path.join(r, f), dist_dir).replace('\\\\', '/'), ctypes.windll.kernel32.GetFileAttributesW(os.path.join(r, f)) & 0xFF if ctypes.windll.kernel32.GetFileAttributesW(os.path.join(r, f)) != -1 else 0), open(os.path.join(r, f), 'rb').read(), compress_type=zipfile.ZIP_DEFLATED) for r, _, files in os.walk(src_dir) for f in files]; zf.close()"
if not exist "dist\RoboSocietario.zip" (
    tar -a -cf "dist\RoboSocietario.zip" -C "dist" "RoboSocietario"
)

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
    echo    - Planilha base: PlanilhasBase\CNPJs.xlsx (com abas CNPJs_consultados e Certificados_Disponiveis)
    echo    - Manual: Manual_do_Usuario_RoboSocietario.pdf (guia passo a passo para o usuario)
    echo    - Evidencias: pasta Resultados\CapturasDeTela-Downloads (com prints e anexos baixados)
    echo    - Relatorio geral: pasta Resultados\Relatorios
) else (
    echo.
    echo  Como distribuir para os usuarios:
    echo    1. Copie ou compacte (ZIP) a pasta "dist\RoboSocietario".
    echo    2. O usuario clica duas vezes direto em "RoboSocietario.exe" para rodar!
)
if not "%1"=="--no-pause" pause
