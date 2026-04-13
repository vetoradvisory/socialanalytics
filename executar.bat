@echo off
chcp 65001 >nul
title Social Media Analytics Scraper

echo ============================================================
echo   Social Media Analytics Scraper  [browser: Edge]
echo ============================================================
echo.

REM ---------- Verifica se a planilha foi informada ----------
if "%~1"=="" (
    echo Arraste sua planilha Excel para cima deste arquivo .bat
    echo ou execute pelo terminal:
    echo.
    echo   executar.bat planilha.xlsx
    echo   executar.bat "https://docs.google.com/spreadsheets/d/..."
    echo.
    pause
    exit /b 1
)

set PLANILHA=%~1
set PORTA=9222

REM ---------- Verifica Python ----------
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale em: https://python.org/downloads
    echo        Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

REM ---------- Instala dependencias se necessario ----------
python -c "import selenium" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias...
    pip install -r "%~dp0requirements.txt"
    echo.
)

REM ---------- Fecha Edge existente e abre com debug ----------
echo Fechando Edge existente (se houver)...
taskkill /f /im msedge.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo Abrindo Edge com porta de debug %PORTA%...
set EDGE_EXE=
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    set EDGE_EXE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
)
if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    set EDGE_EXE=C:\Program Files\Microsoft\Edge\Application\msedge.exe
)

if "%EDGE_EXE%"=="" (
    echo [AVISO] Edge nao encontrado nos caminhos padrao.
    echo O scraper tentara abrir automaticamente.
    goto :run_scraper
)

start "" "%EDGE_EXE%" --remote-debugging-port=%PORTA% --user-data-dir="%LOCALAPPDATA%\Microsoft\Edge\User Data"
echo Edge aberto. Aguardando inicializar...
timeout /t 4 /nobreak >nul

echo.
echo IMPORTANTE: Faca login no LinkedIn / Instagram / TikTok no Edge aberto.
echo Quando estiver logado, pressione qualquer tecla para iniciar a extracao.
echo.
pause

:run_scraper
echo.
echo Iniciando extracao de "%PLANILHA%"...
echo.
python "%~dp0scraper.py" "%PLANILHA%" --porta-debug %PORTA%

echo.
if errorlevel 1 (
    echo [ERRO] O scraper terminou com erro. Veja o arquivo scraper.log para detalhes.
) else (
    echo Concluido! Verifique a planilha para os resultados.
)
echo.
pause
