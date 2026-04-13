@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title Social Media Analytics Scraper

echo ============================================================
echo   Social Media Analytics Scraper  [browser: Edge]
echo ============================================================
echo.

REM ---------- Muda para a pasta do script (garante imports Python) ----------
cd /d "%~dp0"

REM ---------- Verifica se a planilha foi informada ----------
if "%~1"=="" (
    echo Arraste sua planilha ou URL do Google Sheets para cima deste .bat
    echo ou execute pelo terminal:
    echo.
    echo   executar.bat planilha.xlsx
    echo   executar.bat "https://docs.google.com/spreadsheets/d/..."
    echo.
    pause
    exit /b 1
)

set "PLANILHA=%~1"
set "PORTA=9222"

REM ---------- Verifica Python ----------
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale em: https://python.org/downloads
    echo        Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

REM ---------- Instala / atualiza dependencias ----------
python -c "import selenium, openpyxl, gspread, google_auth_oauthlib" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias (aguarde)...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar dependencias. Verifique sua conexao.
        pause
        exit /b 1
    )
    echo.
)

REM ---------- Localiza o Edge ----------
set "EDGE_EXE="

REM 1) Tenta via PATH (mais confiavel)
for /f "delims=" %%i in ('where msedge 2^>nul') do (
    if not defined EDGE_EXE set "EDGE_EXE=%%i"
)

REM 2) Caminhos padrão (Program Files x86 primeiro — instalação mais comum)
if not defined EDGE_EXE (
    if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
        set "EDGE_EXE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
    )
)
if not defined EDGE_EXE (
    if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
        set "EDGE_EXE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
    )
)
if not defined EDGE_EXE (
    if exist "%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe" (
        set "EDGE_EXE=%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"
    )
)

REM ---------- Fecha Edge existente e abre com debug ----------
echo Fechando Edge existente (se houver)...
taskkill /f /im msedge.exe >nul 2>&1
timeout /t 2 /nobreak >nul

if not defined EDGE_EXE (
    echo [AVISO] Edge nao encontrado nos caminhos padrao.
    echo O scraper tentara abrir automaticamente.
    goto :run_scraper
)

echo Abrindo Edge: !EDGE_EXE!
start "" "!EDGE_EXE!" --remote-debugging-port=!PORTA! "--user-data-dir=%LOCALAPPDATA%\Microsoft\Edge\User Data"
echo Edge aberto. Aguardando inicializar...
timeout /t 4 /nobreak >nul

echo.
echo IMPORTANTE: Faca login no LinkedIn / Instagram / TikTok no Edge aberto.
echo Quando estiver logado, pressione qualquer tecla para iniciar a extracao.
echo.
pause

:run_scraper
echo.
echo Iniciando extracao...
echo   Fonte: !PLANILHA!
echo.
python scraper.py "!PLANILHA!" --porta-debug !PORTA!

echo.
if errorlevel 1 (
    echo [ERRO] O scraper terminou com erro. Veja scraper.log para detalhes.
) else (
    echo Concluido! Verifique a planilha para os resultados.
)
echo.
pause
