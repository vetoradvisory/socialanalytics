@echo off
chcp 65001 >nul
title Social Media Analytics Scraper

echo ============================================================
echo   Social Media Analytics Scraper
echo ============================================================
echo.

REM ---------- Verifica se a planilha foi informada ----------
if "%~1"=="" (
    echo Arraste sua planilha Excel para cima deste arquivo .bat
    echo ou execute pelo terminal:
    echo.
    echo   executar.bat planilha.xlsx
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

REM ---------- Fecha Chrome existente e abre com debug ----------
echo Fechando Chrome existente (se houver)...
taskkill /f /im chrome.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo Abrindo Chrome com porta de debug %PORTA%...
set CHROME_PATHS[0]=C:\Program Files\Google\Chrome\Application\chrome.exe
set CHROME_PATHS[1]=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe

set CHROME_EXE=
if exist "%CHROME_PATHS[0]%" set CHROME_EXE=%CHROME_PATHS[0]%
if exist "%CHROME_PATHS[1]%" set CHROME_EXE=%CHROME_PATHS[1]%
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
    set CHROME_EXE=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe
)

if "%CHROME_EXE%"=="" (
    echo [AVISO] Chrome nao encontrado nos caminhos padrao.
    echo O scraper tentara abrir automaticamente.
    goto :run_scraper
)

start "" "%CHROME_EXE%" --remote-debugging-port=%PORTA% --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data"
echo Chrome aberto. Aguardando inicializar...
timeout /t 4 /nobreak >nul

echo.
echo IMPORTANTE: Faca login no LinkedIn / Instagram / TikTok no Chrome aberto.
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
    echo Concluido! Verifique a planilha "%PLANILHA%" para os resultados.
)
echo.
pause
