@echo off
chcp 65001 >nul
title Beemo Office Agent
color 0A

echo.
echo ╔════════════════════════════════════════════════════════════════╗
echo ║                    BEEMO OFFICE AGENT                         ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

REM Verificar se .env existe
if not exist .env (
    echo [ERRO] Arquivo .env não encontrado!
    echo.
    echo Execute CONFIGURAR.bat primeiro para configurar o Beemo.
    echo.
    pause
    exit /b 1
)

REM Verificar se Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python não encontrado!
    echo.
    echo Execute INSTALAR.bat primeiro para instalar as dependências.
    echo.
    pause
    exit /b 1
)

REM Verificar se streamlit está instalado
python -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Streamlit não encontrado!
    echo.
    echo Execute INSTALAR.bat primeiro para instalar as dependências.
    echo.
    pause
    exit /b 1
)

echo [OK] Iniciando Beemo Office Agent...
echo.
echo ════════════════════════════════════════════════════════════════
echo O navegador abrirá automaticamente em alguns segundos...
echo.
echo Para parar o Beemo, feche esta janela ou pressione Ctrl+C
echo ════════════════════════════════════════════════════════════════
echo.

REM Aguardar 3 segundos e abrir navegador
timeout /t 3 /nobreak >nul
start http://localhost:8501

REM Executar Streamlit
streamlit run app.py --server.headless true

REM Se o Streamlit for fechado
echo.
echo Beemo foi encerrado.
pause
