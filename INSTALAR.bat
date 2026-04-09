@echo off
chcp 65001 >nul
title Beemo Office Agent - Instalação
color 0B

echo.
echo ╔════════════════════════════════════════════════════════════════╗
echo ║                  BEEMO OFFICE AGENT - INSTALAÇÃO              ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

REM Verificar se Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python não encontrado!
    echo.
    echo Por favor, instale o Python 3.8 ou superior:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANTE: Marque a opção "Add Python to PATH" durante a instalação
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado!
python --version
echo.

REM Verificar se pip está instalado
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] pip não encontrado!
    echo.
    echo Instalando pip...
    python -m ensurepip --upgrade
)

echo [OK] pip encontrado!
echo.

REM Instalar dependências
echo ════════════════════════════════════════════════════════════════
echo Instalando dependências... (isso pode demorar alguns minutos)
echo ════════════════════════════════════════════════════════════════
echo.

pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao instalar dependências!
    echo.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════════════════
echo [OK] Instalação concluída com sucesso!
echo ════════════════════════════════════════════════════════════════
echo.
echo Próximo passo: Execute CONFIGURAR.bat para configurar sua API key
echo.
pause
