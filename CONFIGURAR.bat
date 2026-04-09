@echo off
chcp 65001 >nul
title Beemo Office Agent - Configuração
color 0E

echo.
echo ╔════════════════════════════════════════════════════════════════╗
echo ║                BEEMO OFFICE AGENT - CONFIGURAÇÃO              ║
echo ╚════════════════════════════════════════════════════════════════╝
echo.

REM Verificar se .env já existe
if exist .env (
    echo [AVISO] Arquivo .env já existe!
    echo.
    set /p "overwrite=Deseja reconfigurar? (S/N): "
    if /i not "%overwrite%"=="S" (
        echo.
        echo Configuração cancelada.
        echo.
        pause
        exit /b 0
    )
)

echo ════════════════════════════════════════════════════════════════
echo PASSO 1: API Key do Google Gemini
echo ════════════════════════════════════════════════════════════════
echo.
echo Para usar o Beemo, você precisa de uma API key GRATUITA do Google.
echo.
echo 1. Acesse: https://aistudio.google.com/apikey
echo 2. Faça login com sua conta Google
echo 3. Clique em "Create API Key"
echo 4. Copie a chave gerada
echo.
echo Abrindo o navegador...
start https://aistudio.google.com/apikey
echo.
set /p "api_key=Cole sua API key aqui: "

if "%api_key%"=="" (
    echo.
    echo [ERRO] API key não pode estar vazia!
    echo.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════════════════
echo PASSO 2: Pasta de Trabalho
echo ════════════════════════════════════════════════════════════════
echo.
echo Onde você quer salvar seus arquivos Office?
echo.
echo Sugestão: %USERPROFILE%\Documents\BeemoFiles
echo.
set /p "root_path=Digite o caminho (ou pressione ENTER para usar a sugestão): "

if "%root_path%"=="" (
    set "root_path=%USERPROFILE%\Documents\BeemoFiles"
)

REM Criar pasta se não existir
if not exist "%root_path%" (
    mkdir "%root_path%" 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo [ERRO] Não foi possível criar a pasta: %root_path%
        echo.
        pause
        exit /b 1
    )
    echo [OK] Pasta criada: %root_path%
) else (
    echo [OK] Pasta já existe: %root_path%
)

echo.
echo ════════════════════════════════════════════════════════════════
echo PASSO 3: Modelo de IA (Opcional)
echo ════════════════════════════════════════════════════════════════
echo.
echo Modelos disponíveis (GRATUITOS):
echo.
echo 1. gemini-2.5-flash-lite  (Mais rápido - 15 req/min)
echo 2. gemini-2.5-flash       (Equilibrado - 10 req/min)
echo 3. gemini-2.5-pro         (Mais inteligente - 5 req/min) [PADRÃO]
echo.
set /p "model_choice=Escolha (1/2/3 ou ENTER para padrão): "

set "model_name=gemini-2.5-pro"
if "%model_choice%"=="1" set "model_name=gemini-2.5-flash-lite"
if "%model_choice%"=="2" set "model_name=gemini-2.5-flash"

echo.
echo Modelo selecionado: %model_name%

REM Criar arquivo .env
echo # Gemini API Configuration> .env
echo # Obtenha sua API key em: https://aistudio.google.com/apikey>> .env
echo GEMINI_API_KEY=%api_key%>> .env
echo.>> .env
echo # Root Path Configuration>> .env
echo # Caminho da pasta raiz onde os arquivos Office estão localizados>> .env
echo ROOT_PATH=%root_path%>> .env
echo.>> .env
echo # Model Configuration>> .env
echo # Nome do modelo Gemini principal>> .env
echo MODEL_NAME=%model_name%>> .env
echo.>> .env
echo # Version Control Configuration>> .env
echo # Número máximo de versões a manter por arquivo (padrão: 10)>> .env
echo MAX_VERSIONS=10>> .env
echo.>> .env
echo # Cache Configuration>> .env
echo # Habilitar cache de respostas do Gemini (padrão: true)>> .env
echo CACHE_ENABLED=true>> .env
echo.>> .env
echo # Tempo de vida do cache em horas (padrão: 24)>> .env
echo CACHE_TTL_HOURS=24>> .env
echo.>> .env
echo # Número máximo de entradas no cache (padrão: 100)>> .env
echo CACHE_MAX_ENTRIES=100>> .env

echo.
echo ════════════════════════════════════════════════════════════════
echo [OK] Configuração concluída com sucesso!
echo ════════════════════════════════════════════════════════════════
echo.
echo Arquivo .env criado com:
echo   - API Key: %api_key:~0,20%...
echo   - Pasta: %root_path%
echo   - Modelo: %model_name%
echo.
echo Próximo passo: Execute EXECUTAR.bat para iniciar o Beemo
echo.
pause
