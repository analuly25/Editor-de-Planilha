@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Editor de Planilha

set "RAIZ=%~dp0dist\"
set "EXE=%RAIZ%PlanilhaEngine.exe"
set "PERFIL=%RAIZ%perfis\planilha_profile.json"
set "SAIDA=%~dp0resultados"

echo.
echo ========================================
echo        EDITOR DE PLANILHA
echo ========================================
echo.

if not exist "%EXE%" (
    echo [ERRO] PlanilhaEngine.exe nao encontrado em: %EXE%
    echo        Gere o programa uma vez com fonte\CRIAR_EXE.bat ^(precisa de Python^).
    pause & exit /b 1
)
if not exist "%PERFIL%" (
    echo [ERRO] Perfil nao encontrado: %PERFIL%
    pause & exit /b 1
)

rem Aceita arrastar a planilha para cima deste arquivo; senao, pergunta.
set "ARQUIVO=%~1"
if "%ARQUIVO%"=="" set /p "ARQUIVO=Arraste a planilha para esta janela (ou cole o caminho) e tecle Enter: "
set "ARQUIVO=%ARQUIVO:"=%"

if not exist "%ARQUIVO%" (
    echo [ERRO] Arquivo nao encontrado: %ARQUIVO%
    pause & exit /b 1
)

rem Cache local: copia o .exe para o computador (mais rapido que abrir pela rede)
set "CACHE=%LOCALAPPDATA%\PlanilhaEngine_Cache"
if not exist "%CACHE%" mkdir "%CACHE%"
xcopy /D /Y /Q "%EXE%" "%CACHE%\" >nul 2>&1
if exist "%CACHE%\PlanilhaEngine.exe" set "EXE=%CACHE%\PlanilhaEngine.exe"

echo.
echo Processando... aguarde.
echo.
"%EXE%" "%PERFIL%" "%ARQUIVO%" "%SAIDA%"
set "COD=%ERRORLEVEL%"
echo.
if "%COD%"=="0" (
    echo [SUCESSO] Planilha pronta na pasta: %SAIDA%
    set /p "ABRIR=Abrir a pasta agora? (S/N): "
    if /i "!ABRIR!"=="S" start "" "%SAIDA%"
) else (
    echo [ERRO] Nao foi possivel concluir. Veja a mensagem acima.
)
echo.
pause
