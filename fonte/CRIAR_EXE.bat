@echo off
chcp 65001 >nul
rem Rode UMA vez, em um computador Windows com Python instalado, para gerar o PlanilhaEngine.exe.
cd /d "%~dp0"

set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY python --version >nul 2>&1 && set "PY=python"
if not defined PY (
    echo [ERRO] Python nao encontrado neste computador.
    echo.
    echo Instale em https://www.python.org/downloads/ e, na primeira tela do instalador,
    echo MARQUE a opcao "Add python.exe to PATH". Depois feche esta janela e rode de novo.
    pause
    exit /b 1
)

%PY% -m pip install openpyxl pyinstaller || goto erro
%PY% -m PyInstaller --onefile --name PlanilhaEngine planilha_engine.py || goto erro
copy /Y dist\PlanilhaEngine.exe ..\dist\PlanilhaEngine.exe || goto erro
echo.
echo Pronto! O arquivo ficou em ..\dist\PlanilhaEngine.exe
pause
exit /b 0
:erro
echo [ERRO] Algo deu errado. Veja a mensagem acima.
pause
exit /b 1
