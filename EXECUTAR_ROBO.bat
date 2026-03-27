@echo off
cls
echo ======================================================
echo    INICIANDO AUTOMACAO DE NOTAS - DIALOGO ENGENHARIA
echo ======================================================

:: Entra na pasta simplificada sem acentos
cd /d "C:\Robo_Notas"

echo [1/2] Buscando links no Outlook...
powershell -ExecutionPolicy Bypass -File "pegar_links.ps1"

echo.
echo [2/2] Iniciando download das Notas Fiscais...
python "bot_notas.py"

echo.
echo Processo finalizado com sucesso!
pause