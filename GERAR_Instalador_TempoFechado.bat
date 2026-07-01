@echo off
title Tempo Fechado v8.21.38 - Upload Manual PDFs - Gerar Instalador
if not exist "dist\TempoFechado\TempoFechado.exe" (
  echo ERRO: gere primeiro o EXE com GERAR_TempoFechado_EXE.bat
  pause
  exit /b 1
)
set ISCC=C:\Projetos\innosetup-portable\app\ISCC.exe
if not exist "%ISCC%" (
  echo ERRO: ajuste o caminho do ISCC.exe neste BAT.
  pause
  exit /b 1
)
"%ISCC%" TempoFechado_Installer.iss
echo Instalador em installer_output\Setup_TempoFechado_v8_21_38.exe
pause

