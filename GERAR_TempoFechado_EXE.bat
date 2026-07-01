@echo off
title Tempo Fechado v8.21.38 - Upload Manual PDFs - Gerar EXE
echo Tempo Fechado v8.21.38 - Upload Manual PDFs
where py >nul 2>nul
if errorlevel 1 ( echo ERRO: comando py nao encontrado. & pause & exit /b 1 )
py -m pip install --upgrade pyinstaller flask pandas openpyxl numpy python-dateutil pytz werkzeug jinja2 pywin32 pdfplumber pdfminer.six waitress
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo Gerando EXE em modo sem console ^(console=False no TempoFechado.spec^)...
py -m PyInstaller TempoFechado.spec
echo.
echo EXE gerado em dist\TempoFechado\TempoFechado.exe
pause

