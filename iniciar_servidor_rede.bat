@echo off
cd /d "%~dp0"
python -m waitress --host=0.0.0.0 --port=5050 robo_ponto_web:app
pause
