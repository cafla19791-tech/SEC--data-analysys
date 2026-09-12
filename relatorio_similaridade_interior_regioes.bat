@echo off
REM ContAgil: similaridade 2014 x 2018 x 2022 no interior das demais regioes
cd /d "%~dp0"
python relatorio_similaridade_interior_regioes.py %*
if errorlevel 1 pause
