@echo off
REM ContAgil: discriminativo Presidente 2014 x 2018 x 2022
cd /d "%~dp0"
python discriminativo_resultados_presidente.py %*
if errorlevel 1 pause
