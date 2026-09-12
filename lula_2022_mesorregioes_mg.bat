@echo off
REM ContAgil: municipios de MG com vitoria de Lula em 2022 2T, por mesorregiao
cd /d "%~dp0"
python lula_2022_mesorregioes_mg.py %*
if errorlevel 1 pause
