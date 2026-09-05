@echo off
REM ContAgil: UE2020 vs urnas antigas em municipios ate 50 mil habitantes
cd /d "%~dp0"
python discriminativo_modelo_urna_ate_50mil.py %*
if errorlevel 1 pause
