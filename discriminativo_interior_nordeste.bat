@echo off
REM ContAgil: discriminativo Lula x Bolsonaro no interior do Nordeste (2022 2T)
cd /d "%~dp0"
python discriminativo_interior_nordeste.py %*
if errorlevel 1 pause
