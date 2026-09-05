@echo off
REM ContAgil: interior do Nordeste ate 40 mil eleitores (2T 2014, 2018 e 2022)
cd /d "%~dp0"
python discriminativo_interior_nordeste_ate_40mil.py %*
if errorlevel 1 pause
