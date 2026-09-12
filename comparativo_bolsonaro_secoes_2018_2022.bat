@echo off
REM ContAgil: Bolsonaro x Haddad 2018 vs Bolsonaro x Lula 2022 por secao
cd /d "%~dp0"
python comparativo_bolsonaro_secoes_2018_2022.py %*
if errorlevel 1 pause
