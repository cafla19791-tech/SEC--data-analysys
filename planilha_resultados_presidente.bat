@echo off
REM ContAgil: gera as planilhas de Presidente (região / UF / município / zona / urna)
cd /d "%~dp0"
python planilha_resultados_presidente.py %*
if errorlevel 1 pause
