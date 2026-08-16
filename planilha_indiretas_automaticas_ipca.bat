@echo off
REM Gera planilha das indiretas automaticas BNDES (aba por ano + IPCA jul/2026).
REM Uso (cmd, na pasta winpython ou no clone do repo):
REM   planilha_indiretas_automaticas_ipca.bat
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON=python.exe"
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
if exist "%~dp0..\python.exe" set "PYTHON=%~dp0..\python.exe"

set "SCRIPT=%~dp0scripts\planilha_indiretas_automaticas_ipca.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0sec_scripts\planilha_indiretas_automaticas_ipca.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0planilha_indiretas_automaticas_ipca.py"

if not exist "%SCRIPT%" (
  echo [ERRO] Nao achei planilha_indiretas_automaticas_ipca.py
  exit /b 1
)

findstr /C:"indiretas-automaticas-ipca-20260816a" "%SCRIPT%" >nul
if errorlevel 1 (
  echo [ERRO] Script desatualizado
  exit /b 1
)

"%PYTHON%" -m pip install "pandas>=2.0" "openpyxl>=3.1" "xlsxwriter>=3.0" "numpy>=1.24" "requests>=2.28"
"%PYTHON%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
