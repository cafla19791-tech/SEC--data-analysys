@echo off
REM Ranking anual das taxas basicas de juros reais (uma aba por ano, 1995-2026).
REM Uso:
REM   discriminativo_ranking_juros_reais.bat
REM   discriminativo_ranking_juros_reais.bat --ano-inicio 1995 --ano-fim 2026
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON=python.exe"
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
if exist "%~dp0..\python.exe" set "PYTHON=%~dp0..\python.exe"

set "SCRIPT=%~dp0scripts\discriminativo_ranking_juros_reais.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0sec_scripts\discriminativo_ranking_juros_reais.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0discriminativo_ranking_juros_reais.py"

if not exist "%SCRIPT%" (
  echo [ERRO] Nao achei discriminativo_ranking_juros_reais.py
  exit /b 1
)

findstr /C:"ranking-juros-reais-1995-2026-20260829" "%SCRIPT%" >nul
if errorlevel 1 (
  echo [ERRO] Script desatualizado
  exit /b 1
)

"%PYTHON%" -m pip install "pandas>=2.0" "openpyxl>=3.1" "requests>=2.28"
"%PYTHON%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
