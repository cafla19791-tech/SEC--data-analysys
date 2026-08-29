@echo off
REM Discriminativo de taxas basicas de juros reais por pais (BIS CBPOL + CPI).
REM Uso (cmd, na pasta winpython ou no clone do repo):
REM   discriminativo_juros_reais_paises.bat
REM   discriminativo_juros_reais_paises.bat --ano-inicio 2000
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON=python.exe"
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
if exist "%~dp0..\python.exe" set "PYTHON=%~dp0..\python.exe"

set "SCRIPT=%~dp0scripts\discriminativo_juros_reais_paises.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0sec_scripts\discriminativo_juros_reais_paises.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0discriminativo_juros_reais_paises.py"

if not exist "%SCRIPT%" (
  echo [ERRO] Nao achei discriminativo_juros_reais_paises.py
  exit /b 1
)

findstr /C:"juros-reais-paises-20260829" "%SCRIPT%" >nul
if errorlevel 1 (
  echo [ERRO] Script desatualizado
  exit /b 1
)

"%PYTHON%" -m pip install "pandas>=2.0" "openpyxl>=3.1" "requests>=2.28"
"%PYTHON%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
