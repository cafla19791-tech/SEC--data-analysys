@echo off
REM Excel do discriminativo → PDF (marcadores) + HTML (abas clicaveis).
REM Uso:
REM   discriminativo_para_pdf.bat --entrada output\discriminativo_ranking_juros_reais.xlsx
setlocal EnableExtensions
cd /d "%~dp0"

set "PYTHON=python.exe"
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
if exist "%~dp0..\python.exe" set "PYTHON=%~dp0..\python.exe"

set "SCRIPT=%~dp0scripts\discriminativo_para_pdf.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0discriminativo_para_pdf.py"

if not exist "%SCRIPT%" (
  echo [ERRO] Nao achei discriminativo_para_pdf.py
  exit /b 1
)

findstr /C:"discriminativo-pdf-20260829" "%SCRIPT%" >nul
if errorlevel 1 (
  echo [ERRO] Script desatualizado
  exit /b 1
)

"%PYTHON%" -m pip install "reportlab>=4.0" "pypdf>=4.0" "openpyxl>=3.1"
"%PYTHON%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
