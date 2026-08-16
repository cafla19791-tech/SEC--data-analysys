@echo off
REM Discriminativo nao automaticas BNDES (abas por periodo + IPCA 31/07/2026).
setlocal EnableExtensions
cd /d "%~dp0"
set "PYTHON=python.exe"
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
set "SCRIPT=%~dp0scripts\discriminativo_naoautomaticas_ipca.py"
if not exist "%SCRIPT%" set "SCRIPT=%~dp0discriminativo_naoautomaticas_ipca.py"
if not exist "%SCRIPT%" (
  echo [ERRO] Script nao encontrado
  exit /b 1
)
findstr /C:"naoautomaticas-discriminativo-20260816a" "%SCRIPT%" >nul
if errorlevel 1 (
  echo [ERRO] Script desatualizado
  exit /b 1
)
"%PYTHON%" -m pip install "pandas>=2.0" "openpyxl>=3.1" "xlsxwriter>=3.0" "numpy>=1.24" "requests>=2.28"
"%PYTHON%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
