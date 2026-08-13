@echo off
REM Calcula colunas K-N em OPERACOES DIRETAS - 2002 a 2018.xlsx
REM Procura o ZIP extraido: SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9

setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\calcular_diretas_ipca_selic.log"

echo ========================================
echo  OPERACOES DIRETAS — colunas K L M N
echo ========================================
echo Pasta: %WINPY%
echo.

REM ---- localizar repo com scripts\ ----
set "REPO="
if exist "%WINPY%\scripts\calcular_diretas_ipca_selic.py" set "REPO=%WINPY%"
if not defined REPO if exist "%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9\scripts\calcular_diretas_ipca_selic.py" (
  set "REPO=%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9"
)
if not defined REPO (
  for /d %%D in ("%WINPY%\SEC--data-analysys*") do (
    if exist "%%~fD\scripts\calcular_diretas_ipca_selic.py" set "REPO=%%~fD"
  )
)
if not defined REPO (
  echo [ERRO] Nao achei scripts\calcular_diretas_ipca_selic.py
  echo Extraia o ZIP do GitHub nesta pasta winpython.
  goto :FIM
)
echo Repo: %REPO%
set "PYSCRIPT=%REPO%\scripts\calcular_diretas_ipca_selic.py"

REM ---- excel ----
set "EXCEL=%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx"
if not exist "%EXCEL%" set "EXCEL=%WINPY%\OPERACOES DIRETAS.xlsx"
if not exist "%EXCEL%" (
  echo [ERRO] Excel nao encontrado: OPERACOES DIRETAS - 2002 a 2018.xlsx
  goto :FIM
)

if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\OPERACOES DIRETAS - 2002 a 2018_calculado.xlsx"
set "SELIC=%WINPY%\selic_mensal.xlsx"
set "IPCA=%WINPY%\IPCA_MENSAL.xlsx"

REM ---- python WinPython ----
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

echo Entrada : %EXCEL%
echo Saida   : %SAIDA%
echo Python  : %PYTHON%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
echo Repo=%REPO%>> "%LOG%"
echo Excel=%EXCEL%>> "%LOG%"
echo Saida=%SAIDA%>> "%LOG%"

set "PYTHONPATH=%REPO%;%PYTHONPATH%"

if exist "%SELIC%" if exist "%IPCA%" (
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --selic "%SELIC%" --ipca "%IPCA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else if exist "%SELIC%" (
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --selic "%SELIC%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else if exist "%IPCA%" (
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --ipca "%IPCA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
)

echo ---- fim ERRORLEVEL=%ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"
echo.
type "%LOG%"
echo.

if exist "%SAIDA%" (
  echo ========================================
  echo OK — SAIDA:
  echo %SAIDA%
  echo ========================================
  explorer /select,"%SAIDA%"
) else (
  echo FALHOU. Abra o log:
  echo %LOG%
)

:FIM
echo.
pause
endlocal
exit /b 0
