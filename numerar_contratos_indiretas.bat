@echo off
REM Numera contratos BNDES INDIRETAS: 1-2002, 2-2002, ... / 1-2003, ...
REM Gera Excel com uma aba por ano em saida\BNDES_INDIRETAS_NUMERADOS.xlsx
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\numerar_contratos_indiretas.log"

echo ========================================
echo  NUMERAR CONTRATOS INDIRETAS N-AAAA
echo ========================================

set "REPO="
if exist "%WINPY%\scripts\numerar_contratos_indiretas.py" set "REPO=%WINPY%"
if not defined REPO (
  for /d %%D in ("%WINPY%\SEC--data-analysys*") do (
    if exist "%%~fD\scripts\numerar_contratos_indiretas.py" set "REPO=%%~fD"
  )
)
if not defined REPO (
  echo [ERRO] scripts\numerar_contratos_indiretas.py nao encontrado
  echo Extraia o ZIP do branch neste pasta winpython e rode de novo.
  goto :FIM
)

set "PYSCRIPT=%REPO%\scripts\numerar_contratos_indiretas.py"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\BNDES_INDIRETAS_NUMERADOS.xlsx"
set "DADOS=%WINPY%\dados"
set "PYTHONPATH=%REPO%;%PYTHONPATH%"

echo Repo : %REPO%
echo Dados: %DADOS%
echo Saida: %SAIDA%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not exist "%DADOS%" (
  echo [ERRO] Pasta dados nao encontrada: %DADOS% >> "%LOG%"
  echo [ERRO] Pasta dados nao encontrada: %DADOS%
  goto :SHOWLOG
)

"%PYTHON%" "%PYSCRIPT%" --pasta-dados "%DADOS%" --saida "%SAIDA%" --ano-min 2002 >> "%LOG%" 2>&1
echo ---- fim %ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"

:SHOWLOG
type "%LOG%"
echo.

if exist "%SAIDA%" (
  echo OK: %SAIDA%
  explorer /select,"%SAIDA%"
) else (
  echo FALHOU. Veja %LOG%
)

:FIM
pause
endlocal
exit /b 0
