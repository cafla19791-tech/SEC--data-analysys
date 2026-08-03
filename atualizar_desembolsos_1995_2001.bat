@echo off
REM Atualiza pelo IPCA a base BNDES DESEMBOLSO 1995-2001
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\atualizar_desembolsos_1995_2001.log"

echo ========================================
echo  DESEMBOLSOS BNDES 1995-2001 x IPCA
echo ========================================

set "REPO="
if exist "%WINPY%\scripts\atualizar_desembolsos_ipca_1995_2001.py" set "REPO=%WINPY%"
if not defined REPO if exist "%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9\scripts\atualizar_desembolsos_ipca_1995_2001.py" (
  set "REPO=%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9"
)
if not defined REPO (
  for /d %%D in ("%WINPY%\SEC--data-analysys*") do (
    if exist "%%~fD\scripts\atualizar_desembolsos_ipca_1995_2001.py" set "REPO=%%~fD"
  )
)
if not defined REPO (
  echo [ERRO] scripts\atualizar_desembolsos_ipca_1995_2001.py nao encontrado
  goto :FIM
)

set "PYSCRIPT=%REPO%\scripts\atualizar_desembolsos_ipca_1995_2001.py"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "IPCA=%WINPY%\IPCA_MENSAL.xlsx"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\DESEMBOLSOS_1995_2001_IPCA.xlsx"
set "PYTHONPATH=%REPO%;%PYTHONPATH%"

echo Repo : %REPO%
echo Saida: %SAIDA%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
REM Gera Resumo_Anual (subtotais + TOTAL) e Detalhe com cada linha atualizada pelo IPCA
if exist "%IPCA%" (
  "%PYTHON%" "%PYSCRIPT%" --baixar --ipca "%IPCA%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%PYSCRIPT%" --baixar --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
)
echo ---- fim %ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"
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
