@echo off
REM Discriminativos OPERACOES DIRETAS com valor atualizado pelo IPCA (30/06/2026)
REM 4 abas: 2002 | 2003-2018 | 2019-2022 | 2023-atual
REM Em cada aba: ordena CLIENTE+DATA, subtotal por cliente e quebra de pagina.
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\discriminativos_diretas_ipca.log"

echo ========================================
echo  DISCRIMINATIVOS DIRETAS + IPCA
echo ========================================

set "REPO="
if exist "%WINPY%\scripts\discriminativos_diretas_ipca.py" set "REPO=%WINPY%"
if not defined REPO if exist "%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9\scripts\discriminativos_diretas_ipca.py" (
  set "REPO=%WINPY%\SEC--data-analysys-cursor-contagil-fluxos-seguro-e4e9"
)
if not defined REPO if exist "%WINPY%\SEC--data-analysys-cursor-discriminativos-diretas-ipca-e4e9\scripts\discriminativos_diretas_ipca.py" (
  set "REPO=%WINPY%\SEC--data-analysys-cursor-discriminativos-diretas-ipca-e4e9"
)
if not defined REPO (
  for /d %%D in ("%WINPY%\SEC--data-analysys*") do (
    if exist "%%~fD\scripts\discriminativos_diretas_ipca.py" set "REPO=%%~fD"
  )
)
if not defined REPO (
  echo [ERRO] scripts\discriminativos_diretas_ipca.py nao encontrado
  echo Extraia o ZIP do branch neste pasta winpython e rode de novo.
  goto :FIM
)

set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "IPCA=%WINPY%\IPCA_MENSAL.xlsx"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\DISCRIMINATIVOS_DIRETAS_IPCA.xlsx"
set "PYTHONPATH=%REPO%;%PYTHONPATH%"

REM Checagem dos modulos (evita ModuleNotFoundError opaco)
if not exist "%REPO%\scripts\__init__.py" (
  echo [ERRO] Falta scripts\__init__.py - rode baixar_discriminativos.ps1 de novo
  goto :FIM
)
if not exist "%REPO%\scripts\calcular_diretas_ipca_selic.py" (
  echo [ERRO] Falta scripts\calcular_diretas_ipca_selic.py - rode baixar_discriminativos.ps1 de novo
  goto :FIM
)
if not exist "%REPO%\scripts\gerar_fluxos.py" (
  echo [ERRO] Falta scripts\gerar_fluxos.py - rode baixar_discriminativos.ps1 de novo
  goto :FIM
)

REM Localiza a planilha (nome informado pelo usuario ou variantes)
set "EXCEL="
if exist "%WINPY%\OPERAÇÕES DIRETAS2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERAÇÕES DIRETAS2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERAÇÕES DIRETAS 2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERAÇÕES DIRETAS 2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS 2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS 2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx"
if not defined EXCEL (
  for %%F in ("%WINPY%\*DIRETA*.xlsx") do (
    if not defined EXCEL set "EXCEL=%%~fF"
  )
)

echo Repo : %REPO%
echo Excel: %EXCEL%
echo Saida: %SAIDA%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not defined EXCEL (
  echo [ERRO] Planilha OPERACOES DIRETAS nao encontrada na pasta winpython >> "%LOG%"
  echo [ERRO] Planilha OPERACOES DIRETAS nao encontrada na pasta winpython
  goto :SHOWLOG
)

REM Roda como modulo (PYTHONPATH=REPO) para import scripts.* funcionar no ContAgil
cd /d "%REPO%"
if exist "%IPCA%" (
  "%PYTHON%" -m scripts.discriminativos_diretas_ipca --excel "%EXCEL%" --ipca "%IPCA%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else (
  "%PYTHON%" -m scripts.discriminativos_diretas_ipca --excel "%EXCEL%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
)
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
