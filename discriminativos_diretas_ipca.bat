@echo off
REM Discriminativos OPERACOES DIRETAS com valor atualizado pelo IPCA (30/06/2026)
REM 4 abas: 2002 | 2003-2018 | 2019-2022 | 2023-atual
REM ContAgil: usa pasta sec_scripts (NAO scripts) para nao colidir com WinPython\Scripts
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\discriminativos_diretas_ipca.log"

echo ========================================
echo  DISCRIMINATIVOS DIRETAS + IPCA
echo ========================================

set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "IPCA=%WINPY%\IPCA_MENSAL.xlsx"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\DISCRIMINATIVOS_DIRETAS_IPCA.xlsx"
set "RUNNER=%WINPY%\discriminativos_diretas_ipca.py"
set "PKG=%WINPY%\sec_scripts"

if not exist "%RUNNER%" (
  echo [ERRO] Falta discriminativos_diretas_ipca.py
  echo Rode baixar_discriminativos.ps1 de novo.
  goto :FIM
)
if not exist "%PKG%\discriminativos_diretas_ipca.py" (
  echo [ERRO] Falta sec_scripts\discriminativos_diretas_ipca.py
  echo Rode baixar_discriminativos.ps1 de novo.
  goto :FIM
)
if not exist "%PKG%\calcular_diretas_ipca_selic.py" (
  echo [ERRO] Falta sec_scripts\calcular_diretas_ipca_selic.py
  echo Rode baixar_discriminativos.ps1 de novo.
  goto :FIM
)
if not exist "%PKG%\gerar_fluxos.py" (
  echo [ERRO] Falta sec_scripts\gerar_fluxos.py
  echo Rode baixar_discriminativos.ps1 de novo.
  goto :FIM
)

REM Localiza a planilha
set "EXCEL="
if exist "%WINPY%\OPERACOES DIRETAS2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS 2002 A 302026.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS 2002 A 302026.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx"
if not defined EXCEL (
  for %%F in ("%WINPY%\*DIRETA*.xlsx") do (
    if not defined EXCEL set "EXCEL=%%~fF"
  )
)

echo WinPy : %WINPY%
echo Excel : %EXCEL%
echo Saida : %SAIDA%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not defined EXCEL (
  echo [ERRO] Planilha OPERACOES DIRETAS nao encontrada na pasta winpython >> "%LOG%"
  echo [ERRO] Planilha OPERACOES DIRETAS nao encontrada na pasta winpython
  goto :SHOWLOG
)

cd /d "%WINPY%"
if exist "%IPCA%" (
  "%PYTHON%" "%RUNNER%" --excel "%EXCEL%" --ipca "%IPCA%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%RUNNER%" --excel "%EXCEL%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
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
