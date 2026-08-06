@echo off
REM Fluxos e impacto fiscal — BNDES INDIRETAS (ContAgil / WinPython)
REM ContAgil: usa pasta sec_scripts (NAO scripts) para nao colidir com WinPython\Scripts
REM
REM Uso:
REM   .\contagil_fluxos_seguro.bat
REM   .\contagil_fluxos_seguro.bat "dados\Operacoes Indiretas 2002.xlsx"
REM   .\contagil_fluxos_seguro.bat "dados\Operacoes Indiretas 2002.xlsx" 50
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\contagil_fluxos_seguro.log"

echo ========================================
echo  FLUXOS BNDES INDIRETAS (seguro)
echo ========================================

set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
REM Evita UnicodeEncodeError no console Windows (cp1252)
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "DADOS=%WINPY%\dados"
set "SAIDA=%WINPY%\saida"
set "FATORES=%WINPY%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
set "RUNNER=%WINPY%\contagil_fluxos_seguro.py"
set "PKG=%WINPY%\sec_scripts"
set "INPUT=%~1"
set "MAXC=%~2"

if not exist "%RUNNER%" (
  echo [ERRO] Falta contagil_fluxos_seguro.py
  echo Rode baixar_contagil_fluxos_seguro.ps1 neste pasta winpython.
  goto :FIM
)
if not exist "%PKG%\contagil_fluxos_seguro.py" (
  echo [ERRO] Falta sec_scripts\contagil_fluxos_seguro.py
  echo Rode baixar_contagil_fluxos_seguro.ps1 neste pasta winpython.
  goto :FIM
)
if not exist "%PKG%\gerar_fluxos.py" (
  echo [ERRO] Falta sec_scripts\gerar_fluxos.py
  echo Rode baixar_contagil_fluxos_seguro.ps1 neste pasta winpython.
  goto :FIM
)
if not exist "%FATORES%" (
  echo [ERRO] Falta fator_acumulado_SELIC_TJLP_TLP.xlsx na pasta winpython
  echo Coloque o arquivo de fatores SELIC/TJLP/TLP ao lado deste .bat
  goto :FIM
)

echo WinPy  : %WINPY%
echo Dados  : %DADOS%
echo Saida  : %SAIDA%
echo Fatores: %FATORES%
if not "%INPUT%"=="" echo Input  : %INPUT%
if not "%MAXC%"=="" echo Max    : %MAXC% contratos/arquivo
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"

if not "%INPUT%"=="" goto :UM_ARQUIVO
goto :MASSA

:UM_ARQUIVO
if exist "%INPUT%" goto :RUN_INPUT
if exist "%WINPY%\%INPUT%" (
  set "INPUT=%WINPY%\%INPUT%"
  goto :RUN_INPUT
)
echo [ERRO] Arquivo nao encontrado: %INPUT% >> "%LOG%"
echo [ERRO] Arquivo nao encontrado: %INPUT%
goto :SHOWLOG

:RUN_INPUT
if not "%MAXC%"=="" (
  "%PYTHON%" "%RUNNER%" --input "%INPUT%" --pasta-saida "%SAIDA%" --fatores "%FATORES%" --max-contratos %MAXC% >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%RUNNER%" --input "%INPUT%" --pasta-saida "%SAIDA%" --fatores "%FATORES%" >> "%LOG%" 2>&1
)
goto :AFTER

:MASSA
if not exist "%DADOS%" (
  echo [ERRO] Pasta dados nao encontrada: %DADOS% >> "%LOG%"
  echo [ERRO] Pasta dados nao encontrada: %DADOS%
  goto :SHOWLOG
)
if not "%MAXC%"=="" (
  "%PYTHON%" "%RUNNER%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%" --max-contratos %MAXC% >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%RUNNER%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%" >> "%LOG%" 2>&1
)

:AFTER
echo ---- fim %ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"

:SHOWLOG
type "%LOG%"
echo.
echo Log: %LOG%
echo Saida: %SAIDA%

:FIM
pause
endlocal
exit /b 0
