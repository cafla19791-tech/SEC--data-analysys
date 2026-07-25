@echo off
REM ContAgil / WinPython - fluxos BNDES indiretos
REM Execute este .bat no cmd (NAO use: python este.bat)
setlocal EnableExtensions
cd /d "%~dp0"

set "WINPY=%CD%"
set "DADOS=%WINPY%\dados"
set "SAIDA=%WINPY%\saida"
set "FATORES=%WINPY%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
set "SCRIPT=%WINPY%\scripts\contagil_fluxos.py"

REM Localiza python.exe do WinPython (evita stub da Microsoft Store)
set "PY="
if exist "%WINPY%\python.exe" set "PY=%WINPY%\python.exe"
if not defined PY if exist "%WINPY%\python\python.exe" set "PY=%WINPY%\python\python.exe"
for /d %%D in ("%WINPY%\python-*") do (
  if exist "%%~D\python.exe" set "PY=%%~D\python.exe"
)
if not defined PY if exist "%WINPY%\..\python.exe" set "PY=%WINPY%\..\python.exe"

if not defined PY (
  echo [ERRO] python.exe do WinPython nao encontrado em "%WINPY%"
  echo Procure com: dir /s /b "%WINPY%\python.exe"
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo [ERRO] Nao achei "%SCRIPT%"
  echo Rode antes: baixar_e_rodar_fluxos.ps1
  exit /b 1
)
if not exist "%WINPY%\scripts\gerar_fluxos.py" (
  echo [ERRO] Falta scripts\gerar_fluxos.py - rode baixar_e_rodar_fluxos.ps1
  exit /b 1
)

echo Python: %PY%
"%PY%" "%SCRIPT%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --arquivo-fatores "%FATORES%"
endlocal
exit /b %ERRORLEVEL%
