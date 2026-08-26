@echo off
REM Boletins de Urna 2022 (2o turno, Presidente) — 28 UFs
REM ContAgil: copie este .bat para a pasta winpython e de duplo-clique.
REM NAO use a pasta Scripts do WinPython (conflito). Codigo em sec_scripts\.
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\baixar_boletins_urna_2022.log"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "DADOS=%WINPY%\dados"
set "SAIDA=%WINPY%\saida"
set "RAW=%DADOS%\tse2022\raw"
set "OUT=%SAIDA%\tse2022"
set "RUNNER=%WINPY%\baixar_boletins_urna_2022.py"
set "PKG=%WINPY%\sec_scripts\baixar_boletins_urna_2022.py"
if not exist "%PKG%" set "PKG=%WINPY%\scripts\baixar_boletins_urna_2022.py"

echo ========================================
echo  BOLETINS DE URNA 2022 — 28 UFs
echo ========================================
echo WinPy : %WINPY%
echo ZIPs  : %RAW%
echo Saida : %OUT%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not exist "%RUNNER%" if not exist "%PKG%" (
  echo [ERRO] Falta baixar_boletins_urna_2022.py >> "%LOG%"
  echo [ERRO] Falta o script. Rode baixar_boletins_urna_2022.ps1 nesta pasta winpython.
  goto :SHOW
)
if not exist "%DADOS%" mkdir "%DADOS%"
if not exist "%SAIDA%" mkdir "%SAIDA%"
if not exist "%RAW%" mkdir "%RAW%"
if not exist "%OUT%" mkdir "%OUT%"

echo Instalando pandas/requests se preciso... >> "%LOG%"
"%PYTHON%" -m pip install "pandas>=2.0" "requests>=2.28" >> "%LOG%" 2>&1

set "PYTHONPATH=%WINPY%;%PYTHONPATH%"
if exist "%RUNNER%" (
  "%PYTHON%" "%RUNNER%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" %* >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%PKG%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" %* >> "%LOG%" 2>&1
)
echo ---- fim ERRORLEVEL=%ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"

:SHOW
echo.
type "%LOG%"
echo.

if exist "%OUT%\urnas_2t_presidente.csv" (
  echo OK: %OUT%\urnas_2t_presidente.csv
  explorer /select,"%OUT%\urnas_2t_presidente.csv"
) else (
  echo FALHOU. Veja %LOG%
  echo Se o TSE bloqueou o download ^(403^), baixe os ZIPs no navegador:
  echo   https://dadosabertos.tse.jus.br/dataset/resultados-2022-boletim-de-urna
  echo e copie para %RAW%
  echo Depois rode de novo com: baixar_boletins_urna_2022.bat --somente-processar
)

echo.
pause
endlocal
exit /b 0
