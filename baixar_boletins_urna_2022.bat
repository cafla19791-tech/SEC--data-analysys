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

REM Nao use winpython\scripts\ — no Windows isso cai em Scripts\ (pip).
set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tse-boletins-urna-209b"
if not exist "%WINPY%\sec_scripts" mkdir "%WINPY%\sec_scripts"
if not exist "%WINPY%\sec_scripts\baixar_boletins_urna_2022.py" (
  echo Baixando script do GitHub para sec_scripts\ ...
  echo Baixando script do GitHub >> "%LOG%"
  curl.exe -fsSL -o "%WINPY%\sec_scripts\baixar_boletins_urna_2022.py" "%BASE%/scripts/baixar_boletins_urna_2022.py"
  if errorlevel 1 (
    echo [ERRO] Nao consegui baixar o .py. Cole no cmd o bloco curl do README. >> "%LOG%"
    echo [ERRO] Falta o script e o download do GitHub falhou.
    goto :SHOW
  )
)
if not exist "%RUNNER%" (
  curl.exe -fsSL -o "%RUNNER%" "%BASE%/baixar_boletins_urna_2022.py" >nul 2>&1
)
set "PKG=%WINPY%\sec_scripts\baixar_boletins_urna_2022.py"
if not exist "%DADOS%" mkdir "%DADOS%"
if not exist "%SAIDA%" mkdir "%SAIDA%"
if not exist "%RAW%" mkdir "%RAW%"
if not exist "%OUT%" mkdir "%OUT%"

echo Instalando pandas/requests se preciso... >> "%LOG%"
"%PYTHON%" -m pip install "pandas>=2.0" "requests>=2.28" >> "%LOG%" 2>&1

set "PYTHONPATH=%WINPY%;%PYTHONPATH%"
if exist "%RUNNER%" (
  "%PYTHON%" "%RUNNER%" --usar-curl --workers 1 --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" %* >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%PKG%" --usar-curl --workers 1 --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" %* >> "%LOG%" 2>&1
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
  echo Na RFB o Python nao fecha TLS com archive.org.
  echo Rode em vez disso: baixar_zips_urna_curl.bat
  echo Ou abra o HTML de links no Edge e salve os ZIPs em:
  echo   %RAW%
  echo Depois: baixar_boletins_urna_2022.bat --somente-processar
  if exist "%OUT%\baixar_boletins_links.html" start "" "%OUT%\baixar_boletins_links.html"
)

echo.
pause
endlocal
exit /b 0
