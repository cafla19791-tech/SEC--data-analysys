@echo off
REM Discriminativo municipal: UE2020 vs urnas anteriores a 2020
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "RUNNER=%WINPY%\discriminativo_urnas_municipio.py"
set "PKG=%WINPY%\sec_scripts\discriminativo_urnas_municipio.py"
if not exist "%PKG%" set "PKG=%WINPY%\scripts\discriminativo_urnas_municipio.py"
set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tse-boletins-urna-209b"
if not exist "%WINPY%\sec_scripts" mkdir "%WINPY%\sec_scripts"
if not exist "%WINPY%\sec_scripts\discriminativo_urnas_municipio.py" (
  curl.exe -fsSL -o "%WINPY%\sec_scripts\discriminativo_urnas_municipio.py" "%BASE%/scripts/discriminativo_urnas_municipio.py"
)
if not exist "%RUNNER%" (
  curl.exe -fsSL -o "%RUNNER%" "%BASE%/discriminativo_urnas_municipio.py"
)
if not exist "%WINPY%\saida\tse2022\urnas_2t_presidente.csv" (
  if not exist "%WINPY%\saida\tse2022\urnas_2t_presidente.csv.gz" (
    echo Falta urnas_2t_presidente.csv. Rode antes:
    echo   python baixar_boletins_urna_2022.py --somente-resultado-github
    pause
    exit /b 1
  )
)
"%PYTHON%" "%RUNNER%" --pasta-saida "%WINPY%\saida" %*
echo.
if exist "%WINPY%\saida\tse2022\discriminativo_municipio_ue2020.xlsx" (
  echo OK: %WINPY%\saida\tse2022\discriminativo_municipio_ue2020.xlsx
  explorer /select,"%WINPY%\saida\tse2022\discriminativo_municipio_ue2020.xlsx"
)
pause
endlocal
exit /b 0
