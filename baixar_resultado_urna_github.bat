@echo off
REM Baixa o CSV nacional ja consolidado (serie + modelo + votos) do GitHub.
REM Na RFB o GitHub ja funcionou; TSE e Archive.org nao.
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "RUNNER=%WINPY%\baixar_boletins_urna_2022.py"
if not exist "%WINPY%\sec_scripts" mkdir "%WINPY%\sec_scripts"
set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/tse-boletins-urna-209b"
if not exist "%RUNNER%" (
  curl.exe -fsSL -o "%RUNNER%" "%BASE%/baixar_boletins_urna_2022.py"
)
if not exist "%WINPY%\sec_scripts\baixar_boletins_urna_2022.py" (
  curl.exe -fsSL -o "%WINPY%\sec_scripts\baixar_boletins_urna_2022.py" "%BASE%/scripts/baixar_boletins_urna_2022.py"
)
echo Baixando CSV nacional do GitHub (TLS do GitHub ja passou na RFB)...
"%PYTHON%" "%RUNNER%" --somente-resultado-github --massa-dados "%WINPY%\dados" --pasta-saida "%WINPY%\saida"
if exist "%WINPY%\saida\tse2022\urnas_2t_presidente.csv" (
  echo OK: %WINPY%\saida\tse2022\urnas_2t_presidente.csv
) else (
  echo FALHOU. Tente no Edge o HTML em saida\tse2022\baixar_boletins_links.html
)
echo.
pause
endlocal
exit /b 0
