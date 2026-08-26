@echo off
REM Baixa os 28 ZIPs de Boletim de Urna 2022 com curl.exe do Windows.
REM Na RFB o Python/OpenSSL falha no TLS do Archive.org (handshake failure).
REM O curl.exe de C:\Windows\System32 usa SChannel e costuma passar.
REM
REM Copie este arquivo para a pasta winpython e de duplo-clique.
REM NAO use a pasta Scripts do WinPython.
setlocal EnableExtensions
cd /d "%~dp0"
if not exist "%CD%\dados" if exist "%CD%\..\..\dados" cd /d "%CD%\..\.."
if not exist "%CD%\dados" if exist "%CD%\..\dados" cd /d "%CD%\.."

set "WINPY=%CD%"
set "RAW=%WINPY%\dados\tse2022\raw"
set "SAIDA=%WINPY%\saida\tse2022"
set "LOG=%WINPY%\baixar_zips_urna_curl.log"
set "UA=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
set "TSE=https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"
set "IA=https://web.archive.org/web/2023id_/https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"
set "CURL=%SystemRoot%\System32\curl.exe"
if not exist "%CURL%" set "CURL=curl.exe"
if not exist "%RAW%" mkdir "%RAW%"
if not exist "%SAIDA%" mkdir "%SAIDA%"

echo ========================================
echo  BOLETINS DE URNA 2022 — curl.exe
echo ========================================
echo WinPy : %WINPY%
echo curl  : %CURL%
echo ZIPs  : %RAW%
echo.
echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
echo curl=%CURL% >> "%LOG%"

for %%U in (AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO ZZ) do (
  call :BAIXA "%%U" "bweb_2t_%%U_311020221535.zip"
)

echo.
echo Processando ZIPs ja baixados...
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
if exist "%WINPY%\baixar_boletins_urna_2022.py" (
  "%PYTHON%" "%WINPY%\baixar_boletins_urna_2022.py" --somente-processar --massa-dados "%WINPY%\dados" --pasta-saida "%WINPY%\saida"
) else (
  echo [AVISO] Falta baixar_boletins_urna_2022.py na pasta winpython.
  echo         Depois de baixar os ZIPs: python baixar_boletins_urna_2022.py --somente-processar
)
echo ---- fim ERRORLEVEL=%ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"
echo.
if exist "%SAIDA%\urnas_2t_presidente.csv" (
  echo OK: %SAIDA%\urnas_2t_presidente.csv
) else (
  echo Veja %LOG%
  echo Se o curl tambem falhou, baixe no Edge e salve em:
  echo   %RAW%
)
echo.
pause
endlocal
exit /b 0

:BAIXA
set "UF=%~1"
set "NOME=%~2"
set "DEST=%RAW%\%NOME%"
if exist "%DEST%" (
  for %%A in ("%DEST%") do if %%~zA GTR 1000 (
    echo [ok] %UF% ja existe
    echo [ok] %UF% ja existe >> "%LOG%"
    goto :EOF
  )
)
echo [TSE] %UF%
echo [TSE] %UF% >> "%LOG%"
"%CURL%" -L --fail --retry 2 --connect-timeout 45 --max-time 600 --ssl-no-revoke -A "%UA%" -o "%DEST%.part" "%TSE%/%NOME%"
if not errorlevel 1 (
  move /Y "%DEST%.part" "%DEST%" >nul
  echo [ok] %UF% via TSE
  echo [ok] %UF% via TSE >> "%LOG%"
  goto :EOF
)
echo [IA] %UF%
echo [IA] %UF% >> "%LOG%"
"%CURL%" -L --fail --retry 2 --connect-timeout 45 --max-time 600 --ssl-no-revoke -k -A "%UA%" -o "%DEST%.part" "%IA%/%NOME%"
if not errorlevel 1 (
  move /Y "%DEST%.part" "%DEST%" >nul
  echo [ok] %UF% via Archive.org
  echo [ok] %UF% via Archive.org >> "%LOG%"
  goto :EOF
)
echo [ERRO] %UF%
echo [ERRO] %UF% >> "%LOG%"
if exist "%DEST%.part" del "%DEST%.part"
goto :EOF
