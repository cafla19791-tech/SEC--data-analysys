@echo off
REM Impacto fiscal OPERACOES DIRETAS (ContAgil WinPython)
REM Execute no cmd, na pasta winpython (NAO use: python este.bat)
setlocal EnableExtensions
cd /d "%~dp0"

set "WINPY=%CD%"
set "SCRIPT=%WINPY%\sec_scripts\impacto_operacoes_diretas.py"
set "PY="

if exist "%WINPY%\python.exe" set "PY=%WINPY%\python.exe"
if not defined PY if exist "%WINPY%\python\python.exe" set "PY=%WINPY%\python\python.exe"

if not defined PY (
  echo [ERRO] python.exe nao encontrado em "%WINPY%"
  exit /b 1
)
if not exist "%SCRIPT%" (
  echo [ERRO] Falta "%SCRIPT%"
  echo Baixe antes com curl o script impacto_operacoes_diretas.py para sec_scripts\
  exit /b 1
)

echo Python: %PY%
echo Script: %SCRIPT%
"%PY%" "%SCRIPT%" %*
endlocal
exit /b %ERRORLEVEL%
