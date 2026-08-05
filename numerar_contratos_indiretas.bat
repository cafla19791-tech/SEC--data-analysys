@echo off
REM Numera contratos BNDES INDIRETAS: 1-2002, 2-2002, ... / 1-2003, ...
REM ContAgil: usa pasta sec_scripts (NAO scripts) para nao colidir com WinPython\Scripts
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\numerar_contratos_indiretas.log"

echo ========================================
echo  NUMERAR CONTRATOS INDIRETAS N-AAAA
echo ========================================

set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\BNDES_INDIRETAS_NUMERADOS.xlsx"
set "DADOS=%WINPY%\dados"
set "RUNNER=%WINPY%\numerar_contratos_indiretas.py"
set "PKG=%WINPY%\sec_scripts"

if not exist "%RUNNER%" (
  echo [ERRO] Falta numerar_contratos_indiretas.py
  echo Rode baixar_numerar_contratos_indiretas.ps1 neste pasta winpython.
  goto :FIM
)
if not exist "%PKG%\numerar_contratos_indiretas.py" (
  echo [ERRO] Falta sec_scripts\numerar_contratos_indiretas.py
  echo Rode baixar_numerar_contratos_indiretas.ps1 neste pasta winpython.
  goto :FIM
)
if not exist "%PKG%\gerar_fluxos.py" (
  echo [ERRO] Falta sec_scripts\gerar_fluxos.py
  echo Rode baixar_numerar_contratos_indiretas.ps1 neste pasta winpython.
  goto :FIM
)

echo WinPy: %WINPY%
echo Dados: %DADOS%
echo Saida: %SAIDA%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not exist "%DADOS%" (
  echo [ERRO] Pasta dados nao encontrada: %DADOS% >> "%LOG%"
  echo [ERRO] Pasta dados nao encontrada: %DADOS%
  echo Coloque os arquivos BNDES INDIRETAS AAAA.xlsx em dados\
  goto :SHOWLOG
)

"%PYTHON%" "%RUNNER%" --pasta-dados "%DADOS%" --saida "%SAIDA%" --ano-min 2002 >> "%LOG%" 2>&1
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
