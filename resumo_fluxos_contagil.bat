@echo off
REM Relatorio de resumo de fluxos ContAgil (WinPython).
REM Requer scripts\resumo_fluxos_polars.py nesta pasta
REM ^(rode deploy_contagil_winpython.bat a partir do repo antes^).

setlocal
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
cd /d "%WINPY%" 2>nul
if errorlevel 1 (
  echo [ERRO] Pasta ContAgil nao encontrada: "%WINPY%"
  exit /b 1
)

set "SAIDA=%WINPY%\saida"
set "DADOS=%WINPY%\dados"
set "SELIC=%WINPY%\selic_mensal.xlsx"
set "TJLP=%WINPY%\tjlp_mensal.xlsx"

if not exist "scripts\resumo_fluxos_polars.py" if not exist "resumo_fluxos_polars.py" (
  echo [ERRO] scripts\resumo_fluxos_polars.py nao encontrado em:
  echo        %CD%
  echo.
  echo O ContAgil WinPython NAO inclui o repo. Faca um destes:
  echo   1^) No repo clonado: deploy_contagil_winpython.bat
  echo   2^) Ou rode direto do repo:
  echo      cd C:\caminho\SEC--data-analysys
  echo      python scripts\resumo_fluxos_polars.py --pasta "%SAIDA%" ...
  exit /b 1
)

REM Original: primeiro BNDES INDIRETAS em dados\, senao OPERACOES DIRETAS
set "ORIGINAL="
for %%F in ("%DADOS%\BNDES INDIRETAS *.xlsx") do (
  if not defined ORIGINAL set "ORIGINAL=%%~fF"
)
if not defined ORIGINAL if exist "%WINPY%\OPERACOES DIRETAS.xlsx" (
  set "ORIGINAL=%WINPY%\OPERACOES DIRETAS.xlsx"
)
if not defined ORIGINAL (
  echo [ERRO] Nenhum Excel original encontrado em "%DADOS%" nem OPERACOES DIRETAS.xlsx
  exit /b 1
)

if not exist "%SAIDA%" (
  echo [ERRO] Pasta de saida nao encontrada: "%SAIDA%"
  echo        Gere os fluxos antes ^(contagil_fluxos_bndes.bat^).
  exit /b 1
)
if not exist "%SELIC%" (
  echo [ERRO] SELIC mensal nao encontrada: "%SELIC%"
  exit /b 1
)

echo Relatorio ContAgil
echo   pasta     : %SAIDA%
echo   original  : %ORIGINAL%
echo   selic     : %SELIC%
echo.

set "PY=python scripts\resumo_fluxos_polars.py"
if not exist "scripts\resumo_fluxos_polars.py" set "PY=python resumo_fluxos_polars.py"

if exist "%TJLP%" (
  %PY% --pasta "%SAIDA%" --original "%ORIGINAL%" --selic "%SELIC%" --tjlp "%TJLP%" --output-dir "%SAIDA%"
) else (
  echo [AVISO] TJLP mensal ausente: "%TJLP%" — seguindo sem --tjlp
  %PY% --pasta "%SAIDA%" --original "%ORIGINAL%" --selic "%SELIC%" --output-dir "%SAIDA%"
)

echo.
if exist "%SAIDA%\RELATORIO_EXECUTIVO.md" (
  echo OK: %SAIDA%\RELATORIO_EXECUTIVO.md
  echo     %SAIDA%\resumo_fluxos_polars_final.xlsx
)
endlocal
exit /b %ERRORLEVEL%
