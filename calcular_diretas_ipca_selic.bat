@echo off
REM Calcula colunas K-N em OPERACOES DIRETAS - 2002 a 2018.xlsx
REM K=Valor IPCA | L=SELIC cap. mensal | M=Juros contrato cap. | N=L-M
REM
REM Requer scripts\ na pasta winpython (deploy_contagil_winpython.bat).

setlocal
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
cd /d "%WINPY%" 2>nul
if errorlevel 1 (
  echo [ERRO] Pasta ContAgil nao encontrada: "%WINPY%"
  exit /b 1
)

set "EXCEL=%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx"
if not exist "%EXCEL%" set "EXCEL=%WINPY%\OPERACOES DIRETAS.xlsx"
if not exist "%EXCEL%" (
  echo [ERRO] Excel nao encontrado:
  echo   %WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx
  exit /b 1
)

set "SELIC=%WINPY%\selic_mensal.xlsx"
set "PY=python scripts\calcular_diretas_ipca_selic.py"
if not exist "scripts\calcular_diretas_ipca_selic.py" set "PY=python calcular_diretas_ipca_selic.py"
if not exist "scripts\calcular_diretas_ipca_selic.py" if not exist "calcular_diretas_ipca_selic.py" (
  echo [ERRO] Script nao encontrado. Rode deploy_contagil_winpython.bat no repo.
  exit /b 1
)

if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\OPERACOES DIRETAS - 2002 a 2018_calculado.xlsx"

echo Arquivo entrada: %EXCEL%
echo Arquivo saida  : %SAIDA%
echo.

if exist "%SELIC%" (
  %PY% --excel "%EXCEL%" --saida "%SAIDA%" --selic "%SELIC%" --data-ref 2026-06-30
) else (
  echo [AVISO] selic_mensal.xlsx ausente — baixando SELIC do Bacen
  %PY% --excel "%EXCEL%" --saida "%SAIDA%" --data-ref 2026-06-30
)

echo.
if exist "%SAIDA%" (
  echo ========================================
  echo SAIDA PRONTA:
  echo %SAIDA%
  echo ========================================
  explorer /select,"%SAIDA%"
) else (
  echo [ERRO] Arquivo de saida nao foi gerado.
)
endlocal
exit /b %ERRORLEVEL%
