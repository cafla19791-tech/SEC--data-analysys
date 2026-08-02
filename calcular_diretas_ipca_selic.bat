@echo off
REM Calcula colunas K-N em OPERACOES DIRETAS - 2002 a 2018.xlsx
REM K=Valor IPCA | L=SELIC cap. mensal | M=Juros contrato cap. | N=L-M
REM
REM Requer scripts\ na pasta winpython (copie do ZIP do repo).

setlocal EnableExtensions
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
set "LOG=%WINPY%\calcular_diretas_ipca_selic.log"

echo ========================================
echo  OPERACOES DIRETAS — colunas K L M N
echo ========================================
echo.

cd /d "%WINPY%" 2>nul
if errorlevel 1 (
  echo [ERRO] Pasta ContAgil nao encontrada:
  echo   %WINPY%
  goto :FIM
)

echo Pasta: %CD%
echo Log  : %LOG%
echo.

set "EXCEL="
if exist "%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx"
if not defined EXCEL if exist "%WINPY%\OPERACOES DIRETAS.xlsx" set "EXCEL=%WINPY%\OPERACOES DIRETAS.xlsx"
if not defined EXCEL (
  echo [ERRO] Excel nao encontrado. Procurados:
  echo   %WINPY%\OPERACOES DIRETAS - 2002 a 2018.xlsx
  echo   %WINPY%\OPERACOES DIRETAS.xlsx
  echo.
  echo Arquivos .xlsx nesta pasta:
  dir /b "%WINPY%\*.xlsx" 2>nul
  goto :FIM
)

if not exist "%WINPY%\saida" mkdir "%WINPY%\saida"
set "SAIDA=%WINPY%\saida\OPERACOES DIRETAS - 2002 a 2018_calculado.xlsx"
set "SELIC=%WINPY%\selic_mensal.xlsx"

set "PYSCRIPT="
if exist "%WINPY%\scripts\calcular_diretas_ipca_selic.py" set "PYSCRIPT=%WINPY%\scripts\calcular_diretas_ipca_selic.py"
if not defined PYSCRIPT if exist "%WINPY%\calcular_diretas_ipca_selic.py" set "PYSCRIPT=%WINPY%\calcular_diretas_ipca_selic.py"
if not defined PYSCRIPT (
  echo [ERRO] Script Python nao encontrado.
  echo Copie do ZIP do repo para esta pasta:
  echo   - pasta scripts\
  echo   - calcular_diretas_ipca_selic.bat
  echo.
  echo Pasta scripts existe?
  if exist "%WINPY%\scripts" (dir /b "%WINPY%\scripts\*.py") else (echo   NAO)
  goto :FIM
)

where python >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Comando "python" nao encontrado no PATH.
  echo Abra o terminal da WinPython ContAgil e rode de la, ou use o python.exe da pasta.
  if exist "%WINPY%\python.exe" (
    set "PYTHON=%WINPY%\python.exe"
    echo Usando: %PYTHON%
  ) else (
    goto :FIM
  )
) else (
  set "PYTHON=python"
)

echo Entrada : %EXCEL%
echo Saida   : %SAIDA%
echo Script  : %PYSCRIPT%
echo Python  : %PYTHON%
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
echo Entrada: %EXCEL%>> "%LOG%"
echo Saida  : %SAIDA%>> "%LOG%"
echo Script : %PYSCRIPT%>> "%LOG%"

if exist "%SELIC%" (
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --selic "%SELIC%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
) else (
  echo [AVISO] selic_mensal.xlsx ausente — Bacen SGS
  "%PYTHON%" "%PYSCRIPT%" --excel "%EXCEL%" --saida "%SAIDA%" --data-ref 2026-06-30 >> "%LOG%" 2>&1
)
set "ERR=%ERRORLEVEL%"

echo.>> "%LOG%"
echo ---- fim codigo=%ERR% %DATE% %TIME% ---->> "%LOG%"

echo.
type "%LOG%"
echo.

if exist "%SAIDA%" (
  echo ========================================
  echo OK — SAIDA PRONTA:
  echo %SAIDA%
  echo ========================================
  explorer /select,"%SAIDA%"
) else (
  echo ========================================
  echo FALHOU — arquivo de saida nao gerado.
  echo Veja o log:
  echo %LOG%
  echo ========================================
)

:FIM
echo.
echo ^(esta janela fica aberta para voce ler a mensagem^)
pause
endlocal
exit /b 0
