@echo off
REM Gera fluxos por aba de ano a partir de saida\BNDES_INDIRETAS_NUMERADOS.xlsx
REM Cada aba YYYY de contratos → aba/arquivo YYYY de fluxos (ano do contrato)
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\fluxos_por_ano_contrato_numerados.log"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "SAIDA=%WINPY%\saida"
set "NUMERADOS=%SAIDA%\BNDES_INDIRETAS_NUMERADOS.xlsx"
set "FATORES=%WINPY%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
set "RUNNER=%WINPY%\fluxos_por_ano_contrato_numerados.py"
if not exist "%RUNNER%" set "RUNNER=%WINPY%\scripts\fluxos_por_ano_contrato_numerados.py"

echo ========================================
echo  FLUXOS POR ANO DO CONTRATO (NUMERADOS)
echo ========================================
echo WinPy    : %WINPY%
echo Numerados: %NUMERADOS%
echo Saida    : %SAIDA%\FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not exist "%NUMERADOS%" (
  echo [ERRO] Falta %NUMERADOS% >> "%LOG%"
  echo [ERRO] Falta %NUMERADOS%
  echo Rode antes: numerar_contratos_indiretas.bat
  goto :SHOW
)
if not exist "%RUNNER%" (
  echo [ERRO] Runner nao encontrado: %RUNNER% >> "%LOG%"
  echo [ERRO] Atualize o deploy ContAgil ^(fluxos_por_ano_contrato_numerados.py^)
  goto :SHOW
)

if exist "%FATORES%" (
  "%PYTHON%" "%RUNNER%" --numerados "%NUMERADOS%" --saida "%SAIDA%" --fatores "%FATORES%" >> "%LOG%" 2>&1
) else (
  "%PYTHON%" "%RUNNER%" --numerados "%NUMERADOS%" --saida "%SAIDA%" >> "%LOG%" 2>&1
)
echo ---- fim %ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"

:SHOW
type "%LOG%"
echo.
if exist "%SAIDA%\FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx" (
  echo OK: %SAIDA%\FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx
  echo CSVs: %SAIDA%\fluxos_por_ano_contrato\
  explorer /select,"%SAIDA%\FLUXOS_BNDES_INDIRETAS_POR_ANO_CONTRATO.xlsx"
) else (
  echo FALHOU. Veja %LOG%
)
pause
endlocal
exit /b 0
