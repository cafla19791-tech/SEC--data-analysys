@echo off
REM Discriminativos INDIRETAS: fluxos na planilha do ANO DO CONTRATO
REM (impacto fiscal permanece capitalizado na data_fluxo)
setlocal EnableExtensions
cd /d "%~dp0"
set "WINPY=%CD%"
set "LOG=%WINPY%\discriminativos_indiretas_ano_contrato.log"
set "PYTHON=%WINPY%\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
set "SAIDA=%WINPY%\saida"
set "RUNNER=%WINPY%\discriminativos_indiretas_ano_contrato.py"
if not exist "%RUNNER%" set "RUNNER=%WINPY%\scripts\discriminativos_indiretas_ano_contrato.py"

echo ========================================
echo  DISCRIMINATIVOS INDIRETAS POR ANO CONTRATO
echo ========================================
echo WinPy : %WINPY%
echo Saida : %SAIDA%\discriminativos_ano_contrato
echo.

echo ---- inicio %DATE% %TIME% ---- > "%LOG%"
if not exist "%RUNNER%" (
  echo [ERRO] Runner nao encontrado >> "%LOG%"
  echo [ERRO] Falta discriminativos_indiretas_ano_contrato.py
  goto :SHOW
)
if not exist "%SAIDA%" (
  echo [ERRO] Pasta saida nao existe. Rode contagil_fluxos primeiro. >> "%LOG%"
  echo [ERRO] Pasta saida nao existe. Rode contagil_fluxos primeiro.
  goto :SHOW
)

"%PYTHON%" "%RUNNER%" --pasta "%SAIDA%" --saida "%SAIDA%\discriminativos_ano_contrato" >> "%LOG%" 2>&1
echo ---- fim %ERRORLEVEL% %DATE% %TIME% ---->> "%LOG%"

:SHOW
type "%LOG%"
echo.
if exist "%SAIDA%\discriminativos_ano_contrato\RESUMO_POR_ANO_CONTRATO.xlsx" (
  echo OK: %SAIDA%\discriminativos_ano_contrato\
  explorer /select,"%SAIDA%\discriminativos_ano_contrato\RESUMO_POR_ANO_CONTRATO.xlsx"
) else (
  echo FALHOU. Veja %LOG%
)
pause
endlocal
exit /b 0
