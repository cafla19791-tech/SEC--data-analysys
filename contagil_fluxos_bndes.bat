@echo off
REM ContAgil / WinPython — fluxos BNDES (indiretas e/ou OPERACOES DIRETAS)
REM Execute a partir da pasta do repositório (onde estão scripts\ e contagil_fluxos.py)
REM ou copie este .bat + scripts\ + contagil_fluxos.py para a pasta winpython.
REM
REM Uso:
REM   contagil_fluxos_bndes.bat              → massa dados\ (INDIRETAS)
REM   contagil_fluxos_bndes.bat diretas      → OPERACOES DIRETAS.xlsx

setlocal
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
set "DADOS=%WINPY%\dados"
set "SAIDA=%WINPY%\saida"
set "FATORES=%WINPY%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
set "DIREITAS=%WINPY%\OPERACOES DIRETAS.xlsx"
set "MODO=%~1"
set "PYSCRIPT="

if exist "scripts\contagil_fluxos.py" set "PYSCRIPT=scripts\contagil_fluxos.py"
if not defined PYSCRIPT if exist "contagil_fluxos.py" set "PYSCRIPT=contagil_fluxos.py"
if not defined PYSCRIPT (
  echo [ERRO] contagil_fluxos.py nao encontrado.
  echo Clone/atualize o repo SEC--data-analysys ou copie scripts\ para esta pasta.
  exit /b 1
)

if not exist "%FATORES%" (
  echo [ERRO] Arquivo de fatores nao encontrado: "%FATORES%"
  exit /b 1
)

if /I "%MODO%"=="diretas" goto :DIREITAS
if /I "%MODO%"=="direta" goto :DIREITAS
if /I "%MODO%"=="--diretas" goto :DIREITAS

if not exist "%DADOS%" (
  echo [ERRO] Massa de dados nao encontrada: "%DADOS%"
  exit /b 1
)
python "%PYSCRIPT%" --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%"
exit /b %ERRORLEVEL%

:DIREITAS
if not exist "%DIREITAS%" (
  echo [ERRO] Arquivo nao encontrado: "%DIREITAS%"
  exit /b 1
)
python "%PYSCRIPT%" --excel "%DIREITAS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%"
exit /b %ERRORLEVEL%
