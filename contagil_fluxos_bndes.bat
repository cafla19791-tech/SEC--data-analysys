@echo off
REM ContAgil / WinPython — fluxos BNDES indiretos (capitalização mensal)
REM Execute a partir da pasta do repositório (onde estão scripts\ e contagil_fluxos.py)
REM ou copie este .bat + scripts\ + contagil_fluxos.py para a pasta winpython.

setlocal
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
set "DADOS=%WINPY%\dados"
set "SAIDA=%WINPY%\saida"
set "FATORES=%WINPY%\fator_acumulado_SELIC_TJLP_TLP.xlsx"

if not exist "%DADOS%" (
  echo [ERRO] Massa de dados nao encontrada: "%DADOS%"
  exit /b 1
)
if not exist "%FATORES%" (
  echo [ERRO] Arquivo de fatores nao encontrado: "%FATORES%"
  exit /b 1
)

if exist "scripts\contagil_fluxos.py" (
  python scripts\contagil_fluxos.py --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%"
) else if exist "contagil_fluxos.py" (
  python contagil_fluxos.py --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --fatores "%FATORES%"
) else (
  echo [ERRO] contagil_fluxos.py nao encontrado.
  echo Clone/atualize o repo SEC--data-analysys ou copie scripts\ para esta pasta.
  exit /b 1
)

endlocal
exit /b %ERRORLEVEL%
