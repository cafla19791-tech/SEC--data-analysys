@echo off
REM ContAgil / WinPython - fluxos BNDES indiretos (capitalizacao mensal)
REM Este arquivo e .BAT - NAO cole o conteudo dentro de scripts\contagil_fluxos.py
REM Execute com duplo-clique ou: contagil_fluxos_bndes.bat
REM Ou rode o Python diretamente (veja abaixo).

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
  findstr /B /C:"#!/usr/bin/env python" "scripts\contagil_fluxos.py" >nul 2>&1
  if errorlevel 1 (
    echo [ERRO] scripts\contagil_fluxos.py nao parece um arquivo Python.
    echo        Ele pode ter sido sobrescrito com o conteudo deste .bat.
    echo        Restaure o .py do repositorio SEC--data-analysys ^(PR #39^).
    exit /b 1
  )
  python scripts\contagil_fluxos.py --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --arquivo-fatores "%FATORES%"
) else if exist "contagil_fluxos.py" (
  python contagil_fluxos.py --massa-dados "%DADOS%" --pasta-saida "%SAIDA%" --arquivo-fatores "%FATORES%"
) else (
  echo [ERRO] contagil_fluxos.py nao encontrado.
  echo Clone/atualize o repo SEC--data-analysys ou copie a pasta scripts\ para esta pasta.
  exit /b 1
)

endlocal
exit /b %ERRORLEVEL%
