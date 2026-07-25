@echo off
REM Baixa os .py corretos do GitHub para a pasta WinPython ContAgil.
REM Execute ESTE arquivo com duplo-clique ou no cmd (NAO use: python este.bat)
REM NAO cole texto de instrucoes (markdown) no cmd.
REM
REM Uso:
REM   cd /d "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
REM   atualizar_scripts_winpython.bat

setlocal EnableExtensions
cd /d "%~dp0"

set "BRANCH=cursor/normalizar-colunas-6f97"
set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/%BRANCH%"

echo ============================================================
echo Atualizando scripts ContAgil a partir do GitHub
echo Pasta: %CD%
echo Branch: %BRANCH%
echo ============================================================

if not exist "scripts" mkdir "scripts"

if exist "scripts\contagil_fluxos.py" (
  findstr /C:"REM ContAgil" /C:"@echo off" "scripts\contagil_fluxos.py" >nul 2>&1
  if not errorlevel 1 (
    echo [AVISO] scripts\contagil_fluxos.py parece misturado com .bat - fazendo backup...
    copy /Y "scripts\contagil_fluxos.py" "scripts\contagil_fluxos.py.bak_errado" >nul
  )
)

echo.
echo Baixando arquivos...

REM 1) Preferir o .ps1 (mais robusto)
if exist "%~dp0atualizar_scripts_winpython.ps1" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar_scripts_winpython.ps1"
  if not errorlevel 1 goto :ok
  echo [AVISO] Falha no .ps1 local - tentando curl...
)

REM 2) Baixar o .ps1 e executar
curl.exe -fsSL -o "%TEMP%\atualizar_scripts_winpython.ps1" "%BASE%/atualizar_scripts_winpython.ps1"
if not errorlevel 1 (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\atualizar_scripts_winpython.ps1"
  if not errorlevel 1 goto :ok
)

REM 3) Fallback: curl direto em cada arquivo (Windows 10+)
echo Tentando curl.exe arquivo a arquivo...
curl.exe -fsSL -o "scripts\__init__.py" "%BASE%/scripts/__init__.py" || goto :fail
curl.exe -fsSL -o "scripts\contagil_fluxos.py" "%BASE%/scripts/contagil_fluxos.py" || goto :fail
curl.exe -fsSL -o "scripts\contagil_fluxos_seguro.py" "%BASE%/scripts/contagil_fluxos_seguro.py" || goto :fail
curl.exe -fsSL -o "scripts\gerar_fluxos.py" "%BASE%/scripts/gerar_fluxos.py" || goto :fail
curl.exe -fsSL -o "contagil_fluxos.py" "%BASE%/contagil_fluxos.py" || goto :fail
curl.exe -fsSL -o "contagil_fluxos_bndes.bat" "%BASE%/contagil_fluxos_bndes.bat" || goto :fail

findstr /B /C:"#!/usr/bin/env python" "scripts\contagil_fluxos.py" >nul 2>&1
if errorlevel 1 goto :fail

:ok
echo.
echo OK: scripts atualizados.
echo.
echo Agora rode NO CMD esta linha (so a linha abaixo, sem texto extra):
echo.
echo python scripts\contagil_fluxos.py --massa-dados "%CD%\dados" --pasta-saida "%CD%\saida" --arquivo-fatores "%CD%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
echo.
endlocal
exit /b 0

:fail
echo.
echo [ERRO] Falha ao baixar. Tente no PowerShell (copie so o bloco):
echo.
echo   cd "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
echo   mkdir scripts -Force
echo   $b="https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/normalizar-colunas-6f97"
echo   Invoke-WebRequest "$b/scripts/contagil_fluxos.py" -OutFile scripts\contagil_fluxos.py
echo   Invoke-WebRequest "$b/scripts/contagil_fluxos_seguro.py" -OutFile scripts\contagil_fluxos_seguro.py
echo   Invoke-WebRequest "$b/scripts/gerar_fluxos.py" -OutFile scripts\gerar_fluxos.py
echo   Invoke-WebRequest "$b/scripts/__init__.py" -OutFile scripts\__init__.py
echo.
endlocal
exit /b 1
