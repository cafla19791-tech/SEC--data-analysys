@echo off
REM Baixa os .py corretos do GitHub para a pasta WinPython ContAgil.
REM Execute ESTE arquivo com duplo-clique ou no cmd (NAO use: python este.bat)
REM
REM Uso:
REM   cd /d "C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
REM   atualizar_scripts_winpython.bat

setlocal EnableExtensions
set "WINPY=%~dp0"
if "%WINPY:~-1%"=="\" set "WINPY=%WINPY:~0,-1%"
cd /d "%WINPY%"

set "BRANCH=cursor/normalizar-colunas-6f97"
set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/%BRANCH%"

echo ============================================================
echo Atualizando scripts ContAgil a partir do GitHub
echo Pasta: %CD%
echo Branch: %BRANCH%
echo ============================================================

if not exist "scripts" mkdir "scripts"

REM Backup do arquivo corrompido (se ainda tiver REM / @echo)
if exist "scripts\contagil_fluxos.py" (
  findstr /C:"REM ContAgil" /C:"@echo off" "scripts\contagil_fluxos.py" >nul 2>&1
  if not errorlevel 1 (
    echo [AVISO] scripts\contagil_fluxos.py parece misturado com .bat - fazendo backup...
    copy /Y "scripts\contagil_fluxos.py" "scripts\contagil_fluxos.py.bak_errado" >nul
  )
)

echo.
echo Baixando arquivos Python...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$base='%BASE%';" ^
  "$files=@(" ^
  "  @{url='$base/scripts/__init__.py'; out='scripts\__init__.py'}," ^
  "  @{url='$base/scripts/contagil_fluxos.py'; out='scripts\contagil_fluxos.py'}," ^
  "  @{url='$base/scripts/contagil_fluxos_seguro.py'; out='scripts\contagil_fluxos_seguro.py'}," ^
  "  @{url='$base/scripts/gerar_fluxos.py'; out='scripts\gerar_fluxos.py'}," ^
  "  @{url='$base/contagil_fluxos.py'; out='contagil_fluxos.py'}," ^
  "  @{url='$base/contagil_fluxos_bndes.bat'; out='contagil_fluxos_bndes.bat'}" ^
  ");" ^
  "foreach ($f in $files) {" ^
  "  Write-Host ('  -> ' + $f.out);" ^
  "  Invoke-WebRequest -Uri $f.url -OutFile $f.out -UseBasicParsing;" ^
  "}"

if errorlevel 1 (
  echo [ERRO] Falha ao baixar arquivos. Verifique internet / proxy.
  exit /b 1
)

echo.
echo Validando scripts\contagil_fluxos.py ...
findstr /B /C:"#!/usr/bin/env python" "scripts\contagil_fluxos.py" >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Download invalido: o .py nao comeca com shebang Python.
  exit /b 1
)
findstr /C:"REM ContAgil" "scripts\contagil_fluxos.py" >nul 2>&1
if not errorlevel 1 (
  echo [ERRO] O .py ainda contem linhas REM. Abortando.
  exit /b 1
)

echo OK: scripts\contagil_fluxos.py e Python valido.
echo.
echo Proximo passo - rode em UMA linha:
echo.
echo python scripts\contagil_fluxos.py --massa-dados "%CD%\dados" --pasta-saida "%CD%\saida" --arquivo-fatores "%CD%\fator_acumulado_SELIC_TJLP_TLP.xlsx"
echo.
echo Ou execute: contagil_fluxos_bndes.bat
echo ============================================================
endlocal
exit /b 0
