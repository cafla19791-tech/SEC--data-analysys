@echo off
REM Baixa scripts do impacto DIRETAS para sec_scripts e executa.
REM Uso (cmd, na pasta winpython):
REM   curl.exe -L -o baixar_impacto_diretas.bat "URL..."
REM   baixar_impacto_diretas.bat
setlocal EnableExtensions
cd /d "%~dp0"

set "BASE=https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/impacto-diretas-f342"

mkdir sec_scripts 2>nul

echo Baixando scripts DIRETAS...
curl.exe -L -o sec_scripts\__init__.py "%BASE%/scripts/__init__.py"
curl.exe -L -o sec_scripts\gerar_fluxos.py "%BASE%/scripts/gerar_fluxos.py"
curl.exe -L -o sec_scripts\contagil_fluxos_seguro.py "%BASE%/scripts/contagil_fluxos_seguro.py"
curl.exe -L -o sec_scripts\impacto_fiscal_por_ano.py "%BASE%/scripts/impacto_fiscal_por_ano.py"
curl.exe -L -o sec_scripts\agregar_impacto_fluxos.py "%BASE%/scripts/agregar_impacto_fluxos.py"
curl.exe -L -o sec_scripts\apresentacao_impacto_bndes.py "%BASE%/scripts/apresentacao_impacto_bndes.py"
curl.exe -L -o sec_scripts\impacto_operacoes_diretas.py "%BASE%/scripts/impacto_operacoes_diretas.py"
curl.exe -L -o impacto_operacoes_diretas.bat "%BASE%/impacto_operacoes_diretas.bat"

if not exist sec_scripts\impacto_operacoes_diretas.py (
  echo [ERRO] download falhou
  exit /b 1
)

findstr /C:"impacto-operacoes-diretas-20260816a" sec_scripts\impacto_operacoes_diretas.py >nul
if errorlevel 1 (
  echo [ERRO] MARKER ausente - download antigo/cache
  exit /b 1
)

echo OK. Rodando pipeline...
python.exe -m pip install "pandas>=2.0" "openpyxl>=3.1" "numpy>=1.24" "requests>=2.28"
python.exe sec_scripts\impacto_operacoes_diretas.py %*
endlocal
exit /b %ERRORLEVEL%
