@echo off
REM Atalho Windows para a CLI nyse-mcp (apos rodar setup_e_rodar_cli.ps1 uma vez).
REM Exemplos:
REM   nyse_mcp_cli.bat quote JPM
REM   nyse_mcp_cli.bat history XOM --period 1y
REM   nyse_mcp_cli.bat search "Coca Cola"
REM   nyse_mcp_cli.bat status

setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERRO: .venv nao encontrado.
  echo Rode antes:
  echo   powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_e_rodar_cli.ps1
  exit /b 2
)

if "%MARKET_DATA_PROVIDER%"=="" set MARKET_DATA_PROVIDER=yahoo

".venv\Scripts\python.exe" -m nyse_mcp.cli %*
exit /b %ERRORLEVEL%
