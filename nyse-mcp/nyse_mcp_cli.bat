@echo off
REM Atalho Windows para a CLI nyse-mcp (WinPython ContAgil / sem .venv).
REM Exemplos:
REM   nyse_mcp_cli.bat quote JPM
REM   nyse_mcp_cli.bat history XOM --period 1y
REM   nyse_mcp_cli.bat search "Coca Cola"
REM   nyse_mcp_cli.bat status

setlocal
cd /d "%~dp0"

if "%MARKET_DATA_PROVIDER%"=="" set MARKET_DATA_PROVIDER=yahoo
set PYTHONPATH=%~dp0src;%PYTHONPATH%

REM Prefer WinPython na pasta pai; senao python do PATH
set "PY=%~dp0..\python.exe"
if not exist "%PY%" set "PY=python.exe"

"%PY%" -m nyse_mcp.cli %*
exit /b %ERRORLEVEL%
