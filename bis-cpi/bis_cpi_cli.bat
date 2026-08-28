@echo off
setlocal
cd /d "%~dp0"
set PYTHONPATH=%~dp0src;%PYTHONPATH%
where python >nul 2>&1 && (
  python -m bis_cpi.cli %*
) || (
  python3 -m bis_cpi.cli %*
)
