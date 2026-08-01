@echo off
REM Copia os scripts do repo SEC--data-analysys para a pasta ContAgil WinPython.
REM Execute este .bat a partir da pasta do repositório clonado.
REM
REM   cd C:\caminho\SEC--data-analysys
REM   deploy_contagil_winpython.bat

setlocal
set "WINPY=C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
set "REPO=%~dp0"
set "REPO=%REPO:~0,-1%"

if not exist "%REPO%\scripts\resumo_fluxos_polars.py" (
  echo [ERRO] Rode este .bat a partir do repositorio SEC--data-analysys
  echo        ^(precisa existir scripts\resumo_fluxos_polars.py^)
  echo Pasta atual do .bat: "%REPO%"
  exit /b 1
)

if not exist "%WINPY%" (
  echo [ERRO] Pasta ContAgil nao encontrada: "%WINPY%"
  exit /b 1
)

echo Copiando scripts ContAgil para:
echo   %WINPY%
echo.

if not exist "%WINPY%\scripts" mkdir "%WINPY%\scripts"

copy /Y "%REPO%\scripts\__init__.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\gerar_fluxos.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\contagil_fluxos.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\contagil_fluxos_seguro.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\resumo_fluxos.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\resumo_fluxos_avancado.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\resumo_fluxos_polars.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\resumo_por_agente.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\impacto_fiscal_por_ano.py" "%WINPY%\scripts\" >nul

copy /Y "%REPO%\contagil_fluxos.py" "%WINPY%\" >nul
copy /Y "%REPO%\contagil_fluxos_seguro.py" "%WINPY%\" >nul
copy /Y "%REPO%\resumo_fluxos_polars.py" "%WINPY%\" >nul
copy /Y "%REPO%\contagil_fluxos_bndes.bat" "%WINPY%\" >nul
copy /Y "%REPO%\resumo_fluxos_contagil.bat" "%WINPY%\" >nul
copy /Y "%REPO%\requirements.txt" "%WINPY%\requirements_sec.txt" >nul

echo OK. Arquivos copiados.
echo.
echo Dependencias ^(na WinPython^):
echo   python -m pip install -r requirements_sec.txt
echo.
echo Relatorio:
echo   resumo_fluxos_contagil.bat
echo.
endlocal
exit /b 0
