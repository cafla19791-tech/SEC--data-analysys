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
copy /Y "%REPO%\scripts\calcular_diretas_ipca_selic.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\atualizar_desembolsos_ipca_1995_2001.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\numerar_contratos_indiretas.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\discriminativos_indiretas_ano_contrato.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\fluxos_por_ano_contrato_numerados.py" "%WINPY%\scripts\" >nul
copy /Y "%REPO%\scripts\agregar_impacto_fluxos.py" "%WINPY%\scripts\" >nul
if not exist "%WINPY%\sec_scripts" mkdir "%WINPY%\sec_scripts"
copy /Y "%REPO%\scripts\baixar_boletins_urna_2022.py" "%WINPY%\sec_scripts\" >nul
copy /Y "%REPO%\scripts\baixar_boletins_urna_2022.py" "%WINPY%\scripts\" >nul

copy /Y "%REPO%\contagil_fluxos.py" "%WINPY%\" >nul
copy /Y "%REPO%\contagil_fluxos_seguro.py" "%WINPY%\" >nul
copy /Y "%REPO%\contagil_fluxos_seguro.bat" "%WINPY%\" >nul
copy /Y "%REPO%\baixar_contagil_fluxos_seguro.ps1" "%WINPY%\" >nul
copy /Y "%REPO%\resumo_fluxos_polars.py" "%WINPY%\" >nul
copy /Y "%REPO%\calcular_diretas_ipca_selic.py" "%WINPY%\" >nul
copy /Y "%REPO%\atualizar_desembolsos_ipca_1995_2001.py" "%WINPY%\" >nul
copy /Y "%REPO%\numerar_contratos_indiretas.py" "%WINPY%\" >nul
copy /Y "%REPO%\contagil_fluxos_bndes.bat" "%WINPY%\" >nul
copy /Y "%REPO%\resumo_fluxos_contagil.bat" "%WINPY%\" >nul
copy /Y "%REPO%\calcular_diretas_ipca_selic.bat" "%WINPY%\" >nul
copy /Y "%REPO%\atualizar_desembolsos_1995_2001.bat" "%WINPY%\" >nul
copy /Y "%REPO%\numerar_contratos_indiretas.bat" "%WINPY%\" >nul
copy /Y "%REPO%\discriminativos_indiretas_ano_contrato.py" "%WINPY%\" >nul
copy /Y "%REPO%\discriminativos_indiretas_ano_contrato.bat" "%WINPY%\" >nul
copy /Y "%REPO%\fluxos_por_ano_contrato_numerados.py" "%WINPY%\" >nul
copy /Y "%REPO%\fluxos_por_ano_contrato_numerados.bat" "%WINPY%\" >nul
copy /Y "%REPO%\baixar_boletins_urna_2022.py" "%WINPY%\" >nul
copy /Y "%REPO%\baixar_boletins_urna_2022.bat" "%WINPY%\" >nul
copy /Y "%REPO%\baixar_boletins_urna_2022.ps1" "%WINPY%\" >nul
copy /Y "%REPO%\requirements.txt" "%WINPY%\requirements_sec.txt" >nul

echo OK. Arquivos copiados.
echo.
echo Dependencias ^(na WinPython^):
echo   python -m pip install -r requirements_sec.txt
echo.
echo Fluxos a partir de BNDES_INDIRETAS_NUMERADOS.xlsx ^(uma aba por ano^):
echo   numerar_contratos_indiretas.bat
echo   fluxos_por_ano_contrato_numerados.bat
echo.
echo Relatorio:
echo   resumo_fluxos_contagil.bat
echo.
echo Boletins de urna 2022 ^(28 UFs, serie + modelo^):
echo   baixar_boletins_urna_2022.bat
echo   Saida: %WINPY%\saida\tse2022\urnas_2t_presidente.csv
echo.
endlocal
exit /b 0
