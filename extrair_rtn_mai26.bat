@echo off
REM Extrai serie anual do RTN serie_historica_mai26 do ContAgil WinPython.
REM IPCA: aba 1.1-A (valores de Mai/2026).
setlocal
cd /d "%~dp0"

if not exist "serie_historica_mai26 (2).xlsx" (
  if not exist "serie_historica_mai26.xlsx" (
    echo ERRO: coloque serie_historica_mai26 ^(2^).xlsx nesta pasta:
    echo   %CD%
    exit /b 1
  )
  set "XLSX=serie_historica_mai26.xlsx"
) else (
  set "XLSX=serie_historica_mai26 (2).xlsx"
)

if not exist scripts\extrair_serie_historica_rtn.py (
  echo Baixando script do GitHub...
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/cursor/serie-historica-mai26-d84a/scripts/extrair_serie_historica_rtn.py' -OutFile 'scripts\extrair_serie_historica_rtn.py'"
)

if not exist saida mkdir saida

echo === IPCA Mai/2026 (aba 1.1-A) ===
.\python.exe scripts\extrair_serie_historica_rtn.py "%XLSX%" --constantes-ipca --from 2001 --to 2025 --out saida\rtn_anual_ipca_mai26.csv
if errorlevel 1 exit /b 1

echo === Valores correntes (aba 1.1) ===
.\python.exe scripts\extrair_serie_historica_rtn.py "%XLSX%" --from 2001 --to 2025 --out saida\rtn_anual_corrente.csv
if errorlevel 1 exit /b 1

echo.
echo Pronto:
echo   saida\rtn_anual_ipca_mai26.csv
echo   saida\rtn_anual_corrente.csv
endlocal
