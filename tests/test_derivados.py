import subprocess
from pathlib import Path
import openpyxl
import pandas as pd
from derivados import run_pipeline

def test_derivados_pipeline(tmp_path):
    root = Path(__file__).resolve().parents[1]
    prod_file = root / "data" / "raw" / "anp" / "producao-derivados-b.xlsx"
    imp_file = root / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xlsx"
    
    assert prod_file.exists()
    assert imp_file.exists()
    
    out_dir = tmp_path / "saida_teste"
    summary = run_pipeline(
        prod_path=prod_file,
        imp_path=imp_file,
        output_dir=out_dir,
        start_year=2000,
        end_year=2026,
        domestic_cost=25.0,
    )
    
    excel_path = out_dir / "DERIVADOS_ANP_2000_2026.xlsx"
    json_path = out_dir / "DERIVADOS_ANP_2000_2026.json"
    csv_path = out_dir / "DERIVADOS_ANP_2000_2026_mensal.csv"
    
    assert excel_path.exists()
    assert json_path.exists()
    assert csv_path.exists()
    
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    expected_sheets = [
        "1_Producao_Brasil_barris",
        "2_Importacao_Volume_barris",
        "3_Dispendio_Importacao_USD",
        "Resumo_Anual",
        "Serie_Mensal_Completa",
    ]
    for s in expected_sheets:
        assert s in wb.sheetnames
        
    ws_prod = wb["1_Producao_Brasil_barris"]
    # Checar se ano 2000 e 2026 estao presentes
    cols = {ws_prod.cell(1, c).value: c for c in range(2, 35)}
    assert 2000 in cols
    assert 2026 in cols
    
    # 2000 Jan producao ~46.1M barris
    val_jan_2000 = ws_prod.cell(2, cols[2000]).value
    assert 46_000_000 < val_jan_2000 < 47_000_000
    
    # 2026 Jun producao ~65.36M barris
    val_jun_2026 = ws_prod.cell(7, cols[2026]).value
    assert 65_000_000 < val_jun_2026 < 66_000_000
    
    # Importação em 2000 Jan ~8.25M barris
    ws_imp = wb["2_Importacao_Volume_barris"]
    val_imp_jan_2000 = ws_imp.cell(2, cols[2000]).value
    assert 8_200_000 < val_imp_jan_2000 < 8_300_000
    
    # Dispêndio em 2000 Jan US$ ~167.4M
    ws_usd = wb["3_Dispendio_Importacao_USD"]
    val_usd_jan_2000 = ws_usd.cell(2, cols[2000]).value
    assert 167_000_000 < val_usd_jan_2000 < 168_000_000
    
    assert summary["totais_acumulados"]["volume_produzido_brasil_barris"] > 18_000_000_000
    assert summary["totais_acumulados"]["volume_importado_barris"] > 4_000_000_000
    assert summary["totais_acumulados"]["dispendio_importacao_usd"] > 300_000_000_000
