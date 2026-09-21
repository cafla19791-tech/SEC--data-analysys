import openpyxl
from pathlib import Path
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
    csv_detail = out_dir / "DERIVADOS_ANP_2000_2026_por_produto_mensal.csv"
    csv_resumo = out_dir / "DERIVADOS_ANP_2000_2026_resumo_por_derivado.csv"
    
    assert excel_path.exists()
    assert json_path.exists()
    assert csv_path.exists()
    assert csv_detail.exists()
    assert csv_resumo.exists()
    
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    expected_sheets = [
        "Resumo_Por_Derivado",
        "Prod_Por_Derivado_Anual",
        "ImpVol_Por_Derivado_Anual",
        "DispUSD_Por_Derivado_Anual",
        "PrecoImp_Por_Derivado",
        "1_Producao_Total_Mensal",
        "2_Importacao_Vol_Mensal",
        "3_Dispendio_USD_Mensal",
        "Serie_Total_Mensal",
        "Detalhe_Mensal_Por_Produto",
    ]
    for s in expected_sheets:
        assert s in wb.sheetnames, f"Falta a aba {s}"
        
    # Verificar totais acumulados
    totais = summary["totais_nacionais"]
    assert 18_000_000_000 < totais["volume_produzido_brasil_barris"] < 19_000_000_000
    assert 4_000_000_000 < totais["volume_importado_barris"] < 5_000_000_000
    assert 300_000_000_000 < totais["dispendio_importacao_usd"] < 320_000_000_000
    
    # Verificar discriminação por produto
    assert summary["total_derivados_analisados"] == 15
    resumo_prods = {r["Produto"]: r for r in summary["resumo_por_derivado"]}
    assert "ÓLEO DIESEL" in resumo_prods
    assert "GASOLINA A" in resumo_prods
    assert "NAFTA" in resumo_prods
    assert "GLP" in resumo_prods
    
    # Óleo diesel deve ser o maior volume consumido
    diesel = resumo_prods["ÓLEO DIESEL"]
    assert diesel["Volume_Produzido_Brasil_barris"] > 7_000_000_000
    assert diesel["Volume_Importado_barris"] > 1_500_000_000
    assert diesel["Dispendio_Importacao_USD"] > 130_000_000_000
    assert 88.0 < diesel["Preco_Medio_Importacao_USD_bbl"] < 92.0
