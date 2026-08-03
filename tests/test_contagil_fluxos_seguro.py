"""Testes da versão segura ContAgil (mapeamento flexível + fatores mensais)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from scripts.contagil_fluxos_seguro import (
    carregar_fatores_mensais,
    main,
    normalizar_colunas,
    processar_arquivo,
)
from scripts.gerar_fluxos import (
    _excel_tem_colunas_contratos,
    _mapear_colunas_contratos,
    load_from_excel,
    normalizar_colunas as normalizar_colunas_gf,
)


def _df_portal_pt() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Data da contratação": ["15/03/2009"],
            "Valor Desembolsado R$ (*)": [100000.0],
            "Juros": [6.0],
            "Prazo - Carência (meses)": [6],
            "Prazo - Amortização (meses)": [48],
            "Instituição Financeira Credenciada": ["BANCO DO BRASIL SA"],
            "Custo financeiro": ["TAXA FIXA"],
        }
    )


def _df_csv_style() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "data_da_contratacao": ["2009-03-15"],
            "valor_desembolsado_reais": ["100000,0"],
            "juros": ["6,0"],
            "prazo_carencia_meses": ["6"],
            "prazo_amortizacao_meses": ["48"],
            "instituicao_financeira_credenciada": ["BANCO DO BRASIL SA"],
            "custo_financeiro": ["TAXA FIXA"],
        }
    )


def _df_variante_bndes() -> pd.DataFrame:
    """Variação comum em exports 'BNDES INDIRETAS' (sem R$ (*), capitalização diferente)."""
    return pd.DataFrame(
        {
            "Data da Contratação": ["15/03/2009"],
            "Valor desembolsado Reais": [100000.0],
            "Juros": [6.0],
            "Prazo de Carência (meses)": [6],
            "Prazo de Amortização (meses)": [48],
            "Instituicao Financeira Credenciada": ["CAIXA ECONOMICA FEDERAL"],
            "Custo Financeiro": ["TJLP"],
        }
    )


def test_normalizar_colunas_definida_e_funciona():
    """Regressão: NameError 'normalizar_colunas' is not defined no WinPython."""
    assert normalizar_colunas is normalizar_colunas_gf
    df = normalizar_colunas(_df_variante_bndes())
    assert "data_contratacao" in df.columns
    assert "valor_desembolsado" in df.columns
    assert len(df) == 1
    assert float(df.iloc[0]["valor_desembolsado"]) == 100000.0


def test_parse_args_arquivo_fatores():
    from scripts.contagil_fluxos_seguro import parse_args

    args = parse_args(
        [
            "--massa-dados",
            "dados",
            "--pasta-saida",
            "saida",
            "--arquivo-fatores",
            "fator_acumulado_SELIC_TJLP_TLP.xlsx",
        ]
    )
    assert args.fatores.name == "fator_acumulado_SELIC_TJLP_TLP.xlsx"


def test_mapear_colunas_portal_pt():
    mapped, rename = _mapear_colunas_contratos(_df_portal_pt())
    assert "data_contratacao" in mapped.columns
    assert "valor_desembolsado" in mapped.columns
    assert _excel_tem_colunas_contratos(_df_portal_pt())


def test_mapear_colunas_csv_style():
    mapped, _ = _mapear_colunas_contratos(_df_csv_style())
    assert set(
        [
            "data_contratacao",
            "valor_desembolsado",
            "juros",
            "prazo_carencia",
            "prazo_amortizacao",
        ]
    ).issubset(mapped.columns)


def test_mapear_colunas_variante_bndes_indiretas():
    assert _excel_tem_colunas_contratos(_df_variante_bndes())
    mapped, rename = _mapear_colunas_contratos(_df_variante_bndes())
    assert mapped["valor_desembolsado"].iloc[0] == 100000.0
    assert "data_contratacao" in rename.values()


def test_mapear_colunas_valor_historico_contagil():
    """Massa ContAgil BNDES INDIRETAS usa 'Valor histórico' em vez de desembolsado."""
    valor_cols = [
        "Valor histórico",
        "Valor  Histórico",
        "Valor Histórico R$ (*)",
        "Valor Histórico R$ ",
        "Valor Histórico em R$",
    ]
    for valor_col in valor_cols:
        df = pd.DataFrame(
            {
                valor_col: [100000.0],
                "Data da contratacao": ["15/03/2009"],
                "Juros": [6.0],
                "Prazo - Carencia (meses)": [6],
                "Prazo - Amortizacao (meses)": [12],
                "Instituição Financeira Credenciada": ["BANCO DO BRASIL SA"],
                "Encargo Financeiro": ["TJLP"],
            }
        )
        assert _excel_tem_colunas_contratos(df), valor_col
        mapped, _rename = _mapear_colunas_contratos(df)
        assert "valor_desembolsado" in mapped.columns, valor_col
        assert "custo_financeiro" in mapped.columns, valor_col
        assert float(mapped["valor_desembolsado"].iloc[0]) == 100000.0, valor_col


def test_mapear_prefere_valor_historico_sobre_ipca():
    """Se houver Valor histórico e valor IPCA, usa o histórico como principal."""
    df = pd.DataFrame(
        {
            "Data da contratação": ["15/03/2009"],
            "Valor histórico": [50000.0],
            "Valor atualizado pelo IPCA até 31/7/2025": [99999.0],
            "Juros": [6.0],
            "Prazo - Carencia (meses)": [6],
            "Prazo - Amortizacao (meses)": [12],
            "Instituição Financeira Credenciada": ["BB"],
            "encargo financeiro": ["TAXA FIXA"],
        }
    )
    mapped, rename = _mapear_colunas_contratos(df)
    assert rename.get("Valor histórico") == "valor_desembolsado"
    assert float(mapped["valor_desembolsado"].iloc[0]) == 50000.0


def test_load_from_excel_header_offset(tmp_path: Path):
    """Planilha com 2 linhas de título antes do header — comum em exports ContAgil."""
    path = tmp_path / "BNDES INDIRETAS 2009.xlsx"
    # Monta planilha com título nas linhas 0-1 e header na linha 2
    rows = [
        ["BNDES", "Operações Indiretas", None, None, None, None, None],
        ["Período", "2009", None, None, None, None, None],
        [
            "data_da_contratacao",
            "valor_desembolsado_reais",
            "juros",
            "prazo_carencia_meses",
            "prazo_amortizacao_meses",
            "instituicao_financeira_credenciada",
            "custo_financeiro",
        ],
        ["2009-03-15", 50000.0, 5.0, 3, 24, "BANCO TESTE SA", "TAXA FIXA"],
    ]
    pd.DataFrame(rows).to_excel(path, index=False, header=False)

    df = load_from_excel(path)
    assert len(df) == 1
    assert df.iloc[0]["valor_desembolsado"] == 50000.0


def test_load_from_excel_variante_nomes(tmp_path: Path):
    path = tmp_path / "BNDES INDIRETAS 2002.xlsx"
    _df_variante_bndes().to_excel(path, index=False)
    df = load_from_excel(path)
    assert len(df) == 1
    assert df.iloc[0]["agente"] == "CAIXA ECONOMICA FEDERAL"


def test_carregar_fatores_mensais(tmp_path: Path):
    path = tmp_path / "fator_acumulado_SELIC_TJLP_TLP.xlsx"
    pd.DataFrame(
        {
            "Data": pd.date_range("2009-01-01", periods=6, freq="MS"),
            "Taxa_Mensal_%": [0.9] * 6,
            "Fator_Acumulado": [1.01, 1.02, 1.03, 1.04, 1.05, 1.06],
        }
    ).to_excel(path, index=False)
    serie = carregar_fatores_mensais(path)
    assert len(serie.fatores) == 6
    assert serie.fator_referencia == 1.06


def _fatores_mensais(path: Path) -> None:
    datas = pd.date_range("2009-01-01", "2026-06-01", freq="MS")
    taxa = 0.009
    fator = (1 + taxa) ** pd.Series(range(1, len(datas) + 1))
    pd.DataFrame(
        {
            "Data": datas,
            "Taxa_Mensal_%": [0.9] * len(datas),
            "Fator_Acumulado": fator.values,
        }
    ).to_excel(path, index=False)


def test_processar_arquivo_seguro(tmp_path: Path):
    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    saida.mkdir()

    contratos = dados / "BNDES INDIRETAS 2003.xlsx"
    _df_variante_bndes().to_excel(contratos, index=False)

    fatores = tmp_path / "fator_acumulado_SELIC_TJLP_TLP.xlsx"
    _fatores_mensais(fatores)

    serie = carregar_fatores_mensais(fatores)
    out = processar_arquivo(contratos, saida, serie)
    assert out is not None
    assert out.exists()
    fluxos = pd.read_excel(out)
    assert len(fluxos) > 0
    assert "impacto_fiscal" in fluxos.columns or "impacto" in fluxos.columns


def test_processar_arquivo_grande_usa_streaming(tmp_path: Path, monkeypatch):
    """Arquivos com muitos contratos devem gravar via gerar_e_gravar_fluxos."""
    import scripts.contagil_fluxos_seguro as seguro

    monkeypatch.setattr(seguro, "LIMITE_MEMORIA_CONTRATOS", 5)
    monkeypatch.setattr(seguro, "LOTE_FLUXOS", 3)

    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    saida.mkdir()

    base = _df_variante_bndes()
    rows = pd.concat([base] * 12, ignore_index=True)
    contratos = dados / "BNDES INDIRETAS 20112016.xlsx"
    rows.to_excel(contratos, index=False)

    fatores = tmp_path / "fator_acumulado_SELIC_TJLP_TLP.xlsx"
    _fatores_mensais(fatores)
    serie = carregar_fatores_mensais(fatores)

    out = processar_arquivo(contratos, saida, serie)
    assert out is not None
    assert out.exists()
    csv_path = out.with_suffix(".csv")
    assert csv_path.exists()
    assert len(pd.read_csv(csv_path)) > 0


def test_main_massa_segura(tmp_path: Path, monkeypatch):
    dados = tmp_path / "dados"
    saida = tmp_path / "saida"
    dados.mkdir()
    saida.mkdir()
    _df_csv_style().to_excel(dados / "BNDES INDIRETAS 2010.xlsx", index=False)

    fatores = tmp_path / "fator_acumulado_SELIC_TJLP_TLP.xlsx"
    datas = pd.date_range("2009-01-01", "2026-06-01", freq="MS")
    fator = (1.009) ** pd.Series(range(1, len(datas) + 1))
    pd.DataFrame(
        {
            "Data": datas,
            "Taxa_Mensal_%": [0.9] * len(datas),
            "Fator_Acumulado": fator.values,
        }
    ).to_excel(fatores, index=False)

    rc = main(
        [
            "--massa-dados",
            str(dados),
            "--pasta-saida",
            str(saida),
            "--fatores",
            str(fatores),
        ]
    )
    assert rc == 0
    assert list(saida.glob("fluxos_*.xlsx"))
