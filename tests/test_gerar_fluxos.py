"""Testes unitários do gerador de fluxos detalhados (carência + impacto)."""

from datetime import datetime

import numpy as np
import pandas as pd

from scripts.gerar_fluxos import (
    SelicSerie,
    calcular_impacto_fiscal_real,
    gerar_fluxos,
    gerar_fluxos_contrato,
    gerar_fluxos_diarios_contrato,
    limpar_valor,
    meses_ate_impacto,
    parse_args,
    parse_datas,
    taxa_contrato_anual,
    taxa_contrato_efetiva,
    taxa_diaria_composta,
    taxa_mensal_composta,
)


def test_limpar_valor_br_e_us():
    s = pd.Series(["1.234,56", "5.0", "10,5", "1000"])
    out = limpar_valor(s)
    assert list(out) == [1234.56, 5.0, 10.5, 1000.0]


def test_parse_datas_iso_e_br():
    s = pd.Series(["2009-03-15", "15/03/2009", "2009-03-15T00:00:00"])
    out = parse_datas(s)
    assert list(out.dt.strftime("%Y-%m-%d")) == [
        "2009-03-15",
        "2009-03-15",
        "2009-03-15",
    ]


def test_meses_ate_impacto():
    assert meses_ate_impacto(datetime(2026, 6, 30)) == 0
    assert meses_ate_impacto(datetime(2025, 6, 30)) == 12
    assert meses_ate_impacto(datetime(2026, 1, 15)) == 5


def test_taxa_mensal_composta():
    m = taxa_mensal_composta(0.145)
    assert abs(m - ((1.145) ** (1 / 12) - 1)) < 1e-12
    assert m != 0.145 / 12  # composta ≠ linear


def test_carencia_nao_consome_amortizacao():
    """Bug original: data=contr+(carencia+p) E em_carencia=p<=carencia no loop 1..n.

    Com carencia=2 e n=3, o saldo final deve zerar e deve haver 2 meses de carência.
    """
    data = pd.Timestamp("2009-01-31")
    fluxos = gerar_fluxos_contrato(
        data_contr=data,
        valor=300.0,
        taxa_juros_aa=0.06,
        carencia=2,
        n=3,
        contrato_id=0,
        instituicao="BANCO TESTE SA",
    )

    assert len(fluxos) == 5  # 2 carência + 3 amort
    assert [f["em_carencia"] for f in fluxos] == [True, True, False, False, False]
    assert [f["amortizacao"] for f in fluxos] == [0.0, 0.0, 100.0, 100.0, 100.0]
    assert fluxos[0]["saldo_fiscal"] == 300.0
    assert fluxos[2]["saldo_fiscal"] == 300.0  # ainda cheio ao sair da carência
    assert fluxos[-1]["saldo_fiscal"] == 100.0
    # Dual balance: contrato cresce na carência; fiscal fica no principal
    assert fluxos[1]["saldo_contrato"] > fluxos[0]["saldo_contrato"]
    assert fluxos[1]["saldo_fiscal"] == fluxos[0]["saldo_fiscal"]
    # ContAgil: base no dia 15
    assert str(fluxos[0]["data_fluxo"]) == "2009-01-15"
    assert str(fluxos[1]["data_fluxo"]) == "2009-02-15"
    assert str(fluxos[2]["data_fluxo"]) == "2009-03-15"
    assert fluxos[0]["Instituição Financeira"] == "BANCO TESTE SA"
    assert "taxa_selic_mensal" in fluxos[0]
    assert fluxos[0]["taxa_contrato_mensal"] is not None
    assert all(f["taxa_contrato_mensal"] is None for f in fluxos[1:])
    assert "spread" in fluxos[0]
    assert "impacto_fiscal" in fluxos[0]


def test_spread_e_taxas_compostas_constantes_no_contrato():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2010-01-15"),
        valor=1200.0,
        taxa_juros_aa=0.06,
        carencia=0,
        n=12,
        contrato_id=1,
        selic_aa=0.145,
    )
    selic_m = taxa_mensal_composta(0.145)
    contrato_m = taxa_mensal_composta(0.06)
    expected_spread = (1.0 + (selic_m - contrato_m)) ** 12

    assert all(abs(f["taxa_selic_mensal"] - selic_m) < 1e-8 for f in fluxos)
    assert abs(fluxos[0]["taxa_contrato_mensal"] - contrato_m) < 1e-8
    assert all(f["taxa_contrato_mensal"] is None for f in fluxos[1:])
    assert all(abs(f["spread"] - expected_spread) < 1e-6 for f in fluxos)
    # subsídio = saldo_fiscal × (selic_m − contrato_m) antes da amortização
    assert abs(fluxos[0]["subsidio"] - round(1200.0 * (selic_m - contrato_m), 2)) < 0.011


def test_sem_carencia_amortiza_tudo():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2010-01-15"),
        valor=1000.0,
        taxa_juros_aa=0.10,
        carencia=0,
        n=4,
        contrato_id=1,
    )
    assert len(fluxos) == 4
    assert all(not f["em_carencia"] for f in fluxos)
    assert abs(sum(f["amortizacao"] for f in fluxos) - 1000.0) < 0.01


def test_impacto_capitaliza_ate_jun_2026():
    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2026-06-15"),
        valor=1200.0,
        taxa_juros_aa=0.0,  # taxa 0 → subsídio = saldo * selic_m
        carencia=0,
        n=1,
        contrato_id=2,
        selic_aa=0.12,
        data_impacto=datetime(2026, 6, 30),
    )
    assert len(fluxos) == 1
    # data_fluxo = 2026-06-15 → 0 meses até jun/2026 → impacto == subsidio
    assert fluxos[0]["impacto_fiscal"] == fluxos[0]["subsidio"]


def test_impacto_via_fator_selic_stp():
    # Fatores sintéticos: dobro em 12 meses
    datas = np.array(
        [
            np.datetime64("2025-06-15"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 2.0], dtype=float)
    serie = SelicSerie(datas, fatores)

    fluxos = gerar_fluxos_contrato(
        data_contr=pd.Timestamp("2025-06-15"),
        valor=1000.0,
        taxa_juros_aa=0.0,
        carencia=0,
        n=1,
        contrato_id=3,
        selic_aa=0.12,
        data_impacto=datetime(2026, 6, 30),
        selic_serie=serie,
    )
    assert len(fluxos) == 1
    # impacto ≈ subsidio × 2/1
    assert abs(fluxos[0]["impacto_fiscal"] - round(fluxos[0]["subsidio"] * 2.0, 2)) < 0.02


def test_calcular_impacto_fiscal_real_contagil():
    datas = np.array(
        [
            np.datetime64("2020-01-15"),
            np.datetime64("2020-01-16"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 1.1, 1.65], dtype=float)
    serie = SelicSerie(datas, fatores)
    # Regra ContAgil: usa dia seguinte à parcela (16/01 → fator 1.1)
    assert calcular_impacto_fiscal_real(100.0, datetime(2020, 1, 15), serie) == 150.0
    assert calcular_impacto_fiscal_real(0.0, datetime(2020, 1, 15), serie) == 0.0
    # data_impacto anterior/igual → sem capitalização
    assert (
        calcular_impacto_fiscal_real(
            100.0, datetime(2026, 6, 30), serie, data_impacto=datetime(2020, 1, 15)
        )
        == 100.0
    )


def test_impacto_usa_dia_seguinte_da_parcela():
    """idx_inicio = nearest(data_parcela + 1 dia), não a própria data da parcela."""
    datas = np.array(
        [
            np.datetime64("2009-02-15"),
            np.datetime64("2009-02-16"),
            np.datetime64("2026-06-30"),
        ],
        dtype="datetime64[ns]",
    )
    fatores = np.array([1.0, 2.0, 4.0], dtype=float)
    serie = SelicSerie(datas, fatores)
    # 15/02 → início em 16/02 (fator 2) → 4/2 = 2×
    assert calcular_impacto_fiscal_real(100.0, datetime(2009, 2, 15), serie) == 200.0
    # Se usasse a própria parcela (fator 1) daria 400 — garante a regra +1 dia
    assert calcular_impacto_fiscal_real(100.0, datetime(2009, 2, 15), serie) != 400.0


def test_fator_from_taxas_diarias():
    datas = np.array(
        [np.datetime64("2024-01-02"), np.datetime64("2024-01-03")],
        dtype="datetime64[ns]",
    )
    # 1% a.d. → fatores 1.01 e 1.01*1.01
    serie = SelicSerie.from_taxas_diarias(datas, np.array([1.0, 1.0]))
    assert abs(serie.fatores[0] - 1.01) < 1e-12
    assert abs(serie.fatores[1] - 1.01 * 1.01) < 1e-12
    assert serie.idx_proximo(datetime(2024, 1, 2, 12)) in (0, 1)


def test_from_excel_le_fator_coluna_e(tmp_path):
    """ContAgil: col A = data, col E (índice 4) = fator acumulado."""
    path = tmp_path / "STP-20260716182715078 (1).xlsx"
    pd.DataFrame(
        {
            "data": ["15/01/2020", "16/01/2020", "30/06/2026"],
            "b": [0.0, 0.0, 0.0],
            "c": [0.0, 0.0, 0.0],
            "d": [0.01, 0.01, 0.02],
            "fator": [1.0, 1.0, 1.8],
        }
    ).to_excel(path, index=False)

    serie = SelicSerie.from_excel(path)
    assert len(serie.fatores) == 3
    assert abs(serie.fatores[0] - 1.0) < 1e-12
    assert abs(serie.fatores[-1] - 1.8) < 1e-12
    # 15/01 + 1 dia → 16/01 (fator 1.0) → 1.8/1.0
    assert calcular_impacto_fiscal_real(100.0, datetime(2020, 1, 15), serie) == 180.0


def test_gerar_fluxos_aceita_dataframe_selic_contagil():
    """Compat ContAgil: gerar_fluxos(df, selic_df) com col E = fator."""
    contratos = pd.DataFrame(
        {
            "data_contratacao": [pd.Timestamp("2009-02-15")],
            "valor_desembolsado": [1200.0],
            "juros": [6.0],  # % a.a. como no CSV BNDES
            "prazo_carencia": [0],
            "prazo_amortizacao": [1],
            "agente": ["Banco Teste"],
            "contrato": [0],
        }
    )
    selic_df = pd.DataFrame(
        {
            "data": ["15/02/2009", "16/02/2009", "30/06/2026"],
            "b": [0.0, 0.0, 0.0],
            "c": [0.0, 0.0, 0.0],
            "d": [0.01, 0.01, 0.02],
            "fator": [1.0, 2.0, 4.0],
        }
    )
    fluxos = gerar_fluxos(contratos, selic_df)
    assert len(fluxos) == 1
    # 15/02 + 1d → fator 2; fim fator 4 → ×2
    assert abs(fluxos.iloc[0]["impacto_fiscal"] - round(fluxos.iloc[0]["subsidio"] * 2.0, 2)) < 0.02


def test_from_dataframe_coluna_e():
    selic_df = pd.DataFrame(
        {
            "data": ["01/01/2020", "30/06/2026"],
            "b": [0, 0],
            "c": [0, 0],
            "d": [0.01, 0.01],
            "fator": [1.0, 1.5],
        }
    )
    serie = SelicSerie.from_dataframe(selic_df)
    assert abs(serie.fatores[-1] - 1.5) < 1e-12


def test_taxa_contrato_anual_tjlp_tlp():
    """Legado anual: TJLP/TLP = 6% + juros; TAXA FIXA = só juros."""
    assert abs(taxa_contrato_anual("TAXA FIXA", 5.5) - 0.055) < 1e-12
    assert abs(taxa_contrato_anual("TJLP", 2.0) - 0.08) < 1e-12
    assert abs(taxa_contrato_anual("TLP + TAXA FIXA", 1.5) - 0.075) < 1e-12


def test_taxa_contrato_efetiva_tjlp_tlp():
    """Lógica corrigida: TJLP/TLP mensal = (1,06)^(1/12)×(1+juros)^(1/12)−1."""
    juros = 2.0 / 100.0
    esperado_tjlp = (1.06) ** (1 / 12) * (1 + juros) ** (1 / 12) - 1
    assert abs(taxa_contrato_efetiva("TAXA FIXA", 5.5) - ((1.055) ** (1 / 12) - 1)) < 1e-12
    assert abs(taxa_contrato_efetiva("TJLP", 2.0) - esperado_tjlp) < 1e-12
    assert abs(
        taxa_contrato_efetiva("TLP + TAXA FIXA", 1.5)
        - ((1.06) ** (1 / 12) * (1.015) ** (1 / 12) - 1)
    ) < 1e-12


def test_gerar_fluxos_aplica_tjlp():
    contratos = pd.DataFrame(
        {
            "data_contratacao": [pd.Timestamp("2010-01-15")],
            "valor_desembolsado": [1200.0],
            "juros": [2.0],  # spread %; fórmula TJLP corrigida
            "prazo_carencia": [0],
            "prazo_amortizacao": [12],
            "agente": ["Banco TJLP"],
            "custo_financeiro": ["TJLP"],
            "contrato": [0],
        }
    )
    fluxos = gerar_fluxos(contratos)
    esperado = taxa_contrato_efetiva("TJLP", 2.0)
    assert abs(fluxos.iloc[0]["taxa_contrato_mensal"] - esperado) < 1e-8


def test_taxa_diaria_composta():
    d = taxa_diaria_composta(0.145)
    assert abs(d - ((1.145) ** (1 / 365) - 1)) < 1e-15
    assert d < taxa_mensal_composta(0.145)


def test_fluxo_diario_flag_no_parser():
    args = parse_args(["--fluxo-diario", "--sem-selic-fatores", "--input", "x.csv"])
    assert args.fluxo_diario is True
    args_off = parse_args(["--sem-selic-fatores", "--input", "x.csv"])
    assert args_off.fluxo_diario is False


def test_gerar_fluxos_diarios_contrato_dia_a_dia():
    """Entre parcelas ContAgil (dia 15): uma linha por dia; amort só no dia da parcela."""
    diarios = gerar_fluxos_diarios_contrato(
        data_contr=pd.Timestamp("2009-01-31"),
        valor=300.0,
        taxa_juros_aa=0.06,
        carencia=1,
        n=2,
        contrato_id=7,
        instituicao="BANCO DIARIO",
    )
    # 3 meses (1 carência + 2 amort) ≈ ~90 dias
    assert len(diarios) >= 89
    assert len(diarios) <= 93
    assert diarios[0]["data_fluxo"].isoformat() == "2009-01-15"
    assert diarios[0]["dia_parcela"] is True
    assert diarios[0]["amortizacao"] == 0.0  # carência
    assert diarios[1]["dia_parcela"] is False
    assert diarios[1]["amortizacao"] == 0.0
    # 2ª parcela (mês 2) — dia 15/02 — sai da carência
    dia_parcela2 = [d for d in diarios if d["mes"] == 2 and d["dia_parcela"]]
    assert len(dia_parcela2) == 1
    assert dia_parcela2[0]["amortizacao"] == 150.0
    assert all("taxa_selic_diaria" in d for d in diarios)
    assert all(d["Instituição Financeira"] == "BANCO DIARIO" for d in diarios)


def test_gerar_fluxos_escreve_excel_diario(tmp_path):
    contratos = pd.DataFrame(
        {
            "data_contratacao": [pd.Timestamp("2010-01-15")],
            "valor_desembolsado": [1200.0],
            "juros": [6.0],
            "prazo_carencia": [0],
            "prazo_amortizacao": [2],
            "agente": ["Banco X"],
            "contrato": [0],
        }
    )
    out = tmp_path / "fluxos_diarios_detalhados.xlsx"
    mensal = gerar_fluxos(contratos, fluxo_diario=True, saida_diario=out)
    assert len(mensal) == 2
    assert out.exists()
    diarios = pd.read_excel(out)
    assert len(diarios) >= 59  # ~2 meses
    assert "taxa_selic_diaria" in diarios.columns
    assert "subsidio" in diarios.columns


def test_taxa_contrato_efetiva_aceita_row_contagil():
    """Rascunho ContAgil: taxa_contrato_efetiva(row) com colunas em português."""
    row = pd.Series({"Custo financeiro": "TJLP", "Juros": "2,0"})
    assert abs(taxa_contrato_efetiva(row) - taxa_contrato_efetiva("TJLP", 2.0)) < 1e-12
    row_fixa = {"Custo financeiro": "TAXA FIXA", "Juros": "5.5%"}
    assert abs(taxa_contrato_efetiva(row_fixa) - ((1.055) ** (1 / 12) - 1)) < 1e-12


def test_gerar_fluxos_df_df_contagil_paste():
    """ContAgil: gerar_fluxos(df, df) usa df_original para Instituição Financeira."""
    bruto = pd.DataFrame(
        {
            "Data da contratação": ["15/01/2010"],
            "Valor Desembolsado R$ (*)": [1200.0],
            "Juros": ["6,0"],
            "Prazo - Carência (meses)": [1],
            "Prazo - Amortização (meses)": [2],
            "Instituição Financeira Credenciada": ["BANCO PASTE SA"],
            "Custo financeiro": ["TAXA FIXA"],
        }
    )
    fluxos = gerar_fluxos(bruto, bruto)
    # carência(1) + amort(2) — corrige bug p=1..n do rascunho
    assert len(fluxos) == 3
    assert list(fluxos["em_carencia"]) == [True, False, False]
    assert fluxos.iloc[0]["Instituição Financeira"] == "BANCO PASTE SA"
    assert "saldo_fiscal" in fluxos.columns
    assert "saldo_contrato" in fluxos.columns
    assert "impacto_fiscal" in fluxos.columns
    assert fluxos.iloc[0]["taxa_contrato_mensal"] is not None
    assert pd.isna(fluxos.iloc[1]["taxa_contrato_mensal"]) or fluxos.iloc[1][
        "taxa_contrato_mensal"
    ] is None
    # dual balance: contrato cresce na carência
    assert fluxos.iloc[1]["saldo_contrato"] > fluxos.iloc[0]["saldo_contrato"]
    assert fluxos.iloc[1]["saldo_fiscal"] == fluxos.iloc[0]["saldo_fiscal"]


def test_gerar_fluxos_df_df_nao_trata_ops_como_selic():
    """Regressão: gerar_fluxos(df, df) não pode interpretar ops como fatores SELIC."""
    bruto = pd.DataFrame(
        {
            "data_contratacao": [pd.Timestamp("2010-01-15")],
            "valor_desembolsado": [1000.0],
            "juros": [6.0],
            "prazo_carencia": [0],
            "prazo_amortizacao": [2],
            "agente": ["BANCO X"],
            "custo_financeiro": ["TAXA FIXA"],
            "contrato": [0],
        }
    )
    fluxos = gerar_fluxos(bruto, bruto)
    assert len(fluxos) == 2
    # SELIC constante 14,5% composta (não fatores inventados da col E de ops)
    esperado = taxa_mensal_composta(0.145)
    assert abs(fluxos.iloc[0]["taxa_selic_mensal"] - esperado) < 1e-8
