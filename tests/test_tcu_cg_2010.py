"""Testes das tabelas TCU Contas do Governo 2010 e da atualização IPCA."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from scripts.build_tcu_cg_2010 import (
    build,
    df_beneficios,
    df_indicadores,
    df_pac_deson,
    df_pac_eixo,
    df_resumo_ipca,
    fator_dez2010_ref,
)
from scripts.analisar_base_monetaria_tcu import df_detalhe_2009_2010, df_fatores
from scripts.tcu_cg_2010_dados import (
    FONTE_URL,
    autorizacoes_legais,
    beneficios_fin_cred,
    creditos_dlsp,
    fatores_base_monetaria,
    indicadores,
    pac_desoneracoes,
    pac_subsidios_eixo,
    renuncia_previdenciaria,
    renuncia_regional,
    renuncia_tributaria,
)


def _ipca_constante(taxa: float = 0.5) -> pd.DataFrame:
    mes = pd.date_range("2003-01-01", "2026-06-01", freq="MS")
    fator = (1.0 + taxa / 100.0) ** pd.Series(range(1, len(mes) + 1))
    return pd.DataFrame({"mes": mes, "valor": taxa, "fator": fator.to_numpy()})


def test_totais_oficiais():
    trib = next(r for r in renuncia_tributaria() if r["tributo"] == "Total")
    assert abs(trib["y2010"] - 105_843.31) < 1e-6
    assert abs(trib["y2006"] - 65_397.52) < 1e-6

    ir = next(r for r in renuncia_tributaria() if r["tributo"] == "IR — total")
    partes = [
        r
        for r in renuncia_tributaria()
        if r["tributo"] in {"IR — Pessoa Física", "IR — Pessoa Jurídica", "IR — Retido na Fonte"}
    ]
    assert abs(sum(r["y2010"] for r in partes) - ir["y2010"]) < 0.02

    prev = next(r for r in renuncia_previdenciaria() if r["item"] == "Total")
    assert abs(prev["y2010"] - 19_246.0) < 1e-6

    fin = next(r for r in beneficios_fin_cred() if r["item"].startswith("Total"))
    assert abs(fin["y2010"] - 18_877.65) < 1e-6
    assert abs(fin["y2009"] - 16_901.39) < 1e-6

    reg = next(r for r in renuncia_regional() if r["regiao"] == "Total")
    assert abs(reg["total"] - 143.97) < 1e-6
    assert abs(reg["tributarios"] + reg["trib_prev"] + reg["fin_cred"] - 143.97) < 0.02

    pac = next(r for r in pac_desoneracoes() if r["medida"] == "Total")
    partes_pac = [r for r in pac_desoneracoes() if r["medida"] != "Total"]
    assert abs(sum(r["projecao_2010"] for r in partes_pac) - pac["projecao_2010"]) < 0.01
    assert abs(pac["projecao_2010"] - 23_318.0) < 1e-6

    eixo = next(r for r in pac_subsidios_eixo() if r["eixo"] == "Total")
    partes_eixo = [r for r in pac_subsidios_eixo() if r["eixo"] != "Total"]
    assert abs(sum(r["desembolsos"] for r in partes_eixo) - eixo["desembolsos"]) < 0.02
    assert abs(sum(r["contratacoes"] for r in partes_eixo) - eixo["contratacoes"]) < 0.02


def test_creditos_bndes_e_autorizacoes():
    bndes = next(r for r in creditos_dlsp() if r["item"] == "Créditos junto ao BNDES")
    assert bndes["v2009"] == 129_237
    assert bndes["v2010"] == 236_723
    assert abs(bndes["pib2010"] - 6.47) < 1e-9

    ifo = next(
        r for r in creditos_dlsp() if r["item"].startswith("Créditos concedidos a inst")
    )
    assert ifo["v2010"] - ifo["v2009"] == 256_602 - 144_787

    tetos = {r["norma"]: r["limite_r_bi"] for r in autorizacoes_legais()}
    assert tetos["Lei 11.948/2009, art. 1º"] == 100.0
    assert tetos["Lei 12.249/2010, art. 44"] == 180.0
    assert tetos["MP 505/2010 (Lei 12.397/2011)"] == 30.0

    custo = next(
        r for r in indicadores() if "Selic − TJLP" in r["indicador"] or "Selic - TJLP" in r["indicador"]
    )
    assert custo["valor_2010"] == 14_200.0


def test_fator_e_ipca_aplicado():
    ipca = _ipca_constante(0.5)
    fator = fator_dez2010_ref(ipca, data_ref=datetime(2026, 6, 30))
    assert fator > 1.0
    # dez/2010 → jun/2026 = 186 meses de variação após o mês-base
    # fator_ipca_entre usa razão dos fatores acumulados
    resumo = df_resumo_ipca(fator, datetime(2026, 6, 30))
    linha = resumo.loc[
        resumo["indicador"].str.contains("Créditos União"), "ipca_jun2026_r_mi"
    ].iloc[0]
    assert abs(linha - 236_723.0 * fator) < 1e-6

    ind = df_indicadores(fator)
    cred = ind.loc[ind["indicador"].str.contains("Créditos da União junto ao BNDES")].iloc[0]
    assert abs(cred["valor_2010_ipca_jun2026_r_mi"] - 236_723.0 * fator) < 1e-6


def test_build_gera_xlsx_e_md(tmp_path: Path):
    ipca = _ipca_constante(0.4)
    xlsx = tmp_path / "TCU_CG_2010.xlsx"
    md = tmp_path / "TCU_CG_2010_RELATORIO.md"
    p_xlsx, p_md, fator = build(
        ipca,
        data_ref=datetime(2026, 6, 30),
        xlsx=xlsx,
        md=md,
    )
    assert p_xlsx.exists()
    assert p_md.exists()
    assert fator > 1.0

    wb = load_workbook(p_xlsx)
    esperadas = {
        "Fonte",
        "Indicadores",
        "Creditos_DLSP",
        "Autorizacoes_Legais",
        "PAC_Subsidios_Eixo",
        "Resumo_IPCA",
        "Beneficios_Fin_Cred",
        "Base_Monetaria",
        "Base_Monetaria_Acum",
    }
    assert esperadas <= set(wb.sheetnames)
    assert (tmp_path / "TCU_CG_2010_BASE_MONETARIA.md").exists()
    assert (tmp_path / "grafico_base_monetaria_2003_2010.png").exists()

    texto = p_md.read_text(encoding="utf-8")
    assert "236,72" in texto
    assert "14,20" in texto
    assert "tcu.gov.br" in texto
    assert FONTE_URL in texto
    assert "Lei 11.948" in texto


def test_fatores_base_monetaria_identidade():
    rows = fatores_base_monetaria()
    assert [r["ano"] for r in rows] == list(range(2003, 2011))
    assert rows[0]["ano"] == 2003
    y2010 = next(r for r in rows if r["ano"] == 2010)
    assert y2010["titulos_publicos"] == 249_513
    assert y2010["demais_operacoes"] == -233_082
    assert y2010["var_base"] == 40_780

    df = df_fatores()
    assert df["identidade_ok"].all()
    assert abs(df["residuo"]).max() <= 1.0
    # Tesouro é contracionista em todos os anos
    assert (df["tesouro_nacional"] < 0).all()
    # 2008 é o único ano de contração do setor externo
    assert list(df.loc[df["setor_externo"] < 0, "ano"]) == [2008]

    det = df_detalhe_2009_2010()
    d2010 = det.loc[det["ano"] == 2010].iloc[0]
    demais = (
        d2010["depositos_inst_financ"]
        + d2010["derivativos_ajustes"]
        + d2010["outras_contas_ajustes"]
    )
    assert abs(demais - (-233_082)) < 1e-6


def test_dataframes_auxiliares():
    fator = 2.0
    assert abs(df_beneficios(fator)["y2010_ipca_jun2026"].iloc[-1] - 18_877.65 * 2) < 1e-6
    assert abs(df_pac_deson(fator)["projecao_2010_ipca_jun2026"].iloc[-1] - 46_636.0) < 1e-6
    assert abs(df_pac_eixo(fator)["desembolsos"].iloc[-1] - 4_444.24) < 1e-6
