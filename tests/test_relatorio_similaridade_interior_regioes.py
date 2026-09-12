"""Testes do relatório de similaridade das demais regiões."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_interior_nordeste import eh_capital, eh_capital_nordeste
from scripts.relatorio_similaridade_interior_regioes import (
    REGIOES_FOCO,
    enriquecer,
    linha_comparativo,
    padronizar_municipios_pais,
)


def test_capitais_nacionais_e_homonimos():
    assert eh_capital("SP", "SÃO PAULO")
    assert eh_capital("RJ", "Rio de Janeiro")
    assert eh_capital("DF", "Brasília")
    assert eh_capital("GO", "Goiânia")
    assert not eh_capital("SP", "SÃO PAULO DAS MISSÕES")
    assert not eh_capital("MA", "FORTALEZA DOS NOGUEIRAS")
    assert not eh_capital("CE", "FORTALEZA DOS NOGUEIRAS")
    assert eh_capital_nordeste("CE", "FORTALEZA")
    assert not eh_capital_nordeste("SP", "SÃO PAULO")


def test_padroniza_pais_e_enriquece_regiao():
    urnas = pd.DataFrame(
        [
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 1,
                "NM_MUNICIPIO": "SÃO PAULO",
                "NR_SECAO": 1,
                "QT_VOTOS_LULA": 40,
                "QT_VOTOS_BOLSONARO": 60,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 1,
                "QT_VOTOS_NULO": 1,
                "QT_COMPARECIMENTO": 102,
                "QT_APTOS": 200,
            },
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 2,
                "NM_MUNICIPIO": "CAMPINAS",
                "NR_SECAO": 1,
                "QT_VOTOS_LULA": 45,
                "QT_VOTOS_BOLSONARO": 55,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 0,
                "QT_COMPARECIMENTO": 100,
                "QT_APTOS": 80_000,
            },
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 3,
                "NM_MUNICIPIO": "SANTOS",
                "NR_SECAO": 1,
                "QT_VOTOS_LULA": 35,
                "QT_VOTOS_BOLSONARO": 65,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 0,
                "QT_COMPARECIMENTO": 100,
                "QT_APTOS": 40_000,
            },
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 4,
                "NM_MUNICIPIO": "RIBEIRÃO PRETO",
                "NR_SECAO": 1,
                "QT_VOTOS_LULA": 30,
                "QT_VOTOS_BOLSONARO": 70,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 0,
                "QT_COMPARECIMENTO": 100,
                "QT_APTOS": 50_000,
            },
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 5,
                "NM_MUNICIPIO": "SOROCABA",
                "NR_SECAO": 1,
                "QT_VOTOS_LULA": 38,
                "QT_VOTOS_BOLSONARO": 62,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 0,
                "QT_COMPARECIMENTO": 100,
                "QT_APTOS": 45_000,
            },
        ]
    )
    mun = padronizar_municipios_pais(urnas, 2022)
    assert set(mun.loc[mun["EH_CAPITAL"], "NM_MUNICIPIO_2022"]) == {"SÃO PAULO"}
    cruz = mun.rename(columns={"NM_MUNICIPIO_2022": "NM_MUNICIPIO"})
    for ano in (2014, 2018, 2022):
        cruz[f"VENCEDOR_{ano}"] = cruz["VENCEDOR_2022"]
        cruz[f"PCT_PT_VALIDOS_{ano}"] = cruz["PCT_PT_VALIDOS_2022"]
        cruz[f"QT_VOTOS_PT_{ano}"] = cruz["QT_VOTOS_PT_2022"]
        cruz[f"QT_VOTOS_OPP_{ano}"] = cruz["QT_VOTOS_OPP_2022"]
        cruz[f"QT_VOTOS_VALIDOS_{ano}"] = cruz["QT_VOTOS_VALIDOS_2022"]
    enr = enriquecer(cruz)
    interior = enr[~enr["EH_CAPITAL"]]
    assert "SÃO PAULO" not in set(interior["NM_MUNICIPIO"])
    assert set(interior["REGIAO"]) == {"Sudeste"}
    rec = linha_comparativo("Sudeste", interior)
    assert rec["Municípios"] == 4
    assert rec["Mesmo vencedor nas 3 (%)"] == 100.0


def test_regioes_foco_nao_incluem_nordeste():
    assert "Nordeste" not in REGIOES_FOCO
    assert set(REGIOES_FOCO) == {"Norte", "Centro-Oeste", "Sudeste", "Sul"}
