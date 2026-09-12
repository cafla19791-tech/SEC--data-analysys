"""Testes do discriminativo interior NE até 40 mil eleitores (2014/2018/2022)."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_interior_nordeste_ate_40mil import (
    agregar_ufs,
    com_diferencas_pares,
    cruzar_anos,
    filtrar_interior_ate_limite,
    montar_resumo,
    padronizar_municipios,
)


def _urna(ano, uf, cd, nome, secao, pt, opp, branco=0, nulo=0, aptos=None):
    validos = pt + opp
    comparecimento = validos + branco + nulo
    if aptos is None:
        aptos = comparecimento + 20
    base = {
        "SG_UF": uf,
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_ZONA": 1,
        "NR_SECAO": secao,
        "QT_VOTOS_VALIDOS": validos,
        "QT_VOTOS_BRANCO": branco,
        "QT_VOTOS_NULO": nulo,
        "QT_COMPARECIMENTO": comparecimento,
        "QT_APTOS": aptos,
        "QT_ABSTENCOES": aptos - comparecimento,
    }
    if ano == 2014:
        base["QT_VOTOS_DILMA"] = pt
        base["QT_VOTOS_AECIO"] = opp
    elif ano == 2018:
        base["QT_VOTOS_HADDAD"] = pt
        base["QT_VOTOS_BOLSONARO"] = opp
    else:
        base["QT_VOTOS_LULA"] = pt
        base["QT_VOTOS_BOLSONARO"] = opp
    return base


def test_diferenca_validos_e_totais():
    linha = com_diferencas_pares(
        pd.DataFrame(
            [
                {
                    "CAND_PT": "Dilma",
                    "CAND_OPP": "Aécio",
                    "QT_VOTOS_PT": 80,
                    "QT_VOTOS_OPP": 20,
                    "QT_VOTOS_VALIDOS": 100,
                    "QT_VOTOS_BRANCO": 5,
                    "QT_VOTOS_NULO": 5,
                    "QT_COMPARECIMENTO": 110,
                }
            ]
        )
    ).iloc[0]
    assert linha["DIF_PCT_VALIDOS"] == 60.0
    assert linha["DIF_PCT_TOTAIS"] == 54.55
    assert linha["VENCEDOR"] == "Dilma"


def test_2018_totais_sem_comparecimento():
    linha = com_diferencas_pares(
        pd.DataFrame(
            [
                {
                    "CAND_PT": "Haddad",
                    "CAND_OPP": "Bolsonaro",
                    "QT_VOTOS_PT": 60,
                    "QT_VOTOS_OPP": 40,
                    "QT_VOTOS_VALIDOS": 100,
                    "QT_VOTOS_BRANCO": 8,
                    "QT_VOTOS_NULO": 2,
                }
            ]
        )
    ).iloc[0]
    assert linha["QT_VOTOS_TOTAIS"] == 110
    assert linha["DIF_PCT_TOTAIS"] == 18.18


def test_filtra_capital_e_acima_de_40_mil():
    d14 = padronizar_municipios(
        pd.DataFrame(
            [
                _urna(2014, "BA", 1, "SALVADOR", 1, 90, 10, aptos=1_000_000),
                _urna(2014, "BA", 2, "FEIRA DE SANTANA", 1, 70, 30, aptos=250_000),
                _urna(2014, "BA", 3, "XIQUE-XIQUE", 1, 80, 20, aptos=28_000),
                _urna(2014, "MA", 4, "FORTALEZA DOS NOGUEIRAS", 1, 75, 25, aptos=12_000),
            ]
        ),
        2014,
    )
    d18 = padronizar_municipios(
        pd.DataFrame(
            [
                _urna(2018, "BA", 1, "SALVADOR", 1, 60, 40, aptos=1_100_000),
                _urna(2018, "BA", 2, "FEIRA DE SANTANA", 1, 55, 45, aptos=260_000),
                _urna(2018, "BA", 3, "XIQUE-XIQUE", 1, 85, 15, aptos=29_000),
                _urna(2018, "MA", 4, "FORTALEZA DOS NOGUEIRAS", 1, 70, 30, aptos=13_000),
            ]
        ),
        2018,
    )
    d22 = padronizar_municipios(
        pd.DataFrame(
            [
                _urna(2022, "BA", 1, "SALVADOR", 1, 65, 35, aptos=1_200_000),
                _urna(2022, "BA", 2, "FEIRA DE SANTANA", 1, 58, 42, aptos=270_000),
                _urna(2022, "BA", 3, "XIQUE-XIQUE", 1, 88, 12, aptos=30_000),
                _urna(2022, "MA", 4, "FORTALEZA DOS NOGUEIRAS", 1, 68, 32, aptos=14_000),
            ]
        ),
        2022,
    )
    cruz = cruzar_anos({2014: d14, 2018: d18, 2022: d22})
    recorte = filtrar_interior_ate_limite(cruz, 40_000)
    assert set(recorte["NM_MUNICIPIO"]) == {"XIQUE-XIQUE", "FORTALEZA DOS NOGUEIRAS"}
    assert recorte["QT_APTOS_2022"].max() <= 40_000
    xique = recorte[recorte["NM_MUNICIPIO"] == "XIQUE-XIQUE"].iloc[0]
    assert xique["VENCEDOR_2014"] == "Dilma"
    assert xique["VENCEDOR_2018"] == "Haddad"
    assert xique["VENCEDOR_2022"] == "Lula"
    assert xique["DIF_PCT_VALIDOS_2022"] == 76.0
    resumo = montar_resumo(recorte)
    assert list(resumo["Pleito"]) == ["2º turno 2014", "2º turno 2018", "2º turno 2022"]
    assert resumo.iloc[2]["Municípios"] == 2
    por_uf = agregar_ufs(recorte)
    assert set(por_uf["SG_UF"]) == {"BA", "MA"}
