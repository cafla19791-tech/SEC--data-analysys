"""Testes do discriminativo Lula × Bolsonaro no interior do Nordeste."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_interior_nordeste import (
    agregar_municipios,
    agregar_ufs,
    com_diferencas,
    eh_capital_nordeste,
    montar_resumo,
    normalizar_nome,
    pct,
    recorte_nordeste,
)


def _urna(uf: str, cd: int, nome: str, secao: int, lula: int, bolo: int, branco=0, nulo=0, aptos=None):
    validos = lula + bolo
    comparecimento = validos + branco + nulo
    if aptos is None:
        aptos = comparecimento + 10
    return {
        "SG_UF": uf,
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_ZONA": 1,
        "NR_SECAO": secao,
        "QT_VOTOS_LULA": lula,
        "QT_VOTOS_BOLSONARO": bolo,
        "QT_VOTOS_VALIDOS": validos,
        "QT_VOTOS_BRANCO": branco,
        "QT_VOTOS_NULO": nulo,
        "QT_COMPARECIMENTO": comparecimento,
        "QT_APTOS": aptos,
        "QT_ABSTENCOES": aptos - comparecimento,
    }


def test_normaliza_e_reconhece_so_a_capital():
    assert normalizar_nome("São Luís") == "SAO LUIS"
    assert eh_capital_nordeste("MA", "SÃO LUÍS")
    assert not eh_capital_nordeste("MA", "SÃO LUÍS GONZAGA DO MARANHÃO")
    assert not eh_capital_nordeste("CE", "FORTALEZA DOS NOGUEIRAS")
    assert not eh_capital_nordeste("RN", "CORONEL JOÃO PESSOA")
    assert eh_capital_nordeste("pb", "João Pessoa")
    assert not eh_capital_nordeste("SP", "SÃO PAULO")


def test_pct_e_diferencas():
    assert pct(70, 100) == 70.0
    assert pct(10, 0) is None
    linha = com_diferencas(
        pd.DataFrame(
            [
                {
                    "QT_VOTOS_LULA": 70,
                    "QT_VOTOS_BOLSONARO": 30,
                    "QT_VOTOS_VALIDOS": 100,
                    "QT_VOTOS_BRANCO": 4,
                    "QT_VOTOS_NULO": 6,
                    "QT_COMPARECIMENTO": 110,
                    "QT_APTOS": 130,
                    "QT_ABSTENCOES": 20,
                }
            ]
        )
    ).iloc[0]
    assert linha["QT_DIF_LULA_BOLSONARO"] == 40
    assert linha["DIF_PCT_VALIDOS"] == 40.0
    assert linha["DIF_PCT_TOTAIS"] == 36.36
    assert linha["PCT_LULA_TOTAIS"] == 63.64
    assert linha["VENCEDOR"] == "Lula"


def test_recorte_exclui_capitais_e_outras_regioes():
    urnas = pd.DataFrame(
        [
            _urna("BA", 38490, "SALVADOR", 1, 80, 20, branco=2, nulo=3),
            _urna("BA", 35157, "FEIRA DE SANTANA", 1, 60, 40),
            _urna("PE", 25313, "RECIFE", 1, 55, 45),
            _urna("PE", 23833, "CARUARU", 1, 90, 10, branco=5, nulo=5),
            _urna("MA", 7790, "FORTALEZA DOS NOGUEIRAS", 1, 70, 30),
            _urna("SP", 71072, "SÃO PAULO", 1, 10, 90),
        ]
    )
    ne = recorte_nordeste(urnas)
    assert set(ne["SG_UF"]) == {"BA", "PE", "MA"}
    mun = agregar_municipios(ne)
    interiores = mun[~mun["EH_CAPITAL"]]
    capitais = mun[mun["EH_CAPITAL"]]
    assert set(interiores["NM_MUNICIPIO"]) == {
        "FEIRA DE SANTANA",
        "CARUARU",
        "FORTALEZA DOS NOGUEIRAS",
    }
    assert set(capitais["NM_MUNICIPIO"]) == {"SALVADOR", "RECIFE"}
    caruaru = interiores[interiores["NM_MUNICIPIO"] == "CARUARU"].iloc[0]
    assert caruaru["DIF_PCT_VALIDOS"] == 80.0
    assert caruaru["QT_VOTOS_TOTAIS"] == 110
    assert caruaru["DIF_PCT_TOTAIS"] == 72.73


def test_resumo_e_uf_somam_so_interior():
    urnas = pd.DataFrame(
        [
            _urna("BA", 1, "SALVADOR", 1, 100, 50),
            _urna("BA", 2, "FEIRA DE SANTANA", 1, 80, 20),
            _urna("AL", 3, "MACEIÓ", 1, 10, 90),
            _urna("AL", 4, "ARAPIRACA", 1, 40, 60),
        ]
    )
    mun = agregar_municipios(recorte_nordeste(urnas))
    interior = mun[~mun["EH_CAPITAL"]]
    por_uf = agregar_ufs(interior)
    ba = por_uf[por_uf["SG_UF"] == "BA"].iloc[0]
    assert ba["QT_VOTOS_LULA"] == 80
    assert ba["QT_VOTOS_BOLSONARO"] == 20
    assert ba["NM_MUNICIPIO"] == "1 município do interior"
    resumo = montar_resumo(mun)
    int_row = resumo[resumo["Recorte"] == "Interior do Nordeste"].iloc[0]
    assert int_row["Municípios"] == 2
    assert int_row["Lula"] == 120
    assert int_row["Bolsonaro"] == 80
    assert int_row["Diferença p.p. (válidos)"] == 20.0
