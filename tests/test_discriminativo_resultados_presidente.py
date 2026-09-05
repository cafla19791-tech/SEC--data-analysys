"""Testes do discriminativo 2014 × 2018 × 2022."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_resultados_presidente import (
    agregar_2t,
    cruzar_anos,
    inverteram,
    lado_vencedor,
    linha_brasil_2t,
    nome_vencedor,
    preparar_2t,
)


def _urna(ano: int, uf: str, cd: int, nome: str, secao: int, pt: int, opp: int) -> dict:
    if ano == 2014:
        votos = {"QT_VOTOS_DILMA": pt, "QT_VOTOS_AECIO": opp}
    elif ano == 2018:
        votos = {"QT_VOTOS_HADDAD": pt, "QT_VOTOS_BOLSONARO": opp}
    else:
        votos = {"QT_VOTOS_LULA": pt, "QT_VOTOS_BOLSONARO": opp}
    return {
        "SG_UF": uf,
        "REGIAO": "Nordeste" if uf == "BA" else "Sudeste",
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_ZONA": 1,
        "NR_SECAO": secao,
        "QT_VOTOS_VALIDOS": pt + opp,
        **votos,
    }


def test_lado_e_inversao():
    assert lado_vencedor(10, 8) == "PT"
    assert lado_vencedor(8, 10) == "OPP"
    assert lado_vencedor(5, 5) == "EMPATE"
    assert nome_vencedor(2014, "PT") == "Dilma"
    assert nome_vencedor(2018, "OPP") == "Bolsonaro"
    assert nome_vencedor(2022, "PT") == "Lula"
    assert inverteram("PT", "OPP") == "S"
    assert inverteram("PT", "PT") == "N"
    assert inverteram("EMPATE", "OPP") == "N"


def test_cruza_municipio_e_detecta_inversao():
    d14 = preparar_2t(
        pd.DataFrame(
            [
                _urna(2014, "BA", 1, "SALVADOR", 10, 80, 20),
                _urna(2014, "BA", 1, "SALVADOR", 11, 70, 30),
                _urna(2014, "SP", 2, "SANTOS", 1, 40, 60),
            ]
        ),
        2014,
    )
    d18 = preparar_2t(
        pd.DataFrame(
            [
                _urna(2018, "BA", 1, "SALVADOR", 10, 55, 45),
                _urna(2018, "SP", 2, "Santos", 1, 30, 70),
            ]
        ),
        2018,
    )
    d22 = preparar_2t(
        pd.DataFrame(
            [
                _urna(2022, "BA", 1, "SALVADOR", 10, 40, 60),
                _urna(2022, "SP", 2, "SANTOS", 1, 65, 35),
            ]
        ),
        2022,
    )
    chaves = ["REGIAO", "SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"]
    blocos = {
        2014: agregar_2t(d14, chaves, 2014),
        2018: agregar_2t(d18, chaves, 2018),
        2022: agregar_2t(d22, chaves, 2022),
    }
    out = cruzar_anos(blocos, chaves)
    assert len(out) == 2
    ba = out.loc[out["SG_UF"] == "BA"].iloc[0]
    assert ba["NM_MUNICIPIO"] == "SALVADOR"
    assert int(ba["QT_VOTOS_PT_2014"]) == 150
    assert ba["VENCEDOR_2014"] == "Dilma"
    assert ba["VENCEDOR_2022"] == "Bolsonaro"
    assert ba["INVERTEU_2014_2022"] == "S"
    assert ba["COMPARAVEL"] == "S"
    assert ba["PCT_PT_2014"] == 75.0
    assert ba["PCT_PT_2022"] == 40.0
    assert ba["DIF_PCT_PT_2014_2022"] == -35.0

    sp = out.loc[out["SG_UF"] == "SP"].iloc[0]
    assert sp["VENCEDOR_2014"] == "Aécio"
    assert sp["VENCEDOR_2022"] == "Lula"
    assert sp["INVERTEU_2018_2022"] == "S"


def test_brasil_2t_soma():
    df = preparar_2t(
        pd.DataFrame(
            [
                _urna(2022, "BA", 1, "SALVADOR", 1, 60, 40),
                _urna(2022, "SP", 2, "SANTOS", 1, 20, 80),
            ]
        ),
        2022,
    )
    linha = linha_brasil_2t(df, 2022)
    assert linha["QT_VOTOS_PT"] == 80
    assert linha["QT_VOTOS_OPP"] == 120
    assert linha["VENCEDOR"] == "Bolsonaro"
    assert linha["PCT_PT"] == 40.0
