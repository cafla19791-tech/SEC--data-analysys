"""Testes do comparativo Bolsonaro×Haddad 2018 vs Bolsonaro×Lula 2022 por seção."""

from __future__ import annotations

import pandas as pd

from scripts.comparativo_bolsonaro_secoes_2018_2022 import (
    cruzar_secoes,
    preparar_ano,
    resumo_recorte,
    tabela_municipio,
)


def _secao(uf, cd, nome, zona, secao, bolso, adv):
    return {
        "SG_UF": uf,
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_ZONA": zona,
        "NR_SECAO": secao,
        "QT_VOTOS_BOLSONARO": bolso,
        "QT_VOTOS_HADDAD": adv,
        "QT_VOTOS_LULA": adv,
        "QT_VOTOS_VALIDOS": bolso + adv,
    }


def test_prepara_margem_bolsonaro_menos_adversario():
    df = preparar_ano(
        pd.DataFrame([_secao("MG", 1, "POVOADO", 1, 10, 70, 30)]),
        ano=2018,
        col_adv="QT_VOTOS_HADDAD",
        nome_adv="Haddad",
    )
    row = df.iloc[0]
    assert row["PCT_BOLSONARO_2018"] == 70.0
    assert row["PCT_HADDAD_2018"] == 30.0
    assert row["DIF_BOLSO_HADDAD_2018"] == 40.0
    assert row["VENCEDOR_2018"] == "Bolsonaro"


def test_cruza_secao_a_secao_e_marca_incomparaveis():
    a18 = preparar_ano(
        pd.DataFrame(
            [
                _secao("MG", 1, "POVOADO", 1, 10, 70, 30),
                _secao("MG", 1, "POVOADO", 1, 11, 40, 60),
                _secao("MG", 1, "POVOADO", 1, 12, 55, 45),
            ]
        ),
        ano=2018,
        col_adv="QT_VOTOS_HADDAD",
        nome_adv="Haddad",
    )
    a22 = preparar_ano(
        pd.DataFrame(
            [
                _secao("MG", 1, "POVOADO", 1, 10, 48, 52),
                _secao("MG", 1, "POVOADO", 1, 11, 35, 65),
                _secao("MG", 1, "POVOADO", 1, 99, 80, 20),
            ]
        ),
        ano=2022,
        col_adv="QT_VOTOS_LULA",
        nome_adv="Lula",
    )
    out = cruzar_secoes(a18, a22)
    assert len(out) == 4
    par = out[out["COMPARAVEL"] == "S"]
    assert len(par) == 2
    s10 = out[(out["NR_SECAO"] == 10)].iloc[0]
    assert s10["DIF_BOLSO_HADDAD_2018"] == 40.0
    assert s10["DIF_BOLSO_LULA_2022"] == -4.0
    assert s10["DIF_MARGEM_BOLSO"] == -44.0
    assert s10["INVERTEU"] == "S"
    assert s10["VENCEDOR_2022"] == "Lula"
    so12 = out[out["NR_SECAO"] == 12].iloc[0]
    assert so12["COMPARAVEL"] == "N"
    assert pd.isna(so12["DIF_MARGEM_BOLSO"])
    so99 = out[out["NR_SECAO"] == 99].iloc[0]
    assert so99["COMPARAVEL"] == "N"
    res = resumo_recorte("teste", out)
    assert res["Seções comparáveis"] == 2
    assert res["Só em 2018"] == 1
    assert res["Só em 2022"] == 1
    assert res["Inverteram vencedor"] == 1
    mun = tabela_municipio(out)
    assert len(mun) == 1
    assert mun.iloc[0]["QT_SECOES"] == 2
