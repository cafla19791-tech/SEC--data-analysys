"""Testes das margens UE2020 vs antigas em municípios pequenos."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_modelo_urna_ate_50mil import (
    apurar_municipios,
    enriquecer_urnas,
    resumo_recorte,
)


def _urna(uf, cd, nome, modelo, lula, bolo):
    return {
        "SG_UF": uf,
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_MODELO": modelo,
        "QT_VOTOS_LULA": lula,
        "QT_VOTOS_BOLSONARO": bolo,
        "QT_VOTOS_VALIDOS": lula + bolo,
    }


def test_margens_por_modelo_e_comparacao():
    urnas = enriquecer_urnas(
        pd.DataFrame(
            [
                # antigas Lula: margens 60 e 20 → média 40
                _urna("BA", 1, "POVOADO", 2015, 80, 20),
                _urna("BA", 1, "POVOADO", 2013, 60, 40),
                # novas Lula: margem 80
                _urna("BA", 1, "POVOADO", 2020, 90, 10),
                # antigas Bolsonaro: margem 20
                _urna("BA", 1, "POVOADO", 2011, 40, 60),
                # novas Bolsonaro: margem 40
                _urna("BA", 1, "POVOADO", 2020, 30, 70),
                # cidade grande (fica de fora)
                _urna("SP", 2, "METROPOLE", 2020, 10, 90),
            ]
        )
    )
    pop = pd.DataFrame(
        {
            "codigo_tse": [1, 2],
            "uf": ["BA", "SP"],
            "nome_municipio": ["POVOADO", "METROPOLE"],
            "codigo_ibge": [1, 2],
            "POP_2022": [12_000, 200_000],
        }
    )
    mun = apurar_municipios(urnas, pop, 50_000)
    assert list(mun["NM_MUNICIPIO"]) == ["POVOADO"]
    row = mun.iloc[0]
    assert row["COMPARA_LULA"] == "S"
    assert row["COMPARA_BOLSO"] == "S"
    assert row["MARGEM_LULA_PRE2020"] == 40.0
    assert row["MARGEM_LULA_UE2020"] == 80.0
    assert row["DIF_MARGEM_LULA_NOVA_ANTIGA"] == 40.0
    assert row["MARGEM_BOLSO_PRE2020"] == 20.0
    assert row["MARGEM_BOLSO_UE2020"] == 40.0
    assert row["DIF_MARGEM_BOLSO_NOVA_ANTIGA"] == 20.0
    res = resumo_recorte("Brasil", mun)
    assert res["Diferença Lula pareada (nova − antiga)"] == 40.0
    assert res["Diferença Bolsonaro pareada (nova − antiga)"] == 20.0
