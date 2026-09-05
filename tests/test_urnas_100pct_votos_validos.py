"""Urnas com 100% dos votos válidos para um único candidato."""

from __future__ import annotations

import pandas as pd

from scripts.urnas_100pct_votos_validos import filtrar_100pct


def test_so_um_candidato_com_todos_os_validos():
    df = pd.DataFrame(
        [
            {"SG_UF": "BA", "NR_SECAO": 1, "QT_VOTOS_LULA": 80, "QT_VOTOS_BOLSONARO": 0, "QT_VOTOS_VALIDOS": 80},
            {"SG_UF": "SP", "NR_SECAO": 2, "QT_VOTOS_LULA": 40, "QT_VOTOS_BOLSONARO": 40, "QT_VOTOS_VALIDOS": 80},
            {"SG_UF": "RS", "NR_SECAO": 3, "QT_VOTOS_LULA": 0, "QT_VOTOS_BOLSONARO": 10, "QT_VOTOS_VALIDOS": 10},
            {"SG_UF": "ZZ", "NR_SECAO": 4, "QT_VOTOS_LULA": 0, "QT_VOTOS_BOLSONARO": 0, "QT_VOTOS_VALIDOS": 0},
        ]
    )
    out = filtrar_100pct(df, 2022, 2)
    assert set(out["SG_UF"]) == {"BA", "RS"}
    assert list(out["CANDIDATO"]) == ["LULA", "BOLSONARO"]
    assert int(out.loc[out["SG_UF"] == "BA", "QT_VOTOS_CANDIDATO"].iloc[0]) == 80
