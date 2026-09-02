"""Testes do discriminativo municipal UE2020 vs anteriores."""

from __future__ import annotations

import pandas as pd

from scripts.discriminativo_urnas_municipio import (
    classificar_geracao,
    discriminar_brasil,
    discriminar_municipios,
    discriminar_ufs,
    escrever_planilha,
    preparar_urnas,
    vencedor_votos,
)


def _urna(uf, cd, nome, serie, modelo, lula, bolso):
    return {
        "SG_UF": uf,
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "NR_URNA_EFETIVADA": serie,
        "NR_MODELO": modelo,
        "DS_MODELO_URNA": f"UE{modelo}" if modelo else "sem_faixa",
        "QT_VOTOS_LULA": lula,
        "QT_VOTOS_BOLSONARO": bolso,
        "QT_VOTOS_VALIDOS": lula + bolso,
    }


def test_classifica_geracao():
    assert classificar_geracao(2015) == "ANTERIOR_2020"
    assert classificar_geracao(2020) == "UE2020"
    assert classificar_geracao(None) == "SEM_FAIXA"
    assert vencedor_votos(10, 8) == "Lula"
    assert vencedor_votos(8, 10) == "Bolsonaro"
    assert vencedor_votos(5, 5) == "Empate"


def test_municipio_compara_votos_e_vitorias():
    df = pd.DataFrame(
        [
            _urna("RR", 3000, "BOA VISTA", 1, 2015, 80, 100),
            _urna("RR", 3000, "BOA VISTA", 2, 2020, 40, 120),
            _urna("AC", 1007, "BUJARI", 3, 2013, 100, 50),
            _urna("AC", 1007, "BUJARI", 4, 2020, 40, 80),
            _urna("ZZ", 99999, "EXTERIOR", 5, 2020, 10, 20),
        ]
    )
    mun = discriminar_municipios(df)
    assert set(mun["NM_MUNICIPIO"]) == {"BOA VISTA", "BUJARI", "EXTERIOR"}

    bv = mun.loc[mun["NM_MUNICIPIO"] == "BOA VISTA"].iloc[0]
    assert bv["COMPARAVEL"] == "S"
    assert bv["PCT_LULA_PRE2020"] == 44.44
    assert bv["PCT_LULA_UE2020"] == 25.0
    assert bv["DIF_PCT_LULA"] == -19.44
    assert bv["VENCEDOR_VOTOS_PRE2020"] == "Bolsonaro"
    assert bv["VENCEDOR_VOTOS_UE2020"] == "Bolsonaro"
    assert bv["INVERTEU_VENCEDOR_VOTOS"] == "N"
    assert bv["PCT_VITORIAS_LULA_PRE2020"] == 0.0
    assert bv["PCT_VITORIAS_BOLSONARO_UE2020"] == 100.0

    bj = mun.loc[mun["NM_MUNICIPIO"] == "BUJARI"].iloc[0]
    assert bj["INVERTEU_VENCEDOR_VOTOS"] == "S"
    assert bj["VENCEDOR_VOTOS_PRE2020"] == "Lula"
    assert bj["VENCEDOR_VOTOS_UE2020"] == "Bolsonaro"
    assert bj["PCT_LULA_PRE2020"] == 66.67
    assert bj["PCT_LULA_UE2020"] == 33.33
    assert bj["DIF_PCT_LULA"] == -33.34

    ext = mun.loc[mun["NM_MUNICIPIO"] == "EXTERIOR"].iloc[0]
    assert ext["COMPARAVEL"] == "N"
    assert ext["INVERTEU_VENCEDOR_VOTOS"] == "N"


def test_resumo_uf_e_brasil_so_comparaveis():
    df = pd.DataFrame(
        [
            _urna("AC", 1, "A", 1, 2015, 100, 50),
            _urna("AC", 1, "A", 2, 2020, 40, 80),
            _urna("AC", 2, "B", 3, 2020, 10, 10),
        ]
    )
    mun = discriminar_municipios(df)
    ufs = discriminar_ufs(mun)
    brasil = discriminar_brasil(mun)
    assert len(ufs) == 1
    assert int(ufs.iloc[0]["QT_MUNICIPIOS"]) == 1
    assert int(ufs.iloc[0]["QT_INVERTERAM"]) == 1
    assert brasil.iloc[0]["QT_INVERTERAM"] == 1
    assert brasil.iloc[0]["VENCEDOR_VOTOS_PRE2020"] == "Lula"
    assert brasil.iloc[0]["VENCEDOR_VOTOS_UE2020"] == "Bolsonaro"


def test_escreve_xlsx(tmp_path):
    df = pd.DataFrame(
        [
            _urna("AC", 1, "A", 1, 2015, 100, 50),
            _urna("AC", 1, "A", 2, 2020, 40, 80),
        ]
    )
    mun = discriminar_municipios(df)
    dest = tmp_path / "disc.xlsx"
    escrever_planilha(mun, discriminar_ufs(mun), discriminar_brasil(mun), dest)
    assert dest.exists()
    abas = pd.ExcelFile(dest).sheet_names
    assert "Municipios_comparaveis" in abas
    assert "Municipios_inverteram" in abas
    assert "Leia-me" in abas


def test_preparar_marca_ue2020():
    df = pd.DataFrame([_urna("SP", 1, "X", 9, 2020, 1, 2)])
    out = preparar_urnas(df)
    assert list(out["GERACAO"]) == ["UE2020"]
    assert list(out["VENCEDOR_URNA"]) == ["Bolsonaro"]
