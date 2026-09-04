"""Testes das planilhas por região / UF / município / zona / urna."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.planilha_resultados_presidente import (
    TOTAIS_OFICIAIS,
    agregar,
    colunas_candidatos,
    conferir_totais,
    enriquecer_serie,
    escrever_xlsx,
    pasta_dados,
    preparar,
    recortes,
    resolver_fonte,
    rotulo_regiao,
    vencedor_frame,
)


def _linhas() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "SG_UF": "BA",
                "CD_MUNICIPIO": 1,
                "NM_MUNICIPIO": "SALVADOR",
                "NR_ZONA": 1,
                "NR_SECAO": 10,
                "NR_URNA_EFETIVADA": 2000001,
                "QT_VOTOS_LULA": 80,
                "QT_VOTOS_BOLSONARO": 20,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 1,
                "QT_VOTOS_NULO": 2,
            },
            {
                "SG_UF": "BA",
                "CD_MUNICIPIO": 1,
                "NM_MUNICIPIO": "SALVADOR",
                "NR_ZONA": 1,
                "NR_SECAO": 11,
                "NR_URNA_EFETIVADA": 1800000,
                "QT_VOTOS_LULA": 40,
                "QT_VOTOS_BOLSONARO": 60,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 1,
            },
            {
                "SG_UF": "RS",
                "CD_MUNICIPIO": 2,
                "NM_MUNICIPIO": "PORTO ALEGRE",
                "NR_ZONA": 5,
                "NR_SECAO": 1,
                "NR_URNA_EFETIVADA": 2100000,
                "QT_VOTOS_LULA": 10,
                "QT_VOTOS_BOLSONARO": 90,
                "QT_VOTOS_VALIDOS": 100,
                "QT_VOTOS_BRANCO": 0,
                "QT_VOTOS_NULO": 0,
            },
        ]
    )


def test_regiao_e_vencedor():
    assert rotulo_regiao("BA") == "Nordeste"
    assert rotulo_regiao("rs") == "Sul"
    assert rotulo_regiao("ZZ") == "Exterior"
    df = preparar(_linhas(), 2022, 2)
    assert set(df["REGIAO"]) == {"Nordeste", "Sul"}
    assert colunas_candidatos(df) == ["QT_VOTOS_LULA", "QT_VOTOS_BOLSONARO"]
    venc = vencedor_frame(df, ["QT_VOTOS_LULA", "QT_VOTOS_BOLSONARO"])
    assert list(venc) == ["LULA", "BOLSONARO", "BOLSONARO"]


def test_agrega_uf_bate_com_urnas():
    df = preparar(_linhas(), 2022, 2)
    ufs = agregar(df, ["SG_UF"])
    ba = ufs.loc[ufs["SG_UF"] == "BA"].iloc[0]
    assert int(ba["QT_SECOES"]) == 2
    assert int(ba["QT_VOTOS_LULA"]) == 120
    assert int(ba["QT_VOTOS_BOLSONARO"]) == 80
    assert ba["PCT_LULA"] == 60.0
    assert ba["VENCEDOR"] == "LULA"
    rs = ufs.loc[ufs["SG_UF"] == "RS"].iloc[0]
    assert rs["VENCEDOR"] == "BOLSONARO"
    assert int(ufs["QT_VOTOS_LULA"].sum()) == int(df["QT_VOTOS_LULA"].sum())


def test_enriquece_serie_por_secao():
    secao = pd.DataFrame(
        [
            {
                "SG_UF": "BA",
                "CD_MUNICIPIO": "1",
                "NR_ZONA": "1",
                "NR_SECAO": "10",
                "QT_VOTOS_LULA": 80,
            }
        ]
    )
    bu = pd.DataFrame(
        [
            {
                "SG_UF": "BA",
                "CD_MUNICIPIO": 1,
                "NR_ZONA": 1,
                "NR_SECAO": 10,
                "NR_URNA_EFETIVADA": 2232140,
                "DS_MODELO_URNA": "UE2020",
            }
        ]
    )
    out = enriquecer_serie(secao, bu)
    assert int(out["NR_URNA_EFETIVADA"].iloc[0]) == 2_232_140
    assert out["DS_MODELO_URNA"].iloc[0] == "UE2020"


def test_xlsx_tem_abas(tmp_path: Path):
    df = preparar(_linhas(), 2022, 2)
    abas = recortes(df)
    dest = tmp_path / "t.xlsx"
    escrever_xlsx(dest, [("Fonte", "teste")], abas)
    lidas = pd.read_excel(dest, sheet_name=None)
    assert set(lidas) == {"Leia-me", "Regiao", "UF", "Municipio", "Zona", "Urna"}
    assert len(lidas["Urna"]) == 3
    assert len(lidas["Municipio"]) == 2
    assert int(lidas["Regiao"].loc[lidas["Regiao"]["REGIAO"] == "Nordeste", "QT_VOTOS_LULA"].iloc[0]) == 120


def test_conferir_totais_oficial():
    df = pd.DataFrame({"QT_VOTOS_LULA": [60_345_999], "QT_VOTOS_BOLSONARO": [58_206_354]})
    avisos = conferir_totais(df, 2022, 2)
    assert all("OK" in a for a in avisos)
    df_errado = pd.DataFrame({"QT_VOTOS_LULA": [1], "QT_VOTOS_BOLSONARO": [2]})
    assert any("DIFERE" in a for a in conferir_totais(df_errado, 2022, 2))


# 2014 2º turno: pastas do Drive ainda sem o BU do Ceará.
TOTAIS_PARCIAIS = {
    (2014, 2): {"DILMA": 50_978_234, "AECIO": 49_972_986},
}


def test_totais_oficiais_nos_csv_publicados():
    dados = pasta_dados()
    for (ano, turno), esperados in TOTAIS_OFICIAIS.items():
        try:
            path = resolver_fonte(dados, ano, turno)
        except FileNotFoundError:
            continue
        df = pd.read_csv(path, compression="gzip" if str(path).endswith(".gz") else None)
        alvo = TOTAIS_PARCIAIS.get((ano, turno), esperados)
        for nome, esperado in alvo.items():
            col = f"QT_VOTOS_{nome}"
            if col not in df.columns:
                continue
            assert int(df[col].fillna(0).sum()) == esperado, f"{ano} T{turno} {nome}"
        if (ano, turno) in TOTAIS_PARCIAIS:
            for nome, oficial in esperados.items():
                assert alvo[nome] < oficial, f"{ano} T{turno} {nome} deveria ser parcial"
