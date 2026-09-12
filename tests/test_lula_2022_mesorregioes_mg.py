"""Testes da divisão por mesorregião dos municípios mineiros com vitória de Lula."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.lula_2022_mesorregioes_mg import (
    CATALOGO_MESO,
    _nome_aba,
    chave_nome,
    cruzar_mesorregiao,
    enriquecer_municipios,
    municipios_do_discriminativo,
    municipios_lula,
    resumo_mesorregioes,
    carregar_catalogo,
)


def _mun(cd, nome, lula, bolo):
    return {
        "SG_UF": "MG",
        "CD_MUNICIPIO": cd,
        "NM_MUNICIPIO": nome,
        "QT_SECOES": 2,
        "QT_VOTOS_LULA": lula,
        "QT_VOTOS_BOLSONARO": bolo,
        "QT_VOTOS_VALIDOS": lula + bolo,
    }


def test_nome_aba_sem_barra_do_excel():
    nome = _nome_aba(5, "Triângulo Mineiro / Alto Paranaíba")
    assert "/" not in nome
    assert nome.startswith("05 ")
    assert len(nome) <= 31


def test_chave_nome_ignora_hifen_e_acento():
    assert chave_nome("Olhos-d'Água") == chave_nome("OLHOS D AGUA")
    assert chave_nome("Sem-Peixe") == chave_nome("SEM PEIXE")


def test_catalogo_oficial_tem_12_mesos_e_853_municipios():
    cat = carregar_catalogo(CATALOGO_MESO)
    assert len(cat) == 853
    assert cat["codigo_ibge"].nunique() == 853
    assert cat["codigo_tse"].nunique() == 853
    assert cat["codigo_meso"].nunique() == 12
    por = cat.groupby("nome_meso").size()
    assert int(por["Noroeste de Minas"]) == 19
    assert int(por["Norte de Minas"]) == 89
    assert int(por["Jequitinhonha"]) == 51
    assert int(por["Vale do Mucuri"]) == 23
    assert int(por["Triângulo Mineiro / Alto Paranaíba"]) == 66
    assert int(por["Central Mineira"]) == 30
    assert int(por["Metropolitana de Belo Horizonte"]) == 105
    assert int(por["Vale do Rio Doce"]) == 102
    assert int(por["Oeste de Minas"]) == 44
    assert int(por["Sul / Sudoeste de Minas"]) == 146
    assert int(por["Campo das Vertentes"]) == 36
    assert int(por["Zona da Mata"]) == 142


def test_cidades_conhecidas_nas_mesorregioes_do_pdf():
    cat = carregar_catalogo(CATALOGO_MESO)
    nome = cat.set_index("nome_municipio")["nome_meso"]
    assert nome["Belo Horizonte"] == "Metropolitana de Belo Horizonte"
    assert nome["Uberlândia"] == "Triângulo Mineiro / Alto Paranaíba"
    assert nome["Montes Claros"] == "Norte de Minas"
    assert nome["Juiz de Fora"] == "Zona da Mata"
    assert nome["Unaí"] == "Noroeste de Minas"
    assert nome["Teófilo Otoni"] == "Vale do Mucuri"
    assert nome["Diamantina"] == "Jequitinhonha"
    assert nome["Divinópolis"] == "Oeste de Minas"
    assert nome["Varginha"] == "Sul / Sudoeste de Minas"
    assert nome["São João del Rei"] == "Campo das Vertentes"
    assert nome["Governador Valadares"] == "Vale do Rio Doce"
    assert nome["Curvelo"] == "Central Mineira"


def test_so_municipios_com_vitoria_de_lula_entram_e_caem_na_meso_certa():
    cat = pd.DataFrame(
        {
            "codigo_ibge": [3106200, 3140704, 3170206],
            "codigo_tse": [41238, 48640, 54039],
            "nome_municipio": ["Belo Horizonte", "Juiz de Fora", "Uberlândia"],
            "codigo_meso": [7, 12, 5],
            "nome_meso": [
                "Metropolitana de Belo Horizonte",
                "Zona da Mata",
                "Triângulo Mineiro / Alto Paranaíba",
            ],
            "codigo_micro": [30, 65, 18],
            "nome_micro": ["Belo Horizonte", "Juiz de Fora", "Uberlândia"],
            "chave_nome": [
                chave_nome("Belo Horizonte"),
                chave_nome("Juiz de Fora"),
                chave_nome("Uberlândia"),
            ],
        }
    )
    mun = enriquecer_municipios(
        pd.DataFrame(
            [
                _mun(41238, "BELO HORIZONTE", 80, 20),
                _mun(48640, "JUIZ DE FORA", 55, 45),
                _mun(54039, "UBERLANDIA", 40, 60),
            ]
        )
    )
    lula = municipios_lula(mun, cat)
    assert list(lula["NM_MUNICIPIO"]) == ["BELO HORIZONTE", "JUIZ DE FORA"]
    assert list(lula["nome_meso"]) == [
        "Metropolitana de Belo Horizonte",
        "Zona da Mata",
    ]
    resumo = resumo_mesorregioes(lula, cat)
    assert int(resumo.loc[resumo["Mesorregião"] == "Metropolitana de Belo Horizonte", "Municípios com vitória de Lula"].iloc[0]) == 1
    assert int(resumo.loc[resumo["Mesorregião"] == "Triângulo Mineiro / Alto Paranaíba", "Municípios com vitória de Lula"].iloc[0]) == 0


def test_cruza_por_nome_quando_falta_codigo(tmp_path: Path):
    cat = pd.DataFrame(
        {
            "codigo_ibge": [3145500],
            "codigo_tse": [99999],
            "nome_municipio": ["Olhos-d'Água"],
            "codigo_meso": [2],
            "nome_meso": ["Norte de Minas"],
            "codigo_micro": [9],
            "nome_micro": ["Bocaiúva"],
            "chave_nome": [chave_nome("Olhos-d'Água")],
        }
    )
    mun = enriquecer_municipios(pd.DataFrame([_mun(1, "OLHOS D'ÁGUA", 70, 30)]))
    cruzado = cruzar_mesorregiao(mun, cat)
    assert cruzado.iloc[0]["nome_meso"] == "Norte de Minas"


def test_le_discriminativo_municipal(tmp_path: Path):
    csv = tmp_path / "disc.csv"
    pd.DataFrame(
        [
            {
                "SG_UF": "MG",
                "CD_MUNICIPIO": 41238,
                "NM_MUNICIPIO": "BELO HORIZONTE",
                "QT_SECOES_2022": 10,
                "QT_VOTOS_PT_2022": 80,
                "QT_VOTOS_OPP_2022": 20,
                "QT_VOTOS_VALIDOS_2022": 100,
            },
            {
                "SG_UF": "SP",
                "CD_MUNICIPIO": 71072,
                "NM_MUNICIPIO": "SAO PAULO",
                "QT_SECOES_2022": 99,
                "QT_VOTOS_PT_2022": 10,
                "QT_VOTOS_OPP_2022": 90,
                "QT_VOTOS_VALIDOS_2022": 100,
            },
        ]
    ).to_csv(csv, index=False)
    mg = municipios_do_discriminativo(csv)
    assert list(mg["NM_MUNICIPIO"]) == ["BELO HORIZONTE"]
    assert mg.iloc[0]["VENCEDOR"] == "Lula"
    assert mg.iloc[0]["DIF_PP"] == 60.0
