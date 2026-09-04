"""Testes do consolidador histórico de Boletins de Urna (2014/2018/2022)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd

from scripts.baixar_boletins_urna import (
    CANDIDATOS,
    carregar_catalogo,
    consolidar_urnas,
    filtrar_presidente,
    ler_bweb_2014,
    nome_arquivo_saida,
    normalizar_colunas,
    processar_zip_bweb,
    processar_zip_bweb_2014,
    recurso_catalogo,
    urls_espelho_historico,
)
from scripts.baixar_boletins_urna_2022 import carregar_faixas_modelo

FIXTURES = Path(__file__).parent / "fixtures"
FAIXAS = FIXTURES / "tse2022" / "modelourna_numerointerno.csv"
BU = FIXTURES / "tse_bu"


def test_catalogo_tem_168_zips_oficiais():
    cat = carregar_catalogo()
    assert len(cat) == 168
    rec = recurso_catalogo(cat, 2022, 1, "RR")
    assert rec["arquivo"] == "bweb_1t_RR_051020221321.zip"
    rec18 = recurso_catalogo(cat, 2018, 2, "SP")
    assert rec18["arquivo"].startswith("BWEB_2t_SP_")
    rec14 = recurso_catalogo(cat, 2014, 1, "AC")
    assert rec14["url"].endswith("bweb_1t_AC_14102014131600.zip")
    rec_zz = recurso_catalogo(cat, 2018, 1, "ZZ")
    assert "111020181508" in rec_zz["arquivo"]


def test_urls_historicas_incluem_varios_wayback():
    url = "https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb/x.zip"
    espelhos = urls_espelho_historico(url)
    assert espelhos[0] == url
    assert any("20221105id_" in u for u in espelhos)
    assert any("202409id_" in u for u in espelhos)
    assert any("2023id_" in u for u in espelhos)


def test_filtra_presidente_primeiro_turno():
    df = pd.read_csv(BU / "bweb_1t_XX.csv", sep=";", dtype=str)
    filtrado = filtrar_presidente(df, 1)
    assert set(filtrado["NR_TURNO"].astype(int)) == {1}
    assert set(filtrado["CD_CARGO_PERGUNTA"].astype(int)) == {1}
    assert len(filtrado) == 9


def test_2022_1t_uma_linha_por_urna():
    df = pd.read_csv(BU / "bweb_1t_XX.csv", sep=";", dtype=str)
    faixas = carregar_faixas_modelo(FAIXAS)
    out = consolidar_urnas(df, faixas, ano=2022, turno=1)
    assert len(out) == 2
    assert int(out["QT_VOTOS_LULA"].sum()) == 110
    assert int(out["QT_VOTOS_BOLSONARO"].sum()) == 200
    assert int(out["QT_VOTOS_TEBET"].sum()) == 8
    assert int(out["QT_VOTOS_CIRO"].sum()) == 5
    ue2015 = out.loc[out["DS_MODELO_URNA"] == "UE2015"].iloc[0]
    assert int(ue2015["QT_VOTOS_VALIDOS"]) == 70 + 90 + 8


def test_2018_normaliza_coluna_sg_uf_com_espaco():
    df = pd.read_csv(BU / "bweb_1t_2018_XX.csv", sep=";", dtype=str)
    norm = normalizar_colunas(df)
    assert "SG_UF" in norm.columns
    faixas = carregar_faixas_modelo(FAIXAS)
    out = consolidar_urnas(norm, faixas, ano=2018, turno=1)
    assert len(out) == 1
    assert int(out["QT_VOTOS_BOLSONARO"].iloc[0]) == 100
    assert int(out["QT_VOTOS_HADDAD"].iloc[0]) == 50
    assert int(out["NR_URNA_EFETIVADA"].iloc[0]) == 1_180_872
    assert out["DS_MODELO_URNA"].iloc[0] == "UE2009"


def test_2014_txt_posicional_com_serie():
    df = ler_bweb_2014(BU / "bweb_1t_2014_XX.txt")
    assert df.loc[0, "SG_UF"] == "RR"
    assert df.loc[0, "NR_URNA_EFETIVADA"] == "1054014"
    faixas = carregar_faixas_modelo(FAIXAS)
    out = consolidar_urnas(df, faixas, ano=2014, turno=1)
    assert len(out) == 1
    assert int(out["QT_VOTOS_DILMA"].iloc[0]) == 80
    assert int(out["QT_VOTOS_AECIO"].iloc[0]) == 60
    assert int(out["QT_VOTOS_MARINA"].iloc[0]) == 20
    assert int(out["QT_VOTOS_BRANCO"].iloc[0]) == 5
    assert int(out["NR_URNA_EFETIVADA"].iloc[0]) == 1_054_014
    assert out["DS_MODELO_URNA"].iloc[0] == "UE2009"


def test_processar_zip_2014_em_lotes(tmp_path: Path):
    zip_path = tmp_path / "bweb_1t_RR_14102014140241.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(BU / "bweb_1t_2014_XX.txt", arcname="bweb_1t_RR_14102014140241.txt")
    faixas = carregar_faixas_modelo(FAIXAS)
    tabela = processar_zip_bweb(zip_path, faixas, ano=2014, turno=1)
    assert len(tabela) == 1
    assert int(tabela["QT_VOTOS_DILMA"].iloc[0]) == 80
    assert int(tabela["NR_URNA_EFETIVADA"].iloc[0]) == 1_054_014
    lotes = processar_zip_bweb_2014(zip_path, faixas, turno=1, chunk=1)
    assert list(lotes.columns) == list(tabela.columns)
    assert int(lotes["QT_VOTOS_DILMA"].iloc[0]) == 80
    assert int(lotes["QT_VOTOS_AECIO"].iloc[0]) == 60
    txt = tmp_path / "bweb_1t_RR_14102014140241.txt"
    txt.write_bytes((BU / "bweb_1t_2014_XX.txt").read_bytes())
    direto = processar_zip_bweb(txt, faixas, ano=2014, turno=1)
    assert int(direto["QT_VOTOS_DILMA"].iloc[0]) == 80


def test_processar_zip_1t(tmp_path: Path):
    zip_path = tmp_path / "bweb_1t_RR_051020221321.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(BU / "bweb_1t_XX.csv", arcname="bweb_1t_RR_051020221321.csv")
    faixas = carregar_faixas_modelo(FAIXAS)
    tabela = processar_zip_bweb(zip_path, faixas, ano=2022, turno=1)
    assert len(tabela) == 2
    assert int(tabela["QT_VOTOS_LULA"].sum()) == 110


def test_nome_saida_e_candidatos():
    assert nome_arquivo_saida(2022, 1) == "urnas_1t_presidente.csv"
    assert 17 in CANDIDATOS[(2018, 1)]
    assert CANDIDATOS[(2018, 1)][17] == "BOLSONARO"
    assert CANDIDATOS[(2018, 1)][30] == "AMOEDO"
    assert CANDIDATOS[(2018, 1)][54] == "JOAO_GOULART"
    assert CANDIDATOS[(2014, 2)][13] == "DILMA"
