"""Testes do consolidador de Boletins de Urna 2022 (série + modelo)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import pytest

from scripts.baixar_boletins_urna_2022 import (
    UFS,
    _pasta_raw_de_massa,
    _pasta_saida_contagil,
    _parece_winpython,
    carregar_faixas_modelo,
    classificar_modelo,
    consolidar_urnas,
    descobrir_winpython,
    filtrar_presidente_2t,
    ler_csv_tse,
    normalizar_ufs,
    parse_args,
    pastas_padrao,
    processar_zip_bweb,
    resolver_pastas,
    resumo_por_modelo,
    rotulo_modelo,
    url_bweb,
)

FIXTURES = Path(__file__).parent / "fixtures" / "tse2022"


def test_28_ufs_incluindo_exterior():
    assert len(UFS) == 28
    assert UFS[0] == "AC"
    assert UFS[-1] == "ZZ"
    assert "SP" in UFS and "DF" in UFS


def test_url_oficial_por_uf():
    url = url_bweb("rr")
    assert url.endswith("bweb_2t_RR_311020221535.zip")
    assert "eleicoes2022/buweb" in url


def test_normalizar_ufs_rejeita_invalida():
    with pytest.raises(ValueError, match="XX"):
        normalizar_ufs(["RR", "XX"])
    assert normalizar_ufs(["rr", "AC", "ac"]) == ["RR", "AC"]


def test_faixas_oficiais_e_modelo_por_serie():
    faixas = carregar_faixas_modelo(FIXTURES / "modelourna_numerointerno.csv")
    assert set(faixas["NR_MODELO"]) == {2009, 2010, 2011, 2013, 2015, 2020}
    nums = pd.Series([1_800_000, 2_100_000, 1_000_000, 50])
    modelos = classificar_modelo(nums, faixas)
    assert list(modelos.astype("object").where(modelos.notna(), None)) == [
        2015,
        2020,
        2009,
        None,
    ]
    assert rotulo_modelo(2015) == "UE2015"
    assert rotulo_modelo(None) == "sem_faixa"


def test_faixas_embutidas_iguais_ao_csv_oficial():
    oficiais = carregar_faixas_modelo(FIXTURES / "modelourna_numerointerno.csv")
    fallback = carregar_faixas_modelo(None)
    pd.testing.assert_frame_equal(oficiais.reset_index(drop=True), fallback)


def test_filtra_so_presidente_segundo_turno():
    df = ler_csv_tse(FIXTURES / "bweb_2t_XX.csv")
    filtrado = filtrar_presidente_2t(df)
    assert set(filtrado["NR_TURNO"].astype(int)) == {2}
    assert set(filtrado["CD_CARGO_PERGUNTA"].astype(int)) == {1}
    assert len(filtrado) == 8


def test_uma_linha_por_urna_com_serie_modelo_e_votos():
    df = ler_csv_tse(FIXTURES / "bweb_2t_XX.csv")
    faixas = carregar_faixas_modelo(FIXTURES / "modelourna_numerointerno.csv")
    out = consolidar_urnas(df, faixas)

    assert len(out) == 2
    assert set(out["NR_URNA_EFETIVADA"]) == {1_800_000, 2_100_000}
    assert set(out["DS_MODELO_URNA"]) == {"UE2015", "UE2020"}

    ue2015 = out.loc[out["DS_MODELO_URNA"] == "UE2015"].iloc[0]
    assert int(ue2015["QT_VOTOS_LULA"]) == 80
    assert int(ue2015["QT_VOTOS_BOLSONARO"]) == 100
    assert int(ue2015["QT_VOTOS_BRANCO"]) == 5
    assert int(ue2015["QT_VOTOS_NULO"]) == 5
    assert int(ue2015["QT_VOTOS_VALIDOS"]) == 180

    ue2020 = out.loc[out["DS_MODELO_URNA"] == "UE2020"].iloc[0]
    assert int(ue2020["QT_VOTOS_LULA"]) == 40
    assert int(ue2020["QT_VOTOS_BOLSONARO"]) == 120


def test_processar_zip_oficial_e_resumos(tmp_path: Path):
    csv_path = FIXTURES / "bweb_2t_XX.csv"
    zip_path = tmp_path / "bweb_2t_RR_311020221535.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(csv_path, arcname="bweb_2t_RR_311020221535.csv")

    faixas = carregar_faixas_modelo(FIXTURES / "modelourna_numerointerno.csv")
    tabela = processar_zip_bweb(zip_path, faixas)
    por_modelo = resumo_por_modelo(tabela)

    assert len(tabela) == 2
    assert int(tabela["QT_VOTOS_LULA"].sum()) == 120
    assert int(tabela["QT_VOTOS_BOLSONARO"].sum()) == 220
    assert set(por_modelo["DS_MODELO_URNA"]) == {"UE2015", "UE2020"}


def test_pastas_contagil_massa_e_saida(tmp_path: Path):
    winpy = tmp_path / "ContAgilAppBeta64" / "python_jep" / "winpython"
    winpy.mkdir(parents=True)
    (winpy / "python.exe").write_bytes(b"")
    assert _parece_winpython(winpy)
    assert descobrir_winpython(winpy) == winpy.resolve()

    raw, saida = pastas_padrao(winpy)
    assert raw == winpy / "dados" / "tse2022" / "raw"
    assert saida == winpy / "saida" / "tse2022"

    args = parse_args(
        [
            "--massa-dados",
            str(winpy / "dados"),
            "--pasta-saida",
            str(winpy / "saida"),
        ]
    )
    raw2, saida2 = resolver_pastas(args)
    assert raw2 == winpy / "dados" / "tse2022" / "raw"
    assert saida2 == winpy / "saida" / "tse2022"
    assert _pasta_raw_de_massa(winpy / "dados") == winpy / "dados" / "tse2022" / "raw"
    assert _pasta_saida_contagil(winpy / "saida") == winpy / "saida" / "tse2022"


def test_entrypoint_contagil_carrega_scripts(tmp_path: Path, monkeypatch):
    import baixar_boletins_urna_2022 as entry

    assert callable(entry._load_main())
