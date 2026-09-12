"""Testes do consolidador de Boletins de Urna 2022 (série + modelo)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import pytest

import requests

from scripts.baixar_boletins_urna_2022 import (
    UFS,
    _pasta_raw_de_massa,
    _pasta_saida_contagil,
    _parece_winpython,
    baixar_arquivo,
    carregar_faixas_modelo,
    classificar_modelo,
    consolidar_urnas,
    descobrir_winpython,
    e_erro_ssl,
    detectar_proxy_windows,
    escrever_pagina_links,
    escrever_script_curl,
    filtrar_presidente_2t,
    ler_csv_tse,
    montar_comando_curl,
    normalizar_ufs,
    parse_args,
    pastas_padrao,
    preferir_curl,
    processar_zip_bweb,
    resolver_pastas,
    resumo_por_modelo,
    resumir_erro_download,
    rotulo_modelo,
    url_bweb,
    urls_espelho,
    user_agent_para,
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
    espelhos = urls_espelho(url)
    assert espelhos[0] == url
    assert any("/20221108000702id_/" in u for u in espelhos)
    assert any("/2023id_/" in u for u in espelhos)
    assert espelhos[-1].endswith(url)


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


def test_pagina_links_tem_28_ufs(tmp_path: Path):
    html = escrever_pagina_links(tmp_path / "links.html")
    texto = html.read_text(encoding="utf-8")
    assert "bweb_2t_SP_311020221535.zip" in texto
    assert "web.archive.org" in texto
    assert "--somente-resultado-github" in texto
    assert texto.count("<li><b>") == 28


def test_ssl_e_curl_windows():
    assert e_erro_ssl(requests.exceptions.SSLError("sslv3 alert handshake failure"))
    assert "TLS" in resumir_erro_download(
        requests.exceptions.SSLError("SSLV3_ALERT_HANDSHAKE_FAILURE")
    )
    assert preferir_curl(
        "https://web.archive.org/web/2023id_/https://cdn.tse.jus.br/x.zip",
        plataforma="win32",
    )
    assert not preferir_curl("https://cdn.tse.jus.br/x.zip", plataforma="win32")
    assert not preferir_curl(
        "https://web.archive.org/web/2023id_/https://cdn.tse.jus.br/x.zip",
        plataforma="linux",
    )

    dest = Path("ac.zip")
    assert user_agent_para("https://web.archive.org/web/x").startswith("ContAgil")
    assert "Mozilla" in user_agent_para("https://cdn.tse.jus.br/x.zip")
    win = montar_comando_curl(
        "https://web.archive.org/web/x.zip", dest, timeout=120, curl="curl.exe", plataforma="win32"
    )
    assert "--ssl-no-revoke" in win
    assert "--tls-max" in win
    assert "1.2" in win
    assert "-k" not in win
    assert "ContAgil-TSE-BU/1.0" in win
    assert win[-1] == "https://web.archive.org/web/x.zip"
    insecure = montar_comando_curl(
        "https://example.test/a.zip",
        dest,
        timeout=120,
        curl="curl.exe",
        plataforma="win32",
        insecure=True,
    )
    assert "-k" in insecure
    linux = montar_comando_curl(
        "https://example.test/a.zip", dest, timeout=90, curl="curl", plataforma="linux"
    )
    assert "--ssl-no-revoke" not in linux


def test_baixar_arquivo_usa_curl_apos_ssl(tmp_path: Path, monkeypatch):
    dest = tmp_path / "bweb.zip"
    chamadas_get: list[str] = []

    class BoomSession:
        def get(self, url, **_kwargs):
            chamadas_get.append(url)
            raise requests.exceptions.SSLError(
                "SSLV3_ALERT_HANDSHAKE_FAILURE sslv3 alert handshake failure"
            )

    def fake_curl(url, destino, **_kwargs):
        destino.write_bytes(b"PK" + b"\x00" * 200)
        return destino

    monkeypatch.setattr(
        "scripts.baixar_boletins_urna_2022.baixar_com_curl", fake_curl
    )
    out = baixar_arquivo(
        "https://cdn.tse.jus.br/estatistica/sead/x.zip",
        dest,
        timeout=10,
        tentativas=4,
        session=BoomSession(),
    )
    assert out == dest
    assert dest.stat().st_size > 100
    assert len(chamadas_get) == 1


def test_script_curl_lista_28_ufs(tmp_path: Path):
    bat = escrever_script_curl(tmp_path / "baixar_zips_urna_curl.bat")
    texto = bat.read_text(encoding="utf-8")
    assert "curl.exe" in texto
    assert "--ssl-no-revoke" in texto
    assert "--tls-max" in texto
    assert "ContAgil-TSE-BU/1.0" in texto
    assert "20221108000702id_" in texto
    assert texto.count("bweb_2t_") >= 28
    assert "bweb_2t_ZZ_311020221535.zip" in texto
    assert "--somente-processar" in texto


def test_padrao_workers_um_e_usar_curl():
    args = parse_args([])
    assert args.workers == 1
    assert args.tentativas == 2
    assert args.usar_curl is False
    assert parse_args(["--usar-curl"]).usar_curl is True
    assert parse_args(["--somente-resultado-github"]).somente_resultado_github is True


def test_urls_github_sem_archive_e_proxy_env(monkeypatch):
    url = "https://raw.githubusercontent.com/org/repo/main/x.csv.gz"
    assert urls_espelho(url) == [url]
    monkeypatch.setenv("HTTPS_PROXY", "http://proxy.rfb:8080")
    assert detectar_proxy_windows() == "http://proxy.rfb:8080"


def test_entrypoint_contagil_carrega_scripts(tmp_path: Path, monkeypatch):
    import baixar_boletins_urna_2022 as entry

    assert callable(entry._load_main())
