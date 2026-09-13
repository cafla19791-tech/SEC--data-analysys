"""Testes do coletor de Portarias SEFIC/MinC da Lei Rouanet no DOU."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

from scripts.portarias_lei_rouanet import (
    classificar_tipo,
    corrigir_frames,
    eh_portaria_sefic,
    extrair_params_leitura,
    html_para_texto,
    numero_portaria,
    parse_artigo,
    parse_data_portaria,
    portarias_para_frame,
    processar,
    valor_referencia_projeto,
    _num,
)

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "dou"


def _html(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


def test_numero_e_data_portaria():
    t = "PORTARIA SEFIC/MINC Nº 502, DE 7 DE AGOSTO DE 2026"
    assert numero_portaria(t) == 502
    assert parse_data_portaria(t) == date(2026, 8, 7)
    assert parse_data_portaria("PORTARIA SEFIC/MINC Nº 189, DE 1º DE ABRIL DE 2026") == date(2026, 4, 1)
    assert numero_portaria("PORTARIA SEFIC/MINC N.º 143, DE 17 DE MARÇO DE 2026") == 143
    assert numero_portaria("PORTARIA SEFIC/MINC N° 25, DE 15 DE JANEIRO DE 2026") == 25
    assert numero_portaria("", "portaria-sefic/minc-n.-165-de-24-de-marco-de-2026-695073580") == 165


def test_num_ptbr():
    assert _num("R$ 7.779.565,50") == 7779565.50
    assert _num("41.888,68") == 41888.68
    assert _num("") is None


def test_classificar_tipos():
    assert (
        classificar_tipo(
            "Art. 1.º - Homologar os projetos culturais relacionados nos anexos desta portaria, "
            "que após terem atendido aos requisitos de admissibilidade estabelecidos pela Lei nº 8.313/91, "
            "passam a fase de obtenção de doações e patrocínios. Art. 2.º"
        )
        == "homologacao_captacao"
    )
    assert (
        classificar_tipo(
            "Art. 1.º - Homologar a prorrogação do prazo de captação de recursos do(s) projeto(s) cultural(is)."
        )
        == "prorrogacao_prazo"
    )
    assert (
        classificar_tipo(
            "Art. 1.º - Homologar o valor complementado de valor em favor do(s) projeto(s) cultural(is)."
        )
        == "complementacao_valor"
    )
    assert (
        classificar_tipo("Art. 1.º - Homologar a redução de valor em favor do(s) projeto(s) cultural(is).")
        == "reducao_valor"
    )
    assert (
        classificar_tipo("Art. 1.º - Homologar a alteração dos projetos culturais relacionados nos anexos.")
        == "alteracao_projeto"
    )
    assert (
        classificar_tipo(
            "Art. 1.º - Homologar a(s) alteração(ões) do(s) resumo(s) do(s) projeto(s) abaixo relacionado(s):"
        )
        == "alteracao_projeto"
    )


def test_indice_filtra_sefic():
    params = extrair_params_leitura(_html("leiturajornal_do1.html"))
    itens = [it for it in params["jsonArray"] if eh_portaria_sefic(it)]
    assert len(itens) == 1
    assert itens[0]["urlTitle"].startswith("portaria-sefic/minc-n-502")


def test_parse_homologacao_502():
    html = _html("portaria_502_homologacao.html")
    parsed = parse_artigo(
        html,
        {
            "title": "PORTARIA SEFIC/MINC Nº 502, DE 7 DE AGOSTO DE 2026",
            "urlTitle": "portaria-sefic/minc-n-502-de-7-de-agosto-de-2026-724197437",
            "pubDate": "10/08/2026",
        },
    )
    assert parsed["tipo"] == "homologacao_captacao"
    assert parsed["libera_captacao_inicial"] is True
    assert parsed["numero"] == 502
    assert parsed["qtd_projetos"] == 3
    pronacs = [p["pronac"] for p in parsed["projetos"]]
    assert pronacs == ["265857", "265860", "265863"]
    ary = parsed["projetos"][0]
    assert "ARY FONTOURA" in ary["nome_projeto"].upper()
    assert ary["valor_aprovado"] == 7779565.50
    assert ary["uf"] == "RJ"
    assert ary["prazo_inicio"] == "10/08/2026"
    assert ary["prazo_fim"] == "31/12/2026"
    assert parsed["url"].endswith("724197437")


def test_parse_prorrogacao_e_complementacao():
    pr = parse_artigo(_html("portaria_431_prorrogacao.html"), {"urlTitle": "x-431"})
    assert pr["tipo"] == "prorrogacao_prazo"
    assert pr["libera_captacao"] is True
    assert pr["libera_captacao_inicial"] is False
    assert pr["projetos"][0]["pronac"] == "2314536"

    cp = parse_artigo(_html("portaria_592_complementacao.html"), {"urlTitle": "x-592"})
    assert cp["tipo"] == "complementacao_valor"
    p0 = cp["projetos"][0]
    assert p0["valor_complementado"] == 41888.68
    assert p0["valor_total_atual"] == 536807.63
    assert p0["cidade"] == "Mariana"


def test_parse_alteracao():
    al = parse_artigo(_html("portaria_504_alteracao.html"), {"urlTitle": "x-504"})
    assert al["tipo"] == "alteracao_projeto"
    assert al["libera_captacao_inicial"] is False
    assert al["projetos"][0]["pronac"] == "256577"


def test_html_para_texto_contem_artigo():
    texto = html_para_texto(_html("portaria_502_homologacao.html"))
    assert "Lei nº 8.313/91" in texto
    assert "DJONGA E SEU BLOCO" in texto


def test_processar_offline(tmp_path: Path):
    html_idx = _html("leiturajornal_do1.html")
    html_art = _html("portaria_502_homologacao.html")
    params = extrair_params_leitura(html_idx)
    url = params["jsonArray"][0]["urlTitle"]
    portarias, projetos = processar(
        tmp_path,
        inicio=date(2026, 8, 1),
        fim=date(2026, 8, 10),
        indice_html=html_idx,
        artigo_html_map={url: html_art},
    )
    assert len(portarias) == 1
    assert int(portarias.iloc[0]["numero"]) == 502
    assert len(projetos) == 3
    assert (tmp_path / "portarias_lei_rouanet.md").exists()
    assert (tmp_path / "portarias_lei_rouanet.csv").exists()
    assert (tmp_path / "projetos_lei_rouanet_captacao.csv").exists()
    md = (tmp_path / "portarias_lei_rouanet.md").read_text(encoding="utf-8")
    assert "Homologações que liberam a captação" in md
    assert "502" in md


def test_valor_referencia_por_tipo():
    assert valor_referencia_projeto({"valor_reduzido": 10, "valor_total_atual": 100}, "reducao_valor") == 10
    assert valor_referencia_projeto({"valor_complementado": 5, "valor_total_atual": 50}, "complementacao_valor") == 5
    assert valor_referencia_projeto({"valor_aprovado": 20, "valor_total_atual": 1}, "homologacao_captacao") == 20


def test_corrigir_frames_numero_e_dedup():
    port = pd.DataFrame(
        [
            {
                "titulo": "PORTARIA SEFIC/MINC N.º 143, DE 17 DE MARÇO DE 2026",
                "numero": None,
                "tipo": "homologacao_captacao",
                "texto_art1": "Homologar os projetos culturais relacionados nos anexos.",
                "data_portaria": "2026-03-17",
                "data_publicacao": "2026-03-18",
                "qtd_projetos": 1,
                "url": "https://www.in.gov.br/web/dou/-/a",
                "url_title": "portaria-sefic/minc-n.-143-de-17-de-marco-de-2026-1",
            },
            {
                "titulo": "PORTARIA SEFIC/MINC Nº 18, DE 14 DE JANEIRO DE 2026",
                "numero": 18,
                "tipo": "homologacao_captacao",
                "texto_art1": "Homologar os projetos culturais.",
                "data_portaria": "2026-01-14",
                "data_publicacao": "2026-01-15",
                "qtd_projetos": 21,
                "url": "https://www.in.gov.br/web/dou/-/b1",
                "url_title": "portaria-sefic/minc-n-18-de-14-de-janeiro-de-2026-1",
            },
            {
                "titulo": "PORTARIA SEFIC/MINC Nº 18, DE 14 DE JANEIRO DE 2026",
                "numero": 18,
                "tipo": "homologacao_captacao",
                "texto_art1": "Homologar os projetos culturais.",
                "data_portaria": "2026-01-14",
                "data_publicacao": "2026-01-15",
                "qtd_projetos": 21,
                "url": "https://www.in.gov.br/web/dou/-/b2",
                "url_title": "portaria-sefic/minc-n-18-de-14-de-janeiro-de-2026-2",
            },
        ]
    )
    proj = pd.DataFrame(
        [
            {"url": "https://www.in.gov.br/web/dou/-/b1", "tipo_portaria": "homologacao_captacao", "valor_aprovado": 10},
            {"url": "https://www.in.gov.br/web/dou/-/b2", "tipo_portaria": "homologacao_captacao", "valor_aprovado": 10},
        ]
    )
    p2, j2 = corrigir_frames(port, proj)
    assert int(p2.loc[p2.titulo.str.contains("143"), "numero"].iloc[0]) == 143
    assert len(p2[p2.numero == 18]) == 1
    assert len(j2) == 1


def test_portarias_para_frame_projeta_tipo():
    recs = [
        {
            "numero": 1,
            "tipo": "homologacao_captacao",
            "qtd_projetos": 1,
            "projetos": [{"pronac": "1", "nome_projeto": "X"}],
        }
    ]
    port, proj = portarias_para_frame(recs)
    assert list(port["numero"]) == [1]
    assert proj.iloc[0]["tipo_portaria"] == "homologacao_captacao"
    assert "projetos" not in port.columns
