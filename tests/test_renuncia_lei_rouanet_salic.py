"""Testes do cruzamento DOU × SALIC e da estimativa de renúncia Rouanet."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.renuncia_lei_rouanet_salic import (
    artigo_norma,
    cruzar,
    norm_pronac,
    processar,
    projeto_do_payload,
    projetos_dou_unicos,
    slim_projeto,
)


def test_norm_pronac():
    assert norm_pronac("265857") == "265857"
    assert norm_pronac(265857.0) == "265857"
    assert norm_pronac("2314536") == "2314536"
    assert norm_pronac(None) is None


def test_artigo_norma():
    assert artigo_norma("Artigo 18 , § 1º") == "artigo_18"
    assert artigo_norma("Artigo 26") == "artigo_26"
    assert artigo_norma("Mecenato") == "indefinido"


def test_slim_e_payload_lista():
    payload = {
        "_embedded": {
            "projetos": [
                {
                    "PRONAC": "265857",
                    "nome": "ARY FONTOURA",
                    "situacao": "Autorizada a captação total dos recursos",
                    "valor_aprovado": 7779565.5,
                    "valor_captado": 0,
                    "enquadradmento": "Artigo 18",
                    "mecanisnmo": "Mecenato",
                    "UF": "RJ",
                }
            ]
        },
        "total": 1,
    }
    p = projeto_do_payload(payload)
    assert p["PRONAC"] == "265857"
    assert p["valor_captado"] == 0.0
    assert p["enquadramento_salic"] == "Artigo 18"
    assert p["mecanismo"] == "Mecenato"


def test_projetos_dou_unicos_fica_com_ultimo():
    df = pd.DataFrame(
        [
            {
                "pronac": "1",
                "tipo_portaria": "homologacao_captacao",
                "data_publicacao": "2026-01-01",
                "valor_referencia": 10,
                "artigo_enquadramento": "Artigo 18",
                "nome_projeto": "A",
            },
            {
                "pronac": "1",
                "tipo_portaria": "homologacao_captacao",
                "data_publicacao": "2026-02-01",
                "valor_referencia": 20,
                "artigo_enquadramento": "Artigo 18",
                "nome_projeto": "A2",
            },
            {
                "pronac": "2",
                "tipo_portaria": "prorrogacao_prazo",
                "data_publicacao": "2026-01-01",
                "valor_referencia": 99,
                "artigo_enquadramento": "Artigo 18",
                "nome_projeto": "B",
            },
        ]
    )
    out = projetos_dou_unicos(df, somente_homologacao=True)
    assert list(out["pronac"]) == ["1"]
    assert float(out.iloc[0]["teto_dou"]) == 20.0


def test_cruzar_renuncia_art18_e_art26():
    dou = pd.DataFrame(
        [
            {"pronac": "10", "teto_dou": 1000, "artigo": "artigo_18", "nome_projeto": "X", "uf": "SP"},
            {"pronac": "20", "teto_dou": 1000, "artigo": "artigo_26", "nome_projeto": "Y", "uf": "RJ"},
        ]
    )
    salic = {
        "10": {"valor_captado": 400, "valor_aprovado": 1000, "situacao": "ok", "enquadramento_salic": "Artigo 18"},
        "20": {"valor_captado": 400, "valor_aprovado": 1000, "situacao": "ok", "enquadramento_salic": "Artigo 26"},
    }
    out = cruzar(dou, salic)
    a18 = out[out.pronac == "10"].iloc[0]
    a26 = out[out.pronac == "20"].iloc[0]
    assert a18["renuncia_estimada"] == 400
    assert a18["taxa_captacao"] == 0.4
    assert a26["renuncia_estimada"] == 120  # 30%
    assert a26["renuncia_art26_doacao_pj"] == 160  # 40%


def test_teto_cai_no_aprovado_salic_se_dou_zerado():
    dou = pd.DataFrame([{"pronac": "1", "teto_dou": 0, "artigo": "artigo_18", "nome_projeto": "Z", "uf": "SP"}])
    salic = {"1": {"valor_captado": 50, "valor_aprovado": 200}}
    out = cruzar(dou, salic).iloc[0]
    assert out["teto_dou"] == 200
    assert out["taxa_captacao"] == 0.25


def test_processar_offline(tmp_path: Path):
    entrada = tmp_path / "dou.csv"
    pd.DataFrame(
        [
            {
                "pronac": "265857",
                "tipo_portaria": "homologacao_captacao",
                "data_publicacao": "2026-08-10",
                "valor_aprovado": 7779565.5,
                "valor_referencia": 7779565.5,
                "artigo_enquadramento": "Artigo 18 , § 1º",
                "nome_projeto": "ARY",
                "uf": "RJ",
                "proponente": "CINE",
                "numero_portaria": 502,
                "url": "https://www.in.gov.br/x",
            }
        ]
    ).to_csv(entrada, index=False)
    salic = {
        "265857": slim_projeto(
            {
                "PRONAC": "265857",
                "nome": "ARY",
                "situacao": "Autorizada a captação total dos recursos",
                "valor_aprovado": 7779565.5,
                "valor_captado": 0,
                "enquadradmento": "Artigo 18",
                "UF": "RJ",
            }
        )
    }
    out = processar(
        entrada,
        tmp_path,
        cache_path=tmp_path / "cache.jsonl",
        salic_map=salic,
    )
    assert len(out) == 1
    assert float(out.iloc[0]["valor_captado_salic"]) == 0.0
    assert (tmp_path / "renuncia_lei_rouanet.md").exists()
    md = (tmp_path / "renuncia_lei_rouanet.md").read_text(encoding="utf-8")
    assert "Renúncia estimada" in md
    assert "art. 18" in md
