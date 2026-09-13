"""Testes do extrato de captações da Lei Rouanet (SALIC)."""

from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from scripts.captacoes_lei_rouanet_salic import (
    agregar_por_ano,
    contar_com_captacao,
    extrair_captacoes_do_detalhe,
    listar_projetos_com_captacao,
    no_periodo,
    norm_cnpj_cpf,
    norm_data,
    norm_pronac,
    processar,
    recibos_periodo,
    slim_projeto_lista,
    slim_recibo,
)

FIXTURE = Path(__file__).parent / "fixtures" / "salic" / "projeto_captacoes.json"


def test_norm_pronac_e_documento():
    assert norm_pronac("261004 ") == "261004"
    assert norm_pronac(261004.0) == "261004"
    assert norm_cnpj_cpf("00.000.000/0001-91") == "00000000000191"
    assert norm_cnpj_cpf("12345678901") == "12345678901"
    assert norm_cnpj_cpf("***919425**") == "***919425**"
    assert norm_data("2026-03-30 00:00:00") == "2026-03-30"
    assert no_periodo("2003-01-01", "2003-01-01", "2026-12-31")
    assert not no_periodo("2002-12-31", "2003-01-01", "2026-12-31")
    assert not no_periodo(None, "2003-01-01", "2026-12-31")


def test_slim_projeto_lista():
    p = slim_projeto_lista(
        {
            "PRONAC": "203525",
            "nome": " Inhotim ",
            "cgccpf": "05422243000131",
            "proponente": "INSTITUTO INHOTIM",
            "valor_captado": "147548414.29",
            "justificativa": "texto longo ignorado",
        }
    )
    assert p["PRONAC"] == "203525"
    assert p["nome"] == "Inhotim"
    assert p["valor_captado"] == 147548414.29
    assert "justificativa" not in p


def test_extrair_captacoes_e_filtro_periodo():
    payload = json.loads(FIXTURE.read_text(encoding="utf-8"))
    rec = extrair_captacoes_do_detalhe(payload)
    assert rec["PRONAC"] == "261004"
    assert rec["cgccpf"] == "00584628000181"
    assert rec["proponente"].startswith("ARTE PRODUCOES")
    assert len(rec["captacoes"]) == 3
    assert rec["captacoes"][0]["PRONAC"] == "261004"
    assert rec["captacoes"][0]["data_recibo"] == "2026-03-30"
    assert rec["captacoes"][0]["cgccpf_doador"] == "06057223000171"

    df = recibos_periodo({"261004": rec}, inicio="2003-01-01", fim="2026-12-31")
    assert len(df) == 2
    assert set(df["data_captacao"]) == {"2003-01-01", "2026-03-30"}
    assert df["cnpj_cpf_beneficiario"].nunique() == 1
    assert "SENDAS" in df["nome_incentivador"].to_string()
    assert float(df["valor_captado"].sum()) == 490000.0
    por_ano = agregar_por_ano(df)
    assert list(por_ano["ano"]) == ["2003", "2026"]


def test_slim_recibo():
    r = slim_recibo(
        {
            "PRONAC": "950962",
            "valor": 200000,
            "data_recibo": "1996-11-29",
            "nome_projeto": "Via Brasil",
            "cgccpf": "00000000000191",
            "nome_doador": "BANCO DO BRASIL SA",
        }
    )
    assert r["data_recibo"] == "1996-11-29"
    assert r["valor"] == 200000.0


def test_contar_com_captacao_binario():
    def get(url: str) -> dict:
        q = parse_qs(urlparse(url).query)
        offset = int(q.get("offset", ["0"])[0])
        valor = 100 if offset < 5 else 0
        return {
            "total": 10,
            "_embedded": {"projetos": [{"PRONAC": str(1000 + offset), "valor_captado": valor}]},
        }

    n, total = contar_com_captacao(get)
    assert total == 10
    assert n == 5


def test_listar_projetos_com_captacao_para_no_zero():
    def get(url: str) -> dict:
        q = parse_qs(urlparse(url).query)
        offset = int(q.get("offset", ["0"])[0])
        limit = int(q.get("limit", ["100"])[0])
        if limit == 1:
            return {
                "total": 150,
                "_embedded": {
                    "projetos": [
                        {
                            "PRONAC": str(offset),
                            "valor_captado": 10 if offset < 3 else 0,
                            "cgccpf": "05422243000131",
                            "proponente": "A",
                            "nome": "P",
                        }
                    ]
                },
            }
        itens = []
        for i in range(offset, min(offset + limit, 3)):
            itens.append(
                {
                    "PRONAC": f"20{i:04d}",
                    "valor_captado": 50 - i,
                    "cgccpf": "05422243000131",
                    "proponente": "INSTITUTO",
                    "nome": f"Proj {i}",
                    "resumo": "ignorar",
                }
            )
        return {"total": 150, "_embedded": {"projetos": itens}}

    projetos = listar_projetos_com_captacao(get_json=get, workers=2, n_com_captacao=3)
    assert [p["PRONAC"] for p in projetos] == ["200000", "200001", "200002"]
    assert all("resumo" not in p for p in projetos)


def test_processar_com_mapa_local(tmp_path: Path):
    payload = json.loads(FIXTURE.read_text(encoding="utf-8"))
    rec = extrair_captacoes_do_detalhe(payload)
    projetos = [
        {
            "PRONAC": "261004",
            "nome": rec["nome"],
            "cgccpf": rec["cgccpf"],
            "proponente": rec["proponente"],
            "UF": "PB",
            "valor_captado": 500000,
        }
    ]
    df = processar(
        tmp_path,
        cache_lista=tmp_path / "lista.jsonl",
        cache_captacoes=tmp_path / "cap.jsonl",
        projetos=projetos,
        captacoes_map={"261004": rec},
    )
    assert len(df) == 2
    assert (tmp_path / "captacoes_lei_rouanet_2003_2026.csv").exists()
    assert (tmp_path / "captacoes_lei_rouanet_2003_2026.md").exists()
    assert (tmp_path / "captacoes_lei_rouanet_por_ano.csv").exists()
    csv = (tmp_path / "captacoes_lei_rouanet_2003_2026.csv").read_text(encoding="utf-8")
    assert "cnpj_cpf_beneficiario" in csv
    assert "00584628000181" in csv
    assert "data_captacao" in csv
    md = (tmp_path / "captacoes_lei_rouanet_2003_2026.md").read_text(encoding="utf-8")
    assert "proponente" in md.lower()
