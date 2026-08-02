from __future__ import annotations

import json
from pathlib import Path

import pytest

from bndes_mcp import excel_export, providers


SAMPLE_DOC = {
    "id": "abc",
    "cliente": "EMBRAER S.A.",
    "documentoCliente": "07689002000189",
    "tipoDocumento": "cnpj",
    "numeroContrato": "13213711",
    "dataContratacao": "2013-12-23T00:00:00Z",
    "anoPosicao": 2013,
    "descricaoProjeto": "TESTE",
    "valorContratacao": 1000.0,
    "valorDesembolsado": 900.0,
    "faixaValorContratacao": "(3)",
    "produtoBndes": ["BNDES FINEM"],
    "instrumentoBndes": ["PSI"],
    "tipoOperacao": "NÃO AUTOMÁTICA",
    "tope": "DIRETA",
    "operacaoDireta": ["DIRETA"],
    "liquidada": ["LIQUIDADO"],
    "uf": ["SP"],
    "municipio": ["SAO JOSE DOS CAMPOS"],
    "setorApoiado": ["INDUSTRIA"],
    "subsetorApoiado": ["MATERIAL DE TRANSPORTE"],
    "cnae": ["C3041500"],
    "porteCliente": ["(3) GRANDE"],
    "naturezaCliente": ["PRIVADA"],
    "fonteRecursos": ["FAT"],
    "custoFinanceiro": ["TJLP"],
    "taxaJuros": ["3.500"],
    "prazoCarencia": ["12"],
    "prazoAmortizacao": ["60"],
    "inovacao": ["SIM"],
    "reembolsavel": ["REEMBOLSÁVEL"],
    "agenteFinanceiro": ["NÃO DISPONÍVEL"],
    "cnpjAgenteFinanceiro": ["NÃO DISPONÍVEL"],
    "tipoGarantia": ["REAL"],
    "areaOperacional": ["AREA"],
    "paisDestino": ["NÃO SE APLICA"],
    "modalidadeOperacional": ["NÃO SE APLICA"],
    "moeda": ["R$"],
    "subcreditos": (
        "<subcreditos><subcredito>"
        "<cliente>EMBRAER S.A.</cliente>"
        "<documentoCliente>07689002000189</documentoCliente>"
        "<numeroContrato>13213711</numeroContrato>"
        "<dataContratacao>23/12/13</dataContratacao>"
        "<valorContratacao>600,00</valorContratacao>"
        "<valorDesembolsado>500,00</valorDesembolsado>"
        "<produtoBndes>BNDES FINEM</produtoBndes>"
        "</subcredito>"
        "<subcredito>"
        "<cliente>EMBRAER S.A.</cliente>"
        "<documentoCliente>07689002000189</documentoCliente>"
        "<numeroContrato>13213711</numeroContrato>"
        "<dataContratacao>23/12/13</dataContratacao>"
        "<valorContratacao>400,00</valorContratacao>"
        "<valorDesembolsado>400,00</valorDesembolsado>"
        "<produtoBndes>BNDES FINEM</produtoBndes>"
        "</subcredito></subcreditos>"
    ),
}


def test_digits_and_tipo():
    assert providers.digits_only("07.689.002/0001-89") == "07689002000189"
    assert providers.detect_tipo_documento("07689002000189") == "cnpj"
    assert providers.detect_tipo_documento("12345678901") == "cpf"
    with pytest.raises(ValueError):
        providers.detect_tipo_documento("123")


def test_parse_subcreditos():
    rows = providers.parse_subcreditos_xml(SAMPLE_DOC["subcreditos"])
    assert len(rows) == 2
    assert rows[0]["valorContratacao"] == "600,00"


def test_expand_and_excel(tmp_path: Path):
    docs = [SAMPLE_DOC]
    summary = providers.summarize(docs)
    assert summary["num_operacoes"] == 1
    assert summary["soma_valor_contratacao"] == 1000.0

    subs = providers.expand_subcreditos(docs)
    assert len(subs) == 2
    assert subs[0]["valorContratacao"] == 600.0
    assert subs[0]["dataContratacao"] == "2013-12-23"

    out = tmp_path / "t.xlsx"
    excel_export.write_excel(docs, out)
    assert out.exists() and out.stat().st_size > 1000


def test_json_roundtrip(tmp_path: Path):
    payload = {"response": {"numFound": 1, "docs": [SAMPLE_DOC]}}
    jp = tmp_path / "a.json"
    jp.write_text(json.dumps(payload), encoding="utf-8")
    docs = payload["response"]["docs"]
    excel_export.write_excel(docs, tmp_path / "b.xlsx")
