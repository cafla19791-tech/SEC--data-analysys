"""Cliente HTTP da consulta publica de operacoes do BNDES."""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from typing import Any
from xml.sax.saxutils import unescape

import httpx

DEFAULT_BASE = "https://gateway.apis.bndes.gov.br/operacoes/web"
DEFAULT_UA = "SEC-data-analysys-bndes-mcp/0.1 (cafla19791@gmail.com)"

LIST_FIELDS = (
    "paisDestino",
    "reembolsavel",
    "modalidadeOperacional",
    "uf",
    "liquidada",
    "setorApoiado",
    "porteCliente",
    "municipio",
    "naturezaCliente",
    "fonteRecursos",
    "taxaJuros",
    "operacaoDireta",
    "inovacao",
    "subsetorApoiado",
    "produtoBndes",
    "custoFinanceiro",
    "cnae",
    "agenteFinanceiro",
    "prazoAmortizacao",
    "tipoGarantia",
    "categoria",
    "instrumentoBndes",
    "prazoCarencia",
    "moeda",
    "cnpjAgenteFinanceiro",
    "areaOperacional",
)


def digits_only(value: str) -> str:
    return re.sub(r"\D+", "", value or "")


def detect_tipo_documento(documento: str) -> str:
    d = digits_only(documento)
    if len(d) == 11:
        return "cpf"
    if len(d) == 14:
        return "cnpj"
    raise ValueError(f"Documento invalido (esperado CPF 11 ou CNPJ 14 digitos): {documento!r}")


def join_list(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        return " | ".join(str(x) for x in value if x is not None and str(x) != "")
    return str(value)


def _client_timeout() -> float:
    try:
        return float(os.environ.get("BNDES_TIMEOUT", "60"))
    except ValueError:
        return 60.0


def fetch_operacoes(
    documento: str,
    *,
    rows: int = 10000,
    start: int = 0,
    base_url: str | None = None,
    user_agent: str | None = None,
) -> dict[str, Any]:
    """Consulta Solr publica do BNDES por documentoClienteIndex."""
    doc = digits_only(documento)
    tipo = detect_tipo_documento(doc)
    base = (base_url or os.environ.get("BNDES_OPERACOES_BASE") or DEFAULT_BASE).rstrip("/")
    ua = user_agent or os.environ.get("BNDES_USER_AGENT") or DEFAULT_UA
    params = {
        "q": f"documentoClienteIndex:({doc})",
        "defType": "lucene",
        "tdoc": f"tipoDocumento:{tipo}",
        "omitHeader": "false",
        "start": str(start),
        "rows": str(rows),
        "wt": "json",
        "sort": "product(fr,query({!edismax v=$q})) desc",
    }
    url = f"{base}/select"
    with httpx.Client(timeout=_client_timeout(), headers={"User-Agent": ua, "Accept": "application/json"}) as client:
        r = client.get(url, params=params)
        r.raise_for_status()
        data = r.json()
    if not isinstance(data, dict) or "response" not in data:
        raise RuntimeError("Resposta inesperada da API BNDES (sem response)")
    return data


def flatten_operacao(doc: dict[str, Any]) -> dict[str, Any]:
    """Uma linha por contrato/documento Solr (sem XML de subcreditos)."""
    out: dict[str, Any] = {
        "id": doc.get("id"),
        "cliente": doc.get("cliente"),
        "documentoCliente": doc.get("documentoCliente"),
        "tipoDocumento": doc.get("tipoDocumento"),
        "numeroContrato": doc.get("numeroContrato"),
        "numeroOperacao": doc.get("numeroOperacao"),
        "dataContratacao": (doc.get("dataContratacao") or "")[:10],
        "anoPosicao": doc.get("anoPosicao"),
        "descricaoProjeto": doc.get("descricaoProjeto"),
        "valorContratacao": doc.get("valorContratacao"),
        "valorDesembolsado": doc.get("valorDesembolsado"),
        "valorContratacaoMoeda": doc.get("valorContratacaoMoeda"),
        "valorDesembolsadoMoeda": doc.get("valorDesembolsadoMoeda"),
        "faixaValorContratacao": doc.get("faixaValorContratacao"),
        "tipoOperacao": doc.get("tipoOperacao"),
        "tope": doc.get("tope"),
        "tipoAtivo": doc.get("tipoAtivo"),
        "objetivoPredominante": doc.get("objetivoPredominante"),
        "mutuario": doc.get("mutuario"),
        "fr": doc.get("fr"),
    }
    for field in LIST_FIELDS:
        out[field] = join_list(doc.get(field))
    return out


def parse_subcreditos_xml(xml_text: str | None) -> list[dict[str, str]]:
    if not xml_text or not str(xml_text).strip():
        return []
    text = unescape(str(xml_text))
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        # Alguns registros escapam & como &amp; ja; tenta wrap
        try:
            root = ET.fromstring(f"<root>{text}</root>")
        except ET.ParseError:
            return []

    rows: list[dict[str, str]] = []
    nodes = root.findall(".//subcredito")
    if not nodes and root.tag.lower().endswith("subcredito"):
        nodes = [root]
    for node in nodes:
        row: dict[str, str] = {}
        for child in list(node):
            tag = child.tag.split("}")[-1]
            row[tag] = (child.text or "").strip()
        if row:
            rows.append(row)
    return rows


def expand_subcreditos(docs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Expande XML subcreditos; se ausente, gera 1 linha a partir do doc pai."""
    out: list[dict[str, Any]] = []
    for doc in docs:
        parent_id = doc.get("id")
        parent_contrato = doc.get("numeroContrato")
        parent_cliente = doc.get("cliente")
        subs = parse_subcreditos_xml(doc.get("subcreditos"))
        if not subs:
            out.append(
                {
                    "idOperacao": parent_id,
                    "numeroContratoPai": parent_contrato,
                    "cliente": parent_cliente,
                    "documentoCliente": doc.get("documentoCliente"),
                    "descricaoProjeto": doc.get("descricaoProjeto"),
                    "uf": join_list(doc.get("uf")),
                    "municipio": join_list(doc.get("municipio")),
                    "numeroContrato": parent_contrato,
                    "dataContratacao": (doc.get("dataContratacao") or "")[:10],
                    "valorContratacao": doc.get("valorContratacao"),
                    "valorDesembolsado": doc.get("valorDesembolsado"),
                    "fonteRecursos": join_list(doc.get("fonteRecursos")),
                    "custoFinanceiro": join_list(doc.get("custoFinanceiro")),
                    "taxaJuros": join_list(doc.get("taxaJuros")),
                    "prazoCarencia": join_list(doc.get("prazoCarencia")),
                    "prazoAmortizacao": join_list(doc.get("prazoAmortizacao")),
                    "reembolsavel": join_list(doc.get("reembolsavel")),
                    "operacaoDireta": join_list(doc.get("operacaoDireta")),
                    "produtoBndes": join_list(doc.get("produtoBndes")),
                    "instrumentoBndes": join_list(doc.get("instrumentoBndes")),
                    "inovacao": join_list(doc.get("inovacao")),
                    "areaOperacional": join_list(doc.get("areaOperacional")),
                    "cnae": join_list(doc.get("cnae")),
                    "setorApoiado": join_list(doc.get("setorApoiado")),
                    "subsetorApoiado": join_list(doc.get("subsetorApoiado")),
                    "porteCliente": join_list(doc.get("porteCliente")),
                    "naturezaCliente": join_list(doc.get("naturezaCliente")),
                    "cnpjAgenteFinanceiro": join_list(doc.get("cnpjAgenteFinanceiro")),
                    "agenteFinanceiro": join_list(doc.get("agenteFinanceiro")),
                    "tipoGarantia": join_list(doc.get("tipoGarantia")),
                    "liquidada": join_list(doc.get("liquidada")),
                    "paisDestino": join_list(doc.get("paisDestino")),
                    "origemLinha": "documento",
                }
            )
            continue
        for i, sub in enumerate(subs, start=1):
            row = {
                "idOperacao": parent_id,
                "numeroContratoPai": parent_contrato,
                "origemLinha": f"subcredito:{i}",
            }
            row.update(sub)
            # normaliza data BR -> ISO quando possivel
            dc = row.get("dataContratacao") or ""
            if re.fullmatch(r"\d{2}/\d{2}/\d{2}", dc):
                dd, mm, yy = dc.split("/")
                century = "20" if int(yy) <= 70 else "19"
                row["dataContratacao"] = f"{century}{yy}-{mm}-{dd}"
            # valores BR com virgula
            for money in ("valorContratacao", "valorDesembolsado"):
                if money in row and isinstance(row[money], str) and row[money]:
                    row[money] = _parse_br_number(row[money])
            out.append(row)
    return out


def _parse_br_number(value: str) -> float | str:
    s = value.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return value


def summarize(docs: list[dict[str, Any]]) -> dict[str, Any]:
    n = len(docs)
    contratacao = sum(float(d["valorContratacao"]) for d in docs if d.get("valorContratacao") is not None)
    desembolso = sum(float(d["valorDesembolsado"]) for d in docs if d.get("valorDesembolsado") is not None)
    com_valor = sum(1 for d in docs if d.get("valorContratacao") is not None)
    cliente = next((d.get("cliente") for d in docs if d.get("cliente")), "")
    anos = sorted({d.get("anoPosicao") for d in docs if d.get("anoPosicao") is not None})
    return {
        "cliente": cliente,
        "num_operacoes": n,
        "com_valor_contratacao": com_valor,
        "sem_valor_contratacao": n - com_valor,
        "soma_valor_contratacao": contratacao,
        "soma_valor_desembolsado": desembolso,
        "ano_min": anos[0] if anos else None,
        "ano_max": anos[-1] if anos else None,
    }
