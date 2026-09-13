#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Portarias do DOU que autorizam captação via incentivos da Lei Rouanet.

Consulta a leitura diária da Seção 1 do Diário Oficial da União
(https://www.in.gov.br/leiturajornal), identifica as Portarias SEFIC/MinC
e extrai os projetos homologados para captação de doações e patrocínios
com incentivo fiscal da Lei nº 8.313/1991 (arts. 18 e 26).

Uso::

  python scripts/portarias_lei_rouanet.py
  python scripts/portarias_lei_rouanet.py --inicio 2026-01-01 --fim 2026-09-13
  python scripts/portarias_lei_rouanet.py --somente-liberacao
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

BASE_IN = "https://www.in.gov.br"
LEITURA = f"{BASE_IN}/leiturajornal"
ARTIGO = f"{BASE_IN}/web/dou/-/"
UA = "SEC-data-analysys/portarias-lei-rouanet (pesquisa fiscal; +https://github.com/cafla19791-tech/SEC--data-analysys)"

SECOES = ("do1", "do1e")
ORGAOS_SEFIC = (
    "secretaria de fomento e incentivo à cultura",
    "secretaria de economia criativa e fomento cultural",
)
RE_TITULO_SEFIC = re.compile(
    r"PORTARIA\s+SEFIC(?:/MINC)?\s+N[ºO°]?\s*(\d+)",
    re.I,
)
RE_ARTIGO_URL = re.compile(r"portaria-sefic", re.I)

TIPOS_LIBERACAO = frozenset(
    {
        "homologacao_captacao",
        "complementacao_valor",
        "prorrogacao_prazo",
    }
)

MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "março": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}

RE_PUB = re.compile(
    r"Publicado em:\s*(\d{2}/\d{2}/\d{4})\s*\|\s*Edição:\s*(\d+)\s*\|\s*Seção:\s*(\d+)\s*\|\s*Página:\s*(\d+)",
    re.I,
)
RE_PRONAC = re.compile(r"^(\d{5,7})\s*[-–—]\s*(.+)$")
RE_LABEL = re.compile(
    r"^(CNPJ/?CPF|Processo|Cidade|Valor Aprovado|Valor Complementado|"
    r"Valor total atual|Valor Homologado|Valor Reduzido|Valor reduzido|"
    r"Prazo de Capta[cç][aã]o|Resumo do Projeto|UF)\s*:\s*(.*)$",
    re.I,
)
RE_AREA = re.compile(
    r"^ÁREA:\s*(.+?)(?:\s*\(\s*(Artigo\s+\d+[^)]*)\))?\s*$",
    re.I,
)


class _HTMLText(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip += 1
        if tag in {"p", "br", "div", "li", "tr", "h1", "h2", "h3"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style", "noscript"} and self._skip:
            self._skip -= 1
        if tag in {"p", "div", "li", "h1", "h2", "h3"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip:
            self.parts.append(data)


def html_para_texto(html: str) -> str:
    """Converte o bloco da matéria DOU em texto com quebras de parágrafo."""
    parser = _HTMLText()
    i = html.find('<div class="texto-dou">')
    j = html.find('<div class="rodape', i if i >= 0 else 0)
    chunk = html[i : j if j > i else None] if i >= 0 else html
    parser.feed(chunk)
    texto = "".join(parser.parts)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n[ \t]+", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def _norm(s: str) -> str:
    trans = str.maketrans(
        "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇáàâãäéèêëíìîïóòôõöúùûüç",
        "AAAAAEEEEIIIIOOOOOUUUUcaaaaaeeeeiiiiooooouuuuc",
    )
    return re.sub(r"\s+", " ", (s or "").translate(trans)).strip().lower()


def _num(valor: Any) -> float | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    if not s:
        return None
    s = re.sub(r"[R$\s]", "", s)
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_data_br(s: str | None) -> date | None:
    if not s:
        return None
    s = s.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_data_portaria(titulo: str) -> date | None:
    m = re.search(
        r"DE\s+(\d{1,2}|1[ºo°])\s+DE\s+([A-ZÇÃÕÁÉÍÓÚÂÊÔ]+)\s+DE\s+(\d{4})",
        titulo or "",
        re.I,
    )
    if not m:
        return None
    dia_raw = m.group(1)
    dia = 1 if re.match(r"1[ºo°]$", dia_raw, re.I) else int(dia_raw)
    mes = MESES.get(_norm(m.group(2)))
    if not mes:
        return None
    return date(int(m.group(3)), mes, dia)


def numero_portaria(titulo: str) -> int | None:
    m = RE_TITULO_SEFIC.search(titulo or "")
    return int(m.group(1)) if m else None


def classificar_tipo(texto: str) -> str:
    """Classifica a Portaria SEFIC a partir do art. 1º (ou do texto integral)."""
    n = _norm(texto)
    m = re.search(r"art\.?\s*1\.?[ºo]?\s*[-–—:]?\s*(.+?)(?:\s+art\.?\s*2|\s+anexo|\Z)", n)
    art1 = m.group(1) if m else n[:800]
    if "prorrogacao do prazo de captacao" in art1 or "prorrogacao do prazo" in art1:
        return "prorrogacao_prazo"
    if "valor complementado" in art1 or "complementacao" in art1:
        return "complementacao_valor"
    if "reducao de valor" in art1 or "valor reduzido" in art1:
        return "reducao_valor"
    if "alteracao dos projetos" in art1 or "alteracao do projeto" in art1:
        return "alteracao_projeto"
    if "homologar os projetos culturais" in art1 and (
        "doacoes" in art1 or "patrocinios" in art1 or "admissibilidade" in art1
    ):
        return "homologacao_captacao"
    if "homologar os projetos culturais" in art1:
        return "homologacao_captacao"
    if "8.313" in art1 or "doacoes ou patrocinios" in art1:
        return "outra_rouanet"
    return "outra"


def eh_portaria_sefic(item: dict[str, Any]) -> bool:
    titulo = str(item.get("title") or item.get("titulo") or "")
    hier = _norm(str(item.get("hierarchyStr") or " ".join(item.get("hierarchyList") or [])))
    url = str(item.get("urlTitle") or "")
    if not RE_TITULO_SEFIC.search(titulo) and not RE_ARTIGO_URL.search(url):
        return False
    if any(org in hier for org in ORGAOS_SEFIC):
        return True
    return "ministerio da cultura" in hier


def extrair_params_leitura(html: str) -> dict[str, Any]:
    m = re.search(
        r'<script[^>]*id="params"[^>]*>(.*?)</script>',
        html,
        re.S | re.I,
    )
    if not m:
        return {}
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return {}


def parse_prazo(valor: str) -> tuple[str | None, str | None]:
    m = re.search(
        r"(\d{2}/\d{2}/\d{4})\s*(?:à|a|ate|até|-|–|—)\s*(\d{2}/\d{2}/\d{4})",
        valor or "",
        re.I,
    )
    if not m:
        return None, None
    return m.group(1), m.group(2)


def parse_projetos(texto: str) -> list[dict[str, Any]]:
    """Extrai os projetos dos anexos (PRONAC, proponente, valores, prazo)."""
    linhas = [ln.strip() for ln in texto.splitlines()]
    projetos: list[dict[str, Any]] = []
    atual: dict[str, Any] | None = None
    area = None
    artigo = None

    def flush() -> None:
        nonlocal atual
        if atual and atual.get("pronac"):
            projetos.append(atual)
        atual = None

    i = 0
    while i < len(linhas):
        ln = linhas[i]
        if not ln:
            i += 1
            continue
        m_area = RE_AREA.match(ln)
        if m_area:
            area = m_area.group(1).strip()
            artigo = (m_area.group(2) or "").strip() or None
            if atual:
                atual.setdefault("area", area)
                atual.setdefault("artigo_enquadramento", artigo)
            i += 1
            continue
        m_pr = RE_PRONAC.match(ln)
        if m_pr:
            flush()
            atual = {
                "pronac": m_pr.group(1),
                "nome_projeto": m_pr.group(2).strip().strip('"'),
                "area": area,
                "artigo_enquadramento": artigo,
                "proponente": None,
            }
            i += 1
            continue
        if atual is None:
            i += 1
            continue
        m_lab = RE_LABEL.match(ln)
        if m_lab:
            chave = _norm(m_lab.group(1)).replace(" ", "_")
            val = (m_lab.group(2) or "").strip()
            # resumo pode ocupar várias linhas
            if chave == "resumo_do_projeto":
                partes = [val] if val else []
                j = i + 1
                while j < len(linhas):
                    nxt = linhas[j].strip()
                    if not nxt:
                        j += 1
                        continue
                    if RE_AREA.match(nxt) or RE_PRONAC.match(nxt) or RE_LABEL.match(nxt):
                        break
                    if nxt.startswith("Este conteúdo não substitui"):
                        break
                    partes.append(nxt)
                    j += 1
                atual["resumo"] = " ".join(partes).strip()
                i = j
                continue
            if chave in {"cnpj/cpf", "cnpjcpf"}:
                atual["cnpj_cpf"] = val
            elif chave == "processo":
                atual["processo"] = val
            elif chave == "cidade":
                cidade = val.rstrip(";")
                uf = None
                cm = re.search(r"^(.*?)\s*-\s*([A-Z]{2})$", cidade)
                if cm:
                    cidade, uf = cm.group(1).strip(), cm.group(2)
                atual["cidade"] = cidade
                atual["uf"] = uf
            elif chave == "uf":
                atual["uf"] = val
            elif chave == "valor_aprovado":
                atual["valor_aprovado"] = _num(val)
            elif chave == "valor_complementado":
                atual["valor_complementado"] = _num(val)
            elif chave == "valor_total_atual":
                atual["valor_total_atual"] = _num(val)
            elif chave in {"valor_reduzido", "valor_homologado"}:
                atual[chave] = _num(val)
            elif chave.startswith("prazo_de_captacao"):
                ini, fim = parse_prazo(val)
                atual["prazo_inicio"] = ini
                atual["prazo_fim"] = fim
                atual["prazo_captacao"] = val
            i += 1
            continue
        if atual.get("proponente") is None and not ln.startswith("ANEXO") and ln.upper() != ln[:5]:
            # primeira linha livre após o nome = proponente
            if not ln.startswith("Este conteúdo"):
                atual["proponente"] = ln
        i += 1
    flush()
    return projetos


def metadados_html(html: str, fallback: dict[str, Any] | None = None) -> dict[str, Any]:
    fb = fallback or {}
    texto_pagina = html
    pub = RE_PUB.search(texto_pagina)
    orgao = None
    m_org = re.search(
        r'class="orgao-dou-data"[^>]*>([^<]+)',
        html,
        re.I,
    )
    if m_org:
        orgao = m_org.group(1).strip()
    titulo = fb.get("title") or ""
    m_tit = re.search(r'<p class="identifica">([^<]+)', html, re.I)
    if m_tit:
        titulo = m_tit.group(1).strip()
    return {
        "titulo": titulo,
        "numero": numero_portaria(titulo),
        "data_portaria": parse_data_portaria(titulo),
        "data_publicacao": parse_data_br(pub.group(1) if pub else fb.get("pubDate")),
        "edicao": int(pub.group(2)) if pub else _safe_int(fb.get("editionNumber")),
        "secao": int(pub.group(3)) if pub else None,
        "pagina": int(pub.group(4)) if pub else _safe_int(fb.get("numberPage")),
        "orgao": orgao or fb.get("hierarchyStr"),
        "url_title": fb.get("urlTitle"),
    }


def _safe_int(v: Any) -> int | None:
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def url_artigo(url_title: str) -> str:
    slug = str(url_title).lstrip("/")
    return ARTIGO + slug


_RATE = threading.Lock()
_LAST_REQ = 0.0
_MIN_INTERVAL = 0.07


def _throttle() -> None:
    global _LAST_REQ
    with _RATE:
        wait = _MIN_INTERVAL - (time.monotonic() - _LAST_REQ)
        if wait > 0:
            time.sleep(wait)
        _LAST_REQ = time.monotonic()


def _get(url: str, timeout: int = 60, tentativas: int = 5) -> str:
    last: Exception | None = None
    for i in range(tentativas):
        try:
            _throttle()
            req = Request(url, headers={"User-Agent": UA, "Accept": "text/html,application/json"})
            with urlopen(req, timeout=timeout) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except (HTTPError, URLError, TimeoutError, OSError) as exc:
            last = exc
            time.sleep(min(2 ** i, 16))
    raise RuntimeError(f"Falha ao consultar {url}: {last}") from last


def datas_intervalo(inicio: date, fim: date) -> list[date]:
    out = []
    d = inicio
    while d <= fim:
        out.append(d)
        d += timedelta(days=1)
    return out


def coletar_indice_dia(dia: date, secao: str = "do1") -> list[dict[str, Any]]:
    url = f"{LEITURA}?data={dia.strftime('%d-%m-%Y')}&secao={secao}"
    html = _get(url)
    params = extrair_params_leitura(html)
    itens = params.get("jsonArray") or []
    sefic = []
    for raw in itens:
        if not isinstance(raw, dict):
            continue
        if eh_portaria_sefic(raw):
            raw = dict(raw)
            raw["_secao_leitura"] = secao
            raw["_data_consulta"] = dia.isoformat()
            sefic.append(raw)
    return sefic


def coletar_indices(
    inicio: date,
    fim: date,
    *,
    workers: int = 6,
    secoes: Iterable[str] = SECOES,
) -> list[dict[str, Any]]:
    tarefas = [(d, s) for d in datas_intervalo(inicio, fim) for s in secoes]
    encontrados: list[dict[str, Any]] = []
    print(f"[1/3] Índice DOU {inicio.isoformat()} → {fim.isoformat()} ({len(tarefas)} edições)…", flush=True)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(coletar_indice_dia, d, s): (d, s) for d, s in tarefas}
        feitos = 0
        for fut in as_completed(futs):
            feitos += 1
            dia, secao = futs[fut]
            try:
                lote = fut.result()
            except Exception as exc:
                print(f"  aviso: índice {dia} {secao}: {exc}", flush=True)
                continue
            encontrados.extend(lote)
            if feitos % 40 == 0 or feitos == len(tarefas):
                print(f"  {feitos}/{len(tarefas)} edições · {len(encontrados)} portarias SEFIC", flush=True)
    # dedup por urlTitle
    seen: set[str] = set()
    uniq = []
    for it in encontrados:
        key = str(it.get("urlTitle") or "")
        if not key or key in seen:
            continue
        seen.add(key)
        uniq.append(it)
    uniq.sort(key=lambda x: (parse_data_br(x.get("pubDate")) or date.min, numero_portaria(x.get("title") or "") or 0))
    return uniq


def parse_artigo(html: str, item: dict[str, Any] | None = None) -> dict[str, Any]:
    texto = html_para_texto(html)
    meta = metadados_html(html, item or {})
    tipo = classificar_tipo(texto)
    projetos = parse_projetos(texto)
    valor_total = 0.0
    for p in projetos:
        v = p.get("valor_aprovado")
        if v is None:
            v = p.get("valor_complementado")
        if v is None:
            v = p.get("valor_total_atual")
        if v is not None:
            valor_total += float(v)
            p["valor_referencia"] = float(v)
        else:
            p["valor_referencia"] = None
    url = url_artigo(meta.get("url_title") or "") if meta.get("url_title") else None
    return {
        "titulo": meta.get("titulo"),
        "numero": meta.get("numero"),
        "tipo": tipo,
        "libera_captacao": tipo in TIPOS_LIBERACAO,
        "libera_captacao_inicial": tipo == "homologacao_captacao",
        "data_portaria": meta.get("data_portaria").isoformat() if meta.get("data_portaria") else None,
        "data_publicacao": meta.get("data_publicacao").isoformat() if meta.get("data_publicacao") else None,
        "edicao": meta.get("edicao"),
        "secao": meta.get("secao"),
        "pagina": meta.get("pagina"),
        "orgao": meta.get("orgao"),
        "url": url,
        "url_title": meta.get("url_title"),
        "qtd_projetos": len(projetos),
        "valor_total_anexo": round(valor_total, 2) if projetos else None,
        "lei_8313": "8.313" in texto or "8.313" in _norm(texto),
        "texto_art1": _art1(texto),
        "projetos": projetos,
    }


def _art1(texto: str) -> str | None:
    m = re.search(
        r"Art\.?\s*1\.?[ºo]?\s*[-–—:]?\s*(.+?)(?:\s+Art\.?\s*2|\s+ANEXO|\Z)",
        texto,
        re.S | re.I,
    )
    if not m:
        return None
    return re.sub(r"\s+", " ", m.group(1)).strip()


def baixar_e_parsear(item: dict[str, Any]) -> dict[str, Any]:
    url = url_artigo(str(item.get("urlTitle") or ""))
    html = _get(url)
    parsed = parse_artigo(html, item)
    parsed["url"] = url
    return parsed


def coletar_artigos(itens: list[dict[str, Any]], *, workers: int = 4) -> list[dict[str, Any]]:
    print(f"[2/3] Baixando {len(itens)} matérias no DOU…", flush=True)
    out: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(baixar_e_parsear, it): it for it in itens}
        feitos = 0
        for fut in as_completed(futs):
            feitos += 1
            it = futs[fut]
            try:
                out.append(fut.result())
            except Exception as exc:
                print(f"  aviso: {it.get('urlTitle')}: {exc}", flush=True)
            if feitos % 20 == 0 or feitos == len(itens):
                print(f"  {feitos}/{len(itens)} matérias", flush=True)
    out.sort(
        key=lambda r: (
            r.get("data_publicacao") or "",
            r.get("numero") or 0,
        )
    )
    return out


def _fmt_brl(valor: float | None) -> str:
    if valor is None:
        return "—"
    s = f"{valor:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def escrever_markdown(
    portarias: pd.DataFrame,
    projetos: pd.DataFrame,
    path: Path,
    *,
    inicio: date,
    fim: date,
    gerado_em: str,
) -> None:
    lib = portarias[portarias["libera_captacao_inicial"] == True]  # noqa: E712
    cap = portarias[portarias["libera_captacao"] == True]  # noqa: E712
    valor_lib = float(projetos.loc[projetos["tipo_portaria"] == "homologacao_captacao", "valor_referencia"].fillna(0).sum()) if not projetos.empty else 0.0
    linhas = [
        "# Portarias SEFIC/MinC — captação Lei Rouanet (DOU)",
        "",
        f"**Gerado em:** {gerado_em}",
        "",
        f"Período de publicação no Diário Oficial da União: **{inicio.isoformat()}** a **{fim.isoformat()}**.",
        "",
        "Fonte: [Imprensa Nacional — leitura do jornal](https://www.in.gov.br/leiturajornal) "
        "(Seção 1 e extra). As Portarias da Secretaria de Fomento e Incentivo à Cultura "
        "homologam projetos culturais da Lei nº 8.313/1991 (Lei Rouanet) para captação "
        "de doações e patrocínios com incentivo fiscal (arts. 18 e 26).",
        "",
        "## Totais",
        "",
        f"- Portarias SEFIC/MinC localizadas: **{len(portarias)}**",
        f"- Homologação inicial (liberação da captação): **{len(lib)}** "
        f"({int(lib['qtd_projetos'].sum()) if not lib.empty else 0} projetos; {_fmt_brl(valor_lib)})",
        f"- Demais atos de captação (prorrogação, complementação, redução, alteração): "
        f"**{len(portarias) - len(lib)}**",
        f"- Projetos extraídos dos anexos: **{len(projetos)}**",
        "",
        "### Por tipo",
        "",
        "| Tipo | Portarias | Projetos | Valor de referência |",
        "|------|----------:|---------:|--------------------:|",
    ]
    if portarias.empty:
        linhas.append("| — | 0 | 0 | — |")
    else:
        grp = (
            portarias.groupby("tipo", dropna=False)
            .agg(qtd=("numero", "count"), projetos=("qtd_projetos", "sum"), valor=("valor_total_anexo", "sum"))
            .reset_index()
            .sort_values("qtd", ascending=False)
        )
        for row in grp.itertuples(index=False):
            linhas.append(
                f"| `{row.tipo}` | {int(row.qtd)} | {int(row.projetos or 0)} | {_fmt_brl(row.valor)} |"
            )
    linhas.extend(
        [
            "",
            "## Homologações que liberam a captação (fase de doações e patrocínios)",
            "",
            "Art. 1º típico: *Homologar os projetos culturais relacionados nos anexos desta "
            "portaria, que após terem atendido aos requisitos de admissibilidade estabelecidos "
            "pela Lei nº 8.313/91, passam à fase de obtenção de doações e patrocínios.*",
            "",
            "| Nº | Data da portaria | Publicação DOU | Projetos | Valor aprovado | Edição | URL |",
            "|---:|-----------------:|---------------:|---------:|---------------:|-------:|-----|",
        ]
    )
    cols = [
        "numero",
        "data_portaria",
        "data_publicacao",
        "qtd_projetos",
        "valor_total_anexo",
        "edicao",
        "url",
    ]
    if not lib.empty:
        for row in lib.sort_values(["data_publicacao", "numero"]).itertuples(index=False):
            url = getattr(row, "url", "") or ""
            num = getattr(row, "numero", "")
            linhas.append(
                f"| {num} | {row.data_portaria or ''} | {row.data_publicacao or ''} | "
                f"{int(row.qtd_projetos or 0)} | {_fmt_brl(row.valor_total_anexo)} | "
                f"{row.edicao or ''} | [{num}]({url}) |"
            )
    else:
        linhas.append("| — |  |  | 0 | — |  |  |")

    top = (
        projetos[projetos["tipo_portaria"] == "homologacao_captacao"]
        .sort_values("valor_referencia", ascending=False)
        .head(20)
        if not projetos.empty
        else projetos
    )
    linhas.extend(
        [
            "",
            "## Maiores valores homologados para captação inicial",
            "",
            "| PRONAC | Projeto | Proponente | UF | Valor aprovado | Portaria |",
            "|-------:|---------|------------|----|---------------:|---------:|",
        ]
    )
    if top is not None and not top.empty:
        for row in top.itertuples(index=False):
            nome = (row.nome_projeto or "")[:80]
            prop = (row.proponente or "")[:50]
            linhas.append(
                f"| {row.pronac} | {nome} | {prop} | {row.uf or ''} | "
                f"{_fmt_brl(row.valor_referencia)} | {row.numero_portaria} |"
            )
    linhas.extend(
        [
            "",
            "## Arquivos",
            "",
            "- `portarias_lei_rouanet.csv` — uma linha por Portaria SEFIC/MinC",
            "- `projetos_lei_rouanet_captacao.csv` — uma linha por projeto do anexo",
            "- `portarias_lei_rouanet.xlsx` — as duas abas",
            "",
            "A captação junto a pessoas físicas e empresas tributadas pelo lucro real "
            "só é válida após a publicação no DOU, pelo sistema SALIC, no prazo "
            "homologado na Portaria.",
            "",
        ]
    )
    path.write_text("\n".join(linhas), encoding="utf-8")


def portarias_para_frame(registros: list[dict[str, Any]]) -> tuple[pd.DataFrame, pd.DataFrame]:
    port_rows = []
    proj_rows = []
    for r in registros:
        base = {k: v for k, v in r.items() if k != "projetos"}
        port_rows.append(base)
        for p in r.get("projetos") or []:
            proj_rows.append(
                {
                    "numero_portaria": r.get("numero"),
                    "tipo_portaria": r.get("tipo"),
                    "data_portaria": r.get("data_portaria"),
                    "data_publicacao": r.get("data_publicacao"),
                    "url": r.get("url"),
                    **p,
                }
            )
    portarias = pd.DataFrame(port_rows)
    projetos = pd.DataFrame(proj_rows)
    return portarias, projetos


def processar(
    saida_dir: Path,
    *,
    inicio: date,
    fim: date,
    somente_liberacao: bool = False,
    workers: int = 6,
    indice_html: str | None = None,
    artigo_html_map: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    saida_dir.mkdir(parents=True, exist_ok=True)
    if indice_html is not None:
        params = extrair_params_leitura(indice_html)
        itens = [it for it in (params.get("jsonArray") or []) if eh_portaria_sefic(it)]
    else:
        itens = coletar_indices(inicio, fim, workers=workers)
    print(f"    {len(itens)} Portarias SEFIC/MinC no índice.", flush=True)

    if artigo_html_map is not None:
        registros = [parse_artigo(artigo_html_map[it["urlTitle"]], it) for it in itens if it.get("urlTitle") in artigo_html_map]
        # também aceita parse direto se o mapa usa o próprio html único
        if not registros and len(artigo_html_map) == 1 and itens:
            html = next(iter(artigo_html_map.values()))
            registros = [parse_artigo(html, itens[0])]
    else:
        registros = coletar_artigos(itens, workers=max(2, min(workers, 5)))

    if somente_liberacao:
        registros = [r for r in registros if r.get("libera_captacao_inicial")]

    print("[3/3] Gravando saídas…", flush=True)
    portarias, projetos = portarias_para_frame(registros)
    stem = "portarias_lei_rouanet"
    csv_p = saida_dir / f"{stem}.csv"
    csv_j = saida_dir / "projetos_lei_rouanet_captacao.csv"
    md = saida_dir / f"{stem}.md"
    xlsx = saida_dir / f"{stem}.xlsx"
    if not portarias.empty:
        portarias.to_csv(csv_p, index=False)
    else:
        pd.DataFrame(columns=["titulo", "numero", "tipo"]).to_csv(csv_p, index=False)
    if not projetos.empty:
        projetos.to_csv(csv_j, index=False)
    else:
        pd.DataFrame(columns=["pronac", "nome_projeto"]).to_csv(csv_j, index=False)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    escrever_markdown(portarias, projetos, md, inicio=inicio, fim=fim, gerado_em=gerado)
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xl:
        (portarias if not portarias.empty else pd.DataFrame({"aviso": ["nenhuma portaria"]})).to_excel(
            xl, sheet_name="portarias", index=False
        )
        (projetos if not projetos.empty else pd.DataFrame({"aviso": ["nenhum projeto"]})).to_excel(
            xl, sheet_name="projetos", index=False
        )
    print(f"[OK] {csv_p}")
    print(f"[OK] {csv_j}")
    print(f"[OK] {md}")
    print(f"[OK] {xlsx}")
    return portarias, projetos


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    hoje = date.today()
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--inicio", default=f"{hoje.year}-01-01", help="YYYY-MM-DD (publicação no DOU)")
    p.add_argument("--fim", default=hoje.isoformat(), help="YYYY-MM-DD (publicação no DOU)")
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    p.add_argument("--workers", type=int, default=6)
    p.add_argument(
        "--somente-liberacao",
        action="store_true",
        help="Mantém só homologação inicial (fase de doações e patrocínios)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    inicio = parse_data_br(args.inicio)
    fim = parse_data_br(args.fim)
    if not inicio or not fim or inicio > fim:
        print("ERRO: --inicio/--fim inválidos (use YYYY-MM-DD).", file=sys.stderr)
        return 2
    try:
        processar(
            args.saida_dir,
            inicio=inicio,
            fim=fim,
            somente_liberacao=args.somente_liberacao,
            workers=args.workers,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
