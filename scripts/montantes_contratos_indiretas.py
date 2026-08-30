#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Montantes anuais dos contratos BNDES na modalidade indireta (2002–2026).

Consolida:
  1. Microdados de operações indiretas automáticas (valor da operação);
  2. Microdados de operações não automáticas com forma_de_apoio = INDIRETA
     (valor contratado; cada linha é um subcrédito);
  3. Série oficial de aprovações indiretas (R$ milhões) — conferência.

Fontes (Portal de Dados Abertos do BNDES):
  https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento
  https://dadosabertos.bndes.gov.br/dataset/aprovacoes

Uso::

  python scripts/montantes_contratos_indiretas.py
  python scripts/montantes_contratos_indiretas.py --saida-dir output
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

CKAN = "https://dadosabertos.bndes.gov.br/api/3/action"
RESOURCE_AUTOMATICAS = "612faa0b-b6be-4b2c-9317-da5dc2c0b901"
RESOURCE_NAO_AUTOMATICAS = "6f56b78c-510f-44b6-8274-78a5b7e931f4"
URL_APROVACOES = (
    "https://dadosabertos.bndes.gov.br/dataset/"
    "dc522a8d-51f8-443c-9a42-16574843d4e3/resource/"
    "0ac8adcb-5276-4132-a79d-d876763e1ee0/download/"
    "aprovacoes-por-forma-de-apoio-indiretas-e-produto-aprovacoes.csv"
)

ANO_MIN = 2002
ANO_MAX = 2026
PAGE_SIZE = 32_000
UA = "SEC-data-analysys/montantes-contratos-indiretas"


def _get_json(url: str, timeout: int = 120, tentativas: int = 5) -> dict:
    last: Exception | None = None
    for i in range(tentativas):
        try:
            req = Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
            last = exc
            time.sleep(2 ** i)
    raise RuntimeError(f"Falha ao consultar {url}: {last}") from last


def datastore_paginas(
    resource_id: str,
    fields: list[str],
    *,
    page_size: int = PAGE_SIZE,
    filters: dict[str, Any] | None = None,
    timeout: int = 180,
) -> Iterable[list[dict]]:
    """Gera páginas de registros do CKAN datastore_search."""
    offset = 0
    total: int | None = None
    while True:
        params: dict[str, Any] = {
            "resource_id": resource_id,
            "fields": ",".join(fields),
            "limit": page_size,
            "offset": offset,
        }
        if filters:
            params["filters"] = json.dumps(filters, ensure_ascii=False)
        url = f"{CKAN}/datastore_search?{urlencode(params)}"
        payload = _get_json(url, timeout=timeout)
        if not payload.get("success"):
            raise RuntimeError(f"CKAN recusou {resource_id}: {payload.get('error')}")
        result = payload["result"]
        if total is None:
            total = int(result.get("total") or 0)
        records = result.get("records") or []
        yield records
        offset += len(records)
        if not records or (total is not None and offset >= total):
            break


def ano_da_data(valor: Any) -> int | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    texto = str(valor).strip()
    if len(texto) < 4 or not texto[:4].isdigit():
        return None
    ano = int(texto[:4])
    if ano < 1900 or ano > 2100:
        return None
    return ano


def _num(valor: Any) -> float:
    if valor is None or valor == "":
        return 0.0
    if isinstance(valor, (int, float)):
        return 0.0 if pd.isna(valor) else float(valor)
    texto = str(valor).strip()
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def agregar_registros(
    records: Iterable[dict],
    *,
    col_data: str,
    col_contratado: str,
    col_desembolsado: str,
    acc: dict[int, dict[str, float]] | None = None,
) -> dict[int, dict[str, float]]:
    """Soma quantidade, valor contratado e desembolsado por ano da contratação."""
    if acc is None:
        acc = defaultdict(lambda: {"qtd": 0.0, "contratado": 0.0, "desembolsado": 0.0})
    for rec in records:
        ano = ano_da_data(rec.get(col_data))
        if ano is None or ano < ANO_MIN or ano > ANO_MAX:
            continue
        acc[ano]["qtd"] += 1
        acc[ano]["contratado"] += _num(rec.get(col_contratado))
        acc[ano]["desembolsado"] += _num(rec.get(col_desembolsado))
    return acc


def agregar_automaticas(*, page_size: int = PAGE_SIZE) -> dict[int, dict[str, float]]:
    acc: dict[int, dict[str, float]] = defaultdict(
        lambda: {"qtd": 0.0, "contratado": 0.0, "desembolsado": 0.0}
    )
    n = 0
    for pagina in datastore_paginas(
        RESOURCE_AUTOMATICAS,
        [
            "data_da_contratacao",
            "valor_da_operacao_em_reais",
            "valor_desembolsado_reais",
        ],
        page_size=page_size,
    ):
        agregar_registros(
            pagina,
            col_data="data_da_contratacao",
            col_contratado="valor_da_operacao_em_reais",
            col_desembolsado="valor_desembolsado_reais",
            acc=acc,
        )
        n += len(pagina)
        print(f"  [automáticas] {n:,} linhas …", flush=True)
    return acc


def agregar_nao_automaticas_indiretas(*, page_size: int = PAGE_SIZE) -> dict[int, dict[str, float]]:
    acc: dict[int, dict[str, float]] = defaultdict(
        lambda: {"qtd": 0.0, "contratado": 0.0, "desembolsado": 0.0}
    )
    n = 0
    for pagina in datastore_paginas(
        RESOURCE_NAO_AUTOMATICAS,
        [
            "data_da_contratacao",
            "valor_contratado_reais",
            "valor_desembolsado_reais",
            "forma_de_apoio",
        ],
        page_size=page_size,
        filters={"forma_de_apoio": "INDIRETA"},
    ):
        indiretas = [
            r
            for r in pagina
            if str(r.get("forma_de_apoio") or "").strip().upper() == "INDIRETA"
        ]
        agregar_registros(
            indiretas,
            col_data="data_da_contratacao",
            col_contratado="valor_contratado_reais",
            col_desembolsado="valor_desembolsado_reais",
            acc=acc,
        )
        n += len(indiretas)
        print(f"  [não automáticas INDIRETA] {n:,} linhas …", flush=True)
    return acc


def carregar_aprovacoes_oficiais(url: str = URL_APROVACOES) -> pd.DataFrame:
    df = pd.read_csv(url, sep=";", decimal=",", encoding="latin-1")
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "ano" not in df.columns:
        raise ValueError(f"CSV de aprovações sem coluna ano: {list(df.columns)}")
    prod = [c for c in df.columns if c not in {"ano", "mes"}]
    for c in prod:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    g = df.groupby("ano", as_index=False)[prod].sum()
    g["aprovado_oficial_r_milhoes"] = g[prod].sum(axis=1)
    g["aprovado_oficial_reais"] = g["aprovado_oficial_r_milhoes"] * 1_000_000.0
    return g[["ano", "aprovado_oficial_r_milhoes", "aprovado_oficial_reais"] + prod]


def _acc_to_frame(acc: dict[int, dict[str, float]], prefixo: str) -> pd.DataFrame:
    rows = []
    for ano in range(ANO_MIN, ANO_MAX + 1):
        d = acc.get(ano, {"qtd": 0.0, "contratado": 0.0, "desembolsado": 0.0})
        rows.append(
            {
                "ano": ano,
                f"qtd_{prefixo}": int(d["qtd"]),
                f"contratado_{prefixo}": float(d["contratado"]),
                f"desembolsado_{prefixo}": float(d["desembolsado"]),
            }
        )
    return pd.DataFrame(rows)


def montar_resumo(
    auto: dict[int, dict[str, float]],
    nao_auto: dict[int, dict[str, float]],
    aprovacoes: pd.DataFrame,
) -> pd.DataFrame:
    a = _acc_to_frame(auto, "automaticas")
    n = _acc_to_frame(nao_auto, "nao_automaticas")
    out = a.merge(n, on="ano", how="outer").merge(
        aprovacoes[["ano", "aprovado_oficial_r_milhoes", "aprovado_oficial_reais"]],
        on="ano",
        how="left",
    )
    out["qtd_total"] = out["qtd_automaticas"] + out["qtd_nao_automaticas"]
    out["contratado_total"] = out["contratado_automaticas"] + out["contratado_nao_automaticas"]
    out["desembolsado_total"] = (
        out["desembolsado_automaticas"] + out["desembolsado_nao_automaticas"]
    )
    out["contratado_total_r_milhoes"] = out["contratado_total"] / 1_000_000.0
    out["desembolsado_total_r_milhoes"] = out["desembolsado_total"] / 1_000_000.0
    return out.sort_values("ano").reset_index(drop=True)


def _fmt_bi(valor: float) -> str:
    return f"R$ {valor / 1e9:,.2f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_mi(valor: float) -> str:
    return f"R$ {valor:,.1f} mi".replace(",", "X").replace(".", ",").replace("X", ".")


def escrever_markdown(resumo: pd.DataFrame, path: Path, gerado_em: str) -> None:
    linhas = [
        "# Montantes dos contratos BNDES — modalidade indireta (2002–2026)",
        "",
        f"**Gerado em:** {gerado_em}",
        "",
        "Valores em reais correntes da data do contrato (microdados) e, na última coluna, "
        "a série oficial de **aprovações** indiretas do BNDES (R$ milhões).",
        "",
        "Fontes:",
        "",
        "- [Operações de financiamento](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento) "
        "(indiretas automáticas + não automáticas com `forma_de_apoio = INDIRETA`).",
        "- [Estatísticas de aprovações](https://dadosabertos.bndes.gov.br/dataset/aprovacoes) "
        "(por forma de apoio indireta e produto).",
        "",
        "Observações:",
        "",
        "- Cada linha das automáticas é uma operação; o montante do contrato é "
        "`valor_da_operacao_em_reais`.",
        "- Nas não automáticas, cada linha é um subcrédito; a soma dos subcréditos "
        "equivale ao valor do contrato (`valor_contratado_reais`).",
        "- A listagem de automáticas **não inclui** Cartão BNDES nem operações com pessoas físicas.",
        "- 2026 está incompleto (última atualização mensal do portal).",
        "",
        "| Ano | Operações | Contratado (R$ bi) | Desembolsado (R$ bi) | Aprovações oficiais (R$ mi) |",
        "|----:|----------:|-------------------:|---------------------:|----------------------------:|",
    ]
    for row in resumo.itertuples(index=False):
        linhas.append(
            f"| {int(row.ano)} | {int(row.qtd_total):,} | "
            f"{row.contratado_total / 1e9:,.2f} | "
            f"{row.desembolsado_total / 1e9:,.2f} | "
            f"{(row.aprovado_oficial_r_milhoes or 0):,.1f} |".replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    tot_q = int(resumo["qtd_total"].sum())
    tot_c = float(resumo["contratado_total"].sum())
    tot_d = float(resumo["desembolsado_total"].sum())
    tot_a = float(resumo["aprovado_oficial_r_milhoes"].fillna(0).sum())
    linhas.append(
        f"| **Total** | **{tot_q:,}** | **{tot_c / 1e9:,.2f}** | "
        f"**{tot_d / 1e9:,.2f}** | **{tot_a:,.1f}** |".replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )
    linhas.extend(
        [
            "",
            f"Total contratado (microdados): {_fmt_bi(tot_c)}.",
            f"Total desembolsado nas mesmas operações: {_fmt_bi(tot_d)}.",
            f"Total aprovado (série oficial): {_fmt_mi(tot_a)}.",
            "",
        ]
    )
    path.write_text("\n".join(linhas), encoding="utf-8")


def processar(saida_dir: Path, *, page_size: int = PAGE_SIZE) -> pd.DataFrame:
    saida_dir.mkdir(parents=True, exist_ok=True)
    print("[1/3] Operações indiretas automáticas …", flush=True)
    auto = agregar_automaticas(page_size=page_size)
    print("[2/3] Operações não automáticas (INDIRETA) …", flush=True)
    nao_auto = agregar_nao_automaticas_indiretas(page_size=page_size)
    print("[3/3] Série oficial de aprovações indiretas …", flush=True)
    aprovacoes = carregar_aprovacoes_oficiais()
    resumo = montar_resumo(auto, nao_auto, aprovacoes)

    csv_path = saida_dir / "montantes_contratos_indiretas_2002_2026.csv"
    md_path = saida_dir / "montantes_contratos_indiretas_2002_2026.md"
    xlsx_path = saida_dir / "montantes_contratos_indiretas_2002_2026.xlsx"
    resumo.to_csv(csv_path, index=False, float_format="%.2f")
    resumo.to_excel(xlsx_path, index=False)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    escrever_markdown(resumo, md_path, gerado)
    print(f"[OK] {csv_path}")
    print(f"[OK] {md_path}")
    print(f"[OK] {xlsx_path}")
    return resumo


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--saida-dir",
        type=Path,
        default=ROOT / "output",
        help="Pasta das saídas CSV/MD/XLSX",
    )
    p.add_argument("--page-size", type=int, default=PAGE_SIZE)
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        processar(args.saida_dir, page_size=args.page_size)
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
