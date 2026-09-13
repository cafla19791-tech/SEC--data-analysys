#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cruza PRONACs das Portarias SEFIC/MinC com o SALIC e estima a renúncia.

Para cada projeto homologado no DOU consulta
``https://api.salic.cultura.gov.br/api/v1/projetos?PRONAC=`` e compara
``valor_captado`` (inflow já registrado) com o teto homologado.

Renúncia específica da Lei nº 8.313/1991 (não inclui eventual despesa
operacional do patrocínio):

- art. 18: 100% do valor captado (dedução integral do IR);
- art. 26: 30% (patrocínio PJ, inciso II) e, em coluna à parte, 40%
  (doação PJ, inciso I). Sem o tipo do incentivador, o cenário-base é 30%.

Uso::

  python scripts/renuncia_lei_rouanet_salic.py
  python scripts/renuncia_lei_rouanet_salic.py --max-pronacs 20
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
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

SALIC = "https://api.salic.cultura.gov.br/api/v1/projetos"
UA = "SEC-data-analysys/renuncia-lei-rouanet-salic"
ALIQUOTA_ART18 = 1.0
ALIQUOTA_ART26_PATROCINIO_PJ = 0.30
ALIQUOTA_ART26_DOACAO_PJ = 0.40

_RATE = threading.Lock()
_LAST = 0.0
_MIN_INTERVAL = 0.03


def _throttle() -> None:
    global _LAST
    with _RATE:
        wait = _MIN_INTERVAL - (time.monotonic() - _LAST)
        if wait > 0:
            time.sleep(wait)
        _LAST = time.monotonic()


def _get_json(url: str, timeout: int = 60, tentativas: int = 5) -> dict[str, Any]:
    last: Exception | None = None
    for i in range(tentativas):
        try:
            _throttle()
            req = Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8", errors="replace"))
        except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError) as exc:
            last = exc
            time.sleep(min(2 ** i, 16))
    raise RuntimeError(f"Falha ao consultar {url}: {last}") from last


def norm_pronac(valor: Any) -> str | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    s = str(valor).strip()
    if not s or s.lower() == "nan":
        return None
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    s = s.replace(".0", "") if s.endswith(".0") else s
    s = re.sub(r"[^\d]", "", s)
    return s or None


def artigo_norma(texto: Any) -> str:
    n = re.sub(r"\s+", " ", str(texto or "")).lower()
    if re.search(r"artigo\s*26|\bart\.?\s*26\b", n):
        return "artigo_26"
    if re.search(r"artigo\s*18|\bart\.?\s*18\b", n):
        return "artigo_18"
    return "indefinido"


def slim_projeto(raw: dict[str, Any]) -> dict[str, Any]:
    def _num(chave: str) -> float | None:
        v = raw.get(chave)
        try:
            return float(v) if v is not None and v != "" else None
        except (TypeError, ValueError):
            return None

    return {
        "PRONAC": norm_pronac(raw.get("PRONAC")),
        "nome": raw.get("nome"),
        "situacao": raw.get("situacao"),
        "proponente": raw.get("proponente"),
        "UF": raw.get("UF"),
        "municipio": raw.get("municipio"),
        "segmento": raw.get("segmento"),
        "enquadramento_salic": raw.get("enquadradmento") or raw.get("enquadramento"),
        "mecanismo": raw.get("mecanisnmo") or raw.get("mecanismo"),
        "valor_aprovado": _num("valor_aprovado"),
        "valor_captado": _num("valor_captado"),
        "valor_projeto": _num("valor_projeto"),
        "valor_solicitado": _num("valor_solicitado"),
        "valor_proposta": _num("valor_proposta"),
        "data_inicio": raw.get("data_inicio"),
        "data_termino": raw.get("data_termino"),
        "ano_projeto": raw.get("ano_projeto"),
    }


def projeto_do_payload(payload: dict[str, Any]) -> dict[str, Any] | None:
    if payload.get("PRONAC"):
        return slim_projeto(payload)
    emb = payload.get("_embedded") or {}
    lista = emb.get("projetos") or []
    if lista:
        return slim_projeto(lista[0])
    return None


def url_pronac(pronac: str) -> str:
    return f"{SALIC}?{urlencode({'PRONAC': pronac, 'limit': 1})}"


def cache_carregar(path: Path) -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    if not path.exists():
        return out
    with path.open(encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            key = rec.get("PRONAC")
            if key:
                out[str(key)] = rec
    return out


def cache_append(path: Path, rec: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rec, ensure_ascii=False) + "\n")


def buscar_pronac(pronac: str) -> dict[str, Any] | None:
    payload = _get_json(url_pronac(pronac))
    proj = projeto_do_payload(payload)
    if proj and proj.get("PRONAC"):
        return proj
    return None


def coletar_salic(
    pronacs: Iterable[str],
    *,
    cache_path: Path,
    workers: int = 8,
) -> dict[str, dict[str, Any]]:
    cache = cache_carregar(cache_path)
    faltam = [p for p in pronacs if p not in cache]
    print(f"[SALIC] cache={len(cache)} a_buscar={len(faltam)}", flush=True)
    if not faltam:
        return cache
    lock = threading.Lock()
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(buscar_pronac, p): p for p in faltam}
        feitos = 0
        for fut in as_completed(futs):
            p = futs[fut]
            feitos += 1
            try:
                rec = fut.result()
            except Exception as exc:
                print(f"  aviso PRONAC {p}: {exc}", flush=True)
                rec = {"PRONAC": p, "erro": str(exc)}
            if rec is None:
                rec = {"PRONAC": p, "erro": "nao_encontrado"}
            with lock:
                cache[p] = rec
                cache_append(cache_path, rec)
            if feitos % 100 == 0 or feitos == len(faltam):
                print(f"  {feitos}/{len(faltam)} consultas", flush=True)
    return cache


def projetos_dou_unicos(df: pd.DataFrame, *, somente_homologacao: bool = True) -> pd.DataFrame:
    dados = df.copy()
    dados["pronac"] = dados["pronac"].map(norm_pronac)
    dados = dados[dados["pronac"].notna()]
    if somente_homologacao and "tipo_portaria" in dados.columns:
        dados = dados[dados["tipo_portaria"] == "homologacao_captacao"]
    dados["data_publicacao"] = pd.to_datetime(dados.get("data_publicacao"), errors="coerce")
    dados["artigo"] = dados.get("artigo_enquadramento").map(artigo_norma) if "artigo_enquadramento" in dados.columns else "indefinido"
    dados = dados.sort_values(["pronac", "data_publicacao"])
    keep = {
        "nome_projeto": "last",
        "proponente": "last",
        "uf": "last",
        "cidade": "last",
        "area": "last",
        "artigo": "last",
        "artigo_enquadramento": "last",
        "numero_portaria": "last",
        "data_portaria": "last",
        "data_publicacao": "last",
        "prazo_inicio": "last",
        "prazo_fim": "last",
        "url": "last",
        "valor_aprovado": "last",
        "valor_referencia": "last",
    }
    use = {k: v for k, v in keep.items() if k in dados.columns}
    out = dados.groupby("pronac", as_index=False).agg(use)
    if "valor_aprovado" in out.columns and "valor_referencia" in out.columns:
        out["teto_dou"] = out["valor_aprovado"].fillna(out["valor_referencia"])
    elif "valor_referencia" in out.columns:
        out["teto_dou"] = out["valor_referencia"]
    else:
        out["teto_dou"] = None
    return out


def aliquota_base(artigo: str) -> float:
    if artigo == "artigo_18":
        return ALIQUOTA_ART18
    if artigo == "artigo_26":
        return ALIQUOTA_ART26_PATROCINIO_PJ
    return 0.0


def cruzar(dou: pd.DataFrame, salic: dict[str, dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for rec in dou.to_dict("records"):
        p = rec["pronac"]
        s = salic.get(p) or {}
        artigo = rec.get("artigo") or artigo_norma(s.get("enquadramento_salic"))
        if artigo == "indefinido":
            artigo = artigo_norma(s.get("enquadramento_salic"))
        captado = s.get("valor_captado")
        try:
            captado_f = float(captado) if captado is not None else None
        except (TypeError, ValueError):
            captado_f = None
        teto = rec.get("teto_dou")
        try:
            teto_f = float(teto) if teto is not None and not (isinstance(teto, float) and pd.isna(teto)) else None
        except (TypeError, ValueError):
            teto_f = None
        aprov_salic = s.get("valor_aprovado")
        try:
            aprov_salic_f = float(aprov_salic) if aprov_salic is not None else None
        except (TypeError, ValueError):
            aprov_salic_f = None
        base = teto_f if teto_f and teto_f > 0 else aprov_salic_f
        taxa = (captado_f / base) if captado_f is not None and base and base > 0 else None
        aliq = aliquota_base(artigo)
        renuncia = (captado_f * aliq) if captado_f is not None else None
        renuncia_40 = (captado_f * ALIQUOTA_ART26_DOACAO_PJ) if captado_f is not None and artigo == "artigo_26" else (
            captado_f if artigo == "artigo_18" and captado_f is not None else None
        )
        rows.append(
            {
                "pronac": p,
                "nome_projeto": rec.get("nome_projeto") or s.get("nome"),
                "proponente": rec.get("proponente") or s.get("proponente"),
                "uf": rec.get("uf") or s.get("UF"),
                "cidade": rec.get("cidade") or s.get("municipio"),
                "area": rec.get("area"),
                "artigo": artigo,
                "numero_portaria": rec.get("numero_portaria"),
                "data_portaria": rec.get("data_portaria"),
                "data_publicacao": rec.get("data_publicacao").date().isoformat()
                if hasattr(rec.get("data_publicacao"), "date")
                else rec.get("data_publicacao"),
                "prazo_inicio": rec.get("prazo_inicio"),
                "prazo_fim": rec.get("prazo_fim"),
                "url_dou": rec.get("url"),
                "teto_dou": teto_f,
                "valor_aprovado_salic": aprov_salic_f,
                "valor_captado_salic": captado_f,
                "taxa_captacao": taxa,
                "situacao_salic": s.get("situacao"),
                "enquadramento_salic": s.get("enquadramento_salic"),
                "encontrado_salic": bool(s) and not s.get("erro"),
                "aliquota_renuncia": aliq if artigo in {"artigo_18", "artigo_26"} else None,
                "renuncia_estimada": renuncia,
                "renuncia_art26_doacao_pj": renuncia_40 if artigo == "artigo_26" else None,
            }
        )
    return pd.DataFrame(rows)


def _fmt_brl(valor: float | None) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    s = f"{valor:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(valor: float | None) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def agregar(cruzado: pd.DataFrame, col: str) -> pd.DataFrame:
    g = (
        cruzado.groupby(col, dropna=False)
        .agg(
            projetos=("pronac", "count"),
            com_captacao=("valor_captado_salic", lambda s: int((s.fillna(0) > 0).sum())),
            teto_dou=("teto_dou", "sum"),
            captado=("valor_captado_salic", "sum"),
            renuncia_estimada=("renuncia_estimada", "sum"),
        )
        .reset_index()
    )
    g["taxa_captacao"] = g["captado"] / g["teto_dou"].replace(0, pd.NA)
    return g.sort_values("teto_dou", ascending=False)


def escrever_markdown(
    cruzado: pd.DataFrame,
    por_uf: pd.DataFrame,
    por_mes: pd.DataFrame,
    por_artigo: pd.DataFrame,
    path: Path,
    gerado_em: str,
) -> None:
    n = len(cruzado)
    n_ok = int(cruzado["encontrado_salic"].fillna(False).sum()) if n else 0
    n_cap = int((cruzado["valor_captado_salic"].fillna(0) > 0).sum()) if n else 0
    teto = float(cruzado["teto_dou"].fillna(0).sum()) if n else 0.0
    captado = float(cruzado["valor_captado_salic"].fillna(0).sum()) if n else 0.0
    renuncia = float(cruzado["renuncia_estimada"].fillna(0).sum()) if n else 0.0
    linhas = [
        "# Renúncia Lei Rouanet — homologações DOU 2026 × SALIC",
        "",
        f"**Gerado em:** {gerado_em}",
        "",
        "Cruza os projetos das Portarias SEFIC/MinC de **homologação inicial** "
        "(fase de doações e patrocínios no DOU) com o cadastro aberto do "
        "[SALIC](https://api.salic.cultura.gov.br/docs).",
        "",
        "**Teto** = valor aprovado na Portaria (autorização para captar). "
        "**Captado** = inflow já lançado no SALIC. A janela de captação da "
        "maioria dos projetos de 2026 segue aberta até 31/12/2026; o captado "
        "é estoque até a data da consulta, não o exercício fechado.",
        "",
        "Renúncia estimada (benefício específico da Lei nº 8.313/1991):",
        "",
        "- art. 18: 100% do captado;",
        "- art. 26: 30% do captado (patrocínio de pessoa jurídica, art. 26, II). "
        "Doação PJ seria 40% (inciso I); PF teria 60%/80%. Sem o tipo de cada "
        "incentivador, o cenário-base usa 30% no art. 26.",
        "",
        "## Totais",
        "",
        f"- Projetos homologados (PRONAC únicos): **{n}**",
        f"- Encontrados no SALIC: **{n_ok}**",
        f"- Com alguma captação lançada: **{n_cap}**",
        f"- Teto homologado no DOU: **{_fmt_brl(teto)}**",
        f"- Já captado (SALIC): **{_fmt_brl(captado)}** ({_fmt_pct(captado / teto if teto else None)} do teto)",
        f"- Renúncia estimada sobre o captado: **{_fmt_brl(renuncia)}**",
        "",
        "### Por artigo",
        "",
        "| Artigo | Projetos | Com captação | Teto DOU | Captado SALIC | Renúncia estimada |",
        "|--------|---------:|-------------:|---------:|--------------:|------------------:|",
    ]
    for row in por_artigo.itertuples(index=False):
        linhas.append(
            f"| `{row.artigo}` | {int(row.projetos)} | {int(row.com_captacao)} | "
            f"{_fmt_brl(row.teto_dou)} | {_fmt_brl(row.captado)} | {_fmt_brl(row.renuncia_estimada)} |"
        )
    linhas.extend(
        [
            "",
            "### Por UF (teto)",
            "",
            "| UF | Projetos | Com captação | Teto DOU | Captado | Renúncia estimada | Taxa |",
            "|----|---------:|-------------:|---------:|--------:|------------------:|-----:|",
        ]
    )
    for row in por_uf.head(15).itertuples(index=False):
        linhas.append(
            f"| {row.uf or '—'} | {int(row.projetos)} | {int(row.com_captacao)} | "
            f"{_fmt_brl(row.teto_dou)} | {_fmt_brl(row.captado)} | "
            f"{_fmt_brl(row.renuncia_estimada)} | {_fmt_pct(row.taxa_captacao)} |"
        )
    linhas.extend(
        [
            "",
            "### Por mês de publicação no DOU",
            "",
            "| Mês | Projetos | Teto DOU | Captado | Renúncia estimada |",
            "|-----|---------:|---------:|--------:|------------------:|",
        ]
    )
    for row in por_mes.itertuples(index=False):
        linhas.append(
            f"| {row.mes} | {int(row.projetos)} | {_fmt_brl(row.teto_dou)} | "
            f"{_fmt_brl(row.captado)} | {_fmt_brl(row.renuncia_estimada)} |"
        )
    top = cruzado.sort_values("valor_captado_salic", ascending=False).head(15)
    linhas.extend(
        [
            "",
            "## Maiores captações já lançadas no SALIC",
            "",
            "| PRONAC | Projeto | UF | Captado | Teto | Art. | Situação |",
            "|-------:|---------|----|--------:|-----:|------|----------|",
        ]
    )
    if not top.empty:
        for row in top.itertuples(index=False):
            nome = str(row.nome_projeto or "")[:70]
            linhas.append(
                f"| {row.pronac} | {nome} | {row.uf or ''} | {_fmt_brl(row.valor_captado_salic)} | "
                f"{_fmt_brl(row.teto_dou)} | {row.artigo} | {str(row.situacao_salic or '')[:40]} |"
            )
    linhas.extend(
        [
            "",
            "## Arquivos",
            "",
            "- `projetos_lei_rouanet_salic.csv` — um PRONAC por linha (DOU + SALIC)",
            "- `renuncia_lei_rouanet_por_uf.csv` / `_por_mes.csv` / `_por_artigo.csv`",
            "- `renuncia_lei_rouanet.xlsx`",
            "",
        ]
    )
    path.write_text("\n".join(linhas), encoding="utf-8")


def processar(
    entrada: Path,
    saida_dir: Path,
    *,
    cache_path: Path,
    somente_homologacao: bool = True,
    workers: int = 8,
    max_pronacs: int | None = None,
    salic_map: dict[str, dict[str, Any]] | None = None,
) -> pd.DataFrame:
    saida_dir.mkdir(parents=True, exist_ok=True)
    df = pd.read_csv(entrada)
    dou = projetos_dou_unicos(df, somente_homologacao=somente_homologacao)
    if max_pronacs:
        dou = dou.head(max_pronacs)
    pronacs = [str(p) for p in dou["pronac"].tolist()]
    print(f"[DOU] {len(pronacs)} PRONACs únicos", flush=True)
    if salic_map is None:
        salic_map = coletar_salic(pronacs, cache_path=cache_path, workers=workers)
    cruzado = cruzar(dou, salic_map)
    cruzado["mes"] = pd.to_datetime(cruzado["data_publicacao"], errors="coerce").dt.to_period("M").astype(str)
    por_uf = agregar(cruzado, "uf")
    por_artigo = agregar(cruzado, "artigo")
    por_mes = agregar(cruzado, "mes").sort_values("mes")
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    cruzado.to_csv(saida_dir / "projetos_lei_rouanet_salic.csv", index=False)
    por_uf.to_csv(saida_dir / "renuncia_lei_rouanet_por_uf.csv", index=False)
    por_mes.to_csv(saida_dir / "renuncia_lei_rouanet_por_mes.csv", index=False)
    por_artigo.to_csv(saida_dir / "renuncia_lei_rouanet_por_artigo.csv", index=False)
    escrever_markdown(
        cruzado, por_uf, por_mes, por_artigo,
        saida_dir / "renuncia_lei_rouanet.md", gerado,
    )
    with pd.ExcelWriter(saida_dir / "renuncia_lei_rouanet.xlsx", engine="openpyxl") as xl:
        cruzado.to_excel(xl, sheet_name="projetos", index=False)
        por_uf.to_excel(xl, sheet_name="por_uf", index=False)
        por_mes.to_excel(xl, sheet_name="por_mes", index=False)
        por_artigo.to_excel(xl, sheet_name="por_artigo", index=False)
    print(f"[OK] {saida_dir / 'renuncia_lei_rouanet.md'}", flush=True)
    return cruzado


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--entrada", type=Path, default=ROOT / "output" / "projetos_lei_rouanet_captacao.csv")
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    p.add_argument("--cache", type=Path, default=ROOT / "output" / ".cache" / "salic_projetos.jsonl")
    p.add_argument("--workers", type=int, default=8)
    p.add_argument("--max-pronacs", type=int, default=None)
    p.add_argument("--todos-tipos", action="store_true", help="Inclui prorrogação/complementação além da homologação inicial")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not args.entrada.exists():
        print(f"ERRO: não achei {args.entrada}", file=sys.stderr)
        return 2
    try:
        processar(
            args.entrada,
            args.saida_dir,
            cache_path=args.cache,
            somente_homologacao=not args.todos_tipos,
            workers=args.workers,
            max_pronacs=args.max_pronacs,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
