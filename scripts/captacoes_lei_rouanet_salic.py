#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Extrato de captações da Lei Rouanet (SALIC), 2003–2026.

Uma linha por recibo de captação. O beneficiário é o **proponente** do
projeto (quem recebe o recurso). O incentivador (doador) vem no recibo.

Fonte: API SALIC ``https://api.salic.cultura.gov.br/api/v1``

1. Lista projetos com ``sort=valor_captado:desc`` até o primeiro zero.
2. Para cada PRONAC com captação, lê ``GET /projetos/{PRONAC}`` e usa
   ``_embedded.captacoes`` (data_recibo, valor, CNPJ/CPF do doador).

Uso::

  python scripts/captacoes_lei_rouanet_salic.py
  python scripts/captacoes_lei_rouanet_salic.py --max-projetos 20
"""

from __future__ import annotations

import argparse
import gzip
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

SALIC = "https://api.salic.cultura.gov.br/api/v1"
UA = "SEC-data-analysys/captacoes-lei-rouanet-salic"
INICIO_PADRAO = "2003-01-01"
FIM_PADRAO = "2026-12-31"
LISTA_LIMITE = 100

_RATE = threading.Lock()
_LAST = 0.0
_MIN_INTERVAL = 0.03

GetJson = Callable[[str], dict[str, Any]]


def _throttle() -> None:
    global _LAST
    with _RATE:
        wait = _MIN_INTERVAL - (time.monotonic() - _LAST)
        if wait > 0:
            time.sleep(wait)
        _LAST = time.monotonic()


def _get_json(url: str, timeout: int = 120, tentativas: int = 5) -> dict[str, Any]:
    last: Exception | None = None
    for i in range(tentativas):
        try:
            _throttle()
            req = Request(
                url,
                headers={
                    "User-Agent": UA,
                    "Accept": "application/json",
                    "Accept-Encoding": "gzip",
                },
            )
            with urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
                encoding = (resp.headers.get("Content-Encoding") or "").lower()
                if "gzip" in encoding or raw[:2] == b"\x1f\x8b":
                    raw = gzip.decompress(raw)
                return json.loads(raw.decode("utf-8", errors="replace"))
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
    s = re.sub(r"[^\d]", "", s)
    return s or None


def norm_cnpj_cpf(valor: Any) -> str | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    s = str(valor).strip()
    if not s or s.lower() == "nan":
        return None
    digits = re.sub(r"\D", "", s)
    if len(digits) in (11, 14):
        return digits
    return s or None


def norm_data(valor: Any) -> str | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    s = str(valor).strip()
    if len(s) >= 10 and re.match(r"\d{4}-\d{2}-\d{2}", s):
        return s[:10]
    return None


def _num(valor: Any) -> float:
    try:
        if valor is None or valor == "":
            return 0.0
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def slim_projeto_lista(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "PRONAC": norm_pronac(raw.get("PRONAC")),
        "nome": (str(raw.get("nome") or "").strip() or None),
        "cgccpf": norm_cnpj_cpf(raw.get("cgccpf")),
        "proponente": (str(raw.get("proponente") or "").strip() or None),
        "UF": raw.get("UF"),
        "municipio": raw.get("municipio"),
        "situacao": raw.get("situacao"),
        "ano_projeto": raw.get("ano_projeto"),
        "valor_captado": _num(raw.get("valor_captado")),
        "valor_aprovado": _num(raw.get("valor_aprovado")),
    }


def slim_recibo(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "PRONAC": norm_pronac(raw.get("PRONAC")),
        "valor": _num(raw.get("valor")),
        "data_recibo": norm_data(raw.get("data_recibo")),
        "nome_projeto": (str(raw.get("nome_projeto") or "").strip() or None),
        "cgccpf_doador": norm_cnpj_cpf(raw.get("cgccpf")),
        "nome_doador": (str(raw.get("nome_doador") or "").strip() or None),
    }


def no_periodo(data: str | None, inicio: str, fim: str) -> bool:
    if not data:
        return False
    return inicio <= data <= fim


def url_lista(*, offset: int, limit: int = LISTA_LIMITE, sort: bool = True) -> str:
    q: dict[str, Any] = {"limit": limit, "offset": offset}
    if sort:
        q["sort"] = "valor_captado:desc"
    return f"{SALIC}/projetos?{urlencode(q)}"


def url_detalhe(pronac: str) -> str:
    return f"{SALIC}/projetos/{quote(str(pronac), safe='')}"


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


def cache_append(path: Path, rec: dict[str, Any], lock: threading.Lock | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(rec, ensure_ascii=False) + "\n"
    if lock is None:
        with path.open("a", encoding="utf-8") as fh:
            fh.write(payload)
        return
    with lock:
        with path.open("a", encoding="utf-8") as fh:
            fh.write(payload)


def contar_com_captacao(get_json: GetJson | None = None) -> tuple[int, int]:
    """Retorna (qtd com valor_captado>0, total de projetos) via busca binária."""
    get = get_json or _get_json
    primeiro = get(url_lista(offset=0, limit=1))
    total = int(primeiro.get("total") or 0)
    if total <= 0:
        return 0, 0
    lista0 = (primeiro.get("_embedded") or {}).get("projetos") or []
    if not lista0 or _num(lista0[0].get("valor_captado")) <= 0:
        return 0, total

    lo, hi = 0, total
    while lo < hi:
        mid = (lo + hi) // 2
        payload = get(url_lista(offset=mid, limit=1))
        projetos = (payload.get("_embedded") or {}).get("projetos") or []
        valor = _num(projetos[0].get("valor_captado")) if projetos else 0.0
        if valor > 0:
            lo = mid + 1
        else:
            hi = mid
    return lo, total


def listar_projetos_com_captacao(
    *,
    get_json: GetJson | None = None,
    workers: int = 8,
    max_projetos: int | None = None,
    n_com_captacao: int | None = None,
) -> list[dict[str, Any]]:
    get = get_json or _get_json
    if n_com_captacao is None:
        n_com_captacao, total = contar_com_captacao(get)
        print(f"[SALIC] projetos_total={total} com_captacao≈{n_com_captacao}", flush=True)
    else:
        print(f"[SALIC] com_captacao≈{n_com_captacao} (informado)", flush=True)
    if n_com_captacao <= 0:
        return []
    limite = n_com_captacao
    if max_projetos:
        limite = min(limite, max_projetos)
    offsets = list(range(0, limite, LISTA_LIMITE))
    por_offset: dict[int, list[dict[str, Any]]] = {}

    def _pagina(offset: int) -> tuple[int, list[dict[str, Any]]]:
        payload = get(url_lista(offset=offset, limit=LISTA_LIMITE))
        raw = (payload.get("_embedded") or {}).get("projetos") or []
        slim = [slim_projeto_lista(p) for p in raw]
        slim = [p for p in slim if p.get("PRONAC") and (p.get("valor_captado") or 0) > 0]
        return offset, slim

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_pagina, off): off for off in offsets}
        feitos = 0
        for fut in as_completed(futs):
            off, slim = fut.result()
            por_offset[off] = slim
            feitos += 1
            if feitos % 10 == 0 or feitos == len(offsets):
                print(f"  lista {feitos}/{len(offsets)} páginas", flush=True)

    projetos: list[dict[str, Any]] = []
    vistos: set[str] = set()
    for off in offsets:
        for p in por_offset.get(off, []):
            pid = p["PRONAC"]
            if pid in vistos:
                continue
            vistos.add(pid)
            projetos.append(p)
            if max_projetos and len(projetos) >= max_projetos:
                return projetos
    return projetos


def extrair_captacoes_do_detalhe(
    payload: dict[str, Any],
    projeto: dict[str, Any] | None = None,
) -> dict[str, Any]:
    proj = projeto or {}
    pronac = norm_pronac(payload.get("PRONAC")) or proj.get("PRONAC")
    rec = {
        "PRONAC": pronac,
        "nome": (str(payload.get("nome") or "").strip() or proj.get("nome")),
        "cgccpf": norm_cnpj_cpf(payload.get("cgccpf")) or proj.get("cgccpf"),
        "proponente": (str(payload.get("proponente") or "").strip() or proj.get("proponente")),
        "UF": payload.get("UF") or proj.get("UF"),
        "valor_captado": _num(payload.get("valor_captado")) or _num(proj.get("valor_captado")),
        "captacoes": [
            slim_recibo(c) for c in ((payload.get("_embedded") or {}).get("captacoes") or [])
        ],
    }
    return rec


def buscar_captacoes_pronac(
    projeto: dict[str, Any],
    *,
    get_json: GetJson | None = None,
) -> dict[str, Any]:
    get = get_json or _get_json
    pronac = projeto["PRONAC"]
    payload = get(url_detalhe(pronac))
    return extrair_captacoes_do_detalhe(payload, projeto)


def coletar_captacoes(
    projetos: Iterable[dict[str, Any]],
    *,
    cache_path: Path,
    workers: int = 8,
    get_json: GetJson | None = None,
) -> dict[str, dict[str, Any]]:
    cache = cache_carregar(cache_path)
    lista = list(projetos)
    faltam = [p for p in lista if str(p.get("PRONAC")) not in cache]
    print(f"[SALIC] cache_captacoes={len(cache)} a_buscar={len(faltam)}", flush=True)
    if not faltam:
        return cache
    lock = threading.Lock()
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(buscar_captacoes_pronac, p, get_json=get_json): p for p in faltam}
        feitos = 0
        for fut in as_completed(futs):
            p = futs[fut]
            feitos += 1
            pid = str(p.get("PRONAC"))
            try:
                rec = fut.result()
            except Exception as exc:
                print(f"  aviso PRONAC {pid}: {exc}", flush=True)
                rec = {"PRONAC": pid, "erro": str(exc), "captacoes": []}
            with lock:
                cache[pid] = rec
            cache_append(cache_path, rec, lock)
            if feitos % 50 == 0 or feitos == len(faltam):
                print(f"  detalhe {feitos}/{len(faltam)}", flush=True)
    return cache


def recibos_periodo(
    cache: dict[str, dict[str, Any]],
    projetos: list[dict[str, Any]] | None = None,
    *,
    inicio: str = INICIO_PADRAO,
    fim: str = FIM_PADRAO,
) -> pd.DataFrame:
    por_pronac = {str(p["PRONAC"]): p for p in (projetos or []) if p.get("PRONAC")}
    rows: list[dict[str, Any]] = []
    for pid, rec in cache.items():
        if rec.get("erro") and not rec.get("captacoes"):
            continue
        base = por_pronac.get(str(pid), {})
        benef_doc = rec.get("cgccpf") or base.get("cgccpf")
        benef_nome = rec.get("proponente") or base.get("proponente")
        nome_proj = rec.get("nome") or base.get("nome")
        uf = rec.get("UF") or base.get("UF")
        for cap in rec.get("captacoes") or []:
            data = cap.get("data_recibo")
            if not no_periodo(data, inicio, fim):
                continue
            rows.append(
                {
                    "pronac": cap.get("PRONAC") or pid,
                    "nome_projeto": cap.get("nome_projeto") or nome_proj,
                    "cnpj_cpf_beneficiario": benef_doc,
                    "nome_beneficiario": benef_nome,
                    "uf": uf,
                    "cnpj_cpf_incentivador": cap.get("cgccpf_doador"),
                    "nome_incentivador": cap.get("nome_doador"),
                    "valor_captado": cap.get("valor") or 0.0,
                    "data_captacao": data,
                }
            )
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(
            columns=[
                "pronac",
                "nome_projeto",
                "cnpj_cpf_beneficiario",
                "nome_beneficiario",
                "uf",
                "cnpj_cpf_incentivador",
                "nome_incentivador",
                "valor_captado",
                "data_captacao",
            ]
        )
    df["data_captacao"] = df["data_captacao"].astype(str)
    df = df.sort_values(["data_captacao", "pronac", "valor_captado"], ascending=[True, True, False])
    return df.reset_index(drop=True)


def _fmt_brl(valor: Any) -> str:
    try:
        n = float(valor or 0)
    except (TypeError, ValueError):
        n = 0.0
    inteiro, frac = f"{n:,.2f}".split(".")
    inteiro = inteiro.replace(",", ".")
    return f"R$ {inteiro},{frac}"


def agregar_por_ano(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["ano", "recibos", "projetos", "beneficiarios", "valor_captado"])
    tmp = df.copy()
    tmp["ano"] = tmp["data_captacao"].str.slice(0, 4)
    g = tmp.groupby("ano", dropna=False).agg(
        recibos=("valor_captado", "size"),
        projetos=("pronac", "nunique"),
        beneficiarios=("cnpj_cpf_beneficiario", "nunique"),
        valor_captado=("valor_captado", "sum"),
    )
    return g.reset_index().sort_values("ano")


def agregar_por_beneficiario(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=["cnpj_cpf_beneficiario", "nome_beneficiario", "recibos", "projetos", "valor_captado"]
        )
    g = df.groupby(["cnpj_cpf_beneficiario", "nome_beneficiario"], dropna=False).agg(
        recibos=("valor_captado", "size"),
        projetos=("pronac", "nunique"),
        valor_captado=("valor_captado", "sum"),
    )
    return g.reset_index().sort_values("valor_captado", ascending=False)


def escrever_markdown(
    df: pd.DataFrame,
    por_ano: pd.DataFrame,
    por_benef: pd.DataFrame,
    path: Path,
    gerado: str,
    *,
    inicio: str,
    fim: str,
    n_projetos_salic: int,
) -> None:
    total = float(df["valor_captado"].sum()) if not df.empty else 0.0
    n_rec = len(df)
    n_proj = int(df["pronac"].nunique()) if not df.empty else 0
    n_ben = int(df["cnpj_cpf_beneficiario"].nunique()) if not df.empty else 0
    linhas = [
        "# Captações da Lei Rouanet (SALIC), 2003–2026",
        "",
        f"Gerado em {gerado}.",
        "",
        "Uma linha por **recibo de captação** (`data_recibo` no SALIC).",
        "Beneficiário = **proponente** do projeto (CNPJ/CPF de quem recebe).",
        "Incentivador = doador registrado no recibo.",
        "",
        f"- Período dos recibos: **{inicio}** a **{fim}**",
        f"- Projetos SALIC com `valor_captado > 0`: **{n_projetos_salic}**",
        f"- Recibos no período: **{n_rec:,}**".replace(",", "."),
        f"- Projetos com recibo no período: **{n_proj:,}**".replace(",", "."),
        f"- Beneficiários distintos: **{n_ben:,}**".replace(",", "."),
        f"- Valor captado no período: **{_fmt_brl(total)}**",
        "",
        "## Por ano do recibo",
        "",
        "| Ano | Recibos | Projetos | Beneficiários | Valor captado |",
        "|----:|--------:|---------:|--------------:|--------------:|",
    ]
    for row in por_ano.itertuples(index=False):
        linhas.append(
            f"| {row.ano} | {int(row.recibos)} | {int(row.projetos)} | "
            f"{int(row.beneficiarios)} | {_fmt_brl(row.valor_captado)} |"
        )
    top = por_benef.head(20)
    linhas.extend(
        [
            "",
            "## Maiores beneficiários (proponentes) no período",
            "",
            "| CNPJ/CPF | Nome | Recibos | Projetos | Valor captado |",
            "|----------|------|--------:|---------:|--------------:|",
        ]
    )
    for row in top.itertuples(index=False):
        nome = str(row.nome_beneficiario or "")[:70]
        linhas.append(
            f"| {row.cnpj_cpf_beneficiario or '—'} | {nome} | {int(row.recibos)} | "
            f"{int(row.projetos)} | {_fmt_brl(row.valor_captado)} |"
        )
    linhas.extend(
        [
            "",
            "## Arquivos",
            "",
            "- `captacoes_lei_rouanet_2003_2026.csv` — um recibo por linha",
            "- `captacoes_lei_rouanet_por_ano.csv`",
            "- `captacoes_lei_rouanet_por_beneficiario.csv`",
            "- `captacoes_lei_rouanet_2003_2026.xlsx`",
            "",
            "Fonte: [API SALIC](https://api.salic.cultura.gov.br/docs) "
            "(`/projetos` + `/projetos/{PRONAC}` `_embedded.captacoes`).",
            "",
        ]
    )
    path.write_text("\n".join(linhas), encoding="utf-8")


def processar(
    saida_dir: Path,
    *,
    cache_lista: Path,
    cache_captacoes: Path,
    workers: int = 8,
    max_projetos: int | None = None,
    inicio: str = INICIO_PADRAO,
    fim: str = FIM_PADRAO,
    get_json: GetJson | None = None,
    projetos: list[dict[str, Any]] | None = None,
    captacoes_map: dict[str, dict[str, Any]] | None = None,
) -> pd.DataFrame:
    saida_dir.mkdir(parents=True, exist_ok=True)
    if projetos is None:
        cached = cache_carregar(cache_lista)
        if cached and not max_projetos:
            projetos = list(cached.values())
            print(f"[SALIC] lista em cache ({len(projetos)} projetos)", flush=True)
        else:
            projetos = listar_projetos_com_captacao(
                get_json=get_json, workers=workers, max_projetos=max_projetos
            )
            if max_projetos is None:
                cache_lista.parent.mkdir(parents=True, exist_ok=True)
                with cache_lista.open("w", encoding="utf-8") as fh:
                    for p in projetos:
                        fh.write(json.dumps(p, ensure_ascii=False) + "\n")
    if captacoes_map is None:
        captacoes_map = coletar_captacoes(
            projetos, cache_path=cache_captacoes, workers=workers, get_json=get_json
        )
    df = recibos_periodo(captacoes_map, projetos, inicio=inicio, fim=fim)
    por_ano = agregar_por_ano(df)
    por_benef = agregar_por_beneficiario(df)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    csv_path = saida_dir / "captacoes_lei_rouanet_2003_2026.csv"
    df.to_csv(csv_path, index=False)
    por_ano.to_csv(saida_dir / "captacoes_lei_rouanet_por_ano.csv", index=False)
    por_benef.to_csv(saida_dir / "captacoes_lei_rouanet_por_beneficiario.csv", index=False)
    escrever_markdown(
        df,
        por_ano,
        por_benef,
        saida_dir / "captacoes_lei_rouanet_2003_2026.md",
        gerado,
        inicio=inicio,
        fim=fim,
        n_projetos_salic=len(projetos),
    )
    amostra = df.head(10_000)
    with pd.ExcelWriter(saida_dir / "captacoes_lei_rouanet_2003_2026.xlsx", engine="openpyxl") as xl:
        por_ano.to_excel(xl, sheet_name="por_ano", index=False)
        por_benef.head(5_000).to_excel(xl, sheet_name="por_beneficiario", index=False)
        amostra.to_excel(xl, sheet_name="amostra_recibos", index=False)
    print(
        f"[OK] {len(df)} recibos | {df['pronac'].nunique() if not df.empty else 0} projetos | "
        f"{csv_path}",
        flush=True,
    )
    return df


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    p.add_argument(
        "--cache-lista",
        type=Path,
        default=ROOT / "output" / ".cache" / "salic_projetos_com_captacao.jsonl",
    )
    p.add_argument(
        "--cache-captacoes",
        type=Path,
        default=ROOT / "output" / ".cache" / "salic_captacoes.jsonl",
    )
    p.add_argument("--workers", type=int, default=8)
    p.add_argument("--max-projetos", type=int, default=None)
    p.add_argument("--inicio", default=INICIO_PADRAO)
    p.add_argument("--fim", default=FIM_PADRAO)
    p.add_argument(
        "--refazer-lista",
        action="store_true",
        help="Ignora o cache da lista de projetos com captação",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.refazer_lista and args.cache_lista.exists():
            args.cache_lista.unlink()
        processar(
            args.saida_dir,
            cache_lista=args.cache_lista,
            cache_captacoes=args.cache_captacoes,
            workers=args.workers,
            max_projetos=args.max_projetos,
            inicio=args.inicio,
            fim=args.fim,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
