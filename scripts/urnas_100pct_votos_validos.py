#!/usr/bin/env python3
"""Lista urnas em que um único candidato teve 100% dos votos válidos.

Varre os discriminativos por urna de 2014, 2018 e 2022 (1º e 2º turnos).
Critério: há votos válidos, só um candidato recebeu votos, e esses votos
coincidem com QT_VOTOS_VALIDOS.

Saída:
  output/tse_planilhas/urnas_100pct_votos_validos.xlsx

Uso:
  python3 scripts/urnas_100pct_votos_validos.py
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

try:
    from scripts.planilha_resultados_presidente import (
        SKIP_VOTOS,
        escrever_xlsx,
        pastas_saida,
    )
except ImportError:
    from planilha_resultados_presidente import (  # type: ignore
        SKIP_VOTOS,
        escrever_xlsx,
        pastas_saida,
    )

ANOS = (2014, 2018, 2022)
SKIP_EXTRA = SKIP_VOTOS | {"PT", "OPP"}
CHAVES = (
    "ANO_ELEICAO",
    "NR_TURNO",
    "REGIAO",
    "SG_UF",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "NR_LOCAL_VOTACAO",
    "NR_URNA_EFETIVADA",
    "NR_MODELO",
    "DS_MODELO_URNA",
    "QT_APTOS",
    "QT_COMPARECIMENTO",
    "QT_VOTOS_BRANCO",
    "QT_VOTOS_NULO",
    "QT_VOTOS_VALIDOS",
)


def colunas_candidatos(df: pd.DataFrame) -> list[str]:
    out: list[str] = []
    for col in df.columns:
        if not str(col).startswith("QT_VOTOS_"):
            continue
        if str(col)[len("QT_VOTOS_") :] in SKIP_EXTRA:
            continue
        out.append(str(col))
    return out


def filtrar_100pct(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    cols = colunas_candidatos(df)
    if not cols or "QT_VOTOS_VALIDOS" not in df.columns:
        return pd.DataFrame()
    trabalho = df.copy()
    for c in cols + ["QT_VOTOS_VALIDOS"]:
        trabalho[c] = pd.to_numeric(trabalho[c], errors="coerce").fillna(0)
    votos = trabalho[cols]
    validos = trabalho["QT_VOTOS_VALIDOS"]
    mx = votos.max(axis=1)
    n_pos = (votos > 0).sum(axis=1)
    hit = trabalho.loc[(validos > 0) & (n_pos == 1) & (mx == validos)].copy()
    if hit.empty:
        return hit
    idx = hit[cols].to_numpy().argmax(axis=1)
    hit["CANDIDATO"] = [cols[i][len("QT_VOTOS_") :] for i in idx]
    hit["QT_VOTOS_CANDIDATO"] = mx.loc[hit.index].astype(int)
    hit["ANO_ELEICAO"] = ano
    hit["NR_TURNO"] = turno
    keep = [c for c in CHAVES if c in hit.columns] + ["CANDIDATO", "QT_VOTOS_CANDIDATO"]
    return hit[keep]


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dados", type=Path, default=None)
    p.add_argument("--saida", type=Path, default=None)
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    dados = Path(args.dados) if args.dados else pastas_saida()
    saida = Path(args.saida) if args.saida else pastas_saida()
    saida.mkdir(parents=True, exist_ok=True)

    partes: list[pd.DataFrame] = []
    for ano in ANOS:
        for turno in (1, 2):
            path = dados / f"discriminativo_urnas_{ano}_{turno}t.csv.gz"
            if not path.exists():
                path = saida / path.name
            if not path.exists():
                print(f"  falta {path.name}", flush=True)
                continue
            df = pd.read_csv(path, low_memory=False)
            hit = filtrar_100pct(df, ano, turno)
            print(f"  {ano} T{turno}: {len(hit):,} urnas ← {path.name}", flush=True)
            if not hit.empty:
                partes.append(hit)

    if not partes:
        raise FileNotFoundError("Nenhuma urna 100% encontrada. Gere os discriminativos por urna.")
    out = pd.concat(partes, ignore_index=True)
    sort_cols = [c for c in ("ANO_ELEICAO", "NR_TURNO", "SG_UF", "NM_MUNICIPIO", "NR_ZONA", "NR_SECAO") if c in out.columns]
    out = out.sort_values(sort_cols, kind="mergesort").reset_index(drop=True)
    por_pleito = (
        out.groupby(["ANO_ELEICAO", "NR_TURNO", "CANDIDATO"], dropna=False)
        .agg(URNAS=("NR_SECAO", "size"), VOTOS=("QT_VOTOS_CANDIDATO", "sum"))
        .reset_index()
    )
    por_uf = (
        out.groupby(["ANO_ELEICAO", "NR_TURNO", "SG_UF", "CANDIDATO"], dropna=False)
        .agg(URNAS=("NR_SECAO", "size"), VOTOS=("QT_VOTOS_CANDIDATO", "sum"))
        .reset_index()
    )
    dest = saida / "urnas_100pct_votos_validos.xlsx"
    escrever_xlsx(
        dest,
        [
            ("Conteúdo", "Urnas com 100% dos votos válidos para um único candidato"),
            (
                "Critério",
                "Há votos válidos, só um candidato recebeu votos, e esses votos = 100% dos válidos",
            ),
            ("Pleitos", "2014, 2018 e 2022 — 1º e 2º turnos"),
            ("Urnas", f"{len(out):,}".replace(",", ".")),
            ("Votos válidos", f"{int(out['QT_VOTOS_VALIDOS'].sum()):,}".replace(",", ".")),
        ],
        {"Urnas": out, "Por_pleito": por_pleito, "Por_UF": por_uf},
    )
    print(f"  {dest} ({dest.stat().st_size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
