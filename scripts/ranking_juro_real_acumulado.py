#!/usr/bin/env python3
"""Ranking da taxa básica de juros real acumulada (ex post, Fisher).

Fonte: BIS (taxa de política monetária mensal, fim de período; CPI mensal).
Juro nominal acumulado: Π (1 + i_m/100)^(1/12) − 1.
Inflação acumulada: CPI(dezembro final) / CPI(dezembro anterior ao início) − 1.
Juro real (Fisher): (1 + i) / (1 + π) − 1.
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

CPI_INDEX = 628
ANO_INICIO_DEFAULT = 2019
ANO_FIM_DEFAULT = 2022

NOMES = {
    "AR": "Argentina",
    "AU": "Austrália",
    "BR": "Brasil",
    "CA": "Canadá",
    "CH": "Suíça",
    "CL": "Chile",
    "CN": "China",
    "CO": "Colômbia",
    "CZ": "Chéquia",
    "DK": "Dinamarca",
    "GB": "Reino Unido",
    "HK": "Hong Kong",
    "HR": "Croácia",
    "HU": "Hungria",
    "ID": "Indonésia",
    "IL": "Israel",
    "IN": "Índia",
    "IS": "Islândia",
    "JP": "Japão",
    "KR": "Coreia do Sul",
    "KW": "Kuwait",
    "MA": "Marrocos",
    "MK": "Macedônia do Norte",
    "MX": "México",
    "MY": "Malásia",
    "NO": "Noruega",
    "NZ": "Nova Zelândia",
    "PE": "Peru",
    "PH": "Filipinas",
    "PL": "Polônia",
    "RO": "Romênia",
    "RS": "Sérvia",
    "RU": "Rússia",
    "SA": "Arábia Saudita",
    "SE": "Suécia",
    "TH": "Tailândia",
    "TR": "Turquia",
    "US": "Estados Unidos",
    "XM": "Área do euro",
    "ZA": "África do Sul",
}


def periodo(ano_inicio: int, ano_fim: int) -> tuple[pd.DatetimeIndex, pd.Timestamp, pd.Timestamp]:
    meses = pd.date_range(f"{ano_inicio}-01-01", f"{ano_fim}-12-01", freq="MS")
    cpi_ini = pd.Timestamp(year=ano_inicio - 1, month=12, day=1)
    cpi_fim = pd.Timestamp(year=ano_fim, month=12, day=1)
    return meses, cpi_ini, cpi_fim


def url_cbpol(ano_inicio: int, ano_fim: int) -> str:
    return (
        "https://stats.bis.org/api/v1/data/WS_CBPOL/M..?"
        f"startPeriod={ano_inicio-1}-12&endPeriod={ano_fim}-12&format=csvfile"
    )


def url_cpi(ano_inicio: int, ano_fim: int) -> str:
    return (
        "https://stats.bis.org/api/v1/data/WS_LONG_CPI/M..?"
        f"startPeriod={ano_inicio-1}-12&endPeriod={ano_fim}-12&format=csvfile"
    )


def baixar_csv(url: str) -> pd.DataFrame:
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    return pd.read_csv(io.StringIO(resp.text))


def _serie_mensal(df: pd.DataFrame, area: str, valor="OBS_VALUE") -> pd.Series:
    g = df.loc[df["REF_AREA"] == area, ["TIME_PERIOD", valor]].copy()
    g["TIME_PERIOD"] = pd.to_datetime(g["TIME_PERIOD"].astype(str) + "-01")
    return (
        g.drop_duplicates("TIME_PERIOD")
        .sort_values("TIME_PERIOD")
        .set_index("TIME_PERIOD")[valor]
        .astype(float)
    )


def juro_nominal_acumulado(taxas_aa: pd.Series, meses: pd.DatetimeIndex) -> float:
    return float(np.prod((1.0 + taxas_aa.loc[meses] / 100.0) ** (1.0 / 12.0)) - 1.0)


def inflacao_acumulada(cpi: pd.Series, cpi_ini: pd.Timestamp, cpi_fim: pd.Timestamp) -> float:
    return float(cpi.loc[cpi_fim] / cpi.loc[cpi_ini] - 1.0)


def juro_real_fisher(i_nom: float, inflacao: float) -> float:
    return (1.0 + i_nom) / (1.0 + inflacao) - 1.0


def montar_ranking(
    cbpol: pd.DataFrame,
    cpi: pd.DataFrame,
    ano_inicio: int = ANO_INICIO_DEFAULT,
    ano_fim: int = ANO_FIM_DEFAULT,
) -> pd.DataFrame:
    meses, cpi_ini, cpi_fim = periodo(ano_inicio, ano_fim)
    cpi_ix = cpi.loc[cpi["UNIT_MEASURE"] == CPI_INDEX].copy()
    rows = []
    for area in sorted(cbpol["REF_AREA"].unique()):
        taxas = _serie_mensal(cbpol, area)
        if not all(m in taxas.index for m in meses):
            continue
        precos = _serie_mensal(cpi_ix, area)
        if cpi_ini not in precos.index or cpi_fim not in precos.index:
            continue
        i_nom = juro_nominal_acumulado(taxas, meses)
        inf = inflacao_acumulada(precos, cpi_ini, cpi_fim)
        rows.append(
            {
                "codigo": area,
                "pais": NOMES.get(area, area),
                "juro_real_acumulado_%": juro_real_fisher(i_nom, inf) * 100.0,
                "taxa_basica_acumulada_%": i_nom * 100.0,
                "inflacao_acumulada_%": inf * 100.0,
                "taxa_ini": float(taxas.get(cpi_ini, np.nan)),
                "taxa_fim": float(taxas.loc[cpi_fim]),
            }
        )
    out = pd.DataFrame(rows).sort_values(
        "juro_real_acumulado_%", ascending=False
    ).reset_index(drop=True)
    out.insert(0, "rank", np.arange(1, len(out) + 1))
    return out


def formatar_pct(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def markdown_ranking(
    df: pd.DataFrame,
    ano_inicio: int = ANO_INICIO_DEFAULT,
    ano_fim: int = ANO_FIM_DEFAULT,
) -> str:
    linhas = [
        f"# Ranking da taxa básica de juros real acumulada (1/1/{ano_inicio}–31/12/{ano_fim})",
        "",
        "Fonte: BIS — taxa de política monetária (fim de período, mensal) e CPI.",
        "Juro real *ex post* (Fisher): `(1 + i_acum) / (1 + π_acum) − 1`.",
        f"i acumulado: produto mensal `(1 + i/100)^(1/12)` de jan/{ano_inicio} a dez/{ano_fim}.",
        f"π acumulada: CPI dez/{ano_fim} ÷ CPI dez/{ano_inicio - 1} − 1.",
        "",
        "| Rank | País | Juro real acum. | Taxa básica acum. | Inflação acum. |",
        "|-----:|------|----------------:|------------------:|---------------:|",
    ]
    for _, r in df.iterrows():
        linhas.append(
            f"| {int(r['rank'])} | {r['pais']} | "
            f"{formatar_pct(r['juro_real_acumulado_%'])}% | "
            f"{formatar_pct(r['taxa_basica_acumulada_%'])}% | "
            f"{formatar_pct(r['inflacao_acumulada_%'])}% |"
        )
    linhas.extend(
        [
            "",
            "Área do euro: o BIS usa a taxa principal de refinanciamento do BCE "
            "(0% até julho/2022; alta a partir de então), não a facilidade de depósito.",
            "",
        ]
    )
    return "\n".join(linhas)


def gravar(
    df: pd.DataFrame,
    pasta: Path,
    ano_inicio: int,
    ano_fim: int,
) -> dict[str, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    stem = f"ranking_juro_real_{ano_inicio}_{ano_fim}"
    csv = pasta / f"{stem}.csv"
    xlsx = pasta / f"{stem}.xlsx"
    md = pasta / f"{stem}.md"
    df.to_csv(csv, index=False)
    df.to_excel(xlsx, index=False)
    md.write_text(markdown_ranking(df, ano_inicio, ano_fim), encoding="utf-8")
    return {"csv": csv, "xlsx": xlsx, "md": md}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Ranking do juro real básico acumulado (BIS, Fisher)."
    )
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO_DEFAULT)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM_DEFAULT)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "output")
    parser.add_argument("--cbpol", type=Path, default=None)
    parser.add_argument("--cpi", type=Path, default=None)
    args = parser.parse_args(argv)

    cbpol = (
        pd.read_csv(args.cbpol)
        if args.cbpol
        else baixar_csv(url_cbpol(args.ano_inicio, args.ano_fim))
    )
    cpi = (
        pd.read_csv(args.cpi)
        if args.cpi
        else baixar_csv(url_cpi(args.ano_inicio, args.ano_fim))
    )
    ranking = montar_ranking(cbpol, cpi, args.ano_inicio, args.ano_fim)
    caminhos = gravar(ranking, args.output_dir, args.ano_inicio, args.ano_fim)
    print(markdown_ranking(ranking, args.ano_inicio, args.ano_fim))
    print(f"[INFO] {len(ranking)} países com série completa")
    for nome, path in caminhos.items():
        print(f"[OK] {nome}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
