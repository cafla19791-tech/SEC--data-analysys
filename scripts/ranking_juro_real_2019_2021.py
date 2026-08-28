#!/usr/bin/env python3
"""Ranking da taxa básica de juros real acumulada: 1/1/2019–31/12/2021.

Fonte: BIS (taxa de política monetária mensal, fim de período; CPI mensal).
Juro nominal acumulado: Π (1 + i_m/100)^(1/12) − 1, m = jan/2019 … dez/2021.
Inflação acumulada: CPI(dez/2021) / CPI(dez/2018) − 1.
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

BIS_CBPOL = (
    "https://stats.bis.org/api/v1/data/WS_CBPOL/M..?"
    "startPeriod=2018-12&endPeriod=2021-12&format=csvfile"
)
BIS_CPI = (
    "https://stats.bis.org/api/v1/data/WS_LONG_CPI/M..?"
    "startPeriod=2018-12&endPeriod=2021-12&format=csvfile"
)
CPI_INDEX = 628
MESES = pd.date_range("2019-01-01", "2021-12-01", freq="MS")
CPI_INI = pd.Timestamp("2018-12-01")
CPI_FIM = pd.Timestamp("2021-12-01")

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


def juro_nominal_acumulado(taxas_aa: pd.Series) -> float:
    return float(np.prod((1.0 + taxas_aa.loc[MESES] / 100.0) ** (1.0 / 12.0)) - 1.0)


def inflacao_acumulada(cpi: pd.Series) -> float:
    return float(cpi.loc[CPI_FIM] / cpi.loc[CPI_INI] - 1.0)


def juro_real_fisher(i_nom: float, inflacao: float) -> float:
    return (1.0 + i_nom) / (1.0 + inflacao) - 1.0


def montar_ranking(cbpol: pd.DataFrame, cpi: pd.DataFrame) -> pd.DataFrame:
    cpi_ix = cpi.loc[cpi["UNIT_MEASURE"] == CPI_INDEX].copy()
    rows = []
    for area in sorted(cbpol["REF_AREA"].unique()):
        taxas = _serie_mensal(cbpol, area)
        if not all(m in taxas.index for m in MESES):
            continue
        precos = _serie_mensal(cpi_ix, area)
        if CPI_INI not in precos.index or CPI_FIM not in precos.index:
            continue
        i_nom = juro_nominal_acumulado(taxas)
        inf = inflacao_acumulada(precos)
        rows.append(
            {
                "codigo": area,
                "pais": NOMES.get(area, area),
                "juro_real_acumulado_%": juro_real_fisher(i_nom, inf) * 100.0,
                "taxa_basica_acumulada_%": i_nom * 100.0,
                "inflacao_acumulada_%": inf * 100.0,
                "taxa_dez2018": float(taxas.get(pd.Timestamp("2018-12-01"), np.nan)),
                "taxa_dez2021": float(taxas.loc[pd.Timestamp("2021-12-01")]),
            }
        )
    out = pd.DataFrame(rows).sort_values(
        "juro_real_acumulado_%", ascending=False
    ).reset_index(drop=True)
    out.insert(0, "rank", np.arange(1, len(out) + 1))
    return out


def formatar_pct(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def markdown_ranking(df: pd.DataFrame) -> str:
    linhas = [
        "# Ranking da taxa básica de juros real acumulada (1/1/2019–31/12/2021)",
        "",
        "Fonte: BIS — taxa de política monetária (fim de período, mensal) e CPI.",
        "Juro real *ex post* (Fisher): `(1 + i_acum) / (1 + π_acum) − 1`.",
        "i acumulado: produto mensal `(1 + i/100)^(1/12)` de jan/2019 a dez/2021.",
        "π acumulada: CPI dez/2021 ÷ CPI dez/2018 − 1.",
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
            "Área do euro: o BIS usa a taxa principal de refinanciamento do BCE (0% no período), "
            "não a facilidade de depósito (−0,50%).",
            "Brasil: Selic-meta BIS; IPCA implícito ~20,0% (2019–2021), alinhado ao IBGE.",
            "",
        ]
    )
    return "\n".join(linhas)


def gravar(df: pd.DataFrame, pasta: Path) -> dict[str, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    stem = "ranking_juro_real_2019_2021"
    csv = pasta / f"{stem}.csv"
    xlsx = pasta / f"{stem}.xlsx"
    md = pasta / f"{stem}.md"
    df.to_csv(csv, index=False)
    df.to_excel(xlsx, index=False)
    md.write_text(markdown_ranking(df), encoding="utf-8")
    return {"csv": csv, "xlsx": xlsx, "md": md}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Ranking do juro real básico acumulado 2019–2021 (BIS)."
    )
    parser.add_argument("--output-dir", type=Path, default=ROOT / "output")
    parser.add_argument("--cbpol", type=Path, default=None)
    parser.add_argument("--cpi", type=Path, default=None)
    args = parser.parse_args(argv)

    cbpol = pd.read_csv(args.cbpol) if args.cbpol else baixar_csv(BIS_CBPOL)
    cpi = pd.read_csv(args.cpi) if args.cpi else baixar_csv(BIS_CPI)
    ranking = montar_ranking(cbpol, cpi)
    caminhos = gravar(ranking, args.output_dir)
    print(markdown_ranking(ranking))
    for nome, path in caminhos.items():
        print(f"[OK] {nome}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
