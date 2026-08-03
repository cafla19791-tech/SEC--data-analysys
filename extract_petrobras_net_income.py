#!/usr/bin/env python3
"""
Extrai lucros líquidos da Petrobras (últimos N anos) a partir da SEC EDGAR.

Exemplos:
  python extract_petrobras_net_income.py
  python extract_petrobras_net_income.py --years 10 --refresh
  python extract_petrobras_net_income.py --no-chart
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

from petrobras import extract_net_income
from petrobras.charts import plot_matplotlib, plot_plotly
from petrobras.report import format_markdown_table, write_analysis_stub


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extrai lucro líquido anual da Petrobras via SEC EDGAR CompanyFacts "
            "e converte para R$ com câmbio médio BCB."
        )
    )
    parser.add_argument("--years", type=int, default=10, help="Anos anuais (padrão: 10)")
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="Ignora cache local e busca novamente SEC + BCB",
    )
    parser.add_argument(
        "--user-agent",
        default="SEC-Data-Analysis cafla19791@gmail.com",
        help="User-Agent exigido pela SEC (Nome email@dominio)",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path("data"),
        help="Diretório de saída dos dados (padrão: data/)",
    )
    parser.add_argument(
        "--reports-dir",
        type=Path,
        default=Path("reports"),
        help="Diretório de relatórios/gráficos (padrão: reports/)",
    )
    parser.add_argument(
        "--no-chart",
        action="store_true",
        help="Não gera gráficos",
    )
    parser.add_argument(
        "--facts-cache",
        type=Path,
        default=Path("data/raw/petrobras_CIK0001119639_companyfacts.json"),
    )
    parser.add_argument(
        "--fx-cache",
        type=Path,
        default=Path("data/usdbrl_annual_avg.json"),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    args.reports_dir.mkdir(parents=True, exist_ok=True)

    result = extract_net_income(
        years=args.years,
        user_agent=args.user_agent,
        facts_cache=str(args.facts_cache),
        fx_cache=str(args.fx_cache),
        refresh=args.refresh,
    )

    rows = result["years"]
    df = pd.DataFrame(rows)

    csv_path = args.out_dir / "petrobras_net_income.csv"
    json_path = args.out_dir / "petrobras_net_income.json"
    md_path = args.reports_dir / "petrobras_net_income_table.md"
    analysis_path = args.reports_dir / "petrobras_net_income_analysis.md"

    export_cols = [
        "year",
        "net_income_brl",
        "net_income_usd",
        "usd_brl_avg",
        "yoy_brl_pct",
        "yoy_usd_pct",
        "form",
        "filed",
        "concept",
        "taxonomy",
    ]
    df[export_cols].to_csv(csv_path, index=False)
    json_path.write_text(
        json.dumps(result, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    md_path.write_text(format_markdown_table(result), encoding="utf-8")
    write_analysis_stub(result, analysis_path)

    print(format_markdown_table(result))
    print(f"\nCSV  → {csv_path}")
    print(f"JSON → {json_path}")
    print(f"MD   → {md_path}")

    if not args.no_chart:
        png = args.reports_dir / "petrobras_net_income_chart.png"
        html = args.reports_dir / "petrobras_net_income_chart.html"
        plot_matplotlib(df, png)
        plot_plotly(df, html)
        print(f"PNG  → {png}")
        print(f"HTML → {html}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
