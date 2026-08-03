#!/usr/bin/env python3
"""
Análise financeira de empresas com dados da SEC EDGAR (ou CSV local).

Exemplos:
  python analyze_finance.py --ticker AAPL
  python analyze_finance.py --ticker MSFT --years 7 --export-json saida.json
  python analyze_finance.py --csv data/sample_companies.csv --company AAPL
  python analyze_finance.py --demo
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
import requests

from financial_analyzer import SecClient, compute_metrics, format_report
from financial_analyzer.metrics import snapshot_from_dataframe
from financial_analyzer.report import export_csv, export_json


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Analisa dados financeiros de empresas (SEC EDGAR ou CSV).",
    )
    source = parser.add_mutually_exclusive_group()
    source.add_argument(
        "--ticker",
        help="Símbolo da ação (ex.: AAPL, MSFT). Busca dados na SEC EDGAR.",
    )
    source.add_argument(
        "--cik",
        help="CIK numérico da empresa na SEC (alternativa ao ticker).",
    )
    source.add_argument(
        "--csv",
        type=Path,
        help="Arquivo CSV local com colunas year, revenue, net_income, etc.",
    )
    source.add_argument(
        "--demo",
        action="store_true",
        help="Roda análise de demonstração com data/sample_companies.csv.",
    )

    parser.add_argument(
        "--company",
        help="Quando --csv/--demo, filtra pela coluna ticker (ex.: AAPL).",
    )
    parser.add_argument(
        "--years",
        type=int,
        default=5,
        help="Quantidade de anos anuais a considerar (padrão: 5).",
    )
    parser.add_argument(
        "--user-agent",
        default="FinancialAnalyzer research@example.com",
        help="User-Agent obrigatório pela SEC (Nome email@dominio).",
    )
    parser.add_argument(
        "--export-json",
        type=Path,
        help="Salva o resultado em JSON.",
    )
    parser.add_argument(
        "--export-csv",
        type=Path,
        help="Salva a tabela anual em CSV.",
    )
    parser.add_argument(
        "--compare",
        nargs="+",
        metavar="TICKER",
        help="Compara vários tickers lado a lado (busca SEC).",
    )
    return parser


def analyze_sec(
    ticker: str | None,
    cik: str | None,
    years: int,
    user_agent: str,
):
    client = SecClient(user_agent=user_agent)
    snapshot = client.get_financial_snapshot(ticker=ticker, cik=cik, years=years)
    return compute_metrics(snapshot)


def analyze_csv(path: Path, company: str | None):
    df = pd.read_csv(path)
    if "ticker" in df.columns and company:
        filtered = df[df["ticker"].str.upper() == company.upper()]
        if filtered.empty:
            available = sorted(df["ticker"].dropna().unique())
            raise SystemExit(
                f"Empresa '{company}' não encontrada no CSV. "
                f"Disponíveis: {', '.join(available)}"
            )
        df = filtered
        ticker = company.upper()
        name = str(df["name"].iloc[0]) if "name" in df.columns else ticker
    elif "ticker" in df.columns and company is None:
        # analisa a primeira empresa do arquivo
        first = str(df["ticker"].iloc[0])
        df = df[df["ticker"] == first]
        ticker = first.upper()
        name = str(df["name"].iloc[0]) if "name" in df.columns else ticker
    else:
        ticker = company.upper() if company else "CSV"
        name = "Empresa (CSV)"

    snapshot = snapshot_from_dataframe(df, ticker=ticker, name=name)
    return compute_metrics(snapshot)


def compare_tickers(tickers: list[str], years: int, user_agent: str) -> str:
    client = SecClient(user_agent=user_agent)
    rows: list[dict] = []
    for ticker in tickers:
        try:
            snap = client.get_financial_snapshot(ticker=ticker, years=years)
            metrics = compute_metrics(snap)
            s = metrics.summary
            rows.append(
                {
                    "Ticker": metrics.ticker,
                    "Empresa": metrics.name[:28],
                    "Ano": s.get("latest_year", "—"),
                    "Receita": s.get("revenue"),
                    "Lucro": s.get("net_income"),
                    "Margem": s.get("profit_margin"),
                    "ROE": s.get("roe"),
                    "D/PL": s.get("debt_to_equity"),
                    "CAGR Rec.": s.get("revenue_cagr"),
                }
            )
        except Exception as exc:  # noqa: BLE001
            rows.append(
                {
                    "Ticker": ticker.upper(),
                    "Empresa": f"ERRO: {exc}",
                    "Ano": "—",
                    "Receita": None,
                    "Lucro": None,
                    "Margem": None,
                    "ROE": None,
                    "D/PL": None,
                    "CAGR Rec.": None,
                }
            )

    df = pd.DataFrame(rows)

    def money(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        if abs(v) >= 1e9:
            return f"${v / 1e9:,.1f}B"
        if abs(v) >= 1e6:
            return f"${v / 1e6:,.1f}M"
        return f"${v:,.0f}"

    def pct(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "—"
        return f"{v * 100:.1f}%"

    display = df.copy()
    display["Receita"] = display["Receita"].map(money)
    display["Lucro"] = display["Lucro"].map(money)
    display["Margem"] = display["Margem"].map(pct)
    display["ROE"] = display["ROE"].map(pct)
    display["D/PL"] = display["D/PL"].map(
        lambda v: "—" if v is None or (isinstance(v, float) and pd.isna(v)) else f"{v:.2f}x"
    )
    display["CAGR Rec."] = display["CAGR Rec."].map(pct)

    lines = [
        "=" * 88,
        "COMPARAÇÃO DE EMPRESAS (SEC EDGAR)",
        "=" * 88,
        display.to_string(index=False),
        "=" * 88,
    ]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    demo_path = Path(__file__).resolve().parent / "data" / "sample_companies.csv"

    try:
        if args.compare:
            print(compare_tickers(args.compare, args.years, args.user_agent))
            return 0

        if args.demo:
            metrics = analyze_csv(demo_path, args.company or "AAPL")
        elif args.csv:
            metrics = analyze_csv(args.csv, args.company)
        elif args.ticker or args.cik:
            metrics = analyze_sec(args.ticker, args.cik, args.years, args.user_agent)
        else:
            # padrão: demo se não houver argumentos de fonte
            if demo_path.exists():
                print(
                    "Nenhuma fonte informada — usando modo demo "
                    f"({demo_path.name}).\n"
                    "Use --ticker AAPL para dados ao vivo da SEC.\n"
                )
                metrics = analyze_csv(demo_path, args.company or "AAPL")
            else:
                parser.print_help()
                return 1

        print(format_report(metrics))

        if args.export_json:
            export_json(metrics, args.export_json)
            print(f"\nJSON salvo em: {args.export_json}")
        if args.export_csv:
            export_csv(metrics, args.export_csv)
            print(f"CSV salvo em: {args.export_csv}")

        return 0
    except requests.RequestException as exc:
        print(f"Erro de rede/SEC: {exc}", file=sys.stderr)
        print(
            "Dica: use --demo ou --csv para análise offline.",
            file=sys.stderr,
        )
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"Erro: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
