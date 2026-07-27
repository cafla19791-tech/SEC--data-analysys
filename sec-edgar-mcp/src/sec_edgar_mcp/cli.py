"""CLI fallback when MCP UI is unavailable (ContAgil / corporate)."""

from __future__ import annotations

import argparse
import json
import sys

from . import providers


def _print(data: object) -> None:
    if isinstance(data, dict) and "_data" in data:
        data = {k: v for k, v in data.items() if k != "_data"}
    print(json.dumps(data, indent=2, default=str))


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="sec-edgar-cli",
        description="SEC EDGAR filings / XBRL facts (data.sec.gov) without Cursor Desktop.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("lookup", help="Ticker -> CIK")
    s.add_argument("ticker")

    s = sub.add_parser("profile", help="Company profile")
    s.add_argument("ticker_or_cik")

    s = sub.add_parser("filings", help="List recent filings")
    s.add_argument("ticker_or_cik")
    s.add_argument("--form", default="", help="10-K, 10-Q, 8-K, ...")
    s.add_argument("--limit", type=int, default=20)

    s = sub.add_parser("facts", help="Key XBRL company facts")
    s.add_argument("ticker_or_cik")
    s.add_argument(
        "--concepts",
        default="",
        help="Comma-separated us-gaap tags (optional)",
    )
    s.add_argument("--limit", type=int, default=8)

    s = sub.add_parser("concept", help="One XBRL concept series")
    s.add_argument("ticker_or_cik")
    s.add_argument("concept")
    s.add_argument(
        "--taxonomy",
        default="auto",
        help="auto (padrao), us-gaap, ifrs-full ou both (mescla 2008-2025)",
    )
    s.add_argument("--limit", type=int, default=20)
    s.add_argument(
        "--annual",
        action="store_true",
        help="So periodos anuais (FY / CYYYYY)",
    )
    s.add_argument(
        "--from",
        dest="year_from",
        type=int,
        default=None,
        help="Ano inicial (com --taxonomy both; default 2008)",
    )
    s.add_argument(
        "--to",
        dest="year_to",
        type=int,
        default=None,
        help="Ano final (com --taxonomy both; default 2025)",
    )

    s = sub.add_parser(
        "series",
        help="Serie anual 2008-2025 mesclando us-gaap + ifrs (ex.: PBR NetIncomeLoss)",
    )
    s.add_argument("ticker_or_cik")
    s.add_argument("concept")
    s.add_argument("--from", dest="year_from", type=int, default=2008)
    s.add_argument("--to", dest="year_to", type=int, default=2025)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "lookup":
            _print(providers.lookup_ticker(args.ticker))
        elif args.command == "profile":
            meta = providers.get_submissions(args.ticker_or_cik)
            meta.pop("_data", None)
            meta.pop("raw_keys", None)
            _print(meta)
        elif args.command == "filings":
            _print(
                providers.list_filings(
                    args.ticker_or_cik,
                    form=args.form or None,
                    limit=args.limit,
                )
            )
        elif args.command == "facts":
            concepts = (
                [c.strip() for c in args.concepts.split(",") if c.strip()]
                if args.concepts
                else None
            )
            _print(
                providers.get_company_facts(
                    args.ticker_or_cik,
                    concepts=concepts,
                    limit_per_concept=args.limit,
                )
            )
        elif args.command == "concept":
            tax = (args.taxonomy or "auto").lower()
            if tax in {"both", "merge", "all"}:
                _print(
                    providers.get_concept_range(
                        args.ticker_or_cik,
                        args.concept,
                        year_from=args.year_from or 2008,
                        year_to=args.year_to or 2025,
                        annual_only=True,
                    )
                )
            else:
                _print(
                    providers.get_concept(
                        args.ticker_or_cik,
                        args.concept,
                        taxonomy=args.taxonomy,
                        limit=args.limit,
                        annual_only=args.annual,
                    )
                )
        elif args.command == "series":
            _print(
                providers.get_concept_range(
                    args.ticker_or_cik,
                    args.concept,
                    year_from=args.year_from,
                    year_to=args.year_to,
                    annual_only=True,
                )
            )
        else:
            raise SystemExit(f"unknown command: {args.command}")
    except Exception as exc:  # noqa: BLE001
        _print({"error": str(exc)})
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
