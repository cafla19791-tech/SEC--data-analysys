"""CLI fallback when MCP registration is unavailable (corporate / cloud-only)."""

from __future__ import annotations

import argparse
import json
import sys

from . import providers


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="nyse-mcp-cli",
        description="Query US equity market data without Cursor Desktop / MCP UI.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_quote = sub.add_parser("quote", help="Latest delayed quote")
    p_quote.add_argument("symbol")

    p_hist = sub.add_parser("history", help="OHLCV history")
    p_hist.add_argument("symbol")
    p_hist.add_argument("--period", default="1mo")
    p_hist.add_argument("--interval", default="1d")

    p_fund = sub.add_parser("fundamentals", help="Basic fundamentals")
    p_fund.add_argument("symbol")

    p_search = sub.add_parser("search", help="Search ticker / company name")
    p_search.add_argument("query")

    sub.add_parser("status", help="Indicative US market status")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "quote":
            _print(providers.get_quote(args.symbol))
        elif args.command == "history":
            _print(
                providers.get_history(
                    args.symbol, period=args.period, interval=args.interval
                )
            )
        elif args.command == "fundamentals":
            _print(providers.get_fundamentals(args.symbol))
        elif args.command == "search":
            _print(providers.search_ticker(args.query))
        elif args.command == "status":
            _print(providers.market_status())
        else:
            raise SystemExit(f"unknown command: {args.command}")
    except Exception as exc:  # noqa: BLE001
        _print({"error": str(exc)})
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
