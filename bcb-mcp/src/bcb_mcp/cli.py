"""CLI fallback when MCP UI is unavailable (ContAgil / corporate)."""

from __future__ import annotations

import argparse
import json
import sys

from . import providers


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str, ensure_ascii=True))


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="bcb-cli",
        description="Banco Central do Brasil open data (SGS / OLINDA) without Cursor Desktop.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    sub.add_parser("catalog", help="Aliases locais de series SGS comuns")

    s = sub.add_parser("serie", help="Serie SGS por alias ou codigo")
    s.add_argument("code_or_alias", help="Ex.: selic, ipca, dolar, 11, 433")
    s.add_argument("--from", dest="date_from", default="", help="YYYY-MM-DD ou DD/MM/YYYY")
    s.add_argument("--to", dest="date_to", default="", help="YYYY-MM-DD ou DD/MM/YYYY")
    s.add_argument(
        "--last",
        type=int,
        default=None,
        help="Ultimos N pontos (ignora --from/--to)",
    )

    s = sub.add_parser("ptax", help="Cotacao USD/BRL (OLINDA PTAX)")
    s.add_argument("--from", dest="date_from", default="")
    s.add_argument("--to", dest="date_to", default="")
    s.add_argument("--days", type=int, default=30, help="Janela se --from omitido")

    s = sub.add_parser("expectativas", help="Expectativas Focus (OLINDA)")
    s.add_argument("indicator", nargs="?", default="IPCA", help="IPCA, Selic, Cambio...")
    s.add_argument("--top", type=int, default=20)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "catalog":
            _print(providers.list_known_series())
        elif args.command == "serie":
            _print(
                providers.get_sgs_series(
                    args.code_or_alias,
                    date_from=args.date_from or None,
                    date_to=args.date_to or None,
                    last=args.last,
                )
            )
        elif args.command == "ptax":
            _print(
                providers.get_ptax(
                    date_from=args.date_from or None,
                    date_to=args.date_to or None,
                    last_days=args.days,
                )
            )
        elif args.command == "expectativas":
            _print(
                providers.get_expectativas(
                    args.indicator,
                    top=args.top,
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
