"""CLI fallback when MCP UI is unavailable (ContAgil / corporate)."""

from __future__ import annotations

import argparse
import json

from . import providers


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str, ensure_ascii=True))


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="tesouro-cli",
        description="Estatisticas fiscais do Tesouro Nacional (RTN/ARIA) sem Cursor Desktop.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    sub.add_parser("temas", help="Temas RTN disponiveis (10/13/20)")
    sub.add_parser("aliases", help="Aliases locais de series fiscais")

    s = sub.add_parser("catalog", help="Catalogo de series ARIA")
    s.add_argument("--tema", default="", help="Filtrar por tema 10/13/20")

    s = sub.add_parser("search", help="Busca series por nome")
    s.add_argument("query")
    s.add_argument("--tema", default="")
    s.add_argument("--limit", type=int, default=30)

    s = sub.add_parser("serie", help="Serie mensal por alias ou codigo")
    s.add_argument("alias_or_code", help="Ex.: resultado_primario, receita_total, 10.04.1")
    s.add_argument("--from", dest="date_from", default="", help="MM/AAAA ou YYYY-MM")
    s.add_argument("--to", dest="date_to", default="", help="MM/AAAA ou YYYY-MM")
    s.add_argument("--ipca", action="store_true", help="Corrigir valores pelo IPCA")

    s = sub.add_parser("rtn", help="Consulta resultado-fiscal (tema completo ou filtrado)")
    s.add_argument("--tema", default="10", help="10, 13, 20 (ou aliases)")
    s.add_argument("--serie", default="", help="codigo_da_serie opcional")
    s.add_argument("--from", dest="date_from", default="")
    s.add_argument("--to", dest="date_to", default="")
    s.add_argument("--ipca", action="store_true")

    s = sub.add_parser("headline", help="Grandes numeros (capa Tesouro Transparente)")
    s.add_argument("metric", nargs="?", default="", help="resultado_primario, estoque_dpf...")

    s = sub.add_parser("ckan-search", help="Busca datasets no Tesouro Transparente")
    s.add_argument("query", nargs="?", default="resultado do tesouro")
    s.add_argument("--rows", type=int, default=10)

    s = sub.add_parser("ckan-show", help="Detalha um pacote CKAN")
    s.add_argument("package_id", nargs="?", default="resultado-do-tesouro-nacional")

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "temas":
            _print(providers.list_temas())
        elif args.command == "aliases":
            _print(providers.list_known_aliases())
        elif args.command == "catalog":
            _print(providers.list_series(tema=args.tema or None))
        elif args.command == "search":
            _print(
                providers.search_series(
                    args.query,
                    tema=args.tema or None,
                    limit=args.limit,
                )
            )
        elif args.command == "serie":
            _print(
                providers.get_serie(
                    args.alias_or_code,
                    data_inicio=args.date_from or None,
                    data_fim=args.date_to or None,
                    correcao_ipca=args.ipca,
                )
            )
        elif args.command == "rtn":
            _print(
                providers.get_resultado_fiscal(
                    tema=args.tema,
                    data_inicio=args.date_from or None,
                    data_fim=args.date_to or None,
                    codigo_serie=args.serie or None,
                    correcao_ipca=args.ipca,
                )
            )
        elif args.command == "headline":
            _print(providers.get_grandes_numeros(args.metric or None))
        elif args.command == "ckan-search":
            _print(providers.ckan_package_search(args.query, rows=args.rows))
        elif args.command == "ckan-show":
            _print(providers.ckan_package_show(args.package_id))
        else:
            raise SystemExit(f"unknown command: {args.command}")
    except Exception as exc:  # noqa: BLE001
        _print({"error": str(exc)})
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
