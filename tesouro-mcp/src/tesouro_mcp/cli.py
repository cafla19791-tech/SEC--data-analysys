"""CLI fallback when MCP UI is unavailable (ContAgil / corporate)."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from . import collector, providers


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

    s = sub.add_parser(
        "coletar-anual",
        help=(
            "Tabela anual DBGG/RTN/emissoes/resgates/BNDES "
            "(+ merge opcional DGT e FNO/FNE/FCO)"
        ),
    )
    s.add_argument("--from", dest="year_from", type=int, default=2001)
    s.add_argument("--to", dest="year_to", type=int, default=2025)
    s.add_argument(
        "--out",
        default="",
        help="Caminho CSV de saida (se vazio, imprime JSON resumido)",
    )
    s.add_argument(
        "--dgt",
        default="",
        help="CSV de renuncias (template data/templates/dgt_renuncias_anual.csv)",
    )
    s.add_argument(
        "--fundos",
        default="",
        help="CSV FNO/FNE/FCO (template data/templates/fundos_constitucionais_anual.csv)",
    )
    s.add_argument(
        "--no-emissoes",
        action="store_true",
        help="Nao baixar planilha de emissoes/resgates da DPF",
    )
    s.add_argument(
        "--print-csv",
        action="store_true",
        help="Imprime CSV no stdout (alem de --out, se houver)",
    )

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
        elif args.command == "coletar-anual":
            table = collector.collect_annual_table(
                year_from=args.year_from,
                year_to=args.year_to,
                dgt_csv=args.dgt or None,
                fundos_csv=args.fundos or None,
                include_emissoes=not args.no_emissoes,
            )
            csv_text = collector.rows_to_csv(table["rows"])
            if args.out:
                out_path = Path(args.out)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path.write_text(csv_text, encoding="utf-8")
            if args.print_csv or not args.out:
                if args.print_csv:
                    print(csv_text, end="")
                else:
                    _print(
                        {
                            "year_from": table["year_from"],
                            "year_to": table["year_to"],
                            "unit": table["unit"],
                            "count": table["count"],
                            "columns": table["columns"],
                            "sources": table["sources"],
                            "notes": table["notes"],
                            "out": str(Path(args.out).resolve()) if args.out else None,
                            "rows_preview": table["rows"][:3],
                            "rows_tail": table["rows"][-2:],
                            "provider": table["provider"],
                        }
                    )
            elif args.out:
                _print(
                    {
                        "ok": True,
                        "out": str(Path(args.out).resolve()),
                        "count": table["count"],
                        "notes": table["notes"],
                        "sources": table["sources"],
                    }
                )
        else:
            raise SystemExit(f"unknown command: {args.command}")
    except Exception as exc:  # noqa: BLE001
        _print({"error": str(exc)})
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
