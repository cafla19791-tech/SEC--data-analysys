"""CLI for ContAgil WinPython (no Cursor Desktop required)."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import providers


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str, ensure_ascii=True))


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="bis-cli",
        description=(
            "BIS central bank policy rates (WS_CBPOL). "
            "SDMX API + local WS_CBPOL_csv_flat.csv (ContAgil)."
        ),
    )
    sub = p.add_subparsers(dest="command", required=True)

    sub.add_parser("catalog", help="Aliases locais de paises / REF_AREA")

    s = sub.add_parser("serie", help="Taxa de politica monetaria por pais")
    s.add_argument(
        "areas",
        help="Codigo/alias: BR, brasil, selic, US, euro, ou lista BR,US,XM",
    )
    s.add_argument(
        "--freq",
        default="M",
        help="M (mensal, padrao), D (diario) ou A (anual)",
    )
    s.add_argument("--from", dest="date_from", default="", help="YYYY-MM ou YYYY-MM-DD")
    s.add_argument("--to", dest="date_to", default="", help="YYYY-MM ou YYYY-MM-DD")
    s.add_argument("--last", type=int, default=None, help="Ultimos N pontos por pais")
    s.add_argument(
        "--local",
        action="store_true",
        help="Ler WS_CBPOL_csv_flat.csv local (ContAgil) em vez do SDMX",
    )
    s.add_argument(
        "--csv",
        default="",
        help="Caminho explicito do CSV flat (default: winpython/CWD)",
    )

    c = sub.add_parser("compare", help="Ultima taxa por pais (snapshot)")
    c.add_argument(
        "areas",
        nargs="?",
        default="BR,US,XM,GB,JP",
        help="Lista de areas (default: BR,US,XM,GB,JP)",
    )
    c.add_argument("--freq", default="M")
    c.add_argument("--local", action="store_true")
    c.add_argument("--csv", default="")

    d = sub.add_parser(
        "download",
        help="Baixa WS_CBPOL_csv_flat.zip e extrai o CSV (ContAgil winpython)",
    )
    d.add_argument(
        "--dir",
        default=".",
        help="Pasta destino (default: atual; no ContAgil use winpython)",
    )
    d.add_argument(
        "--keep",
        action="store_true",
        help="Nao sobrescrever se o CSV ja existir",
    )

    l = sub.add_parser("local-info", help="Localiza WS_CBPOL_csv_flat.csv")
    l.add_argument("--csv", default="")

    e = sub.add_parser(
        "extract",
        help="Extrai paises do CSV flat (~450MB) para um CSV leve (ContAgil)",
    )
    e.add_argument("areas", help="Ex.: BR ou BR,US,XM")
    e.add_argument(
        "--out",
        required=True,
        help="Arquivo de saida (ex.: ..\\cbpol_BR_US_XM.csv)",
    )
    e.add_argument("--csv", default="", help="CSV flat de origem")
    e.add_argument(
        "--freq",
        default="M",
        help="M/D/A, ou vazio para todas as frequencias",
    )

    x = sub.add_parser(
        "excel-diario",
        help=(
            "Excel com 1 aba/pais: Dia | Taxa (%% a.d.) | Taxa acumulada (%%) "
            "(juros compostos, ano com 252 dias uteis)"
        ),
    )
    x.add_argument(
        "--out",
        default="cbpol_taxas_diarias_compostas.xlsx",
        help="Arquivo .xlsx de saida",
    )
    x.add_argument("--csv", default="", help="WS_CBPOL_csv_flat.csv de origem")
    x.add_argument(
        "--areas",
        default="",
        help="Filtrar paises (ex.: BR,US,XM). Vazio = todos com serie diaria",
    )
    x.add_argument("--from", dest="date_from", default="", help="YYYY-MM-DD")
    x.add_argument("--to", dest="date_to", default="", help="YYYY-MM-DD")
    x.add_argument(
        "--sdmx",
        action="store_true",
        help="Forcar download SDMX (nao usa CSV local)",
    )

    p_per = sub.add_parser(
        "excel-periodos",
        help=(
            "Excel com abas de taxa acumulada por periodo "
            "(D: sem sab/dom; M: mensal; ordenado crescente)"
        ),
    )
    p_per.add_argument(
        "--out",
        default="cbpol_taxas_acumuladas_periodos.xlsx",
        help="Arquivo .xlsx de saida",
    )
    p_per.add_argument("--csv", default="", help="WS_CBPOL_csv_flat.csv de origem")
    p_per.add_argument(
        "--freq",
        default="D",
        help="D = diaria (padrao) ou M = mensal",
    )
    p_per.add_argument(
        "--sdmx",
        action="store_true",
        help="Forcar SDMX (nao usa CSV local)",
    )

    xm = sub.add_parser(
        "excel-mensal",
        help=(
            "Excel com 1 aba/pais: Mes | Taxa (%% a.m.) | Taxa acumulada (%%) "
            "(juros compostos, 1/12)"
        ),
    )
    xm.add_argument(
        "--out",
        default="cbpol_taxas_mensais_compostas.xlsx",
        help="Arquivo .xlsx de saida",
    )
    xm.add_argument("--csv", default="", help="WS_CBPOL_csv_flat.csv de origem")
    xm.add_argument(
        "--areas",
        default="",
        help="Filtrar paises (ex.: BR,US,XM). Vazio = todos com serie mensal",
    )
    xm.add_argument("--from", dest="date_from", default="", help="YYYY-MM ou YYYY-MM-DD")
    xm.add_argument("--to", dest="date_to", default="", help="YYYY-MM ou YYYY-MM-DD")
    xm.add_argument(
        "--sdmx",
        action="store_true",
        help="Forcar download SDMX (nao usa CSV local)",
    )

    pdf = sub.add_parser(
        "para-pdf",
        help="Converte um .xlsx em PDF (requer LibreOffice/soffice)",
    )
    pdf.add_argument("xlsx", help="Caminho do arquivo .xlsx")
    pdf.add_argument(
        "--outdir",
        default="",
        help="Pasta de saida (default: mesma pasta do xlsx)",
    )

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "catalog":
            _print(providers.list_known_areas())
        elif args.command == "serie":
            _print(
                providers.get_policy_rates(
                    args.areas,
                    freq=args.freq,
                    date_from=args.date_from or None,
                    date_to=args.date_to or None,
                    last=args.last,
                    prefer_local=args.local,
                    csv_path=args.csv or None,
                )
            )
        elif args.command == "compare":
            _print(
                providers.compare_latest(
                    args.areas,
                    freq=args.freq,
                    prefer_local=args.local,
                    csv_path=args.csv or None,
                )
            )
        elif args.command == "download":
            _print(
                providers.download_flat_csv(
                    args.dir,
                    overwrite=not args.keep,
                )
            )
        elif args.command == "local-info":
            path = providers.find_local_flat_csv(args.csv or None)
            _print(
                {
                    "found": path is not None,
                    "path": str(path) if path else None,
                    "expected_name": providers.FLAT_CSV_NAME,
                    "contagil_hint": str(
                        Path(
                            r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
                        )
                        / providers.FLAT_CSV_NAME
                    ),
                }
            )
        elif args.command == "extract":
            freq = args.freq.strip() if args.freq is not None else "M"
            _print(
                providers.extract_areas_csv(
                    args.areas,
                    args.out,
                    csv_path=args.csv or None,
                    freq=freq if freq else None,
                )
            )
        elif args.command == "excel-diario":
            from . import excel_diario

            _print(
                excel_diario.gerar_excel_diario(
                    args.out,
                    csv_path=args.csv or None,
                    areas=args.areas or None,
                    date_from=args.date_from or None,
                    date_to=args.date_to or None,
                    prefer_local=not args.sdmx,
                )
            )
        elif args.command == "excel-periodos":
            from . import excel_periodos

            _print(
                excel_periodos.gerar_excel_periodos(
                    args.out,
                    csv_path=args.csv or None,
                    prefer_local=not args.sdmx,
                    freq=args.freq,
                )
            )
        elif args.command == "excel-mensal":
            from . import excel_mensal

            _print(
                excel_mensal.gerar_excel_mensal(
                    args.out,
                    csv_path=args.csv or None,
                    areas=args.areas or None,
                    date_from=args.date_from or None,
                    date_to=args.date_to or None,
                    prefer_local=not args.sdmx,
                )
            )
        elif args.command == "para-pdf":
            from . import pdf_export

            _print(
                pdf_export.xlsx_para_pdf(
                    args.xlsx,
                    out_dir=args.outdir or None,
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
