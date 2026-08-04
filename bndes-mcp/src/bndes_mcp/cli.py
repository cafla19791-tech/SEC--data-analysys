"""CLI ContAgil WinPython para operacoes BNDES."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import excel_export, providers


def _print(data: object) -> None:
    print(json.dumps(data, indent=2, default=str, ensure_ascii=False))


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="bndes-cli",
        description="Consulta operacoes de financiamento do BNDES por CNPJ/CPF.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("cnpj", help="Busca por CNPJ/CPF e gera Excel")
    c.add_argument("documento", help="CNPJ (14) ou CPF (11), com ou sem mascara")
    c.add_argument(
        "--out",
        default="",
        help="Arquivo xlsx de saida (default: bndes_<doc>.xlsx no CWD)",
    )
    c.add_argument("--json-out", default="", help="Salva JSON bruto da API")
    c.add_argument("--rows", type=int, default=10000)

    s = sub.add_parser("resumo", help="Resumo JSON por CNPJ/CPF (sem Excel)")
    s.add_argument("documento")
    s.add_argument("--rows", type=int, default=10000)

    j = sub.add_parser("json-para-excel", help="Converte JSON ja baixado em Excel")
    j.add_argument("json_path", help="Arquivo JSON (resposta Solr completa ou lista docs)")
    j.add_argument("--out", default="", help="Arquivo xlsx de saida")

    sub.add_parser("endpoint", help="Mostra URL base da API")

    return p


def _docs_from_payload(payload: object) -> list[dict]:
    if isinstance(payload, dict):
        if "response" in payload and isinstance(payload["response"], dict):
            docs = payload["response"].get("docs") or []
            return list(docs)
        if "docs" in payload:
            return list(payload["docs"] or [])
    if isinstance(payload, list):
        return list(payload)
    raise ValueError("JSON nao reconhecido (esperado response.docs)")


def _default_out(documento: str) -> Path:
    dig = providers.digits_only(documento)
    return Path.cwd() / f"bndes_{dig}.xlsx"


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.command == "endpoint":
        _print({"base": providers.DEFAULT_BASE, "select": f"{providers.DEFAULT_BASE}/select"})
        return 0

    if args.command == "json-para-excel":
        raw = json.loads(Path(args.json_path).read_text(encoding="utf-8"))
        docs = _docs_from_payload(raw)
        out = Path(args.out) if args.out else Path(args.json_path).with_suffix(".xlsx")
        path = excel_export.write_excel(docs, out)
        _print({"ok": True, "operacoes": len(docs), "excel": str(path)})
        return 0

    if args.command in ("cnpj", "resumo"):
        data = providers.fetch_operacoes(args.documento, rows=args.rows)
        docs = list(data.get("response", {}).get("docs") or [])
        summary = providers.summarize(docs)
        if args.command == "resumo":
            _print(summary)
            return 0

        if args.json_out:
            Path(args.json_out).write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        out = Path(args.out) if args.out else _default_out(args.documento)
        path = excel_export.write_excel(docs, out)
        _print({**summary, "excel": str(path), "numFound": data.get("response", {}).get("numFound")})
        return 0

    return 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"ERRO: {exc}", file=sys.stderr)
        raise SystemExit(1)
