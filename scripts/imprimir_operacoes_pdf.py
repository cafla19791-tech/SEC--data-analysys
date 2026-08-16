#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera PDFs resumidos das operações BNDES — Diretas e Indiretas.

Fontes:
  - Não automáticas (SITE): Forma de apoio DIRETA / INDIRETA
  - Indiretas automáticas: resumo anual já calculado (opcional)

Cada PDF traz capa, totais, tabela por ano (desembolso + IPCA até 31/07/2026)
e top clientes.

Uso::

  python scripts/imprimir_operacoes_pdf.py
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from scripts.discriminativo_naoautomaticas_ipca import (
    COL_CLIENTE,
    COL_DATA,
    COL_DESEMBOLSO,
    COL_IPCA,
    DATA_REF_IPCA,
    aplicar_ipca,
    baixar_fonte,
    carregar_contratos,
    carregar_ipca_desde_2002,
)

MARKER = "imprimir-operacoes-pdf-20260816a"

PAGE = landscape(A4)


def _brl(v: float) -> str:
    s = f"{v:,.2f}"
    return "R$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def _int_br(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "t",
            parent=base["Title"],
            fontSize=18,
            alignment=TA_CENTER,
            spaceAfter=8,
        ),
        "h": ParagraphStyle(
            "h",
            parent=base["Heading2"],
            fontSize=12,
            spaceBefore=10,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "b",
            parent=base["Normal"],
            fontSize=9,
            leading=12,
            alignment=TA_LEFT,
        ),
        "small": ParagraphStyle(
            "s",
            parent=base["Normal"],
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#444444"),
        ),
        "center": ParagraphStyle(
            "c",
            parent=base["Normal"],
            fontSize=10,
            alignment=TA_CENTER,
            leading=13,
        ),
    }


def _table(data: list[list], col_widths=None) -> Table:
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#AAAAAA")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return t


def resumo_por_ano(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby(df[COL_DATA].dt.year)
        .agg(
            qtd=(COL_CLIENTE, "size"),
            desembolso=(COL_DESEMBOLSO, "sum"),
            ipca=(COL_IPCA, "sum"),
        )
        .reset_index()
        .rename(columns={COL_DATA: "ano"})
    )
    g["ano"] = g["ano"].astype(int)
    return g.sort_values("ano")


def top_clientes(df: pd.DataFrame, n: int = 20) -> pd.DataFrame:
    g = (
        df.groupby(COL_CLIENTE, sort=False)
        .agg(
            qtd=(COL_CLIENTE, "size"),
            desembolso=(COL_DESEMBOLSO, "sum"),
            ipca=(COL_IPCA, "sum"),
        )
        .reset_index()
        .sort_values("desembolso", ascending=False)
        .head(n)
    )
    return g


def carregar_resumo_automaticas(path: Path | None) -> pd.DataFrame | None:
    if path is None or not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name="Resumo_Anual")
    except Exception:
        return None
    # normaliza
    cols = {str(c).strip(): c for c in df.columns}
    ano_c = next((cols[k] for k in cols if k.lower() == "ano"), None)
    if ano_c is None:
        return None
    df = df[pd.to_numeric(df[ano_c], errors="coerce").notna()].copy()
    df["ano"] = pd.to_numeric(df[ano_c], errors="coerce").astype(int)
    qtd_c = next((cols[k] for k in cols if "qtd" in k.lower() and "opera" in k.lower()), None)
    des_c = next(
        (cols[k] for k in cols if "desembolso" in k.lower() and "ipca" not in k.lower()),
        None,
    )
    ipca_c = next((cols[k] for k in cols if "ipca" in k.lower()), None)
    out = pd.DataFrame({"ano": df["ano"]})
    out["qtd"] = pd.to_numeric(df[qtd_c], errors="coerce") if qtd_c else 0
    out["desembolso"] = pd.to_numeric(df[des_c], errors="coerce") if des_c else 0.0
    out["ipca"] = pd.to_numeric(df[ipca_c], errors="coerce") if ipca_c else 0.0
    return out.dropna(subset=["ano"]).sort_values("ano")


def montar_pdf(
    saida: Path,
    *,
    titulo: str,
    subtitulo: str,
    df: pd.DataFrame | None,
    extra_ano: pd.DataFrame | None = None,
    extra_titulo: str | None = None,
    notas: list[str] | None = None,
) -> Path:
    saida.parent.mkdir(parents=True, exist_ok=True)
    styles = _styles()
    doc = SimpleDocTemplate(
        str(saida),
        pagesize=PAGE,
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title=titulo,
    )
    story = []
    story.append(Paragraph(titulo, styles["title"]))
    story.append(Paragraph(subtitulo, styles["center"]))
    story.append(Spacer(1, 6))
    story.append(
        Paragraph(
            f"Gerado em {datetime.now():%d/%m/%Y %H:%M} · {MARKER} · "
            f"IPCA até {DATA_REF_IPCA:%d/%m/%Y}",
            styles["small"],
        )
    )
    story.append(Spacer(1, 10))

    if df is not None and len(df):
        tot_q = len(df)
        tot_d = float(df[COL_DESEMBOLSO].sum())
        tot_i = float(df[COL_IPCA].sum())
        story.append(Paragraph("Totais", styles["h"]))
        story.append(
            _table(
                [
                    ["Métrica", "Valor"],
                    ["Quantidade de contratos", _int_br(tot_q)],
                    ["Valor desembolsado (corrente)", _brl(tot_d)],
                    [COL_IPCA, _brl(tot_i)],
                ],
                col_widths=[12 * cm, 10 * cm],
            )
        )

        por_ano = resumo_por_ano(df)
        story.append(Paragraph("Desembolso por ano de contratação", styles["h"]))
        rows = [["Ano", "Qtd", "Desembolso corrente", "Desembolso IPCA 31/07/2026"]]
        for _, r in por_ano.iterrows():
            rows.append(
                [
                    str(int(r["ano"])),
                    _int_br(int(r["qtd"])),
                    _brl(float(r["desembolso"])),
                    _brl(float(r["ipca"])),
                ]
            )
        rows.append(
            [
                "TOTAL",
                _int_br(int(por_ano["qtd"].sum())),
                _brl(float(por_ano["desembolso"].sum())),
                _brl(float(por_ano["ipca"].sum())),
            ]
        )
        story.append(_table(rows, col_widths=[2.5 * cm, 3 * cm, 8 * cm, 9 * cm]))

        story.append(PageBreak())
        story.append(Paragraph("Top 20 clientes por desembolso corrente", styles["h"]))
        top = top_clientes(df, 20)
        rows = [["#", "Cliente", "Qtd", "Desembolso corrente", "IPCA 31/07/2026"]]
        for i, r in enumerate(top.itertuples(index=False), start=1):
            nome = str(r[0])
            if len(nome) > 55:
                nome = nome[:52] + "..."
            rows.append(
                [
                    str(i),
                    nome,
                    _int_br(int(r[1])),
                    _brl(float(r[2])),
                    _brl(float(r[3])),
                ]
            )
        story.append(_table(rows, col_widths=[1.2 * cm, 11 * cm, 2 * cm, 5.5 * cm, 5.5 * cm]))

    if extra_ano is not None and len(extra_ano):
        story.append(PageBreak() if df is not None and len(df) else Spacer(1, 1))
        story.append(Paragraph(extra_titulo or "Indiretas automáticas — resumo anual", styles["h"]))
        story.append(
            Paragraph(
                "Valores conforme planilha de indiretas automáticas (unidade da publicação BNDES).",
                styles["small"],
            )
        )
        rows = [["Ano", "Qtd operações", "Desembolso corrente", "Desembolso IPCA jul/2026"]]
        for _, r in extra_ano.iterrows():
            rows.append(
                [
                    str(int(r["ano"])),
                    _int_br(int(r["qtd"])) if pd.notna(r["qtd"]) else "",
                    _brl(float(r["desembolso"])) if pd.notna(r["desembolso"]) else "",
                    _brl(float(r["ipca"])) if pd.notna(r["ipca"]) else "",
                ]
            )
        if pd.notna(extra_ano["desembolso"]).any():
            rows.append(
                [
                    "TOTAL",
                    _int_br(int(extra_ano["qtd"].fillna(0).sum())),
                    _brl(float(extra_ano["desembolso"].fillna(0).sum())),
                    _brl(float(extra_ano["ipca"].fillna(0).sum())),
                ]
            )
        story.append(_table(rows, col_widths=[2.5 * cm, 3.5 * cm, 8 * cm, 9 * cm]))

    if notas:
        story.append(Spacer(1, 12))
        story.append(Paragraph("Notas", styles["h"]))
        for n in notas:
            story.append(Paragraph(f"• {n}", styles["small"]))

    doc.build(story)
    print(f"[OK] PDF: {saida} ({saida.stat().st_size / 1e6:.2f} MB)")
    return saida


def processar(
    *,
    fonte_naoauto: Path,
    saida_dir: Path,
    automaticas_xlsx: Path | None = None,
    ipca_path: Path | None = None,
    hoje: datetime | None = None,
    baixar: bool = True,
) -> dict[str, Path]:
    print(f"[{MARKER}]")
    hoje = hoje or datetime.now()
    if baixar or not fonte_naoauto.exists():
        baixar_fonte(fonte_naoauto)

    df = carregar_contratos(fonte_naoauto, hoje=hoje)
    ipca = carregar_ipca_desde_2002(ipca_path)
    df = aplicar_ipca(df, ipca, data_ref=DATA_REF_IPCA)

    if "Forma de apoio" not in df.columns:
        raise ValueError("Coluna 'Forma de apoio' ausente no Excel de não automáticas.")

    forma = df["Forma de apoio"].astype(str).str.strip().str.upper()
    diretas = df[forma.str.startswith("DIRETA")].copy()
    indiretas_na = df[forma.str.contains("INDIRETA", na=False)].copy()

    print(f"[INFO] Diretas (não auto): {len(diretas):,}")
    print(f"[INFO] Indiretas (não auto): {len(indiretas_na):,}")

    auto = carregar_resumo_automaticas(automaticas_xlsx)
    if auto is not None:
        print(f"[INFO] Automáticas (resumo): {len(auto)} anos")

    saida_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = saida_dir
    out_d = montar_pdf(
        pdf_dir / "OPERACOES_DIRETAS.pdf",
        titulo="Operações Diretas — BNDES",
        subtitulo="Contratos com Forma de Apoio = DIRETA (operações não automáticas)",
        df=diretas,
        notas=[
            "Fonte: naoautomaticas.xlsx (BNDES — central de downloads).",
            "IPCA: Bacen SGS 433; atualização do valor desembolsado até 31/07/2026.",
            "Relatório resumido para impressão (totais, série anual e top 20 clientes).",
        ],
    )
    out_i = montar_pdf(
        pdf_dir / "OPERACOES_INDIRETAS.pdf",
        titulo="Operações Indiretas — BNDES",
        subtitulo="Não automáticas (INDIRETA) + resumo das automáticas",
        df=indiretas_na,
        extra_ano=auto,
        extra_titulo="Indiretas automáticas — desembolso por ano (resumo)",
        notas=[
            "Parte 1: Forma de Apoio = INDIRETA em naoautomaticas.xlsx.",
            "Parte 2: resumo anual de operacoes_indiretas_automaticas (quando disponível).",
            "IPCA nas não automáticas: Bacen SGS 433 até 31/07/2026.",
            "Valores das automáticas seguem a unidade da publicação BNDES.",
        ],
    )
    return {"diretas": out_d, "indiretas": out_i}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--fonte",
        type=Path,
        default=ROOT / "data" / "bndes_naoautomaticas" / "naoautomaticas.xlsx",
    )
    p.add_argument(
        "--automaticas",
        type=Path,
        default=ROOT
        / "output"
        / "indiretas_automaticas_ipca"
        / "INDIRETAS_AUTOMATICAS_IPCA_JUL2026.xlsx",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT / "output" / "pdf_operacoes",
    )
    p.add_argument("--ipca", type=Path, default=None)
    p.add_argument("--hoje", type=str, default=None)
    p.add_argument("--sem-baixar", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    hoje = datetime.strptime(args.hoje, "%Y-%m-%d") if args.hoje else datetime.now()
    try:
        processar(
            fonte_naoauto=args.fonte,
            saida_dir=args.saida,
            automaticas_xlsx=args.automaticas if args.automaticas.exists() else None,
            ipca_path=args.ipca,
            hoje=hoje,
            baixar=not args.sem_baixar,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
