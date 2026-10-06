#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF das conversas Petrobras — sequência: links, dívida, caixa, juros, lucro.

Uso::

  python scripts/gerar_pdf_conversas_petrobras.py
  python scripts/gerar_pdf_conversas_petrobras.py --saida-dir output
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from gerar_relatorio_petrobras_20f import (  # noqa: E402
    EDGAR,
    GREEN,
    GREEN_SOFT,
    INK,
    LINE,
    MUTED,
    extrair_tabela_evolucao,
    md_inline,
    parse_md_table,
    totais_series,
)
from petrobras_caixa_operacional_20f import (  # noqa: E402
    _fmt_mi as fmt_cx,
    escrever_markdown as md_caixa,
    montar_dataframe as df_caixa,
)
from petrobras_divida_bruta_20f import (  # noqa: E402
    LINHAS as LINHAS_DIVIDA,
    _fmt_mi as fmt_div,
    edgar_url,
    escrever_markdown as md_divida,
    montar_dataframe as df_divida,
)
from petrobras_juros_pagos_20f import (  # noqa: E402
    _fmt_mi as fmt_juro,
    escrever_markdown as md_juros,
    montar_dataframe as df_juros,
)
from petrobras_lucro_liquido_20f import (  # noqa: E402
    _fmt_mi as fmt_lucro,
    escrever_markdown as md_lucro,
    montar_dataframe as df_lucro,
)

STEM = "conversas_petrobras_20f_2026-09-01"
PAGE = A4
USABLE = 17.6 * cm


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Times-Bold",
            fontSize=11, textColor=GREEN, alignment=TA_CENTER, spaceAfter=8,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Title"], fontName="Times-Bold",
            fontSize=18, leading=22, textColor=INK, alignment=TA_CENTER, spaceAfter=6,
        ),
        "sub": ParagraphStyle(
            "sub", parent=base["Normal"], fontName="Times-Italic",
            fontSize=10, leading=13, textColor=MUTED, alignment=TA_CENTER, spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontName="Times-Bold",
            fontSize=13, leading=16, textColor=GREEN, spaceBefore=10, spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Times-Roman",
            fontSize=9.5, leading=13, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=5,
        ),
        "msg": ParagraphStyle(
            "msg", parent=base["Normal"], fontName="Times-Bold",
            fontSize=9, leading=12, textColor=GREEN, spaceBefore=8, spaceAfter=3,
        ),
        "cell": ParagraphStyle(
            "cell", parent=base["Normal"], fontName="Times-Roman",
            fontSize=6.2, leading=8.0, textColor=INK, alignment=TA_LEFT,
        ),
        "cell_h": ParagraphStyle(
            "cell_h", parent=base["Normal"], fontName="Times-Bold",
            fontSize=6.2, leading=8.0, textColor=colors.white,
        ),
        "cell_total": ParagraphStyle(
            "cell_total", parent=base["Normal"], fontName="Times-Bold",
            fontSize=6.2, leading=8.0, textColor=colors.white,
        ),
    }


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = PAGE
    canvas.setFillColor(GREEN)
    canvas.rect(0, h - 10, w, 10, fill=1, stroke=0)
    canvas.setFillColor(MUTED)
    canvas.setFont("Times-Roman", 8)
    canvas.drawString(
        16 * mm, 10 * mm,
        "Petrobras — conversas 1º/set/2026 — Forms 20-F (EDGAR)",
    )
    canvas.drawRightString(w - 16 * mm, 10 * mm, f"Página {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(16 * mm, 13 * mm, w - 16 * mm, 13 * mm)
    canvas.restoreState()


def make_table(rows: list[list[str]], s, weights: list[float] | None = None) -> Table:
    header = [Paragraph(md_inline(c), s["cell_h"]) for c in rows[0]]
    body, total_idx = [], []
    for i, row in enumerate(rows[1:], start=1):
        first = (row[0] if row else "").replace("*", "")
        is_total = first.lower().startswith("total")
        st = s["cell_total"] if is_total else s["cell"]
        body.append([Paragraph(md_inline(c), st) for c in row])
        if is_total:
            total_idx.append(i)
    data = [header] + body
    wts = weights or [1.0] * len(rows[0])
    col_w = [USABLE * w / sum(wts) for w in wts]
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), GREEN),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GREEN_SOFT]),
        ("GRID", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.4),
    ]
    for i in total_idx:
        cmds.append(("BACKGROUND", (0, i), (-1, i), GREEN))
        cmds.append(("TEXTCOLOR", (0, i), (-1, i), colors.white))
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t


def tabela_evolucao(md_fn, df) -> list[list[str]]:
    return parse_md_table(extrair_tabela_evolucao(md_fn(df, "conversas")))


def tabela_links() -> list[list[str]]:
    rows = [["Exercício", "Protocolo", "Documento 20-F"]]
    for item in reversed(LINHAS_DIVIDA):
        url = edgar_url(item["accession"], item["arquivo"])
        rows.append([
            str(item["ano"]),
            item["data_protocolo"],
            f'[{item["arquivo"]}]({url})',
        ])
    return rows


def tabela_sintese(t: dict) -> list[list[str]]:
    return [
        ["Série", "Recorte", "Total (US$ mi)"],
        ["Dívida bruta", "posição 31/12/2002", fmt_div(t["ini"].divida_bruta_usd_milhoes)],
        ["Dívida bruta", "posição 31/12/2025", f"**{fmt_div(t['fim'].divida_bruta_usd_milhoes)}**"],
        ["Dívida bruta", "variação 2002→2025", f"**{fmt_div(t['var'])} (+375,4%)**"],
        ["Caixa operacional", "soma 2002–2025 (24 anos)", f"**{fmt_cx(t['c_anos'])}**"],
        ["Caixa operacional", "24 anos + 1S2026", f"**{fmt_cx(t['c_1s'])}**"],
        ["Juros pagos (caixa)", "soma 2002–2025 (24 anos)", f"**{fmt_juro(t['j_anos'])}**"],
        ["Juros pagos (caixa)", "24 anos + 1S2026", f"**{fmt_juro(t['j_1s'])}**"],
        ["Lucro líquido (acionistas)", "soma 2002–2025 (24 anos)", f"**{fmt_lucro(t['l_anos'])}**"],
        ["Lucro líquido (acionistas)", "24 anos + 1S2026", f"**{fmt_lucro(t['l_1s'])}**"],
    ]


def story(s, div, cx, jur, luc) -> list:
    t = totais_series(div, cx, jur, luc)
    itens = [
        "Links dos Forms 20-F da Petrobras (2002–2025)",
        "Discriminativo da dívida bruta (2002–2025), com página de cada 20-F",
        "Discriminativo da geração operacional de caixa (2002–2026; 2026 = 6-K 1º semestre)",
        "Discriminativo dos juros pagos (2002–2026; 2026 = 6-K 1º semestre)",
        "Discriminativo do lucro líquido (2002–2026; 2026 = 6-K 1º semestre)",
        "Pedido de impressão desta conversa em PDF",
    ]
    out = [
        Spacer(1, 2.2 * cm),
        Paragraph("SEC · EDGAR · CIK 0001119639", s["kicker"]),
        Paragraph("Conversas de 1º de setembro de 2026", s["title"]),
        Paragraph("Petróleo Brasileiro S.A. — Petrobras", s["title"]),
        Paragraph(
            "Transcrição das mensagens relativas aos Forms 20-F:<br/>"
            "links 2002–2025 · dívida bruta · geração operacional de caixa · "
            "juros pagos · lucro líquido.",
            s["sub"],
        ),
        Spacer(1, 8 * mm),
        Paragraph(
            "A sequência de leitura é a dos pedidos: primeiro os links, depois a "
            "<b>evolução da dívida bruta</b> e, logo após, a <b>geração "
            "operacional de caixa</b>; em seguida os juros pagos e o lucro "
            "líquido. Fonte: "
            f'<link href="https://cursor.com/agents/bc-0088e4db-2905-473e-b7f9-c6bdfdcb968b" '
            f'color="#0B4F8A"><u>execução Sec 20-F Petrobras links</u></link>.',
            s["body"],
        ),
        PageBreak(),
        Paragraph("Índice das conversas", s["h1"]),
        ListFlowable(
            [ListItem(Paragraph(item, s["body"]), leftIndent=8) for item in itens],
            bulletType="1",
            start="1",
            leftIndent=16,
            spaceBefore=2,
            spaceAfter=8,
        ),
        Paragraph("Totais das quatro séries (US$ milhões)", s["h1"]),
        Paragraph(
            "Juros, lucro e caixa operacional são fluxos: a linha de total soma "
            "os exercícios. Dívida bruta é estoque em 31 de dezembro: o total é a "
            "variação 2002→2025, não a soma das posições anuais.",
            s["body"],
        ),
        make_table(tabela_sintese(t), s, [1.6, 1.8, 1.0]),
        PageBreak(),
        Paragraph("Conversas — Petrobras (Forms 20-F)", s["h1"]),
        Paragraph("USUÁRIO · Mensagem 1 — Usuário", s["msg"]),
        Paragraph(
            "Traz os links do form 20-da SEC, da, sobre a Petrobrás, de 2002 até 2025",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Form 20-F da Petróleo Brasileiro S.A. — Petrobras (CIK 0001119639) "
            f'no EDGAR da SEC. Lista oficial: <link href="{EDGAR}" color="#0B4F8A">'
            f"<u>todos os 20-F da Petrobras</u></link>. Os anos abaixo são o "
            "exercício fiscal (encerrado em 31 de dezembro).",
            s["body"],
        ),
        make_table(tabela_links(), s, [0.7, 0.9, 2.2]),
        PageBreak(),
        Paragraph("USUÁRIO · Mensagem 2 — Usuário", s["msg"]),
        Paragraph(
            "você vai elaborar um discriminativo com a evolução das dívidas brutas "
            "da petrobras, no período dos anos 2002 a 2025, com a indicação da "
            "página de cada form 20-f , onde está a informação da dívida bruta "
            "do respectivo",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Discriminativo da dívida bruta consolidada da Petrobras em 31 de "
            "dezembro de cada exercício, extraído do Form 20-F original. "
            f"Variação 2002→2025: US$ {fmt_div(t['var'])} milhões (+375,4%). "
            "Pico: US$ 132.158 milhões em 2014. Mínimo recente: US$ 53.799 "
            "milhões em 2022.",
            s["body"],
        ),
        make_table(tabela_evolucao(md_divida, div), s),
        PageBreak(),
        Paragraph("USUÁRIO · Mensagem 3 — Usuário", s["msg"]),
        Paragraph(
            "qual a evolução da geração operacional de caixa no anos 2002 a 2026 "
            "da Petrobras",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Logo após a evolução da dívida bruta, a geração operacional de caixa "
            "— linha <i>Net cash provided by operating activities</i> do fluxo de "
            "caixa consolidado do 20-F original. Não é EBITDA nem lucro líquido. "
            f"Total 2002–2025: US$ {fmt_cx(t['c_anos'])} milhões. "
            f"Total com 1S2026: US$ {fmt_cx(t['c_1s'])} milhões. "
            "Pico em 2022 (US$ 49.717 milhões, F-6); mínimo em 2002 "
            "(US$ 6.287 milhões, F-7). O 20-F de 2005 reapresenta 2004 como "
            "8.155 — usa-se 8.833 do próprio ano. O 20-F de 2011 reapresenta "
            "2010 como 30.110 — usa-se 28.495 US GAAP do 20-F de 2010.",
            s["body"],
        ),
        make_table(tabela_evolucao(md_caixa, cx), s),
        PageBreak(),
        Paragraph("USUÁRIO · Mensagem 4 — Usuário", s["msg"]),
        Paragraph(
            "faça o mesmo procedimento para os juros pagos pela Petrobras nos "
            "anos do período 2002 a 2026",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Juros pagos em caixa (não despesa financeira pelo regime de "
            "competência). 2026 ainda não tem 20-F; o número é o 6-K de "
            f"30/06/2026. Total 2002–2025: US$ {fmt_juro(t['j_anos'])} milhões. "
            f"Total com 1S2026: US$ {fmt_juro(t['j_1s'])} milhões. Pico em 2016 "
            "(US$ 7.308 milhões, F-8).",
            s["body"],
        ),
        make_table(tabela_evolucao(md_juros, jur), s),
        PageBreak(),
        Paragraph("USUÁRIO · Mensagem 5 — Usuário", s["msg"]),
        Paragraph(
            "elabore o discriminativo com a evolução dos lucros líquidos da "
            "Petrobras nos anos 2002 a 2026, conforme os FORM 20-F, com a "
            "indicação da página do FORM-20F do respectivo ano.",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Lucro (prejuízo) líquido atribuível aos acionistas da Petrobras na "
            f"DRE de cada 20-F. Total 2002–2025: US$ {fmt_lucro(t['l_anos'])} "
            f"milhões. Total com 1S2026: US$ {fmt_lucro(t['l_1s'])} milhões. "
            "Pico em 2022 (US$ 36.623 milhões, F-4); prejuízo máximo em 2015 "
            "(US$ −8.450 milhões, F-5).",
            s["body"],
        ),
        make_table(tabela_evolucao(md_lucro, luc), s),
        Spacer(1, 8),
        Paragraph("USUÁRIO · Mensagem 6 — Usuário", s["msg"]),
        Paragraph(
            "IMPRIME TODAS AS CONVERSAS DE HOJE , RELATIVAS À PETROBRAS EM UM PDF",
            s["body"],
        ),
        Paragraph("ASSISTENTE", s["msg"]),
        Paragraph(
            "Este PDF reúne as conversas na sequência corrigida: links, "
            "evolução da dívida bruta e, logo após, geração operacional de "
            "caixa; depois juros pagos e lucro líquido.",
            s["body"],
        ),
    ]
    return out


def gerar(saida_dir: Path) -> Path:
    saida_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = saida_dir / f"{STEM}.pdf"
    div, cx, jur, luc = df_divida(), df_caixa(), df_juros(), df_lucro()
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=PAGE,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Conversas Petrobras — 20-F (dívida, caixa operacional, juros e lucro)",
        author="SEC--data-analysys",
    )
    doc.build(story(styles(), div, cx, jur, luc), onFirstPage=header_footer, onLaterPages=header_footer)
    art = Path("/opt/cursor/artifacts") / pdf_path.name
    if art.parent.is_dir():
        art.write_bytes(pdf_path.read_bytes())
    return pdf_path


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    args = p.parse_args()
    path = gerar(args.saida_dir)
    print(f"pdf: {path} ({path.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
