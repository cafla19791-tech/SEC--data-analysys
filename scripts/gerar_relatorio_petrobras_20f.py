#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Relatório formal — dívida bruta, juros pagos e lucro líquido da Petrobras (20-F).

Uso::

  python scripts/gerar_relatorio_petrobras_20f.py
  python scripts/gerar_relatorio_petrobras_20f.py --saida-dir output
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

PAGE = landscape(A4)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from petrobras_divida_bruta_20f import (  # noqa: E402
    STEM as STEM_DIVIDA,
    _fmt_mi as fmt_div,
    escrever_markdown as md_divida,
    montar_dataframe as df_divida,
)
from petrobras_juros_pagos_20f import (  # noqa: E402
    STEM as STEM_JUROS,
    _fmt_mi as fmt_juro,
    escrever_markdown as md_juros,
    montar_dataframe as df_juros,
)
from petrobras_lucro_liquido_20f import (  # noqa: E402
    STEM as STEM_LUCRO,
    _fmt_mi as fmt_lucro,
    escrever_markdown as md_lucro,
    montar_dataframe as df_lucro,
)

STEM = "relatorio_petrobras_20f_divida_juros_lucro"
GREEN = colors.HexColor("#0B5F2A")
GREEN_SOFT = colors.HexColor("#E6F2EA")
INK = colors.HexColor("#1F2A37")
MUTED = colors.HexColor("#5B6570")
LINE = colors.HexColor("#D0D5DD")
EDGAR = (
    "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany"
    "&CIK=0001119639&type=20-F&dateb=&owner=exclude&count=100"
)


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Times-Bold",
            fontSize=11, textColor=GREEN, alignment=TA_CENTER, spaceAfter=8,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Title"], fontName="Times-Bold",
            fontSize=20, leading=24, textColor=INK, alignment=TA_CENTER, spaceAfter=8,
        ),
        "sub": ParagraphStyle(
            "sub", parent=base["Normal"], fontName="Times-Italic",
            fontSize=11, leading=15, textColor=MUTED, alignment=TA_CENTER, spaceAfter=6,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontName="Times-Bold",
            fontSize=14, leading=18, textColor=GREEN, spaceBefore=12, spaceAfter=7,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Times-Roman",
            fontSize=10, leading=14, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=6,
        ),
        "meta": ParagraphStyle(
            "meta", parent=base["Normal"], fontName="Times-Italic",
            fontSize=9, leading=12, textColor=MUTED, spaceAfter=4,
        ),
        "cell": ParagraphStyle(
            "cell", parent=base["Normal"], fontName="Times-Roman",
            fontSize=6.6, leading=8.4, textColor=INK,
        ),
        "cell_h": ParagraphStyle(
            "cell_h", parent=base["Normal"], fontName="Times-Bold",
            fontSize=6.6, leading=8.4, textColor=colors.white,
        ),
        "cell_total": ParagraphStyle(
            "cell_total", parent=base["Normal"], fontName="Times-Bold",
            fontSize=6.6, leading=8.4, textColor=colors.white,
        ),
    }


def md_inline(text: str) -> str:
    text = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[([^\]]+)\]\((https?://[^)]+)\)",
        lambda m: f'<link href="{m.group(2)}" color="#0B4F8A"><u>{m.group(1)}</u></link>',
        text,
    )
    return text


def _col_weights(headers: list[str]) -> list[float]:
    weights = []
    for h in headers:
        hl = h.lower()
        if hl == "ano":
            weights.append(0.5)
        elif "período" in hl or "periodo" in hl:
            weights.append(0.85)
        elif "protocolo" in hl:
            weights.append(0.8)
        elif "página" in hl:
            weights.append(0.85)
        elif hl == "Δ %":
            weights.append(0.5)
        elif "Δ us$" in hl:
            weights.append(0.7)
        elif "us$ mi" in hl:
            weights.append(1.05)
        elif "métrica" in hl or "metrica" in hl:
            weights.append(2.15)
        elif "norma" in hl:
            weights.append(0.7)
        elif "documento" in hl:
            weights.append(0.95)
        elif "série" in hl or "recorte" in hl:
            weights.append(1.5)
        else:
            weights.append(1.0)
    return weights


def make_table(rows: list[list[str]], s) -> Table:
    header = [Paragraph(md_inline(c), s["cell_h"]) for c in rows[0]]
    body, total_idx = [], []
    for i, row in enumerate(rows[1:], start=1):
        first = re.sub(r"[*]", "", row[0] if row else "")
        is_total = first.lower().startswith("total")
        st = s["cell_total"] if is_total else s["cell"]
        body.append([Paragraph(md_inline(c), st) for c in row])
        if is_total:
            total_idx.append(i)
    data = [header] + body
    usable = 26.5 * cm
    weights = _col_weights(rows[0])
    col_w = [usable * w / sum(weights) for w in weights]
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), GREEN),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GREEN_SOFT]),
        ("GRID", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]
    for i in total_idx:
        cmds.append(("BACKGROUND", (0, i), (-1, i), GREEN))
        cmds.append(("TEXTCOLOR", (0, i), (-1, i), colors.white))
    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t


def header_footer(canvas, doc):
    canvas.saveState()
    w, h = PAGE
    canvas.setFillColor(GREEN)
    canvas.rect(0, h - 10, w, 10, fill=1, stroke=0)
    canvas.setFillColor(MUTED)
    canvas.setFont("Times-Roman", 8)
    canvas.drawString(
        16 * mm, 10 * mm,
        "Relatório — Petrobras · dívida bruta, juros pagos e lucro líquido · Forms 20-F",
    )
    canvas.drawRightString(w - 16 * mm, 10 * mm, f"Página {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(16 * mm, 13 * mm, w - 16 * mm, 13 * mm)
    canvas.restoreState()


def extrair_tabela_evolucao(md: str) -> str:
    """Copia a tabela da seção Evolução, igual ao discriminativo de origem."""
    lines = md.splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith("## Evolução"))
    start += 1
    while start < len(lines) and not lines[start].startswith("|"):
        start += 1
    end = start
    while end < len(lines) and lines[end].startswith("|"):
        end += 1
    return "\n".join(lines[start:end])


def parse_md_table(md_table: str) -> list[list[str]]:
    rows = []
    for line in md_table.splitlines():
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if cells and all(set(c) <= set(":-") and c for c in cells):
            continue
        rows.append(cells)
    return rows


def tabela_divida(div) -> list[list[str]]:
    return parse_md_table(extrair_tabela_evolucao(md_divida(div, "relatorio")))


def tabela_juros(jur) -> list[list[str]]:
    return parse_md_table(extrair_tabela_evolucao(md_juros(jur, "relatorio")))


def tabela_lucro(luc) -> list[list[str]]:
    return parse_md_table(extrair_tabela_evolucao(md_lucro(luc, "relatorio")))


def escrever_markdown(div, jur, luc, gerado: str) -> str:
    ini, fim = div.iloc[0], div.iloc[-1]
    var = int(fim.divida_bruta_usd_milhoes - ini.divida_bruta_usd_milhoes)
    j_anos = int(jur[jur["periodo"] == "ano"]["juros_pagos_usd_milhoes"].sum())
    j_1s = j_anos + int(jur[jur["periodo"] != "ano"]["juros_pagos_usd_milhoes"].sum())
    l_anos = int(luc[luc["periodo"] == "ano"]["lucro_liquido_usd_milhoes"].sum())
    l_1s = l_anos + int(luc[luc["periodo"] != "ano"]["lucro_liquido_usd_milhoes"].sum())
    return "\n".join([
        "# Relatório — Evolução da dívida bruta, dos juros pagos e do lucro líquido da Petrobras",
        "",
        f"**Data:** {gerado}",
        "",
        "Fonte: Forms 20-F originais de Petróleo Brasileiro S.A. — Petrobras "
        f"(CIK 0001119639) na [SEC/EDGAR]({EDGAR}). Valores em US$ milhões. "
        "O exercício de 2026, ainda sem 20-F, usa o 6-K das demonstrações de 30/06/2026.",
        "",
        "## 1. Apresentação",
        "",
        "O presente relatório apresenta as informações acerca da **evolução da "
        "dívida bruta**, em seguida dos **juros pagos** e, adiante, do **lucro "
        "líquido** da Petróleo Brasileiro S.A. — Petrobras, companhia de economia "
        "mista controlada pela União, no período **2002 a 2026**.",
        "",
        "As cifras não são a Dívida Bruta do Governo Geral (DBGG) do Tesouro "
        "Nacional. São os números **da própria companhia**, extraídos do Form "
        "20-F de cada exercício (e, em 2026, do 6-K interino), com a **página** "
        "do formulário em que o valor aparece. As tabelas abaixo são as mesmas "
        "dos discriminativos já publicados (colunas, páginas, links e totais).",
        "",
        "A ordem da exposição é esta:",
        "",
        "1. dívida bruta consolidada em 31 de dezembro (estoque);",
        "2. juros pagos em caixa no exercício (fluxo);",
        "3. lucro (prejuízo) líquido atribuível aos acionistas da Petrobras (fluxo).",
        "",
        "## 2. Fonte e método",
        "",
        "Usa-se o 20-F **original** do próprio ano, não a emenda 20-F/A nem a "
        "reapresentação em formulários posteriores. A coluna Página é o folio "
        "impresso no HTML (rodapé, “N Table of Contents” ou nota F-N).",
        "",
        "## 3. Dívida bruta",
        "",
        "A dívida bruta é **posição em 31 de dezembro**. Não se soma ano a ano. "
        "O total da série é a variação entre 2002 (US$ 14.680 milhões) e 2025 "
        f"(US$ {fmt_div(fim.divida_bruta_usd_milhoes)} milhões), igual a "
        f"**US$ {fmt_div(var)} milhões (+375,4%)**. O pico foi 2014 "
        "(US$ 132.158 milhões); o mínimo recente, 2022 (US$ 53.799 milhões).",
        "",
        "A definição muda ao longo do tempo: US GAAP até 2008 (ST+LT+project "
        "finance+capital leases); Item 5 em 2009–2012; IFRS sem IFRS 16 "
        "operacional em 2013–2018; e, de 2019 a 2025, Gross Debt oficial "
        "(finance debt + lease liabilities).",
        "",
        extrair_tabela_evolucao(md_divida(div, gerado)),
        "",
        "## 4. Juros pagos",
        "",
        "Passa-se agora aos **juros pagos em caixa** — não à despesa financeira "
        "pelo regime de competência. De 2004 a 2010 o 20-F informa o valor "
        "*net of amount capitalized*; de 2011 em diante, *Repayment of interest* "
        "na seção de financiamento. Em 2026 não há 20-F: o 6-K de 07/08/2026 "
        "registra US$ 1.070 milhões no 1º semestre (contra US$ 856 milhões no 1S2025).",
        "",
        f"**Total 2002–2025:** US$ {fmt_juro(j_anos)} milhões. "
        f"**Total com 1S2026:** US$ {fmt_juro(j_1s)} milhões. "
        "Pico em 2016 (US$ 7.308 milhões, F-8).",
        "",
        extrair_tabela_evolucao(md_juros(jur, gerado)),
        "",
        "## 5. Lucro líquido",
        "",
        "Por fim, o **lucro (prejuízo) líquido atribuível aos acionistas da "
        "Petrobras**, na DRE de cada 20-F. Até 2010 a série é US GAAP; a partir "
        "de 2011, IFRS. O 20-F de 2011 reapresenta 2010 como 20.055; a tabela "
        "usa o número US GAAP do próprio 20-F de 2010 (19.184). Em 2019 o lucro "
        "de 10.151 inclui descontinuadas (BR Distribuidora) de 2.491.",
        "",
        f"**Total 2002–2025:** US$ {fmt_lucro(l_anos)} milhões. "
        f"**Total com 1S2026:** US$ {fmt_lucro(l_1s)} milhões. "
        "Pico em 2022 (US$ 36.623 milhões, F-4); prejuízo máximo em 2015 "
        "(US$ −8.450 milhões, F-5).",
        "",
        extrair_tabela_evolucao(md_lucro(luc, gerado)),
        "",
        "## 6. Síntese dos totais",
        "",
        "| Série | Recorte | Total (US$ mi) |",
        "|---|---|---:|",
        f"| Juros pagos (caixa) | soma 2002–2025 (24 anos) | **{fmt_juro(j_anos)}** |",
        f"| Juros pagos (caixa) | 24 anos + 1S2026 | **{fmt_juro(j_1s)}** |",
        f"| Lucro líquido (acionistas) | soma 2002–2025 (24 anos) | **{fmt_lucro(l_anos)}** |",
        f"| Lucro líquido (acionistas) | 24 anos + 1S2026 | **{fmt_lucro(l_1s)}** |",
        f"| Dívida bruta | posição 31/12/2002 | {fmt_div(ini.divida_bruta_usd_milhoes)} |",
        f"| Dívida bruta | posição 31/12/2025 | **{fmt_div(fim.divida_bruta_usd_milhoes)}** |",
        f"| Dívida bruta | variação 2002→2025 | **{fmt_div(var)} (+375,4%)** |",
        "",
        "## 7. Fonte",
        "",
        "SEC EDGAR, CIK 0001119639, Form 20-F anual (2002–2025) e Form 6-K de "
        "07/08/2026 (demonstrações em US$ do 2º trimestre de 2026).",
        "",
    ])


def _img(path: Path, s) -> list:
    if not path.exists():
        return []
    img = Image(str(path), width=24.5 * cm, height=7.6 * cm)
    img.hAlign = "CENTER"
    return [Spacer(1, 4), img, Spacer(1, 6)]


def story_pdf(div, jur, luc, gerado: str, saida: Path, s) -> list:
    ini, fim = div.iloc[0], div.iloc[-1]
    var = int(fim.divida_bruta_usd_milhoes - ini.divida_bruta_usd_milhoes)
    j_anos = int(jur[jur["periodo"] == "ano"]["juros_pagos_usd_milhoes"].sum())
    j_1s = j_anos + int(jur[jur["periodo"] != "ano"]["juros_pagos_usd_milhoes"].sum())
    l_anos = int(luc[luc["periodo"] == "ano"]["lucro_liquido_usd_milhoes"].sum())
    l_1s = l_anos + int(luc[luc["periodo"] != "ano"]["lucro_liquido_usd_milhoes"].sum())
    story = [
        Spacer(1, 2.4 * cm),
        Paragraph("SEC · EDGAR · CIK 0001119639", s["kicker"]),
        Paragraph("Relatório", s["title"]),
        Paragraph(
            "Evolução da dívida bruta, dos juros pagos<br/>e do lucro líquido da Petrobras",
            s["title"],
        ),
        Paragraph(
            "Forms 20-F, exercícios 2002–2025, e 6-K do 1º semestre de 2026<br/>"
            f"Elaborado em {gerado}",
            s["sub"],
        ),
        Spacer(1, 14 * mm),
        Paragraph("1. Apresentação", s["h1"]),
        Paragraph(
            "O presente relatório apresenta as informações acerca da <b>evolução "
            "da dívida bruta</b>, em seguida dos <b>juros pagos</b> e, adiante, "
            "do <b>lucro líquido</b> da Petróleo Brasileiro S.A. — Petrobras, "
            "companhia de economia mista controlada pela União, no período "
            "<b>2002 a 2026</b>.",
            s["body"],
        ),
        Paragraph(
            "As cifras não são a Dívida Bruta do Governo Geral (DBGG) do Tesouro "
            "Nacional. São os números <b>da própria companhia</b>, extraídos do "
            "Form 20-F de cada exercício (e, em 2026, do 6-K interino), com a "
            "<b>página</b> do formulário em que o valor aparece. As tabelas "
            "abaixo são as mesmas dos discriminativos já publicados (colunas, "
            "páginas, links e totais).",
            s["body"],
        ),
        Paragraph(
            "A ordem da exposição é esta: (i) dívida bruta consolidada em 31 de "
            "dezembro (estoque); (ii) juros pagos em caixa no exercício (fluxo); "
            "(iii) lucro (prejuízo) líquido atribuível aos acionistas da Petrobras "
            "(fluxo).",
            s["body"],
        ),
        Paragraph("2. Fonte e método", s["h1"]),
        Paragraph(
            "Usa-se o 20-F <b>original</b> do próprio ano, não a emenda 20-F/A "
            "nem a reapresentação em formulários posteriores. A coluna Página é "
            "o folio impresso no HTML (rodapé, “N Table of Contents” ou nota F-N). "
            f'Lista EDGAR: <link href="{EDGAR}" color="#0B4F8A"><u>20-F da Petrobras</u></link>.',
            s["body"],
        ),
        PageBreak(),
        Paragraph("3. Dívida bruta", s["h1"]),
        Paragraph(
            f"A dívida bruta é <b>posição em 31 de dezembro</b>. Não se soma ano "
            f"a ano. O total da série é a variação entre 2002 (US$ 14.680 milhões) "
            f"e 2025 (US$ {fmt_div(fim.divida_bruta_usd_milhoes)} milhões), igual a "
            f"<b>US$ {fmt_div(var)} milhões (+375,4%)</b>. O pico foi 2014 "
            f"(US$ 132.158 milhões); o mínimo recente, 2022 (US$ 53.799 milhões).",
            s["body"],
        ),
        Paragraph(
            "A definição muda ao longo do tempo: US GAAP até 2008 "
            "(ST+LT+project finance+capital leases); Item 5 em 2009–2012; IFRS "
            "sem IFRS 16 operacional em 2013–2018; e, de 2019 a 2025, Gross Debt "
            "oficial (finance debt + lease liabilities).",
            s["body"],
        ),
    ]
    story.extend(_img(saida / f"{STEM_DIVIDA}.png", s))
    story.append(make_table(tabela_divida(div), s))
    story.append(PageBreak())
    story.append(Paragraph("4. Juros pagos", s["h1"]))
    story.append(Paragraph(
        "Passa-se agora aos <b>juros pagos em caixa</b> — não à despesa "
        "financeira pelo regime de competência. De 2004 a 2010 o 20-F informa "
        "o valor <i>net of amount capitalized</i>; de 2011 em diante, "
        "<i>Repayment of interest</i> na seção de financiamento. Em 2026 não "
        "há 20-F: o 6-K de 07/08/2026 registra US$ 1.070 milhões no 1º semestre "
        "(contra US$ 856 milhões no 1S2025).",
        s["body"],
    ))
    story.append(Paragraph(
        f"<b>Total 2002–2025:</b> US$ {fmt_juro(j_anos)} milhões. "
        f"<b>Total com 1S2026:</b> US$ {fmt_juro(j_1s)} milhões. "
        "Pico em 2016 (US$ 7.308 milhões, F-8).",
        s["body"],
    ))
    story.extend(_img(saida / f"{STEM_JUROS}.png", s))
    story.append(make_table(tabela_juros(jur), s))
    story.append(PageBreak())
    story.append(Paragraph("5. Lucro líquido", s["h1"]))
    story.append(Paragraph(
        "Por fim, o <b>lucro (prejuízo) líquido atribuível aos acionistas da "
        "Petrobras</b>, na DRE de cada 20-F. Até 2010 a série é US GAAP; a "
        "partir de 2011, IFRS. O 20-F de 2011 reapresenta 2010 como 20.055; a "
        "tabela usa o número US GAAP do próprio 20-F de 2010 (19.184). Em 2019 "
        "o lucro de 10.151 inclui descontinuadas (BR Distribuidora) de 2.491.",
        s["body"],
    ))
    story.append(Paragraph(
        f"<b>Total 2002–2025:</b> US$ {fmt_lucro(l_anos)} milhões. "
        f"<b>Total com 1S2026:</b> US$ {fmt_lucro(l_1s)} milhões. "
        "Pico em 2022 (US$ 36.623 milhões, F-4); prejuízo máximo em 2015 "
        "(US$ −8.450 milhões, F-5).",
        s["body"],
    ))
    story.extend(_img(saida / f"{STEM_LUCRO}.png", s))
    story.append(make_table(tabela_lucro(luc), s))
    story.append(Spacer(1, 8))
    story.append(Paragraph("6. Síntese dos totais", s["h1"]))
    story.append(make_table([
        ["Série", "Recorte", "Total (US$ mi)"],
        ["Juros pagos (caixa)", "soma 2002–2025 (24 anos)", f"**{fmt_juro(j_anos)}**"],
        ["Juros pagos (caixa)", "24 anos + 1S2026", f"**{fmt_juro(j_1s)}**"],
        ["Lucro líquido (acionistas)", "soma 2002–2025 (24 anos)", f"**{fmt_lucro(l_anos)}**"],
        ["Lucro líquido (acionistas)", "24 anos + 1S2026", f"**{fmt_lucro(l_1s)}**"],
        ["Dívida bruta", "posição 31/12/2002", fmt_div(ini.divida_bruta_usd_milhoes)],
        ["Dívida bruta", "posição 31/12/2025", f"**{fmt_div(fim.divida_bruta_usd_milhoes)}**"],
        ["Dívida bruta", "variação 2002→2025", f"**{fmt_div(var)} (+375,4%)**"],
    ], s))
    story.append(Paragraph("7. Fonte", s["h1"]))
    story.append(Paragraph(
        "SEC EDGAR, CIK 0001119639, Form 20-F anual (2002–2025) e Form 6-K de "
        "07/08/2026 (demonstrações em US$ do 2º trimestre de 2026).",
        s["body"],
    ))
    return story


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    args = p.parse_args()
    saida = args.saida_dir
    saida.mkdir(parents=True, exist_ok=True)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    div, jur, luc = df_divida(), df_juros(), df_lucro()
    md_path = saida / f"{STEM}.md"
    pdf_path = saida / f"{STEM}.pdf"
    md_path.write_text(escrever_markdown(div, jur, luc, gerado), encoding="utf-8")
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=PAGE,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Relatório Petrobras — dívida bruta, juros pagos e lucro líquido (20-F)",
        author="SEC--data-analysys",
    )
    doc.build(
        story_pdf(div, jur, luc, gerado, saida, styles()),
        onFirstPage=header_footer,
        onLaterPages=header_footer,
    )
    art = Path("/opt/cursor/artifacts") / pdf_path.name
    if art.parent.is_dir():
        art.write_bytes(pdf_path.read_bytes())
    print(f"md: {md_path}")
    print(f"pdf: {pdf_path} ({pdf_path.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
