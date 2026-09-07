#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Relatório do Acórdão TCU 1.839/2018-Plenário, com as tabelas do processo
e o cruzamento com a série dos Forms 20-F da Petrobras.

Uso::

  python scripts/gerar_relatorio_acordao_1839_2018.py
  python scripts/gerar_relatorio_acordao_1839_2018.py --pdf /caminho/acordao.pdf
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from petrobras_divida_bruta_20f import (  # noqa: E402
    _fmt_mi as fmt_div,
    montar_dataframe as df_divida,
)
from petrobras_juros_pagos_20f import (  # noqa: E402
    _fmt_mi as fmt_juro,
    montar_dataframe as df_juros,
)
from petrobras_lucro_liquido_20f import (  # noqa: E402
    _fmt_mi as fmt_lucro,
    montar_dataframe as df_lucro,
)

STEM = "relatorio_acordao_1839_2018_petrobras"
GREEN = colors.HexColor("#0B5F2A")
GREEN_SOFT = colors.HexColor("#E6F2EA")
INK = colors.HexColor("#1F2A37")
MUTED = colors.HexColor("#5B6570")
LINE = colors.HexColor("#D0D5DD")
PAGE = A4

# Tabelas transcritas do Acórdão 1.839/2018-TCU-Plenário (TC 003.502/2016-3).
# Fonte: páginas 9, 90 e 91 do PDF autenticado (código 59581786).
ANOS_CA = list(range(2004, 2016))
TABELA2_REUNIOES = {
    "Reuniões": [12, 13, 15, 16, 15, 15, 20, 14, 14, 13, 14, 27],
    "Arquivos": [66, 64, 59, 64, 61, 63, 89, 84, 71, 79, 78, 179],
    "Páginas": [457, 588, 591, 517, 528, 549, 950, 765, 729, 1288, 1569, 4644],
    "Ocorrências": [0, 0, 1, 8, 13, 17, 32, 30, 32, 88, 188, 310],
}
TABELA2_TOTAIS = {"Reuniões": 188, "Arquivos": 957, "Páginas": 13175, "Ocorrências": 719}

CAPEX_VPL = [
    ["Fase", "Capex (US$ bi)", "VPL (US$ bi)"],
    ["FEL 1 (2006)", "20,43", "4,74"],
    ["FEL 2", "43,84", "5,31"],
    ["FEL 3 (2009–2010)", "81,13", "−1,56"],
    ["Epílogo", "83,69", "−43,32"],
]

IGP = [
    ["Empreendimento", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "Total"],
    ["Rnest", "0", "9", "5", "4", "5", "0", "0", "0", "23"],
    ["Comperj", "0", "0", "0", "0", "1", "0", "0", "0", "1"],
    ["Premium I e II", "0", "0", "0", "0", "0", "0", "0", "0", "0"],
    ["**Total**", "**0**", "**9**", "**5**", "**4**", "**6**", "**0**", "**0**", "**0**", "**24**"],
]

CAPACIDADE_TOTAIS = [
    ["Plano de negócios", "Capacidade projetada (kbpd)"],
    ["PDR (2007)", "1.200"],
    ["PN 08-12", "1.400"],
    ["PN 09-13", "1.280"],
    ["PN 10-14", "1.460"],
    ["PN 11-15", "1.460"],
    ["PNG 12-16", "1.595"],
    ["PNG 13-17", "1.595"],
    ["PNG 14-18", "1.295"],
    ["PNG 15-19 (após cancelamentos)", "425"],
]

PDF_CANDIDATOS = [
    Path("/home/ubuntu/.cursor/projects/workspace/uploads/Ac_rd_o_1839_de_2018_Plen_rio__10__1a64.pdf"),
]


def styles():
    base = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle(
            "kicker", parent=base["Normal"], fontName="Times-Bold",
            fontSize=10, textColor=GREEN, alignment=TA_CENTER, spaceAfter=6,
        ),
        "title": ParagraphStyle(
            "title", parent=base["Title"], fontName="Times-Bold",
            fontSize=16, leading=20, textColor=INK, alignment=TA_CENTER, spaceAfter=6,
        ),
        "sub": ParagraphStyle(
            "sub", parent=base["Normal"], fontName="Times-Italic",
            fontSize=10, leading=13, textColor=MUTED, alignment=TA_CENTER, spaceAfter=8,
        ),
        "h1": ParagraphStyle(
            "h1", parent=base["Heading1"], fontName="Times-Bold",
            fontSize=13, leading=16, textColor=GREEN, spaceBefore=10, spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body", parent=base["Normal"], fontName="Times-Roman",
            fontSize=10, leading=13.5, textColor=INK, alignment=TA_JUSTIFY, spaceAfter=6,
        ),
        "cell": ParagraphStyle(
            "cell", parent=base["Normal"], fontName="Times-Roman",
            fontSize=7.2, leading=9, textColor=INK, alignment=TA_CENTER,
        ),
        "cell_l": ParagraphStyle(
            "cell_l", parent=base["Normal"], fontName="Times-Roman",
            fontSize=7.2, leading=9, textColor=INK, alignment=TA_LEFT,
        ),
        "cell_h": ParagraphStyle(
            "cell_h", parent=base["Normal"], fontName="Times-Bold",
            fontSize=7.0, leading=9, textColor=colors.white, alignment=TA_CENTER,
        ),
        "cell_total": ParagraphStyle(
            "cell_total", parent=base["Normal"], fontName="Times-Bold",
            fontSize=7.2, leading=9, textColor=colors.white, alignment=TA_CENTER,
        ),
        "caption": ParagraphStyle(
            "caption", parent=base["Normal"], fontName="Times-Italic",
            fontSize=8, leading=10, textColor=MUTED, alignment=TA_CENTER, spaceAfter=8,
        ),
    }


def md_inline(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("**", "")
    )


def make_table(rows: list[list[str]], s, left_first: bool = True) -> Table:
    header = [Paragraph(md_inline(c), s["cell_h"]) for c in rows[0]]
    body, totals = [], []
    for i, row in enumerate(rows[1:], start=1):
        is_total = str(row[0]).replace("*", "").lower().startswith("total")
        st = s["cell_total"] if is_total else s["cell"]
        cells = []
        for j, c in enumerate(row):
            use = s["cell_l"] if (left_first and j == 0 and not is_total) else st
            if is_total:
                use = s["cell_total"]
            cells.append(Paragraph(md_inline(c), use))
        body.append(cells)
        if is_total:
            totals.append(i)
    data = [header] + body
    usable = 17.8 * cm
    n = len(rows[0])
    weights = [1.6] + [1.0] * (n - 1)
    col_w = [usable * w / sum(weights) for w in weights]
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), GREEN),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, GREEN_SOFT]),
        ("GRID", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    for i in totals:
        cmds.append(("BACKGROUND", (0, i), (-1, i), GREEN))
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
        "Relatório — Acórdão TCU 1.839/2018-Plenário · Petrobras (Rnest, Comperj, Premium)",
    )
    canvas.drawRightString(w - 16 * mm, 10 * mm, f"Página {doc.page}")
    canvas.setStrokeColor(LINE)
    canvas.line(16 * mm, 13 * mm, w - 16 * mm, 13 * mm)
    canvas.restoreState()


def tabela2_rows() -> list[list[str]]:
    header = ["Indicador"] + [str(a) for a in ANOS_CA] + ["Total"]
    rows = [header]
    for nome, vals in TABELA2_REUNIOES.items():
        fmt = [f"{v:,}".replace(",", ".") for v in vals]
        tot = f"**{TABELA2_TOTAIS[nome]:,}**".replace(",", ".")
        rows.append([nome] + fmt + [tot])
    return rows


def cruzamento_rows(div, jur, luc) -> list[list[str]]:
    eventos = {
        2006: "PE 2015 / plano de ~US$ 12 bi e +1.200 kbpd",
        2008: "Início das obras; TCU passa a fiscalizar",
        2009: "Autorização das obras; Capex sobe para a casa dos US$ 60 bi",
        2010: "Follow-on US$ 70 bi; PN 10-14 de US$ 224 bi",
        2014: "Pico da dívida 20-F; 1º trem da Rnest; início do impairment",
        2015: "Premium canceladas; prejuízo líquido máximo no 20-F",
        2016: "Pico dos juros pagos (caixa); DCs 2014–16 reconhecem ~R$ 43 bi",
    }
    dmap = {int(r.ano): r for r in div.itertuples(index=False)}
    jmap = {int(r.ano): r for r in jur.itertuples(index=False) if r.periodo == "ano"}
    lmap = {int(r.ano): r for r in luc.itertuples(index=False) if r.periodo == "ano"}
    rows = [[
        "Ano", "Dívida bruta 20-F (US$ mi)", "Juros pagos (US$ mi)",
        "Lucro líquido (US$ mi)", "Marco no Acórdão 1.839/2018",
    ]]
    for ano in sorted(eventos):
        rows.append([
            str(ano),
            fmt_div(dmap[ano].divida_bruta_usd_milhoes),
            fmt_juro(jmap[ano].juros_pagos_usd_milhoes),
            fmt_lucro(lmap[ano].lucro_liquido_usd_milhoes),
            eventos[ano],
        ])
    return rows


def md_table(rows: list[list[str]]) -> str:
    lines = ["| " + " | ".join(rows[0]) + " |"]
    lines.append("|" + "|".join("---" for _ in rows[0]) + "|")
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def localizar_pdf(explicit: Path | None) -> Path | None:
    if explicit and explicit.exists():
        return explicit
    for cand in PDF_CANDIDATOS:
        if cand.exists():
            return cand
    return None


def recortar_paginas(pdf: Path, saida: Path) -> dict[int, Path]:
    import fitz

    doc = fitz.open(pdf)
    clips = {}
    for n in (7, 9, 49, 80, 90, 91, 130):
        page = doc[n - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        dest = saida / f"acordao1839_pagina_{n:03d}.png"
        pix.save(str(dest))
        clips[n] = dest
    return clips


def escrever_markdown(div, jur, luc, gerado: str) -> str:
    t2, igp, capex, cap, cruz = (
        tabela2_rows(), IGP, CAPEX_VPL, CAPACIDADE_TOTAIS, cruzamento_rows(div, jur, luc),
    )
    return "\n".join([
        "# Relatório — Acórdão TCU 1.839/2018-Plenário e a evolução da dívida bruta, dos juros pagos e do lucro líquido da Petrobras",
        "",
        f"**Data:** {gerado}",
        "",
        "Fonte principal: Acórdão nº 1.839/2018 – TCU – Plenário, TC-003.502/2016-3, "
        "relator Ministro José Múcio Monteiro, sessão ordinária de 8/8/2018 "
        "(código eletrônico AC-1839-30/18-P; autenticidade 59581786). "
        "As tabelas abaixo são as do próprio acórdão. Os números de dívida, juros e lucro "
        "vêm dos Forms 20-F originais da Petrobras (CIK 0001119639).",
        "",
        "## 1. Apresentação",
        "",
        "O presente relatório apresenta as informações acerca da **auditoria do TCU** "
        "sobre a conduta do Conselho de Administração da Petrobras nos projetos das "
        "refinarias **Rnest, Comperj e Premium I e II**, e, em seguida, dos **juros "
        "pagos** e do **lucro líquido** da companhia nos mesmos anos em que o "
        "endividamento disparou.",
        "",
        "As cifras do acórdão **não** são a Dívida Bruta do Governo Geral (DBGG). "
        "São o diagnóstico do TCU sobre a **própria Petrobras** — sociedade de "
        "economia mista controlada pela União — e o cruzamento com o Form 20-F.",
        "",
        "A ordem da exposição é esta:",
        "",
        "1. síntese do que o TCU apurou (orçamento, desembolso, capacidade e perdas);",
        "2. tabelas do acórdão, conforme constam do PDF autenticado;",
        "3. cruzamento com a dívida bruta, os juros pagos e o lucro líquido dos 20-F;",
        "4. dispositivo do Acórdão 1.839/2018.",
        "",
        "## 2. O que o TCU apurou",
        "",
        "Entre o início da década de 2000 e 2015 a Petrobras planejou quatro "
        "empreendimentos para ampliar o refino nacional. Por volta de **2006**, o "
        "plano era investir pouco mais de **US$ 12 bilhões** para acrescer cerca de "
        "**1.200 kbpd** (~70% da capacidade então instalada). Passados dez anos, o "
        "orçamento ultrapassou **US$ 80 bilhões**, dos quais cerca de **US$ 30 "
        "bilhões** foram desembolsados, e apenas cerca de **100 kbpd** entraram em "
        "operação — cerca de 10% da capacidade esperada, por duas vezes e meia o "
        "custo inicialmente previsto.",
        "",
        "As Premium I e II foram canceladas em 2015. A Rnest opera em regime "
        "parcial (um dos dois trens). O Comperj foi postergado. Nas "
        "demonstrações de 2014, 2015 e 2016 a companhia reconheceu baixas e "
        "impairment da ordem de **R$ 43 bilhões** nesses ativos (R$ 2,8 bi das "
        "Premium, R$ 3,4 bi de propina capitalizada e cerca de R$ 40 bi de "
        "irrecuperabilidade).",
        "",
        "O TCU liga esse insucesso ao **endividamento** da estatal, então próximo "
        "de **US$ 100 bilhões**, e aponta gestão temerária da Diretoria, omissão "
        "do Conselho no dever de se informar e de investigar, e fragilidades na "
        "governança da União sobre a companhia (ministros no board, ausência de "
        "política de indicação, falta de accountability da função de propriedade).",
        "",
        "## 3. Tabelas do acórdão, conforme estão no PDF",
        "",
        "### Tabela 2 — Reuniões do Conselho de Administração (2004–2015)",
        "",
        "Página 9 do acórdão. “Ocorrências” conta as menções a Comperj, Rnest e "
        "Premium nas atas. Fonte no original: Petrobras/Segep (evidência 2).",
        "",
        md_table(t2),
        "",
        "### Figura 22 — Evolução do Capex e do VPL (Comperj, Rnest, Premium I e II)",
        "",
        "Página 90. Visão expedita do TCU, sem ajuste temporal das datas de fase. "
        "Valores em US$ bilhões.",
        "",
        md_table(capex),
        "",
        "O Capex quadruplica entre FEL 1 e FEL 3 (2006–2010). O VPL vira negativo "
        "na autorização das obras e fecha o epílogo em **−US$ 43,32 bilhões**.",
        "",
        "### Tabela 3 — Propostas de irregularidade grave (IG-P) do TCU, 2008–2015",
        "",
        "Página 91. Vinte e quatro propostas ao Congresso, 23 delas na Rnest. "
        "Fonte: portal do TCU (acesso em 4/10/2017).",
        "",
        md_table(igp),
        "",
        "### Capacidade de refino projetada nos planos de negócios (totais)",
        "",
        "Página 91. O quadro original reparte a entrada por ano; aqui ficam os "
        "**totais de cada plano**, iguais aos do acórdão. Depois do cancelamento "
        "das Premium, o PNG 15-19 cai para 425 kbpd.",
        "",
        md_table(cap),
        "",
        "A tabela de composição do Conselho (também chamada Tabela 3 no "
        "acórdão, páginas 49 e seguintes) lista, ano a ano, os indicados da União "
        "e o cargo no governo (Casa Civil, Fazenda, MME, Relações Institucionais). "
        "Reproduz-se a página original no PDF deste relatório.",
        "",
        "## 4. Cruzamento com os Forms 20-F",
        "",
        "Os mesmos anos em que o TCU registra o salto do Capex e a perda de VPL "
        "são aqueles em que a **dívida bruta** do 20-F sai de US$ 21.338 milhões "
        "(2006) para o pico de **US$ 132.158 milhões** (2014); os **juros pagos** "
        "em caixa sobem até **US$ 7.308 milhões** (2016); e o **lucro líquido** "
        "vira prejuízo em 2014–2016.",
        "",
        md_table(cruz),
        "",
        "A dívida bruta do 20-F é estoque em 31/12 (não se soma). Juros e lucro "
        "são fluxos do exercício. O 20-F de 2014 (prejuízo de 7.367) e o de 2015 "
        "(−8.450) são o espelho, nas demonstrações em US$, do impairment e das "
        "baixas que o TCU quantifica em reais nas DCs de 2014–2016.",
        "",
        "## 5. Dispositivo do Acórdão 1.839/2018",
        "",
        "Na sessão de 8/8/2018 o Plenário, relator Ministro José Múcio Monteiro, "
        "decidiu, em síntese:",
        "",
        "1. criar **processo apartado** sobre as vulnerabilidades de governança da "
        "União (achado que não se confunde com a conduta dos conselheiros);",
        "2. ouvir Casa Civil, Ministério da Fazenda, MPDG, MME e CNPE sobre a "
        "função de propriedade, a indicação de conselheiros e a indefinição do "
        "interesse público (art. 238 da Lei 6.404/1976);",
        "3. ouvir a **Petrobras** sobre o descumprimento dos deveres fiduciários "
        "dos Conselhos de Administração e Fiscal;",
        "4. cientificar CVM e MPF; juntar o acórdão aos TCs das três "
        "fiscalizações de gestão (Comperj, Premium e Rnest);",
        "5. recomendar à CVM seção no Formulário de Referência sobre "
        "irregularidades em apuração;",
        "6. classificar o relatório como público;",
        "7. encaminhar o acórdão aos Presidentes do Senado e da Câmara e a todas "
        "as comissões permanentes, para a revisão da Lei 13.303/2017.",
        "",
        "## 6. Fonte",
        "",
        "Tribunal de Contas da União, Acórdão nº 1.839/2018 – Plenário, "
        "TC-003.502/2016-3, Ata 30/2018, sessão de 8/8/2018. Autenticidade: "
        "www.tcu.gov.br/autenticidade, código 59581786. "
        "Séries 20-F: `output/petrobras_divida_bruta_20f_2002_2025.md`, "
        "`output/petrobras_juros_pagos_20f_2002_2026.md` e "
        "`output/petrobras_lucro_liquido_20f_2002_2026.md`.",
        "",
    ])


def _img(path: Path, width_cm: float = 17.6, height_cm: float = 22.0):
    if not path or not path.exists():
        return []
    img = Image(str(path), width=width_cm * cm, height=height_cm * cm)
    img.hAlign = "CENTER"
    return [img, Spacer(1, 6)]


def story_pdf(div, jur, luc, gerado: str, clips: dict[int, Path], s) -> list:
    story = [
        Spacer(1, 1.2 * cm),
        Paragraph("TCU · TC-003.502/2016-3 · Acórdão 1.839/2018-Plenário", s["kicker"]),
        Paragraph("Relatório", s["title"]),
        Paragraph(
            "Auditoria das refinarias Rnest, Comperj e Premium<br/>"
            "e o cruzamento com a dívida, os juros e o lucro da Petrobras",
            s["title"],
        ),
        Paragraph(
            "Tabelas conforme o PDF autenticado do acórdão<br/>"
            f"Elaborado em {gerado}",
            s["sub"],
        ),
        Paragraph("1. Apresentação", s["h1"]),
        Paragraph(
            "O presente relatório apresenta as informações acerca da <b>auditoria "
            "do TCU</b> sobre a conduta do Conselho de Administração da Petrobras "
            "nos projetos das refinarias <b>Rnest, Comperj e Premium I e II</b> e, "
            "em seguida, dos <b>juros pagos</b> e do <b>lucro líquido</b> da "
            "companhia nos mesmos anos em que o endividamento disparou.",
            s["body"],
        ),
        Paragraph(
            "As cifras do acórdão não são a Dívida Bruta do Governo Geral (DBGG). "
            "São o diagnóstico do TCU sobre a <b>própria Petrobras</b> e o "
            "cruzamento com o Form 20-F. As tabelas abaixo são as do PDF "
            "autenticado (código 59581786).",
            s["body"],
        ),
        Paragraph("2. O que o TCU apurou", s["h1"]),
        Paragraph(
            "Por volta de 2006 o plano era investir pouco mais de <b>US$ 12 "
            "bilhões</b> para acrescer cerca de <b>1.200 kbpd</b>. Passados dez "
            "anos, o orçamento ultrapassou <b>US$ 80 bilhões</b>, cerca de "
            "<b>US$ 30 bilhões</b> foram desembolsados e apenas cerca de "
            "<b>100 kbpd</b> entraram em operação. As Premium foram canceladas "
            "em 2015; a Rnest opera em regime parcial; o Comperj foi postergado. "
            "Nas DCs de 2014–2016 a companhia reconheceu cerca de <b>R$ 43 "
            "bilhões</b> em baixas e impairment. O TCU associa o insucesso ao "
            "endividamento então próximo de <b>US$ 100 bilhões</b>.",
            s["body"],
        ),
        Paragraph("3. Tabelas do acórdão, conforme estão no PDF", s["h1"]),
        Paragraph(
            "Tabela 2 — Resumo das reuniões do Conselho de Administração da "
            "Petrobras, 2004 a 2015 (página 9). Ocorrências = menções a Comperj, "
            "Rnest e Premium nas atas. Fonte: Petrobras/Segep (evidência 2).",
            s["body"],
        ),
        make_table(tabela2_rows(), s),
        Paragraph("Acórdão 1.839/2018, página 9 — Tabela 2.", s["caption"]),
    ]
    if 9 in clips:
        story.append(PageBreak())
        story.append(Paragraph("Tabela 2 no original (página 9)", s["h1"]))
        story.extend(_img(clips[9], 17.2, 24.2))
        story.append(Paragraph("Reprodução da página 9 do PDF autenticado.", s["caption"]))
    story.append(Paragraph(
        "Figura 22 — Evolução dos investimentos × viabilidade econômica "
        "(página 90). Valores em US$ bilhões, visão expedita do TCU.",
        s["body"],
    ))
    story.append(make_table(CAPEX_VPL, s))
    story.append(Paragraph(
        "O Capex vai de US$ 20,43 bi (FEL 1) a US$ 83,69 bi (epílogo). "
        "O VPL fecha em −US$ 43,32 bi.",
        s["body"],
    ))
    if 90 in clips:
        story.append(KeepTogether(_img(clips[90], 17.2, 24.2) + [
            Paragraph("Reprodução da página 90 (Figura 22 — Capex × VPL).", s["caption"]),
        ]))
    story.append(Paragraph(
        "Tabela 3 — Propostas de IG-P em obras da Petrobras, 2008–2015 "
        "(página 91). Fonte: portal do TCU, acesso em 4/10/2017.",
        s["body"],
    ))
    story.append(make_table(IGP, s))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Capacidade de refino projetada — totais de cada plano (página 91).",
        s["body"],
    ))
    story.append(make_table(CAPACIDADE_TOTAIS, s))
    if 91 in clips:
        story.append(PageBreak())
        story.append(Paragraph("IG-P e capacidade no original (página 91)", s["h1"]))
        story.extend(_img(clips[91], 17.2, 24.2))
        story.append(Paragraph("Reprodução da página 91 do PDF autenticado.", s["caption"]))
    if 49 in clips:
        story.append(Paragraph("Tabela 3 — indicação de membros do CA (página 49)", s["h1"]))
        story.append(Paragraph(
            "O acórdão reproduz a composição do Conselho entre 2005 e 2015, "
            "com o cargo na União (Casa Civil, Fazenda, MME). A página 49 "
            "abre a série com Dilma Rousseff na presidência do colegiado e "
            "ministros da Fazenda (Palocci / Mantega).",
            s["body"],
        ))
        story.extend(_img(clips[49], 17.2, 24.2))
        story.append(Paragraph("Reprodução da página 49 do PDF autenticado.", s["caption"]))
    story.append(PageBreak())
    story.append(Paragraph("4. Cruzamento com os Forms 20-F", s["h1"]))
    story.append(Paragraph(
        "Os anos em que o TCU registra o salto do Capex e a perda de VPL são "
        "os mesmos em que a dívida bruta do 20-F vai de US$ 21.338 milhões "
        "(2006) ao pico de <b>US$ 132.158 milhões</b> (2014), os juros pagos "
        "em caixa chegam a <b>US$ 7.308 milhões</b> (2016) e o lucro líquido "
        "vira prejuízo em 2014–2016.",
        s["body"],
    ))
    story.append(make_table(cruzamento_rows(div, jur, luc), s, left_first=False))
    story.append(Paragraph(
        "Dívida bruta = estoque em 31/12 (não se soma). Juros e lucro = fluxos "
        "do exercício no 20-F original.",
        s["caption"],
    ))
    if 80 in clips:
        story.append(Paragraph(
            "A Figura 18 do acórdão (página 80) cruza investimento, defasagem "
            "de preços do diesel e endividamento líquido/Ebitda — a mesma "
            "relação que o 20-F mostra pelo estoque da Gross Debt.",
            s["body"],
        ))
        story.extend(_img(clips[80], 17.2, 24.2))
        story.append(Paragraph("Reprodução da página 80 (Figura 18).", s["caption"]))
    story.append(Paragraph("5. Dispositivo do Acórdão 1.839/2018", s["h1"]))
    story.append(Paragraph(
        "Sessão de 8/8/2018, relator Ministro José Múcio Monteiro. O Plenário "
        "criou processo apartado sobre a governança da União; determinou oitivas "
        "da Casa Civil, Fazenda, MPDG, MME, CNPE e da Petrobras; cientificou CVM "
        "e MPF; recomendou seção no Formulário de Referência sobre "
        "irregularidades em apuração; classificou o relatório como público; e "
        "encaminhou o acórdão ao Congresso para a revisão da Lei 13.303/2017.",
        s["body"],
    ))
    if 130 in clips:
        story.extend(_img(clips[130], 17.2, 24.2))
        story.append(Paragraph("Reprodução da página 130 — abertura do dispositivo.", s["caption"]))
    story.append(Paragraph("6. Fonte", s["h1"]))
    story.append(Paragraph(
        "Tribunal de Contas da União, Acórdão nº 1.839/2018 – Plenário, "
        "TC-003.502/2016-3, Ata 30/2018. Autenticidade: "
        "www.tcu.gov.br/autenticidade, código 59581786. "
        "Séries 20-F em output/petrobras_*_20f_*.md.",
        s["body"],
    ))
    return story


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--saida-dir", type=Path, default=ROOT / "output")
    p.add_argument("--pdf", type=Path, default=None)
    args = p.parse_args()
    saida = args.saida_dir
    saida.mkdir(parents=True, exist_ok=True)
    gerado = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    div, jur, luc = df_divida(), df_juros(), df_lucro()
    pdf_src = localizar_pdf(args.pdf)
    clips: dict[int, Path] = {}
    if pdf_src:
        clips = recortar_paginas(pdf_src, saida)
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
        title="Relatório — Acórdão TCU 1.839/2018 e 20-F da Petrobras",
        author="SEC--data-analysys",
    )
    doc.build(
        story_pdf(div, jur, luc, gerado, clips, styles()),
        onFirstPage=header_footer,
        onLaterPages=header_footer,
    )
    art = Path("/opt/cursor/artifacts") / pdf_path.name
    if art.parent.is_dir():
        art.write_bytes(pdf_path.read_bytes())
    print(f"md: {md_path}")
    print(f"pdf: {pdf_path} ({pdf_path.stat().st_size} bytes)")
    print(f"fonte TCU: {pdf_src}")
    print(f"clips: {sorted(clips)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
