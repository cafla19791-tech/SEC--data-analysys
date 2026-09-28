#!/usr/bin/env python3
"""Lucro líquido de Itaú, Unibanco, Itaú Unibanco e Bradesco, atualizado pelo IPCA.

Série nominal (R$ milhões):
- 2002–2009: lucro líquido contábil consolidado publicado pelo próprio banco
  (DRE / earnings release em BR GAAP), depois das participações minoritárias.
- 2010–2025 e 1º semestre de 2026: lucro líquido consolidado atribuído aos
  acionistas da controladora, DFP/ITR da CVM (IFRS), última versão do exercício.

Atualização:
  valor_31_08_2026 = valor_nominal × índice(31/08/2026) / índice(data-base)

O índice é o produto acumulado de (1 + IPCA mensal). A data-base de cada
exercício anual é 31/12 do ano (o IPCA de janeiro do ano seguinte em diante
entra no fator). O 1º semestre de 2026 usa 30/06/2026.

IPCA: Banco Central, SGS 433 (variação mensal, % a.m.), mesma série do IBGE.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BCB_SGS = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados"
IPCA_COD = 433
DATA_ALVO = pd.Timestamp("2026-08-31")
DATA_ALVO_LABEL = "31/08/2026"

ROOT = Path(__file__).resolve().parents[1]
OUT_DEFAULT = ROOT / "output" / "lucros_liquidos_itau_unibanco_bradesco_ipca_2002_2026.xlsx"

# R$ milhões. None = a pessoa jurídica não publicou lucro anual isolado naquele ano.
# Itaú: Banco Itaú Holding Financeira S.A., consolidado, após minoritários.
ITAU = {
    2002: 2376.723,  # DRE 6-K 20/02/2004, R$ mil 2.376.723
    2003: 3151.820,
    2004: 3775.616,  # DRE 6-K 02/03/2005
    2005: 5251.334,  # DRE 6-K 24/02/2006
    2006: 4308.927,  # DRE 6-K 23/02/2007
    2007: 8474.000,  # release 13/02/2008, R$ 8.474 milhões
}
# Unibanco - União de Bancos Brasileiros S.A., consolidado BR GAAP (não é a holding pura).
UNIBANCO = {
    2002: 1010.000,  # release 14/02/2003
    2003: 1052.346,  # DRE comparativa em 2004
    2004: 1283.208,
    2005: 1838.483,
    2006: 1750.011,  # depois da amortização extraordinária de ágio
    2007: 3447.825,
}
# Itaú Unibanco Holding S.A. 2008 é o consolidado legal (Unibanco só no 4º trimestre).
# De 2009 em diante, atribuído à controladora.
ITAU_UNIBANCO = {
    2008: 7803.000,
    2009: 10067.000,
    2010: 11708.000,
    2011: 13837.000,
    2012: 12634.000,
    2013: 16424.000,
    2014: 21555.000,
    2015: 25740.000,
    2016: 23263.000,
    2017: 23903.000,
    2018: 24907.000,
    2019: 27113.000,
    2020: 18896.000,
    2021: 26760.000,
    2022: 29702.000,
    2023: 33105.000,
    2024: 41085.000,
    2025: 44857.000,
}
BRADESCO = {
    2002: 2023.000,
    2003: 2306.000,
    2004: 3060.000,
    2005: 5514.000,
    2006: 5054.000,  # reported; recorrente foi 6.363
    2007: 8010.000,
    2008: 7620.000,
    2009: 8012.000,
    2010: 9939.575,
    2011: 10958.054,
    2012: 11291.570,
    2013: 12395.920,
    2014: 15314.943,
    2015: 18132.906,
    2016: 17894.249,
    2017: 17089.364,
    2018: 16583.915,
    2019: 21023.023,
    2020: 15836.862,
    2021: 23172.322,
    2022: 20983.690,
    2023: 14251.329,
    2024: 17252.900,
    2025: 23672.706,
}
# 1º semestre de 2026, atribuído à controladora (ITR CVM, 01/01 a 30/06/2026).
SEMESTRE_2026 = {
    "itau_unibanco": 23615.000,
    "bradesco": 12317.787,
}

OBS = {
    2002: "Itaú e Unibanco: DRE consolidada BR GAAP. Bradesco: release, net income de R$ 2.023 milhões.",
    2006: "Contábil, não o recorrente. Bradesco recorrente R$ 6.363 mi; Itaú sem efeitos do BankBoston R$ 6.480 mi; Unibanco antes do ágio R$ 2.210 mi.",
    2007: "Contábil. Itaú recorrente R$ 7.179 mi; Unibanco recorrente R$ 2.600 mi; Bradesco ajustado citado depois foi R$ 7.210 mi.",
    2008: "Fusão anunciada em 03/11/2008. Não há lucro anual isolado de Itaú nem de Unibanco em 2008. O valor de Itaú Unibanco (R$ 7.803 mi) é o consolidado legal e inclui o Unibanco só no 4º trimestre. No fato relevante, o acumulado jan–set/2008 era R$ 5,9 bi (Itaú) e R$ 2,2 bi (Unibanco).",
    2009: "Primeiro exercício completo do Itaú Unibanco. Lucro da controladora. O comparativo pró-forma de 2008 (soma Itaú + Unibanco) era R$ 10.004 mi e não entra na série.",
    2010: "A partir daqui: DFP CVM, IFRS, atribuído aos sócios da controladora. O release do Bradesco falou em R$ 10.022 mi; a DFP (versão mais recente) registra R$ 9.939,575 mi na controladora.",
    2020: "No Itaú a controladora (R$ 18.896 mi) supera o consolidado (R$ 15.064 mi) porque a participação de não controladores foi negativa.",
    2024: "Bradesco: o release chama de lucro líquido contábil R$ 19.086 mi (conglomerado). A DFP IFRS da controladora é R$ 17.252,9 mi.",
    2025: "Itaú: recorrente gerencial R$ 46,8 bi; contábil da controladora na DFP R$ 44.857 mi. Bradesco: recorrente R$ 24.652 mi e contábil do release R$ 24.550 mi; DFP da controladora R$ 23.672,7 mi.",
    2026: "Exercício em aberto. Valor do 1º semestre (ITR de 30/06/2026), não anualizado. Data-base da correção: 30/06/2026. Fora da soma 2002–2025.",
}


def baixar_ipca(inicio: str = "01/12/2001", fim: str = "31/08/2026") -> pd.DataFrame:
    """IPCA mensal % a.m. (SGS 433), com índice acumulado no fim de cada mês."""
    url = BCB_SGS.format(cod=IPCA_COD)
    inicio_ts = pd.Timestamp(datetime.strptime(inicio, "%d/%m/%Y"))
    fim_ts = pd.Timestamp(datetime.strptime(fim, "%d/%m/%Y"))
    partes: list[pd.DataFrame] = []
    cursor = inicio_ts
    while cursor <= fim_ts:
        bloco_fim = min(cursor + pd.DateOffset(years=9, months=11), fim_ts)
        resp = requests.get(
            url,
            params={
                "formato": "json",
                "dataInicial": cursor.strftime("%d/%m/%Y"),
                "dataFinal": bloco_fim.strftime("%d/%m/%Y"),
            },
            timeout=120,
        )
        resp.raise_for_status()
        dados = resp.json()
        if dados:
            partes.append(pd.DataFrame(dados))
        cursor = bloco_fim + pd.DateOffset(days=1)
    if not partes:
        raise RuntimeError("Série IPCA (SGS 433) vazia")
    df = pd.concat(partes, ignore_index=True)
    df["data"] = pd.to_datetime(df["data"], dayfirst=True)
    df["ipca_pct"] = pd.to_numeric(df["valor"], errors="coerce")
    df = (
        df.dropna(subset=["data", "ipca_pct"])
        .drop_duplicates("data")
        .sort_values("data")
        .reset_index(drop=True)
    )
    df["mes"] = df["data"].dt.to_period("M").dt.to_timestamp("M")
    df["indice"] = (1.0 + df["ipca_pct"] / 100.0).cumprod()
    return df[["mes", "ipca_pct", "indice"]]


def indice_no_mes(ipca: pd.DataFrame, mes: pd.Timestamp) -> float:
    """Índice no fim do mês (último mês disponível não posterior a `mes`)."""
    mes = pd.Timestamp(mes).to_period("M").to_timestamp("M")
    ate = ipca.loc[ipca["mes"] <= mes, "indice"]
    if ate.empty:
        raise KeyError(f"IPCA sem observação até {mes.date()}")
    return float(ate.iloc[-1])


def fator_ipca(ipca: pd.DataFrame, data_base: pd.Timestamp, data_alvo: pd.Timestamp = DATA_ALVO) -> float:
    """Multiplicador que leva um valor do fim de `data_base` até o fim de `data_alvo`."""
    base = indice_no_mes(ipca, data_base)
    alvo = indice_no_mes(ipca, data_alvo)
    if base <= 0:
        raise ValueError("índice IPCA inválido")
    return alvo / base


def montar_tabela(ipca: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for ano in range(2002, 2026):
        data_base = pd.Timestamp(year=ano, month=12, day=31)
        fator = fator_ipca(ipca, data_base)
        itau = ITAU.get(ano)
        uni = UNIBANCO.get(ano)
        iu = ITAU_UNIBANCO.get(ano)
        brad = BRADESCO.get(ano)
        # Série contínua do grupo que virou Itaú Unibanco, sem somar 2008 duas vezes.
        if ano <= 2007:
            grupo = (itau or 0.0) + (uni or 0.0)
        else:
            grupo = iu
        linhas.append(
            {
                "ano": ano,
                "data_base": data_base,
                "cobertura": "Ano completo",
                "itau": itau,
                "unibanco": uni,
                "itau_unibanco": iu,
                "bradesco": brad,
                "grupo_itau": grupo,
                "fator": fator,
                "obs": OBS.get(ano, ""),
            }
        )
    fator_s1 = fator_ipca(ipca, pd.Timestamp("2026-06-30"))
    linhas.append(
        {
            "ano": 2026,
            "data_base": pd.Timestamp("2026-06-30"),
            "cobertura": "1º semestre",
            "itau": None,
            "unibanco": None,
            "itau_unibanco": SEMESTRE_2026["itau_unibanco"],
            "bradesco": SEMESTRE_2026["bradesco"],
            "grupo_itau": SEMESTRE_2026["itau_unibanco"],
            "fator": fator_s1,
            "obs": OBS[2026],
        }
    )
    df = pd.DataFrame(linhas)
    for col in ("itau", "unibanco", "itau_unibanco", "bradesco", "grupo_itau"):
        df[f"{col}_ipca"] = df[col] * df["fator"]
    df["ipca_acum_pct"] = (df["fator"] - 1.0) * 100.0
    return df


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold: bool = False, color: str = "000000", size: int = 11, name: str = "Calibri") -> Font:
    return Font(bold=bold, color=color, size=size, name=name)


THIN = Border(
    left=Side(style="thin", color="D0D5DD"),
    right=Side(style="thin", color="D0D5DD"),
    top=Side(style="thin", color="D0D5DD"),
    bottom=Side(style="thin", color="D0D5DD"),
)
NUM = '#,##0.000'
PCT = '0.00'
FATOR = '0.0000'


def _escrever_lucros(wb: Workbook, tabela: pd.DataFrame, ipca: pd.DataFrame) -> None:
    ws = wb.active
    ws.title = "Lucros_IPCA"

    ws.merge_cells("A1:R1")
    ws["A1"] = (
        "Lucro líquido contábil — Itaú, Unibanco, Itaú Unibanco e Bradesco — "
        f"atualizado pelo IPCA para {DATA_ALVO_LABEL}"
    )
    ws["A1"].font = _font(True, "FFFFFF", 16)
    ws["A1"].fill = _fill("1B3A4B")
    ws["A1"].alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:R2")
    ultimo = ipca["mes"].max()
    ws["A2"] = (
        "Valores em R$ milhões. Data-base do exercício anual: 31/12. "
        f"Índice de chegada: IPCA de {ultimo.strftime('%m/%Y')} (SGS 433, Bacen). "
        "2026 é o 1º semestre e não entra na soma. "
        "Grupo Itaú = Itaú + Unibanco até 2007 e Itaú Unibanco de 2008 em diante."
    )
    ws["A2"].font = _font(False, "1B3A4B", 10)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 32

    headers = [
        ("A", "Ano", "1B3A4B"),
        ("B", "Data-base", "1B3A4B"),
        ("C", "Cobertura", "1B3A4B"),
        ("D", "Itaú nominal", "EC7000"),
        ("E", "Unibanco nominal", "1F4E79"),
        ("F", "Itaú Unibanco nominal", "C45C26"),
        ("G", "Bradesco nominal", "CC092F"),
        ("H", "Grupo Itaú nominal", "7A4E12"),
        ("I", "Fator IPCA", "0F6E56"),
        ("J", "IPCA acumulado %", "0F6E56"),
        ("K", "Itaú em 31/08/2026", "EC7000"),
        ("L", "Unibanco em 31/08/2026", "1F4E79"),
        ("M", "Itaú Unibanco em 31/08/2026", "C45C26"),
        ("N", "Bradesco em 31/08/2026", "CC092F"),
        ("O", "Grupo Itaú em 31/08/2026", "7A4E12"),
        ("P", "Observação", "1B3A4B"),
    ]
    for col, titulo, cor in headers:
        cell = ws[f"{col}4"]
        cell.value = titulo
        cell.font = _font(True, "FFFFFF", 10)
        cell.fill = _fill(cor)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = THIN
    ws.row_dimensions[4].height = 36
    ws.auto_filter.ref = "A4:P29"
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:P29"

    nominal_cols = ["itau", "unibanco", "itau_unibanco", "bradesco", "grupo_itau"]
    ipca_cols = [f"{c}_ipca" for c in nominal_cols]

    for i, row in enumerate(tabela.itertuples(index=False), start=5):
        ws.cell(i, 1, int(row.ano)).font = _font(True)
        ws.cell(i, 2, row.data_base.strftime("%d/%m/%Y"))
        ws.cell(i, 3, row.cobertura)
        valores = [getattr(row, c) for c in nominal_cols]
        for col_i, valor in enumerate(valores, start=4):
            cell = ws.cell(i, col_i, None if pd.isna(valor) else float(valor))
            cell.number_format = NUM
        ws.cell(i, 9, float(row.fator)).number_format = FATOR
        ws.cell(i, 10, float(row.ipca_acum_pct)).number_format = PCT
        for col_i, nome in enumerate(ipca_cols, start=11):
            valor = getattr(row, nome)
            cell = ws.cell(i, col_i, None if pd.isna(valor) else float(valor))
            cell.number_format = NUM
        ws.cell(i, 16, row.obs).alignment = Alignment(wrap_text=True, vertical="center")
        fundo = "F7F4EF" if row.ano == 2026 else ("F4F7FB" if i % 2 == 0 else "FFFFFF")
        if row.ano == 2008:
            fundo = "FFF6E8"
        for col in range(1, 17):
            ws.cell(i, col).fill = _fill(fundo)
            ws.cell(i, col).border = THIN
            ws.cell(i, col).font = _font(False, "1A1A1A", 10)
            if col in (1, 3):
                ws.cell(i, col).alignment = Alignment(horizontal="center")
        ws.row_dimensions[i].height = 32 if row.obs else 18

    # Soma só dos anos completos (linhas 5 a 28 = 2002–2025).
    total_row = 30
    ws.cell(total_row, 1, "Soma 2002–2025").font = _font(True, "FFFFFF", 11)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    for col in (4, 5, 6, 7, 8, 11, 12, 13, 14, 15):
        letter = get_column_letter(col)
        cell = ws.cell(total_row, col, f"=SUM({letter}5:{letter}28)")
        cell.number_format = NUM
        cell.font = _font(True, "FFFFFF", 10)
    for col in range(1, 17):
        ws.cell(total_row, col).fill = _fill("1B3A4B")
        ws.cell(total_row, col).font = _font(True, "FFFFFF", 10)
        ws.cell(total_row, col).border = THIN
    ws.cell(total_row, 16, "Não inclui o 1º semestre de 2026. Somar Itaú, Unibanco e Itaú Unibanco junto conta a fusão em dobro; use a coluna Grupo Itaú.")
    ws.cell(total_row, 16).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[total_row].height = 32

    ws.cell(32, 1, "Como ler a soma").font = _font(True, "1B3A4B", 12)
    ws.merge_cells("A33:P35")
    ws["A33"] = (
        "A coluna Grupo Itaú junta os predecessores (Itaú + Unibanco) de 2002 a 2007 e, "
        "de 2008 em diante, usa só o Itaú Unibanco. 2008 não é a soma dos dois bancos no ano inteiro: "
        "é o lucro legal da instituição que continuou, com o Unibanco consolidado apenas no 4º trimestre. "
        "Bradesco é a série contínua do Banco Bradesco S.A. O 1º semestre de 2026 está na linha 29 "
        "para não ser lido como lucro anual."
    )
    ws["A33"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[33].height = 20
    ws.row_dimensions[34].height = 20
    ws.row_dimensions[35].height = 20

    larguras = {
        "A": 14, "B": 14, "C": 16, "D": 16, "E": 18, "F": 22, "G": 18, "H": 20,
        "I": 14, "J": 18, "K": 20, "L": 22, "M": 26, "N": 22, "O": 24, "P": 62,
    }
    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "Lucro líquido atualizado pelo IPCA"
    ws.oddFooter.right.text = "Página &P de &N"
    ws.print_title_rows = "1:4"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110

    chart = LineChart()
    chart.title = "Lucro líquido em R$ de 31/08/2026"
    chart.style = 10
    chart.y_axis.title = "R$ milhões de 31/08/2026"
    chart.x_axis.title = None
    chart.height = 8
    chart.width = 18
    chart.legend.position = "b"
    # Anos completos (linhas 5–28): Bradesco e Grupo Itaú já em R$ de 31/08/2026.
    cats = Reference(ws, min_col=1, min_row=5, max_row=28)
    data = Reference(ws, min_col=14, max_col=15, min_row=4, max_row=28)
    chart.add_data(data, from_rows=False, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, "A37")

    # Segundo gráfico só do Bradesco e do grupo, já incluído (N e O).
    # Itaú solo e Unibanco ficam na tabela; o gráfico compara as duas franquias contínuas.
    if chart.series:
        chart.series[0].graphicalProperties.line.solidFill = "CC092F"
        if len(chart.series) > 1:
            chart.series[1].graphicalProperties.line.solidFill = "EC7000"


def _escrever_ipca(wb: Workbook, ipca: pd.DataFrame) -> None:
    ws = wb.create_sheet("IPCA")
    ws["A1"] = "IPCA mensal usado na atualização — Bacen SGS 433"
    ws["A1"].font = _font(True, "FFFFFF", 14)
    ws["A1"].fill = _fill("0F6E56")
    ws.merge_cells("A1:D1")
    ws["A2"] = (
        "O índice do fim do mês é o produto acumulado de (1 + IPCA/100) desde dez/2001. "
        "Fator de um exercício = índice(31/08/2026) / índice(31/12 do ano). "
        "Isso aplica o IPCA de janeiro do ano seguinte até agosto de 2026."
    )
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 32
    for col, titulo in enumerate(("Mês", "IPCA % a.m.", "Índice fim do mês", "Referência"), start=1):
        cell = ws.cell(4, col, titulo)
        cell.font = _font(True, "FFFFFF")
        cell.fill = _fill("0F6E56")
        cell.border = THIN
    indice_alvo = indice_no_mes(ipca, DATA_ALVO)
    for i, row in enumerate(ipca.itertuples(index=False), start=5):
        ws.cell(i, 1, row.mes.strftime("%m/%Y"))
        ws.cell(i, 2, float(row.ipca_pct)).number_format = "0.00"
        ws.cell(i, 3, float(row.indice)).number_format = "0.000000"
        marca = ""
        if row.mes.month == 12:
            marca = "data-base dos lucros anuais"
        if row.mes.year == 2026 and row.mes.month == 8:
            marca = "data de chegada 31/08/2026"
        if row.mes.year == 2026 and row.mes.month == 6:
            marca = "data-base do 1º semestre de 2026"
        ws.cell(i, 4, marca)
        if marca:
            for col in range(1, 5):
                ws.cell(i, col).fill = _fill("E5F6F1")
        for col in range(1, 5):
            ws.cell(i, col).border = THIN
            ws.cell(i, col).font = _font(size=10)
    ws.cell(3, 6, "Índice em 31/08/2026")
    ws.cell(3, 7, indice_alvo).number_format = "0.000000"
    ws.cell(3, 6).font = _font(True)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 42
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:D{4 + len(ipca)}"
    ws.sheet_view.showGridLines = False


def _escrever_fontes(wb: Workbook) -> None:
    ws = wb.create_sheet("Fontes_e_criterio")
    ws["A1"] = "Critério, fontes e o que a série não é"
    ws["A1"].font = _font(True, "FFFFFF", 14)
    ws["A1"].fill = _fill("1B3A4B")
    ws.merge_cells("A1:B1")
    blocos = [
        ("O que está na coluna nominal",
         "Lucro líquido contábil consolidado, em R$ milhões da época. "
         "De 2002 a 2009 é o número que o próprio banco chamou de net income / lucro líquido na DRE, "
         "depois das participações de minoritários — não o lucro recorrente nem o ajustado. "
         "De 2010 em diante é a linha da DFP/ITR consolidada da CVM "
         "“atribuído aos sócios da empresa controladora”, última versão de cada exercício, escala mil convertida para milhões. "
         "Essa é a linha do acionista da holding listada (Itaú Unibanco Holding S.A., CNPJ 60.872.504/0001-23, "
         "e Banco Bradesco S.A., CNPJ 60.746.948/0001-12)."),
        ("Por que não o lucro recorrente",
         "O recorrente muda de definição no tempo e tira eventos que o banco considera extraordinários "
         "(ágio, venda de participação, provisão). O pedido é de lucro líquido. "
         "Onde a diferença é grande, a observação do ano registra o recorrente ao lado."),
        ("Itaú, Unibanco e a fusão",
         "Itaú até 2007 é o Banco Itaú Holding Financeira S.A., a holding listada, não a Itaúsa "
         "(que também consolidava Duratex e Itautec). "
         "Unibanco é o Unibanco - União de Bancos Brasileiros S.A., o banco operacional. "
         "A Unibanco Holdings S.A. tinha lucro menor porque não detinha 100% do banco; "
         "não é a série usada aqui, porque o número comparado com Itaú e Bradesco na época era o do banco. "
         "A associação foi anunciada em 3 de novembro de 2008. O lucro legal de 2008 da instituição que continuou "
         "é R$ 7.803 milhões e só traz o Unibanco no quarto trimestre. "
         "Por isso 2008 não aparece nas colunas Itaú e Unibanco. "
         "2009 é o primeiro ano cheio do Itaú Unibanco."),
        ("Bradesco recente e o release",
         "Em 2024 e 2025 o release do Bradesco destaca um lucro líquido contábil "
         "(R$ 19.086 milhões e R$ 24.550 milhões) maior do que a controladora na DFP IFRS "
         "(R$ 17.252,9 milhões e R$ 23.672,7 milhões). "
         "A planilha fica na DFP, que é a demonstração auditada depositada na CVM, "
         "para a série de 2010 a 2026 usar a mesma linha."),
        ("Correção pelo IPCA",
         "Fonte: Banco Central, série SGS 433, variação mensal do IPCA. "
         "O IPCA de agosto de 2026, divulgado pelo IBGE, foi −0,32%, com alta de 3,11% no ano. "
         "Um lucro de 31/12/AAAA é multiplicado pelo produto dos IPCA de janeiro/AAAA+1 até agosto/2026. "
         "O 1º semestre de 2026 é multiplicado só por julho e agosto de 2026. "
         "Não há correção dentro do próprio ano do lucro: o resultado anual inteiro está posicionado em 31 de dezembro."),
        ("2026",
         "Em 28/09/2026 o exercício de 2026 não está fechado. Os bancos já publicaram o ITR do segundo trimestre. "
         "A linha 2026 é o lucro do primeiro semestre (1º de janeiro a 30 de junho), sem anualizar."),
        ("Itaú 2002–2007",
         "DRE consolidada arquivada na SEC: 6-K de 20/02/2004 (2002 e 2003), "
         "6-K de 02/03/2005 arquivo bi2149 (2004), 6-K de 24/02/2006 arquivo ba4918 (2005), "
         "6-K de 23/02/2007 arquivo bi9005 (2006) e release de 13/02/2008 (2007, R$ 8.474 milhões)."),
        ("Unibanco 2002–2007",
         "Releases e DRE em BR GAAP na SEC, CIK 0001038583: "
         "14/02/2003 (2002, R$ 1.010 milhões), DRE de 2004 (2003 R$ 1.052,346 milhões e 2004 R$ 1.283,208 milhões), "
         "release de 23/02/2006 (2005), DRE de 2006 e de 2007."),
        ("Itaú Unibanco 2008–2009",
         "2008: lucro líquido consolidado de R$ 7.803 milhões no relatório do exercício e na proposta à assembleia. "
         "2009: lucro líquido da controladora de R$ 10.067 milhões no release do 4º trimestre de 2009."),
        ("2010–2025 e 1º semestre de 2026",
         "CVM Dados Abertos, DFP consolidada (dfp_cia_aberta_DRE_con) e ITR 2026. "
         "Conta de lucro atribuído à controladora, ordem do exercício “ÚLTIMO”, maior número de versão. "
         "Arquivos consultados em 28/09/2026."),
        ("Bradesco 2002–2009",
         "Earnings releases na SEC, CIK 0001160330, linha Net Income (não a linha Adjusted / Recurring): "
         "2002 R$ 2.023 mi, 2003 R$ 2.306 mi, 2004 R$ 3.060 mi, 2005 R$ 5.514 mi, "
         "2006 reportado R$ 5.054 mi, 2007 R$ 8.010 mi, 2008 R$ 7.620 mi, 2009 R$ 8.012 mi."),
    ]
    linha = 3
    for titulo, texto in blocos:
        ws.cell(linha, 1, titulo).font = _font(True, "1B3A4B", 12)
        ws.merge_cells(start_row=linha + 1, start_column=1, end_row=linha + 1, end_column=2)
        ws.cell(linha + 1, 1, texto).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[linha + 1].height = 48
        linha += 3
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 110
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.orientation = "landscape"


def gravar_planilha(tabela: pd.DataFrame, ipca: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    _escrever_lucros(wb, tabela, ipca)
    _escrever_ipca(wb, ipca)
    _escrever_fontes(wb)
    wb.properties.title = "Lucros líquidos Itaú, Unibanco e Bradesco atualizados pelo IPCA"
    wb.properties.creator = "scripts/lucros_bancos_ipca.py"
    wb.properties.description = (
        "Lucro líquido contábil 2002–2025 e 1º semestre de 2026, "
        "em R$ de 31/08/2026 pelo IPCA (SGS 433)."
    )
    wb.save(destino)
    return destino


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--saida", type=Path, default=OUT_DEFAULT)
    parser.add_argument(
        "--ipca-csv",
        type=Path,
        default=None,
        help="CSV com colunas mes (AAAA-MM-DD) e ipca_pct, para não baixar o SGS.",
    )
    args = parser.parse_args()
    if args.ipca_csv is not None:
        bruto = pd.read_csv(args.ipca_csv)
        bruto["mes"] = pd.to_datetime(bruto["mes"]).dt.to_period("M").dt.to_timestamp("M")
        bruto["ipca_pct"] = pd.to_numeric(bruto["ipca_pct"])
        bruto = bruto.sort_values("mes").drop_duplicates("mes")
        bruto["indice"] = (1.0 + bruto["ipca_pct"] / 100.0).cumprod()
        ipca = bruto[["mes", "ipca_pct", "indice"]].reset_index(drop=True)
    else:
        print("Baixando IPCA (Bacen SGS 433)...")
        ipca = baixar_ipca()
    if indice_no_mes(ipca, DATA_ALVO) == indice_no_mes(ipca, pd.Timestamp("2026-07-31")):
        raise RuntimeError("IPCA de agosto/2026 não está na série; a data de chegada ficaria em julho.")
    tabela = montar_tabela(ipca)
    destino = gravar_planilha(tabela, ipca, args.saida)
    cheios = tabela[tabela["cobertura"] == "Ano completo"]
    print(f"Planilha: {destino}")
    print(f"IPCA ago/2026: {float(ipca.loc[ipca['mes'].dt.year.eq(2026) & ipca['mes'].dt.month.eq(8), 'ipca_pct'].iloc[0]):.2f}%")
    print(
        "Soma 2002–2025 em R$ milhões de 31/08/2026 — "
        f"Grupo Itaú {cheios['grupo_itau_ipca'].sum():,.1f}; "
        f"Bradesco {cheios['bradesco_ipca'].sum():,.1f}"
    )


if __name__ == "__main__":
    main()
