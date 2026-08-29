#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Discriminativo de taxas básicas de juros reais por país (BIS).

Uma aba por país, com colunas:

  - Mês/ano
  - Taxa básica nominal (% a.a.)  — BIS WS_CBPOL, fim de período
  - Índice de inflação oficial (2010=100) — BIS WS_LONG_CPI
  - Inflação no mês (% a.m.)
  - Taxa básica real no mês (% a.m.)
  - Taxa básica real acumulada no ano (%)  — linha após dezembro

Fórmula de Fisher (composta):

  i_m = (1 + i_aa)^(1/12) − 1
  π_m = IPC_t / IPC_{t−1} − 1
  r_m = (1 + i_m) / (1 + π_m) − 1
  R_ano = Π (1 + r_m) − 1

Fontes (bulk download BIS)::

  https://data.bis.org/bulkdownload
  https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip
  https://data.bis.org/static/bulk/WS_LONG_CPI_csv_flat.zip

Uso::

  python scripts/discriminativo_juros_reais_paises.py
  python scripts/discriminativo_juros_reais_paises.py
  python scripts/discriminativo_juros_reais_paises.py --ano-inicio 1995
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

MARKER = "juros-reais-paises-20260829"
ANO_INICIO_DEFAULT = 1995
BIS_BULK = "https://data.bis.org/bulkdownload"
URL_CBPOL = "https://data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip"
URL_CPI = "https://data.bis.org/static/bulk/WS_LONG_CPI_csv_flat.zip"
ZIP_CBPOL = "WS_CBPOL_csv_flat.zip"
ZIP_CPI = "WS_LONG_CPI_csv_flat.zip"

COL_MES = "Mês/ano"
COL_NOMINAL = "Taxa básica nominal (% a.a.)"
COL_INDICE = "Índice de inflação oficial (2010=100)"
COL_INFLACAO = "Inflação no mês (% a.m.)"
COL_REAL_MES = "Taxa básica real no mês (% a.m.)"
COL_REAL_ANO = "Taxa básica real acumulada no ano (%)"
COL_TIPO = "tipo"

COLS_OUT = (COL_MES, COL_NOMINAL, COL_INDICE, COL_INFLACAO, COL_REAL_MES, COL_REAL_ANO)

PAISES_PT: dict[str, str] = {
    "AR": "Argentina",
    "AT": "Áustria",
    "AU": "Austrália",
    "BE": "Bélgica",
    "BR": "Brasil",
    "CA": "Canadá",
    "CH": "Suíça",
    "CL": "Chile",
    "CN": "China",
    "CO": "Colômbia",
    "CZ": "Tchéquia",
    "DE": "Alemanha",
    "DK": "Dinamarca",
    "ES": "Espanha",
    "FR": "França",
    "GB": "Reino Unido",
    "GR": "Grécia",
    "HK": "Hong Kong",
    "HR": "Croácia",
    "HU": "Hungria",
    "ID": "Indonésia",
    "IL": "Israel",
    "IN": "Índia",
    "IS": "Islândia",
    "IT": "Itália",
    "JP": "Japão",
    "KR": "Coreia",
    "KW": "Kuwait",
    "MA": "Marrocos",
    "MK": "Macedônia do Norte",
    "MX": "México",
    "MY": "Malásia",
    "NL": "Países Baixos",
    "NO": "Noruega",
    "NZ": "Nova Zelândia",
    "PE": "Peru",
    "PH": "Filipinas",
    "PL": "Polônia",
    "PT": "Portugal",
    "RO": "Romênia",
    "RS": "Sérvia",
    "RU": "Rússia",
    "SA": "Arábia Saudita",
    "SE": "Suécia",
    "TH": "Tailândia",
    "TR": "Turquia",
    "US": "Estados Unidos",
    "XM": "Zona do Euro",
    "ZA": "África do Sul",
}

# Euro-adotantes: série nacional costuma encerrar; a taxa vigente é a da Zona do Euro.
PAISES_EURO = frozenset(
    {"AT", "BE", "DE", "ES", "FR", "GR", "IT", "NL", "PT", "HR"}
)


def taxa_mensal_composta(taxa_aa: float) -> float:
    """Equivalente mensal composto de uma taxa anual (decimal)."""
    return (1.0 + float(taxa_aa)) ** (1.0 / 12.0) - 1.0


def taxa_real_fisher(taxa_nominal: float, inflacao: float) -> float:
    """Fisher: (1+i)/(1+π) − 1. Entradas e saída em decimal."""
    den = 1.0 + float(inflacao)
    if den == 0.0:
        return float("nan")
    return (1.0 + float(taxa_nominal)) / den - 1.0


def acumular_fator(taxas: pd.Series) -> float:
    """Π(1+r) − 1 ignorando nulos. ``taxas`` em decimal."""
    validas = pd.to_numeric(taxas, errors="coerce").dropna()
    if validas.empty:
        return float("nan")
    return float((1.0 + validas).prod() - 1.0)


def _parse_area(valor: str) -> tuple[str, str]:
    texto = str(valor).strip()
    if ":" in texto:
        codigo, nome = texto.split(":", 1)
        return codigo.strip(), nome.strip()
    return texto, texto


def nome_pais(codigo: str, fallback: str = "") -> str:
    return PAISES_PT.get(codigo, fallback or codigo)


def nome_aba(codigo: str, fallback: str = "") -> str:
    nome = nome_pais(codigo, fallback)
    nome = re.sub(r"[\\/*?:\[\]]", "-", nome)
    return nome[:31]


def formatar_mes_ano(ts: pd.Timestamp) -> str:
    return pd.Timestamp(ts).strftime("%m/%Y")


def baixar_arquivo(url: str, dest: Path, timeout: int = 300) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 50_000:
        print(f"[CACHE] {dest.name} ({dest.stat().st_size / 1e6:.1f} MB)")
        return dest
    print(f"[DOWNLOAD] {url}")
    r = requests.get(url, timeout=timeout, headers={"User-Agent": "SEC-data-analysys/1.0"})
    r.raise_for_status()
    dest.write_bytes(r.content)
    print(f"[OK] {dest.name} ({dest.stat().st_size / 1e6:.1f} MB)")
    return dest


def _ler_csv_zip(path: Path, **kwargs) -> pd.DataFrame:
    with zipfile.ZipFile(path) as zf:
        nomes = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not nomes:
            raise FileNotFoundError(f"Nenhum CSV em {path.name}")
        with zf.open(nomes[0]) as raw:
            return pd.read_csv(raw, **kwargs)


def carregar_cbpol_mensal(path_zip: Path) -> pd.DataFrame:
    """Taxa básica nominal mensal (fim de período), % a.a."""
    print(f"[LER] CBPOL mensal: {path_zip.name}")
    t0 = time.time()
    df = _ler_csv_zip(
        path_zip,
        usecols=[
            "FREQ:Frequency",
            "REF_AREA:Reference area",
            "TIME_PERIOD:Time period or range",
            "OBS_VALUE:Observation Value",
            "OBS_STATUS:Observation Status",
        ],
        low_memory=False,
    )
    df = df[df["FREQ:Frequency"].astype(str).str.startswith("M")].copy()
    status = df["OBS_STATUS:Observation Status"].astype(str)
    df = df[status.str.startswith("A") | (status == "") | status.eq("nan")]
    parsed = df["REF_AREA:Reference area"].map(_parse_area)
    df["codigo"] = parsed.map(lambda x: x[0])
    df["pais_en"] = parsed.map(lambda x: x[1])
    df["mes"] = pd.to_datetime(df["TIME_PERIOD:Time period or range"], format="%Y-%m", errors="coerce")
    df["nominal_aa"] = pd.to_numeric(df["OBS_VALUE:Observation Value"], errors="coerce") / 100.0
    out = (
        df.dropna(subset=["mes", "nominal_aa"])
        .sort_values(["codigo", "mes"])
        .drop_duplicates(["codigo", "mes"], keep="last")
        [["codigo", "pais_en", "mes", "nominal_aa"]]
        .reset_index(drop=True)
    )
    print(f"  → {len(out):,} observações / {out['codigo'].nunique()} países em {time.time() - t0:.1f}s")
    return out


def carregar_cpi_mensal(path_zip: Path) -> pd.DataFrame:
    """Índice de preços ao consumidor (2010=100), mensal."""
    print(f"[LER] CPI mensal: {path_zip.name}")
    t0 = time.time()
    df = _ler_csv_zip(
        path_zip,
        usecols=[
            "FREQ:Frequency",
            "REF_AREA:Reference area",
            "UNIT_MEASURE:Unit of measure",
            "TIME_PERIOD:Time period or range",
            "OBS_VALUE:Observation Value",
            "OBS_STATUS:Observation Status",
        ],
        low_memory=False,
    )
    df = df[df["FREQ:Frequency"].astype(str).str.startswith("M")].copy()
    unidade = df["UNIT_MEASURE:Unit of measure"].astype(str)
    df = df[unidade.str.startswith("628")]
    status = df["OBS_STATUS:Observation Status"].astype(str)
    df = df[status.str.startswith("A") | (status == "") | status.eq("nan")]
    parsed = df["REF_AREA:Reference area"].map(_parse_area)
    df["codigo"] = parsed.map(lambda x: x[0])
    df["pais_en"] = parsed.map(lambda x: x[1])
    df["mes"] = pd.to_datetime(df["TIME_PERIOD:Time period or range"], format="%Y-%m", errors="coerce")
    df["indice"] = pd.to_numeric(df["OBS_VALUE:Observation Value"], errors="coerce")
    out = (
        df.dropna(subset=["mes", "indice"])
        .sort_values(["codigo", "mes"])
        .drop_duplicates(["codigo", "mes"], keep="last")
        [["codigo", "pais_en", "mes", "indice"]]
        .reset_index(drop=True)
    )
    print(f"  → {len(out):,} observações / {out['codigo'].nunique()} países em {time.time() - t0:.1f}s")
    return out


def montar_serie_pais(
    nominal: pd.DataFrame,
    cpi: pd.DataFrame,
    *,
    codigo: str,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
) -> pd.DataFrame:
    """Cruza taxa nominal e IPC; calcula inflação, real mensal e acumulado anual."""
    nom = nominal[nominal["codigo"] == codigo][["mes", "nominal_aa", "pais_en"]].copy()
    ipc = cpi[cpi["codigo"] == codigo][["mes", "indice"]].copy()
    if nom.empty or ipc.empty:
        return pd.DataFrame()

    base = pd.merge(nom, ipc, on="mes", how="inner")
    if base.empty:
        return pd.DataFrame()
    base = base.sort_values("mes").drop_duplicates("mes").reset_index(drop=True)
    if ano_inicio is not None:
        base = base[base["mes"].dt.year >= ano_inicio]
    if ano_fim is not None:
        base = base[base["mes"].dt.year <= ano_fim]
    if base.empty:
        return pd.DataFrame()

    # IPC do mês anterior (mesmo país) — usa a série completa para o 1º mês do recorte
    ipc_full = ipc.sort_values("mes")
    ipc_map = ipc_full.set_index("mes")["indice"]
    prev_mes = base["mes"] - pd.DateOffset(months=1)
    base["indice_ant"] = prev_mes.map(ipc_map)
    base["inflacao_am"] = base["indice"] / base["indice_ant"] - 1.0
    base.loc[base["indice_ant"].isna() | (base["indice_ant"] == 0), "inflacao_am"] = pd.NA
    base["nominal_am"] = base["nominal_aa"].map(taxa_mensal_composta)
    base["real_am"] = [
        taxa_real_fisher(i, p) if pd.notna(i) and pd.notna(p) else pd.NA
        for i, p in zip(base["nominal_am"], base["inflacao_am"])
    ]
    base["ano"] = base["mes"].dt.year.astype(int)
    base["codigo"] = codigo
    return base.reset_index(drop=True)


def linhas_discriminativo(serie: pd.DataFrame) -> pd.DataFrame:
    """Linhas mensais + linha de acumulado após dezembro (e parcial no último ano)."""
    if serie.empty:
        return pd.DataFrame(columns=list(COLS_OUT) + [COL_TIPO, "ano", "codigo"])

    rows: list[dict] = []
    ultimo_mes = serie["mes"].max()
    for ano, bloco in serie.groupby("ano", sort=True):
        bloco = bloco.sort_values("mes")
        for rec in bloco.itertuples(index=False):
            rows.append(
                {
                    COL_MES: formatar_mes_ano(rec.mes),
                    COL_NOMINAL: rec.nominal_aa,
                    COL_INDICE: rec.indice,
                    COL_INFLACAO: rec.inflacao_am,
                    COL_REAL_MES: rec.real_am,
                    COL_REAL_ANO: pd.NA,
                    COL_TIPO: "mes",
                    "ano": ano,
                    "codigo": rec.codigo,
                    "mes": rec.mes,
                }
            )
        tem_dezembro = (bloco["mes"].dt.month == 12).any()
        eh_ultimo_ano = int(ano) == int(ultimo_mes.year)
        if tem_dezembro or eh_ultimo_ano:
            n = int(bloco["real_am"].notna().sum())
            if n == 0:
                continue
            acum = acumular_fator(bloco["real_am"])
            if tem_dezembro:
                rotulo = f"ACUMULADO {ano}"
                tipo = "acumulado"
            else:
                rotulo = f"ACUMULADO PARCIAL {ano} ({n} meses)"
                tipo = "parcial"
            rows.append(
                {
                    COL_MES: rotulo,
                    COL_NOMINAL: pd.NA,
                    COL_INDICE: pd.NA,
                    COL_INFLACAO: pd.NA,
                    COL_REAL_MES: pd.NA,
                    COL_REAL_ANO: acum,
                    COL_TIPO: tipo,
                    "ano": ano,
                    "codigo": bloco["codigo"].iloc[0],
                    "mes": pd.Timestamp(year=int(ano), month=12, day=1),
                }
            )
    return pd.DataFrame(rows)


def resumo_pais(serie: pd.DataFrame, linhas: pd.DataFrame) -> dict:
    if serie.empty:
        return {}
    ultimo = serie.sort_values("mes").iloc[-1]
    anuais = linhas[linhas[COL_TIPO] == "acumulado"]
    ultimo_ano_completo = None
    real_ultimo_ano = None
    if not anuais.empty:
        last_a = anuais.sort_values("ano").iloc[-1]
        ultimo_ano_completo = int(last_a["ano"])
        val = last_a[COL_REAL_ANO]
        real_ultimo_ano = float(val) if pd.notna(val) else None
    return {
        "codigo": ultimo["codigo"],
        "pais": nome_pais(ultimo["codigo"], str(ultimo.get("pais_en", ""))),
        "primeiro_mes": formatar_mes_ano(serie["mes"].min()),
        "ultimo_mes": formatar_mes_ano(ultimo["mes"]),
        "n_meses": int(len(serie)),
        "nominal_ultimo": float(ultimo["nominal_aa"]) if pd.notna(ultimo["nominal_aa"]) else None,
        "indice_ultimo": float(ultimo["indice"]) if pd.notna(ultimo["indice"]) else None,
        "inflacao_ultimo": float(ultimo["inflacao_am"]) if pd.notna(ultimo["inflacao_am"]) else None,
        "real_mes_ultimo": float(ultimo["real_am"]) if pd.notna(ultimo["real_am"]) else None,
        "ultimo_ano_completo": ultimo_ano_completo,
        "real_acumulada_ultimo_ano": real_ultimo_ano,
        "euro": ultimo["codigo"] in PAISES_EURO,
    }


def pivot_anual(
    por_pais: dict[str, pd.DataFrame],
    *,
    ano_minimo: int = ANO_INICIO_DEFAULT,
) -> pd.DataFrame:
    """Anos × países com a taxa real acumulada no ano (decimal)."""
    frames = []
    for codigo, linhas in por_pais.items():
        anuais = linhas[linhas[COL_TIPO] == "acumulado"][[COL_MES, COL_REAL_ANO, "ano"]].copy()
        if anuais.empty:
            continue
        if ano_minimo is not None:
            anuais = anuais[anuais["ano"] >= int(ano_minimo)]
        if anuais.empty:
            continue
        anuais = anuais.rename(columns={COL_REAL_ANO: nome_pais(codigo)})
        frames.append(anuais[["ano", nome_pais(codigo)]].set_index("ano"))
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, axis=1).sort_index()
    out = out.dropna(axis=1, how="all")
    cols = list(out.columns)
    if "Brasil" in cols:
        cols = ["Brasil"] + sorted([c for c in cols if c != "Brasil"], key=str.casefold)
    else:
        cols = sorted(cols, key=str.casefold)
    return out[cols]


def _estilo_base():
    thin = Border(
        left=Side(style="thin", color="B0BEC5"),
        right=Side(style="thin", color="B0BEC5"),
        top=Side(style="thin", color="B0BEC5"),
        bottom=Side(style="thin", color="B0BEC5"),
    )
    return {
        "header": Font(name="Calibri", bold=True, color="FFFFFF", size=11),
        "header_fill": PatternFill("solid", fgColor="1B4F72"),
        "title": Font(name="Calibri", bold=True, size=14, color="1B4F72"),
        "subtitle": Font(name="Calibri", size=10, color="34495E"),
        "mes": Font(name="Calibri", size=10),
        "acum": Font(name="Calibri", bold=True, size=10, color="1B4F72"),
        "acum_fill": PatternFill("solid", fgColor="D5F5E3"),
        "parcial_fill": PatternFill("solid", fgColor="FCF3CF"),
        "alt": PatternFill("solid", fgColor="EBF5FB"),
        "pct": "0.00%",
        "pct4": "0.0000%",
        "num": "#,##0.0000",
        "thin": thin,
        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "left": Alignment(horizontal="left", vertical="center"),
        "right": Alignment(horizontal="right", vertical="center"),
    }


def _escrever_cabecalho_aba(ws, titulo: str, subtitulo: str, sty: dict) -> int:
    ws.merge_cells("A1:F1")
    ws["A1"] = titulo
    ws["A1"].font = sty["title"]
    ws.merge_cells("A2:F2")
    ws["A2"] = subtitulo
    ws["A2"].font = sty["subtitle"]
    ws.merge_cells("A3:F3")
    ws["A3"] = (
        "Fisher: i_m = (1+i_aa)^(1/12)−1;  π_m = IPC_t/IPC_{t−1}−1;  "
        "r_m = (1+i_m)/(1+π_m)−1;  R_ano = Π(1+r_m)−1. "
        "Fonte: BIS Data Portal (WS_CBPOL + WS_LONG_CPI)."
    )
    ws["A3"].font = Font(name="Calibri", size=9, italic=True, color="5D6D7E")
    return 5


def _aplicar_larguras(ws, larguras: dict[int, float]) -> None:
    for col, w in larguras.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def escrever_aba_pais(
    wb: Workbook,
    codigo: str,
    serie: pd.DataFrame,
    linhas: pd.DataFrame,
    sty: dict,
) -> str:
    pais_en = str(serie["pais_en"].iloc[0]) if "pais_en" in serie.columns else codigo
    aba = nome_aba(codigo, pais_en)
    # nomes duplicados (raro) — sufixa código
    if aba in wb.sheetnames:
        aba = f"{aba[:28]}-{codigo}"[:31]
    ws = wb.create_sheet(aba)
    pais = nome_pais(codigo, pais_en)
    nota_euro = ""
    if codigo in PAISES_EURO:
        nota_euro = " Série nacional; após a adesão ao euro use a aba Zona do Euro."
    elif codigo == "XM":
        nota_euro = " Taxa de política do BCE (Zona do Euro)."
    sub = (
        f"{pais} ({codigo}). Taxa básica: BIS CBPOL mensal, fim de período, % a.a. "
        f"Índice: BIS CPI longo, 2010=100.{nota_euro}"
    )
    header_row = _escrever_cabecalho_aba(ws, f"Taxas básicas de juros reais — {pais}", sub, sty)

    for j, col in enumerate(COLS_OUT, start=1):
        cell = ws.cell(header_row, j, col)
        cell.font = sty["header"]
        cell.fill = sty["header_fill"]
        cell.alignment = sty["center"]
        cell.border = sty["thin"]
    ws.row_dimensions[header_row].height = 32
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:F{header_row + max(len(linhas), 1)}"

    r = header_row + 1
    for rec in linhas.itertuples(index=False):
        tipo = rec.tipo
        valores = (
            rec[0],
            rec[1],
            rec[2],
            rec[3],
            rec[4],
            rec[5],
        )
        fill = None
        font = sty["mes"]
        if tipo == "acumulado":
            fill = sty["acum_fill"]
            font = sty["acum"]
        elif tipo == "parcial":
            fill = sty["parcial_fill"]
            font = sty["acum"]
        elif r % 2 == 0:
            fill = sty["alt"]
        for j, val in enumerate(valores, start=1):
            cell = ws.cell(r, j)
            cell.font = font
            cell.border = sty["thin"]
            if fill is not None:
                cell.fill = fill
            if j == 1:
                cell.value = val
                cell.alignment = sty["left"]
            elif pd.isna(val):
                cell.value = None
            else:
                cell.value = float(val)
                if j == 3:
                    cell.number_format = sty["num"]
                elif j in (4, 5):
                    cell.number_format = sty["pct4"]
                else:
                    cell.number_format = sty["pct"]
                cell.alignment = sty["right"]
        r += 1

    _aplicar_larguras(ws, {1: 36, 2: 28, 3: 36, 4: 24, 5: 32, 6: 40})
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.left.text = f"&B{pais} — juros reais"
    return aba


def escrever_capa(wb: Workbook, resumos: list[dict], sty: dict, gerado: datetime) -> None:
    ws = wb.active
    ws.title = "Capa"
    ws["A1"] = "Discriminativo — taxas básicas de juros reais por país"
    ws["A1"].font = Font(name="Calibri", bold=True, size=16, color="1B4F72")
    ws.merge_cells("A1:B1")

    linhas_capa = [
        ("Fonte", f"BIS Data Portal — {BIS_BULK}"),
        ("Taxa básica nominal", "WS_CBPOL (Central bank policy rates), mensal, fim de período, % a.a."),
        ("Índice de inflação", "WS_LONG_CPI (Consumer prices), índice 2010=100, mensal"),
        ("Fórmula mensal", "r_m = (1 + i_aa^(1/12)) / (1 + IPC_t/IPC_{t−1} − 1) − 1"),
        ("Acumulado no ano", "Após dezembro: R = Π(1+r_m) − 1 dos meses do ano-calendário"),
        ("Ano incompleto", "Linha 'ACUMULADO PARCIAL' após o último mês disponível"),
        ("Países", str(len(resumos))),
        ("Gerado em", gerado.strftime("%d/%m/%Y %H:%M")),
        ("Marker", MARKER),
        (
            "Observação",
            "Países da Zona do Euro mantêm a série nacional histórica; "
            "a taxa vigente do BCE está na aba Zona do Euro.",
        ),
    ]
    ws["A3"] = "Campo"
    ws["B3"] = "Valor"
    ws["A3"].font = sty["header"]
    ws["B3"].font = sty["header"]
    ws["A3"].fill = sty["header_fill"]
    ws["B3"].fill = sty["header_fill"]
    for i, (k, v) in enumerate(linhas_capa, start=4):
        ws.cell(i, 1, k).font = Font(name="Calibri", bold=True, size=10)
        ws.cell(i, 2, v).font = Font(name="Calibri", size=10)
        ws.cell(i, 2).alignment = Alignment(wrap_text=True)
    _aplicar_larguras(ws, {1: 28, 2: 110})
    ws.row_dimensions[1].height = 22
    for i in range(4, 4 + len(linhas_capa)):
        ws.row_dimensions[i].height = 18
    ws.row_dimensions[13].height = 36


def escrever_resumo(wb: Workbook, resumos: list[dict], sty: dict) -> None:
    ws = wb.create_sheet("Resumo")
    ws["A1"] = "Resumo — último mês disponível e real acumulada do último ano completo"
    ws["A1"].font = sty["title"]
    ws.merge_cells("A1:K1")
    headers = [
        "País",
        "Código",
        "Primeiro mês",
        "Último mês",
        "Meses",
        "Taxa básica nominal (% a.a.)",
        "Índice IPC (2010=100)",
        "Inflação no mês (% a.m.)",
        "Taxa real no mês (% a.m.)",
        "Último ano completo",
        "Real acumulada no ano (%)",
    ]
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(3, j, h)
        cell.font = sty["header"]
        cell.fill = sty["header_fill"]
        cell.alignment = sty["center"]
        cell.border = sty["thin"]
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:K{3 + max(len(resumos), 1)}"

    ordenados = sorted(resumos, key=lambda x: (x["pais"] != "Brasil", x["pais"].casefold()))
    for i, info in enumerate(ordenados, start=4):
        vals = [
            info["pais"],
            info["codigo"],
            info["primeiro_mes"],
            info["ultimo_mes"],
            info["n_meses"],
            info["nominal_ultimo"],
            info["indice_ultimo"],
            info["inflacao_ultimo"],
            info["real_mes_ultimo"],
            info["ultimo_ano_completo"],
            info["real_acumulada_ultimo_ano"],
        ]
        fill = sty["alt"] if i % 2 == 0 else None
        for j, val in enumerate(vals, start=1):
            cell = ws.cell(i, j, val if val is not None else None)
            cell.border = sty["thin"]
            cell.font = sty["mes"]
            if fill is not None:
                cell.fill = fill
            if j in (6, 11) and val is not None:
                cell.number_format = sty["pct"]
            elif j in (8, 9) and val is not None:
                cell.number_format = sty["pct4"]
            elif j == 7 and val is not None:
                cell.number_format = sty["num"]
    _aplicar_larguras(
        ws,
        {1: 22, 2: 10, 3: 14, 4: 14, 5: 10, 6: 28, 7: 22, 8: 22, 9: 22, 10: 20, 11: 26},
    )


def escrever_anual(wb: Workbook, pivot: pd.DataFrame, sty: dict) -> None:
    ws = wb.create_sheet("Anual")
    ws["A1"] = "Taxa básica de juros real acumulada no ano — comparação entre países"
    ws["A1"].font = sty["title"]
    if pivot.empty:
        ws["A3"] = "Sem anos completos para comparar."
        return
    n_cols = len(pivot.columns) + 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws["A2"] = f"Período: {int(pivot.index.min())}–{int(pivot.index.max())}. Valores em % a.a. (Fisher composto)."
    ws["A2"].font = sty["subtitle"]
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    ws["A3"] = "Ano"
    ws["A3"].font = sty["header"]
    ws["A3"].fill = sty["header_fill"]
    ws["A3"].alignment = sty["center"]
    ws["A3"].border = sty["thin"]
    alinha_num = Alignment(horizontal="center", vertical="center", wrap_text=False)
    alinha_cab = Alignment(horizontal="center", vertical="center", textRotation=90, wrap_text=True)
    for j, col in enumerate(pivot.columns, start=2):
        cell = ws.cell(3, j, col)
        cell.font = sty["header"]
        cell.fill = sty["header_fill"]
        cell.alignment = alinha_cab
        cell.border = sty["thin"]
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.row_dimensions[3].height = 110
    ws.freeze_panes = "B4"
    for i, (ano, rec) in enumerate(pivot.iterrows(), start=4):
        ws.cell(i, 1, int(ano)).font = Font(name="Calibri", bold=True, size=10)
        ws.cell(i, 1).border = sty["thin"]
        ws.cell(i, 1).alignment = sty["center"]
        fill = sty["alt"] if i % 2 == 0 else None
        if fill:
            ws.cell(i, 1).fill = fill
        ws.row_dimensions[i].height = 18
        for j, val in enumerate(rec.tolist(), start=2):
            cell = ws.cell(i, j)
            cell.border = sty["thin"]
            cell.alignment = alinha_num
            if fill:
                cell.fill = fill
            if pd.notna(val):
                cell.value = float(val)
                cell.number_format = sty["pct"]
    ws.column_dimensions["A"].width = 10
    ws.auto_filter.ref = f"A3:{get_column_letter(len(pivot.columns) + 1)}{3 + len(pivot)}"


def escrever_planilha(
    por_serie: dict[str, pd.DataFrame],
    por_linhas: dict[str, pd.DataFrame],
    saida: Path,
) -> Path:
    saida.parent.mkdir(parents=True, exist_ok=True)
    sty = _estilo_base()
    wb = Workbook()
    resumos = []
    # Brasil primeiro; demais em ordem alfabética PT
    codigos = sorted(por_linhas.keys(), key=lambda c: (c != "BR", nome_pais(c).casefold()))
    for codigo in codigos:
        resumos.append(resumo_pais(por_serie[codigo], por_linhas[codigo]))

    escrever_capa(wb, resumos, sty, datetime.now())
    escrever_resumo(wb, resumos, sty)
    escrever_anual(wb, pivot_anual(por_linhas), sty)
    for codigo in codigos:
        print(f"[ABA] {nome_pais(codigo)} ({codigo}): {len(por_linhas[codigo]):,} linhas")
        escrever_aba_pais(wb, codigo, por_serie[codigo], por_linhas[codigo], sty)

    wb.save(saida)
    print(f"[OK] Planilha: {saida} ({saida.stat().st_size / 1e6:.2f} MB)")
    return saida


def estatisticas_anuais(serie: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por ano-calendário: real, inflação e nominal compostos no ano."""
    if serie.empty:
        return pd.DataFrame(
            columns=[
                "ano",
                "codigo",
                "pais",
                "real_aa",
                "inflacao_aa",
                "nominal_composta",
                "nominal_fim",
                "n_meses",
                "completo",
            ]
        )
    codigo = str(serie["codigo"].iloc[0])
    pais_en = str(serie["pais_en"].iloc[0]) if "pais_en" in serie.columns else codigo
    pais = nome_pais(codigo, pais_en)
    rows: list[dict] = []
    for ano, bloco in serie.groupby("ano", sort=True):
        bloco = bloco.sort_values("mes")
        n = int(bloco["real_am"].notna().sum())
        if n == 0:
            continue
        tem_dezembro = bool((bloco["mes"].dt.month == 12).any())
        ultimo = bloco.dropna(subset=["nominal_aa"]).iloc[-1] if bloco["nominal_aa"].notna().any() else None
        rows.append(
            {
                "ano": int(ano),
                "codigo": codigo,
                "pais": pais,
                "real_aa": acumular_fator(bloco["real_am"]),
                "inflacao_aa": acumular_fator(bloco["inflacao_am"]),
                "nominal_composta": acumular_fator(bloco["nominal_am"]),
                "nominal_fim": float(ultimo["nominal_aa"]) if ultimo is not None else float("nan"),
                "n_meses": n,
                "completo": tem_dezembro and n >= 12,
            }
        )
    return pd.DataFrame(rows)


def carregar_series_paises(
    pasta_cache: Path,
    *,
    baixar: bool = True,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
    paises: set[str] | None = None,
    cbpol_zip: Path | None = None,
    cpi_zip: Path | None = None,
) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame]]:
    """Baixa/lê BIS e devolve ``(por_serie, por_linhas)`` por código de país."""
    print(f"[{MARKER}]")
    if cbpol_zip is None:
        cbpol_zip = pasta_cache / ZIP_CBPOL
        if baixar:
            baixar_arquivo(URL_CBPOL, cbpol_zip)
    if cpi_zip is None:
        cpi_zip = pasta_cache / ZIP_CPI
        if baixar:
            baixar_arquivo(URL_CPI, cpi_zip)
    if not Path(cbpol_zip).exists():
        raise FileNotFoundError(cbpol_zip)
    if not Path(cpi_zip).exists():
        raise FileNotFoundError(cpi_zip)

    nominal = carregar_cbpol_mensal(Path(cbpol_zip))
    cpi = carregar_cpi_mensal(Path(cpi_zip))
    if paises:
        paises = {p.upper() for p in paises}
        nominal = nominal[nominal["codigo"].isin(paises)]
        cpi = cpi[cpi["codigo"].isin(paises)]

    comuns = sorted(set(nominal["codigo"]) & set(cpi["codigo"]))
    if not comuns:
        raise ValueError("Nenhum país com taxa básica e IPC sobrepostos.")

    por_serie: dict[str, pd.DataFrame] = {}
    por_linhas: dict[str, pd.DataFrame] = {}
    for codigo in comuns:
        serie = montar_serie_pais(
            nominal, cpi, codigo=codigo, ano_inicio=ano_inicio, ano_fim=ano_fim
        )
        if serie.empty:
            print(f"[AVISO] {codigo}: sem interseção nominal × IPC no recorte.")
            continue
        linhas = linhas_discriminativo(serie)
        por_serie[codigo] = serie
        por_linhas[codigo] = linhas
        print(
            f"  {nome_pais(codigo):22s} {codigo}  "
            f"{formatar_mes_ano(serie['mes'].min())}–{formatar_mes_ano(serie['mes'].max())}  "
            f"{len(serie):4d} meses  {len(linhas):4d} linhas"
        )

    if not por_linhas:
        raise ValueError("Nenhuma série gerada no recorte pedido.")
    return por_serie, por_linhas


def processar(
    pasta_cache: Path,
    saida: Path,
    *,
    baixar: bool = True,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
    paises: set[str] | None = None,
    cbpol_zip: Path | None = None,
    cpi_zip: Path | None = None,
) -> Path:
    por_serie, por_linhas = carregar_series_paises(
        pasta_cache,
        baixar=baixar,
        ano_inicio=ano_inicio,
        ano_fim=ano_fim,
        paises=paises,
        cbpol_zip=cbpol_zip,
        cpi_zip=cpi_zip,
    )
    return escrever_planilha(por_serie, por_linhas, saida)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Discriminativo de taxas básicas de juros reais por país (BIS)."
    )
    p.add_argument(
        "--pasta-cache",
        type=Path,
        default=ROOT / "data" / "raw" / "bis",
        help="Pasta dos ZIPs BIS (default: data/raw/bis)",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT / "output" / "discriminativo_juros_reais_paises.xlsx",
        help="Excel de saída",
    )
    p.add_argument(
        "--ano-inicio",
        type=int,
        default=ANO_INICIO_DEFAULT,
        help=f"Primeiro ano (inclusive, default {ANO_INICIO_DEFAULT})",
    )
    p.add_argument("--ano-fim", type=int, default=None, help="Último ano (inclusive)")
    p.add_argument(
        "--paises",
        type=str,
        default="",
        help="Códigos ISO-2 separados por vírgula (ex.: BR,US,XM). Vazio = todos.",
    )
    p.add_argument("--cbpol", type=Path, default=None, help="ZIP CBPOL local")
    p.add_argument("--cpi", type=Path, default=None, help="ZIP CPI local")
    p.add_argument("--sem-download", action="store_true", help="Não baixa; usa cache local")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    paises = {x.strip().upper() for x in args.paises.split(",") if x.strip()} or None
    processar(
        args.pasta_cache,
        args.saida,
        baixar=not args.sem_download,
        ano_inicio=args.ano_inicio,
        ano_fim=args.ano_fim,
        paises=paises,
        cbpol_zip=args.cbpol,
        cpi_zip=args.cpi,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
