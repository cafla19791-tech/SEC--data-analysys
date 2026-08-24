"""Planilhas BIS a partir das URLs oficiais: uma aba por país/economia.

Aceita csv_col (largo) e csv_flat (longo). O csv_flat de contas nacionais
é reorganizado em formato largo (uma linha por série, períodos nas colunas).

Uso:
  python3 scripts/bis_urls_paises.py
  python3 scripts/bis_urls_paises.py --sem-download
  python3 scripts/bis_urls_paises.py --apenas WS_DSR,WS_SPP
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
import zipfile
from pathlib import Path

import pandas as pd
import requests
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

DATA_DIR = ROOT / "data" / "bis"
OUTPUT_DIR = ROOT / "output" / "bis_paises"

FONTES = [
    {
        "stem": "WS_NA_SEC_DSS",
        "arquivo": "WS_NA_SEC_DSS_csv_flat.zip",
        "url": "https://data.bis.org/static/bulk/WS_NA_SEC_DSS_csv_flat.zip",
        "formato": "flat",
        "titulo": "Títulos de dívida nas contas nacionais (DSS)",
        "pais": ("REF_AREA:Reference area",),
        "chave": "país/economia de referência",
    },
    {
        "stem": "WS_DEBT_SEC2_PUB",
        "arquivo": "WS_DEBT_SEC2_PUB_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_DEBT_SEC2_PUB_csv_col.zip",
        "formato": "col",
        "titulo": "Estatísticas públicas de títulos de dívida",
        "pais": ("ISSUER_RES", "Issuer residence"),
        "chave": "residência do emissor",
        "partir_agregados": True,
    },
    {
        "stem": "WS_TC",
        "arquivo": "WS_TC_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_TC_csv_col.zip",
        "formato": "col",
        "titulo": "Crédito ao setor não financeiro",
        "pais": ("BORROWERS_CTY", "Borrowers' country"),
        "chave": "país do tomador",
    },
    {
        "stem": "WS_DSR",
        "arquivo": "WS_DSR_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_DSR_csv_col.zip",
        "formato": "col",
        "titulo": "Índice de serviço da dívida",
        "pais": ("BORROWERS_CTY", "Borrowers' country"),
        "chave": "país do tomador",
    },
    {
        "stem": "WS_GLI",
        "arquivo": "WS_GLI_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_GLI_csv_col.zip",
        "formato": "col",
        "titulo": "Indicadores de liquidez global",
        "pais": ("BORROWERS_CTY", "Borrowers' country"),
        "chave": "país do tomador",
    },
    {
        "stem": "WS_XTD_DERIV",
        "arquivo": "WS_XTD_DERIV_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_XTD_DERIV_csv_col.zip",
        "formato": "col",
        "titulo": "Derivativos de bolsa (exchange-traded)",
        "pais": ("XD_EXCHANGE", "Location of trade (Exchange or country)"),
        "chave": "praça/região da bolsa (o BIS não publica este conjunto por país)",
    },
    {
        "stem": "WS_OTC_DERIV2",
        "arquivo": "WS_OTC_DERIV2_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_OTC_DERIV2_csv_col.zip",
        "formato": "col",
        "titulo": "Derivativos de balcão (OTC)",
        "pais": ("DER_CPC", "Counterparty country"),
        "chave": "país da contraparte (o país declarante é sempre 'All countries')",
    },
    {
        "stem": "WS_SPP",
        "arquivo": "WS_SPP_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_SPP_csv_col.zip",
        "formato": "col",
        "titulo": "Preços de imóveis residenciais",
        "pais": ("REF_AREA", "Reference area"),
        "chave": "país/economia de referência",
    },
    {
        "stem": "WS_CPP",
        "arquivo": "WS_CPP_csv_col.zip",
        "url": "https://data.bis.org/static/bulk/WS_CPP_csv_col.zip",
        "formato": "col",
        "titulo": "Preços de imóveis comerciais",
        "pais": ("REF_AREA", "Reference area"),
        "chave": "país/economia de referência",
    },
]

DESCARTAR_FLAT = {
    "STRUCTURE",
    "STRUCTURE_ID",
    "ACTION",
    "COMMENT_DSET:Dataset comment",
    "REF_PERIOD_DETAIL:Reference period detail [deprecated]",
    "REPYEARSTART:Reference year start",
    "REPYEAREND:Reference year end",
    "TIME_FORMAT:Time format",
    "TIME_PER_COLLECT:Time period collection",
    "REF_YEAR_PRICE:Reference year (price)",
    "TABLE_IDENTIFIER:Table identifier",
    "LAST_UPDATE:Last Update Date",
    "COLL_PERIOD:Collection period",
    "COMMENT_TS:Series comment",
    "GFS_ECOFUNC:GFS economic function",
    "GFS_TAXCAT:GFS tax category",
    "DATA_COMP:Underlying compilation",
    "CURRENCY:Currency code used for compilation",
    "DISS_ORG:Dissemination organisation",
    "OBS_PRE_BREAK:Observation pre-break value",
    "CONF_STATUS:Confidentiality status",
    "COMMENT_OBS:Comments to the observation value",
    "EMBARGO_DATE:Embargo date",
    "OBS_EDP_WBB:EDP working balance basis",
    "ADJUSTMENT:Adjustment indicator",
}

PAISES_PT = {
    "1E": "Residentes/local",
    "2A": "Economias avançadas (BIS)",
    "3P": "Todos os países exceto residentes",
    "4T": "Todas as economias declarantes",
    "4U": "AL e Caribe emergentes",
    "4W": "Europa emergente",
    "4Y": "Ásia e Pacífico emergentes",
    "5A": "Todos os países",
    "5C": "Área do euro",
    "5F": "Outros países asiáticos",
    "5J": "Todos os países",
    "5K": "Europa avançada",
    "5P": "Demais países",
    "5R": "Economias avançadas",
    "5U": "América Latina",
    "5Z": "Não residentes/cross-border",
    "8A": "Todas as bolsas",
    "8B": "Bolsas da América do Norte",
    "8C": "Bolsas europeias",
    "8E": "Bolsas da Ásia/Pacífico",
    "8F": "Bolsas asiáticas",
    "8G": "Bolsas da Austrália/Nova Zelândia",
    "8K": "Outras bolsas",
    "AR": "Argentina",
    "AT": "Áustria",
    "AU": "Austrália",
    "BE": "Bélgica",
    "BG": "Bulgária",
    "BR": "Brasil",
    "CA": "Canadá",
    "CH": "Suíça",
    "CL": "Chile",
    "CN": "China",
    "CO": "Colômbia",
    "CY": "Chipre",
    "CZ": "Tchéquia",
    "DE": "Alemanha",
    "DK": "Dinamarca",
    "EE": "Estônia",
    "ES": "Espanha",
    "FI": "Finlândia",
    "FR": "França",
    "GB": "Reino Unido",
    "GR": "Grécia",
    "HK": "Hong Kong",
    "HR": "Croácia",
    "HU": "Hungria",
    "ID": "Indonésia",
    "IE": "Irlanda",
    "IL": "Israel",
    "IN": "Índia",
    "IS": "Islândia",
    "IT": "Itália",
    "JP": "Japão",
    "KR": "Coreia do Sul",
    "KW": "Kuwait",
    "LU": "Luxemburgo",
    "LT": "Lituânia",
    "LV": "Letônia",
    "MA": "Marrocos",
    "MK": "Macedônia do Norte",
    "MT": "Malta",
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
    "SG": "Singapura",
    "SI": "Eslovênia",
    "SK": "Eslováquia",
    "TH": "Tailândia",
    "TR": "Turquia",
    "U2": "Área do euro (composição variável)",
    "US": "Estados Unidos",
    "XM": "Zona do euro",
    "ZA": "África do Sul",
}

PERIODO_RE = re.compile(r"^\d{4}(?:-Q[1-4]|-S[12]|-\d{2})?$")


def fundir_cabecalho(header: list[str]) -> list[str]:
    """Recompõe nomes SDMX quebrados por vírgula sem aspas (continuação com espaço)."""
    out: list[str] = []
    for h in header:
        if out and h.startswith(" "):
            out[-1] = f"{out[-1]},{h}"
        else:
            out.append(h)
    return out


def e_periodo(nome: str) -> bool:
    return bool(PERIODO_RE.match(str(nome).strip()))


def partir_codigo_nome(valor: object) -> tuple[str, str]:
    txt = "" if valor is None or (isinstance(valor, float) and pd.isna(valor)) else str(valor).strip()
    if not txt:
        return "", ""
    if ":" in txt:
        cod, nome = txt.split(":", 1)
        return cod.strip(), nome.strip()
    return txt, txt


def nome_pais(codigo: str, nome_en: str) -> str:
    return PAISES_PT.get(codigo, nome_en or codigo)


def nome_aba(codigo: str, nome: str, usados: set[str]) -> str:
    base = f"{codigo} {nome}".strip() or codigo or "pais"
    for ch in r"\/*?:[]":
        base = base.replace(ch, "-")
    base = " ".join(base.split())[:31] or codigo[:31] or "pais"
    cand = base
    n = 2
    while cand in usados:
        suf = f" {n}"
        cand = base[: 31 - len(suf)] + suf
        n += 1
    usados.add(cand)
    return cand


def _rotulo_coluna(nome: str) -> str:
    return nome.split(":", 1)[-1] if ":" in nome and not nome.startswith("_") else nome


def coluna_pais(colunas: list[str], candidatos: tuple[str, ...]) -> str:
    for cand in candidatos:
        if cand in colunas:
            return cand
    for col in colunas:
        up = col.upper()
        if any(up == c or up.startswith(c + ":") for c in ("REF_AREA", "BORROWERS_CTY", "ISSUER_RES", "DER_CPC")):
            return col
    raise ValueError(f"Sem coluna de país em {colunas[:20]}")


def baixar_zip(fonte: dict, cache_dir: Path, baixar: bool) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    dest = cache_dir / fonte["arquivo"]
    if dest.exists() and dest.stat().st_size > 0:
        return dest
    if not baixar:
        raise FileNotFoundError(dest)
    print(f"  baixando {fonte['url']}", flush=True)
    resp = requests.get(fonte["url"], timeout=300)
    resp.raise_for_status()
    dest.write_bytes(resp.content)
    return dest


def abrir_csv_zip(zip_path: Path):
    zf = zipfile.ZipFile(zip_path)
    nome = next(n for n in zf.namelist() if n.lower().endswith(".csv"))
    raw = zf.open(nome)
    return zf, raw


def ler_csv_col(zip_path: Path) -> pd.DataFrame:
    zf, raw = abrir_csv_zip(zip_path)
    try:
        return pd.read_csv(raw, dtype=str, low_memory=False)
    finally:
        raw.close()
        zf.close()


def ler_csv_flat(zip_path: Path, usecols: list[str] | None = None) -> pd.DataFrame:
    zf, raw = abrir_csv_zip(zip_path)
    try:
        fh = io.TextIOWrapper(raw, encoding="utf-8", newline="")
        reader = csv.reader(fh)
        header = fundir_cabecalho(next(reader))
        if usecols is None:
            usecols = [c for c in header if c not in DESCARTAR_FLAT]
        idx = [header.index(c) for c in usecols]
        linhas = []
        for rec in reader:
            linhas.append([rec[i] if i < len(rec) else "" for i in idx])
        return pd.DataFrame(linhas, columns=usecols)
    finally:
        raw.close()
        zf.close()


def _usecols_flat(header: list[str], col_pais: str) -> list[str]:
    usecols = [c for c in header if c not in DESCARTAR_FLAT]
    if col_pais not in usecols:
        usecols = [col_pais] + usecols
    return usecols


def iterar_flat_por_pais(zip_path: Path, col_pais: str, usecols: list[str] | None = None):
    """Gera (codigo, nome_en, DataFrame). Para arquivos grandes, grave em disco via partir_flat_em_arquivos."""
    zf, raw = abrir_csv_zip(zip_path)
    try:
        fh = io.TextIOWrapper(raw, encoding="utf-8", newline="")
        reader = csv.reader(fh)
        header = fundir_cabecalho(next(reader))
        if usecols is None:
            usecols = _usecols_flat(header, col_pais)
        idx = [header.index(c) for c in usecols]
        i_pais = usecols.index(col_pais)
        grupos: dict[str, list[list[str]]] = {}
        nomes: dict[str, str] = {}
        for rec in reader:
            vals = [rec[i] if i < len(rec) else "" for i in idx]
            codigo, nome_en = partir_codigo_nome(vals[i_pais])
            grupos.setdefault(codigo, []).append(vals)
            nomes.setdefault(codigo, nome_en)
        for codigo, rows in grupos.items():
            yield codigo, nomes[codigo], pd.DataFrame(rows, columns=usecols)
    finally:
        raw.close()
        zf.close()


def partir_flat_em_arquivos(zip_path: Path, col_pais: str, dest_dir: Path) -> list[tuple[str, str, Path]]:
    dest_dir.mkdir(parents=True, exist_ok=True)
    zf, raw = abrir_csv_zip(zip_path)
    writers: dict[str, csv.writer] = {}
    handles: dict[str, object] = {}
    nomes: dict[str, str] = {}
    try:
        fh = io.TextIOWrapper(raw, encoding="utf-8", newline="")
        reader = csv.reader(fh)
        header = fundir_cabecalho(next(reader))
        usecols = _usecols_flat(header, col_pais)
        idx = [header.index(c) for c in usecols]
        i_pais = usecols.index(col_pais)
        for rec in reader:
            vals = [rec[i] if i < len(rec) else "" for i in idx]
            codigo, nome_en = partir_codigo_nome(vals[i_pais])
            if not codigo:
                continue
            if codigo not in writers:
                path = dest_dir / f"{codigo}.csv"
                handle = path.open("w", encoding="utf-8", newline="")
                writer = csv.writer(handle)
                writer.writerow(usecols)
                writers[codigo] = writer
                handles[codigo] = handle
                nomes[codigo] = nome_en
            writers[codigo].writerow(vals)
    finally:
        for handle in handles.values():
            handle.close()
        raw.close()
        zf.close()
    return [(codigo, nomes[codigo], dest_dir / f"{codigo}.csv") for codigo in writers]


def pivotar_flat(df: pd.DataFrame) -> pd.DataFrame:
    tempo = next((c for c in df.columns if c.startswith("TIME_PERIOD")), None)
    valor = next((c for c in df.columns if c.startswith("OBS_VALUE")), None)
    if tempo is None or valor is None:
        return df
    id_cols = [c for c in df.columns if c not in {tempo, valor, "_codigo", "_nome_en", "_pais"}]
    trabalho = df.copy()
    trabalho[valor] = pd.to_numeric(trabalho[valor], errors="coerce")
    for c in id_cols:
        trabalho[c] = trabalho[c].fillna("").astype(str)
    wide = trabalho.pivot_table(index=id_cols, columns=tempo, values=valor, aggfunc="last")
    wide = wide.reset_index()
    wide.columns = [str(c) for c in wide.columns]
    periodos = sorted((c for c in wide.columns if e_periodo(c)), key=_chave_periodo)
    outros = [c for c in wide.columns if c not in periodos]
    return wide[outros + periodos]


def _chave_periodo(txt: str) -> tuple:
    t = str(txt)
    ano = int(t[:4])
    if "-Q" in t:
        return (ano, 1, int(t[-1]))
    if "-S" in t:
        return (ano, 2, int(t[-1]))
    if len(t) == 7 and t[4] == "-":
        return (ano, 3, int(t[5:7]))
    return (ano, 0, 0)


def _rotulo_vizinho(colunas: list[str], col: str) -> str | None:
    i = colunas.index(col)
    if i + 1 >= len(colunas):
        return None
    viz = colunas[i + 1]
    if e_periodo(viz) or viz.upper() == col.upper() or viz.isupper():
        return None
    return viz


def preparar_col(df: pd.DataFrame, candidatos: tuple[str, ...]) -> tuple[pd.DataFrame, str]:
    col = coluna_pais(list(df.columns), candidatos)
    out = df.copy()
    label = _rotulo_vizinho(list(out.columns), col)
    bruto = out[col].fillna("").astype(str).str.strip()
    if label:
        out["_codigo"] = bruto
        out["_nome_en"] = out[label].fillna("").astype(str).str.strip()
    else:
        partes = bruto.map(partir_codigo_nome)
        out["_codigo"] = [p[0] for p in partes]
        out["_nome_en"] = [p[1] for p in partes]
    out["_pais"] = [nome_pais(c, n) for c, n in zip(out["_codigo"], out["_nome_en"])]
    for c in out.columns:
        if e_periodo(c):
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out, col


def anexar_pais_flat(df: pd.DataFrame, col: str) -> pd.DataFrame:
    partes = df[col].map(partir_codigo_nome)
    out = df.copy()
    out["_codigo"] = [p[0] for p in partes]
    out["_nome_en"] = [p[1] for p in partes]
    out["_pais"] = [nome_pais(c, n) for c, n in zip(out["_codigo"], out["_nome_en"])]
    return out


def _escrever_cabecalho(ws, cabecalhos: list[str], formatos: dict) -> None:
    for col, cab in enumerate(cabecalhos):
        ws.write(0, col, _rotulo_coluna(str(cab)), formatos["cab"])
        ws.set_column(col, col, 18 if e_periodo(str(cab)) else 28)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, 0, max(len(cabecalhos) - 1, 0))
    ws.set_row(0, 22)


def _escrever_df(ws, dados: pd.DataFrame, formatos: dict) -> None:
    cols = [str(c) for c in dados.columns]
    _escrever_cabecalho(ws, cols, formatos)
    n = len(dados)
    bordar = n <= 2500
    valores = dados.itertuples(index=False, name=None)
    for i, rec in enumerate(valores, start=1):
        for j, valor in enumerate(rec):
            if valor is None or (isinstance(valor, float) and pd.isna(valor)) or valor == "":
                continue
            if isinstance(valor, float):
                fmt = formatos["num_borda"] if bordar else formatos["num"]
                ws.write_number(i, j, float(valor), fmt)
            else:
                fmt = formatos["cel_borda"] if bordar else None
                if fmt:
                    ws.write(i, j, valor, fmt)
                else:
                    ws.write(i, j, valor)


def gerar_planilha(
    grupos: list[tuple[str, str, pd.DataFrame]],
    titulo: str,
    fonte_url: str,
    chave: str,
    path: Path,
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    grupos = sorted(grupos, key=lambda x: (x[0] != "BR", x[1].casefold(), x[0]))
    usados = {"Notas", "Indice"}
    indice = []
    abas: list[tuple[str, pd.DataFrame]] = []
    drop_interno = {"_codigo", "_nome_en", "_pais"}

    for codigo, pais, g in grupos:
        aba = nome_aba(codigo, pais, usados)
        dados = g.drop(columns=[c for c in drop_interno if c in g.columns], errors="ignore").copy()
        vazias = [c for c in dados.columns if dados[c].isna().all() or (dados[c].astype(str).str.strip() == "").all()]
        dados = dados.drop(columns=vazias, errors="ignore")
        dados.columns = [_rotulo_coluna(str(c)) if ":" in str(c) else str(c) for c in dados.columns]
        periodos = [c for c in dados.columns if e_periodo(c)]
        abas.append((aba, dados))
        indice.append(
            {
                "Código": codigo,
                "País": pais,
                "Aba": aba,
                "Séries/linhas": int(len(dados)),
                "Início": periodos[0] if periodos else "",
                "Fim": periodos[-1] if periodos else "",
            }
        )

    wb = xlsxwriter.Workbook(str(path), {"constant_memory": False, "strings_to_urls": False})
    fmt_titulo = wb.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#1F4E79"})
    fmt_nota = wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True, "valign": "top"})
    formatos = {
        "cab": wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "bg_color": "#E8E8E8",
                "border": 1,
                "align": "center",
                "text_wrap": True,
            }
        ),
        "num": wb.add_format({"font_name": "Calibri", "font_size": 9, "num_format": "#,##0.0000"}),
        "num_borda": wb.add_format(
            {"font_name": "Calibri", "font_size": 9, "num_format": "#,##0.0000", "border": 1}
        ),
        "cel_borda": wb.add_format({"font_name": "Calibri", "font_size": 9, "border": 1}),
        "cel": wb.add_format({"font_name": "Calibri", "font_size": 9, "border": 1}),
    }

    ws0 = wb.add_worksheet("Notas")
    notas = [
        titulo,
        "Uma aba por país/economia (ou praça, quando o conjunto não tem dimensão de país).",
        f"Chave das abas: {chave}.",
        f"Fonte: {fonte_url}",
        "Catálogo: https://data.bis.org/bulkdownload",
        "csv_col: períodos já vêm em colunas. csv_flat: o arquivo longo foi pivotado "
        "(uma linha por série, períodos nas colunas) para caber no Excel.",
        "Agregados do BIS (área do euro, economias avançadas, todas as bolsas etc.) entram como aba.",
    ]
    ws0.write(0, 0, notas[0], fmt_titulo)
    ws0.set_column(0, 0, 120)
    for i, txt in enumerate(notas[1:], start=1):
        ws0.write(i, 0, txt, fmt_nota)
        ws0.set_row(i, 28)

    ws_i = wb.add_worksheet("Indice")
    _escrever_df(ws_i, pd.DataFrame(indice), formatos)

    for aba, dados in abas:
        ws = wb.add_worksheet(aba)
        _escrever_df(ws, dados, formatos)
    wb.close()
    return path


def _eh_agregado(codigo: str) -> bool:
    return bool(codigo) and codigo[0].isdigit()


def processar(fonte: dict, cache_dir: Path, output_dir: Path, baixar: bool) -> list[Path]:
    print(f"{fonte['stem']} — {fonte['titulo']}", flush=True)
    zip_path = baixar_zip(fonte, cache_dir, baixar=baixar)
    grupos: list[tuple[str, str, pd.DataFrame]] = []

    if fonte["formato"] == "flat":
        col_pais = fonte["pais"][0]
        split_dir = cache_dir / f"_split_{fonte['stem']}"
        partes = partir_flat_em_arquivos(zip_path, col_pais, split_dir)
        for codigo, nome_en, csv_pais in partes:
            g = pd.read_csv(csv_pais, dtype=str, low_memory=False)
            pais = nome_pais(codigo, nome_en)
            wide = pivotar_flat(g)
            grupos.append((codigo, pais, wide))
            print(f"  {codigo} {pais}: {len(g)} obs → {len(wide)} séries", flush=True)
        print(f"  {len(partes)} economias lidas do csv_flat", flush=True)
    else:
        bruto = ler_csv_col(zip_path)
        prep, _col = preparar_col(bruto, fonte["pais"])
        for (codigo, pais), g in prep.groupby(["_codigo", "_pais"], sort=False):
            grupos.append((str(codigo), str(pais), g))

    if fonte.get("partir_agregados"):
        paises = [g for g in grupos if not _eh_agregado(g[0])]
        agreg = [g for g in grupos if _eh_agregado(g[0])]
        dests = []
        if paises:
            dest = output_dir / f"bis_{fonte['stem']}_paises.xlsx"
            gerar_planilha(paises, f"{fonte['titulo']} — BIS {fonte['stem']} (países)", fonte["url"], fonte["chave"], dest)
            print(f"  {len(paises)} países → {dest}", flush=True)
            dests.append(dest)
        if agreg:
            dest = output_dir / f"bis_{fonte['stem']}_agregados.xlsx"
            gerar_planilha(
                agreg,
                f"{fonte['titulo']} — BIS {fonte['stem']} (agregados)",
                fonte["url"],
                fonte["chave"] + " — códigos agregados do BIS (3P, 4T, 5R…)",
                dest,
            )
            print(f"  {len(agreg)} agregados → {dest}", flush=True)
            dests.append(dest)
        return dests

    dest = output_dir / f"bis_{fonte['stem']}_paises.xlsx"
    gerar_planilha(grupos, f"{fonte['titulo']} — BIS {fonte['stem']}", fonte["url"], fonte["chave"], dest)
    print(f"  {len(grupos)} abas → {dest}", flush=True)
    return [dest]


def gerar_indice(gerados: list[tuple[dict, Path]], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Indicadores")
    titulo = wb.add_format({"font_name": "Calibri", "font_size": 13, "bold": True, "font_color": "#1F4E79"})
    cab = wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "bg_color": "#E8E8E8", "border": 1})
    cel = wb.add_format({"font_name": "Calibri", "font_size": 9, "border": 1, "text_wrap": True})
    ws.write(0, 0, "Indicadores BIS — uma planilha por URL, uma aba por país/economia", titulo)
    ws.merge_range(0, 0, 0, 4, "Indicadores BIS — uma planilha por URL, uma aba por país/economia", titulo)
    headers = ["Código", "Indicador", "Arquivo", "Chave das abas", "Fonte"]
    for i, h in enumerate(headers):
        ws.write(2, i, h, cab)
    for r, (fonte, dest) in enumerate(gerados, start=3):
        vals = [fonte["stem"], fonte["titulo"], dest.name, fonte["chave"], fonte["url"]]
        for c, v in enumerate(vals):
            ws.write(r, c, v, cel)
    for i, w in enumerate([22, 48, 36, 52, 78]):
        ws.set_column(i, i, w)
    wb.close()
    return path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--sem-download", action="store_true")
    parser.add_argument("--apenas", default="", help="Lista de códigos WS_* separados por vírgula")
    args = parser.parse_args(argv)

    escolhidos = {x.strip().upper() for x in args.apenas.split(",") if x.strip()}
    alvos = [f for f in FONTES if not escolhidos or f["stem"] in escolhidos]
    if not alvos:
        raise SystemExit("Nenhum indicador selecionado.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    gerados: list[tuple[dict, Path]] = []
    for fonte in alvos:
        for dest in processar(fonte, args.cache_dir, args.output_dir, not args.sem_download):
            gerados.append((fonte, dest))

    cat = args.output_dir / "bis_indice_urls.xlsx"
    gerar_indice(gerados, cat)
    print(f"Índice: {cat}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
