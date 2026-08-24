#!/usr/bin/env python3
"""Gera um Excel por tema do BIS Data Portal, com aba por país.

Fonte: https://data.bis.org/bulkdownload

Cada arquivo cobre um tópico do portal. Abas:
  Capa          — metadados, recorte e citação
  Indice        — países/agregados presentes no tema
  Comparativo   — último valor disponível por país
  <CODIGO>      — série(s) daquele país/jurisdição (ex.: BR, US, 5A)

Uso:
  python3 scripts/bis_bulk_excel.py
  python3 scripts/bis_bulk_excel.py --topics CBPOL,CREDIT_GAP --saida output/bis
"""

from __future__ import annotations

import argparse
import io
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Iterable
from urllib.request import Request, urlopen

import pandas as pd

BIS_BULK_PAGE = "https://data.bis.org/bulkdownload"
BIS_BULK_FILE = "https://data.bis.org/static/bulk/{dataset}_csv_flat.zip"
USER_AGENT = "SEC-data-analysys BIS Excel builder (https://data.bis.org)"

DROP_COLS = {
    "STRUCTURE",
    "ACTION",
    "OBS_CONF",
    "OBS_PRE_BREAK",
    "TIME_FORMAT",
    "COLLECTION",
    "AVAILABILITY",
}

COUNTRY_COL_PRIORITY = (
    "REF_AREA",
    "BORROWERS_CTY",
    "L_REP_CTY",
    "REPORTING_COUNTRY",
    "ISSUER_RES",
    "ISSUER_CTY",
    "ISSUER_COUNTRY",
    "COUNTRY",
    "REF_CTY",
    "PARENT_CTY",
)

LOW_FREQ = ("A", "Q", "M")
HIGH_FREQ = ("W", "D", "B", "H")
SHEET_INVALID = re.compile(r"[\[\]:*?/\\]")
YEAR_RE = re.compile(r"^(\d{4})")
CODE_LABEL_RE = re.compile(r"^([^:]{1,12}):\s*(.+)$")

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CACHE = ROOT / "data" / "raw" / "bis"
DEFAULT_OUT = ROOT / "output" / "bis"

# Tópicos na ordem do portal https://data.bis.org/bulkdownload
TOPICOS: list[dict] = [
    {
        "id": "LBS",
        "titulo": "Estatísticas bancárias locacionais",
        "titulo_en": "Locational banking statistics",
        "datasets": ["WS_LBS_D_PUB"],
        "grande": True,
        "anos": 12,
    },
    {
        "id": "CBS",
        "titulo": "Estatísticas bancárias consolidadas",
        "titulo_en": "Consolidated banking statistics",
        "datasets": ["WS_CBS_PUB"],
        "grande": True,
        "anos": 12,
    },
    {
        "id": "DSS",
        "titulo": "Estatísticas de títulos de dívida",
        "titulo_en": "Debt securities statistics",
        "datasets": ["WS_NA_SEC_DSS"],
        "grande": True,
        "anos": 12,
    },
    {
        "id": "IDS",
        "titulo": "Títulos de dívida internacionais (compilação BIS)",
        "titulo_en": "International debt securities (BIS-compiled)",
        "datasets": ["WS_DEBT_SEC2_PUB"],
        "grande": True,
        "anos": 12,
    },
    {
        "id": "TC",
        "titulo": "Crédito ao setor não financeiro",
        "titulo_en": "Credit to the non-financial sector",
        "datasets": ["WS_TC"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CREDIT_GAP",
        "titulo": "Hiato crédito/PIB",
        "titulo_en": "Credit-to-GDP gaps",
        "datasets": ["WS_CREDIT_GAP"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "DSR",
        "titulo": "Índices de serviço da dívida",
        "titulo_en": "Debt service ratios",
        "datasets": ["WS_DSR"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "GLI",
        "titulo": "Liquidez global",
        "titulo_en": "Global liquidity",
        "datasets": ["WS_GLI"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "XTD",
        "titulo": "Derivativos negociados em bolsa",
        "titulo_en": "Exchange-traded derivatives statistics",
        "datasets": ["WS_XTD_DERIV"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "OTC",
        "titulo": "Derivativos de balcão em aberto",
        "titulo_en": "OTC derivatives outstanding",
        "datasets": ["WS_OTC_DERIV2"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "TRIENNIAL",
        "titulo": "Pesquisa trienal (Triennial Survey)",
        "titulo_en": "Triennial Survey",
        "datasets": ["WS_DER_OTC_TOV"],
        "grande": True,
        "anos": None,
    },
    {
        "id": "RPP",
        "titulo": "Preços de imóveis residenciais",
        "titulo_en": "Residential property prices",
        "datasets": ["WS_SPP", "WS_DPP"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CPP",
        "titulo": "Preços de imóveis comerciais",
        "titulo_en": "Commercial property prices",
        "datasets": ["WS_CPP"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CPI",
        "titulo": "Preços ao consumidor",
        "titulo_en": "Consumer prices",
        "datasets": ["WS_LONG_CPI"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "XRU",
        "titulo": "Taxas de câmbio bilaterais",
        "titulo_en": "Bilateral exchange rates",
        "datasets": ["WS_XRU"],
        "grande": True,
        "anos": 15,
    },
    {
        "id": "EER",
        "titulo": "Taxas de câmbio efetivas",
        "titulo_en": "Effective exchange rates",
        "datasets": ["WS_EER"],
        "grande": True,
        "anos": 15,
    },
    {
        "id": "CBTA",
        "titulo": "Ativo total dos bancos centrais",
        "titulo_en": "Central bank total assets",
        "datasets": ["WS_CBTA"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CBPOL",
        "titulo": "Taxas de política dos bancos centrais",
        "titulo_en": "Central bank policy rates",
        "datasets": ["WS_CBPOL"],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CPMI_VAREJO",
        "titulo": "Pagamentos de varejo, numerário e indicadores correlatos",
        "titulo_en": "Retail payments, currency and related indicators",
        "datasets": [
            "WS_CPMI_CT1",
            "WS_CPMI_CASHLESS",
            "WS_CPMI_INSTITUT",
            "WS_CPMI_MACRO",
            "WS_CPMI_DEVICES",
        ],
        "grande": False,
        "anos": None,
    },
    {
        "id": "CPMI_FMI",
        "titulo": "Infraestruturas do mercado financeiro e prestadores críticos",
        "titulo_en": "Financial market infrastructures and critical service providers",
        "datasets": ["WS_CPMI_CT2", "WS_CPMI_PARTICIP", "WS_CPMI_SYSTEMS"],
        "grande": False,
        "anos": None,
    },
]


@dataclass
class Pais:
    codigo: str
    nome: str
    agregado: bool

    @property
    def aba(self) -> str:
        return sanitizar_aba(self.codigo)


@dataclass
class ResultadoTema:
    topico_id: str
    arquivo: Path
    n_paises: int
    n_agregados: int
    n_linhas: int
    recortes: list[str] = field(default_factory=list)


def sanitizar_aba(nome: str) -> str:
    texto = SHEET_INVALID.sub("_", str(nome).strip()) or "NA"
    texto = texto.replace("'", "")
    return texto[:31]


def codigo_coluna(nome: str) -> str:
    return str(nome).split(":", 1)[0].strip()


def rotulo_coluna(nome: str) -> str:
    partes = str(nome).split(":", 1)
    return partes[1].strip() if len(partes) == 2 and partes[1].strip() else partes[0].strip()


def partir_codigo_rotulo(valor: object) -> tuple[str, str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "", ""
    texto = str(valor).strip()
    m = CODE_LABEL_RE.match(texto)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return texto, texto


def ano_periodo(valor: object) -> int | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    m = YEAR_RE.match(str(valor).strip())
    return int(m.group(1)) if m else None


def eh_agregado(codigo: str, nome: str) -> bool:
    c = (codigo or "").strip()
    n = (nome or "").lower()
    if not c:
        return True
    if c[0].isdigit():
        return True
    palavras = (
        "all countries",
        "all issuers",
        "world",
        "euro area",
        "advanced",
        "emerging",
        "offshore",
        "unallocated",
        "total",
        "international organisations",
        "residual",
        "aggregate",
    )
    return any(p in n for p in palavras)


def detectar_coluna_pais(colunas: Iterable[str]) -> str | None:
    mapa = {codigo_coluna(c).upper(): c for c in colunas}
    for chave in COUNTRY_COL_PRIORITY:
        if chave in mapa:
            return mapa[chave]
    for codigo, original in mapa.items():
        if "COUNTRY" in codigo or codigo.endswith("_CTY") or codigo.endswith("_AREA"):
            return original
    return None


def detectar_coluna_freq(colunas: Iterable[str]) -> str | None:
    for c in colunas:
        if codigo_coluna(c).upper() == "FREQ":
            return c
    return None


def detectar_coluna_tempo(colunas: Iterable[str]) -> str | None:
    for c in colunas:
        if codigo_coluna(c).upper() == "TIME_PERIOD":
            return c
    return None


def detectar_coluna_valor(colunas: Iterable[str]) -> str | None:
    for c in colunas:
        if codigo_coluna(c).upper() == "OBS_VALUE":
            return c
    return None


def frequencias_presentes(serie: pd.Series) -> set[str]:
    return {partir_codigo_rotulo(v)[0] for v in serie.dropna().unique()}


def filtrar_frequencia(df: pd.DataFrame, col_freq: str | None) -> tuple[pd.DataFrame, str | None]:
    if not col_freq or col_freq not in df.columns or df.empty:
        return df, None
    freqs = frequencias_presentes(df[col_freq])
    codes = {partir_codigo_rotulo(v)[0] for v in df[col_freq]}
    if freqs & set(LOW_FREQ) and freqs & set(HIGH_FREQ):
        mask = df[col_freq].map(lambda v: partir_codigo_rotulo(v)[0] in LOW_FREQ)
        return df.loc[mask].copy(), (
            "excluídas frequências diárias/semanais (mantidas A/Q/M)"
        )
    if codes & set(LOW_FREQ) and codes & set(HIGH_FREQ):
        mask = df[col_freq].map(lambda v: partir_codigo_rotulo(v)[0] in LOW_FREQ)
        return df.loc[mask].copy(), (
            "excluídas frequências diárias/semanais (mantidas A/Q/M)"
        )
    return df, None


def filtrar_anos(
    df: pd.DataFrame, col_tempo: str | None, anos: int | None, hoje: date | None = None
) -> tuple[pd.DataFrame, str | None]:
    if not anos or not col_tempo or col_tempo not in df.columns or df.empty:
        return df, None
    ref = hoje or date.today()
    minimo = ref.year - anos + 1
    years = df[col_tempo].map(ano_periodo)
    keep = years.isna() | (years >= minimo)
    return df.loc[keep].copy(), f"períodos a partir de {minimo}"


def limpar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    drop = [c for c in out.columns if codigo_coluna(c) in DROP_COLS]
    if drop:
        out = out.drop(columns=drop)
    rename = {c: codigo_coluna(c) for c in out.columns}
    # Evita colisão se duas colunas colapsarem no mesmo código
    used: dict[str, int] = {}
    final = {}
    for old, new in rename.items():
        n = used.get(new, 0)
        used[new] = n + 1
        final[old] = new if n == 0 else f"{new}_{n+1}"
    out = out.rename(columns=final)
    return out


def paises_de_serie(serie: pd.Series) -> dict[str, Pais]:
    encontrados: dict[str, Pais] = {}
    for valor in serie.dropna().unique():
        codigo, nome = partir_codigo_rotulo(valor)
        if not codigo:
            continue
        encontrados[codigo] = Pais(
            codigo=codigo, nome=nome or codigo, agregado=eh_agregado(codigo, nome)
        )
    return encontrados


def preparar_aba_pais(
    df: pd.DataFrame, col_pais: str, pais: Pais, max_linhas: int
) -> pd.DataFrame:
    codigo = pais.codigo
    mask = df[col_pais].map(lambda v: partir_codigo_rotulo(v)[0] == codigo)
    bloco = df.loc[mask].copy()
    if bloco.empty:
        return bloco
    bloco = bloco.drop(columns=[col_pais])
    if "OBS_VALUE" in bloco.columns:
        bloco["OBS_VALUE"] = pd.to_numeric(bloco["OBS_VALUE"], errors="coerce")
    if "TIME_PERIOD" in bloco.columns:
        bloco = bloco.sort_values(
            [c for c in ("TIME_PERIOD",) if c in bloco.columns] + [
                c for c in bloco.columns if c not in {"TIME_PERIOD", "OBS_VALUE"}
            ][:4]
        )
        if len(bloco) > max_linhas:
            bloco = bloco.tail(max_linhas)
    return bloco.reset_index(drop=True)


def comparativo_paises(
    df: pd.DataFrame, col_pais: str, paises: dict[str, Pais]
) -> pd.DataFrame:
    if df.empty or "TIME_PERIOD" not in df.columns or "OBS_VALUE" not in df.columns:
        return pd.DataFrame(
            columns=["codigo", "pais", "tipo", "periodo", "valor", "unidade", "serie"]
        )
    dimensoes = [
        c
        for c in df.columns
        if c not in {col_pais, "TIME_PERIOD", "OBS_VALUE", "STRUCTURE_ID"}
        and codigo_coluna(c) not in {"DECIMALS", "UNIT_MULT"}
    ]
    titulo_col = next((c for c in ("TITLE_TS", "TITLE") if c in df.columns), None)
    unidade_col = next((c for c in df.columns if codigo_coluna(c) == "UNIT_MEASURE"), None)
    linhas = []
    tmp = df.copy()
    tmp["_codigo"] = tmp[col_pais].map(lambda v: partir_codigo_rotulo(v)[0])
    tmp["OBS_VALUE"] = pd.to_numeric(tmp["OBS_VALUE"], errors="coerce")
    tmp = tmp.dropna(subset=["TIME_PERIOD", "OBS_VALUE"])
    for codigo, pais in sorted(paises.items(), key=lambda kv: (kv[1].agregado, kv[1].nome)):
        sub = tmp.loc[tmp["_codigo"] == codigo]
        if sub.empty:
            continue
        ultimo = sub["TIME_PERIOD"].astype(str).max()
        rec = sub.loc[sub["TIME_PERIOD"].astype(str) == ultimo]
        # Uma linha representativa: primeira série no último período
        row = rec.iloc[0]
        serie = ""
        if titulo_col and pd.notna(row.get(titulo_col)):
            serie = str(row[titulo_col])
        elif dimensoes:
            serie = " | ".join(
                str(row[c]) for c in dimensoes[:6] if c in row and pd.notna(row[c])
            )
        unidade = ""
        if unidade_col and pd.notna(row.get(unidade_col)):
            unidade = str(row[unidade_col])
        linhas.append(
            {
                "codigo": codigo,
                "pais": pais.nome,
                "tipo": "agregado" if pais.agregado else "pais",
                "periodo": ultimo,
                "valor": row["OBS_VALUE"],
                "unidade": unidade,
                "serie": serie,
                "n_obs": int(len(sub)),
                "aba": pais.aba,
            }
        )
    return pd.DataFrame(linhas)


def indice_paises(paises: dict[str, Pais], comparativo: pd.DataFrame) -> pd.DataFrame:
    base = pd.DataFrame(
        [
            {
                "codigo": p.codigo,
                "pais": p.nome,
                "tipo": "agregado" if p.agregado else "pais",
                "aba": p.aba,
            }
            for p in sorted(paises.values(), key=lambda x: (x.agregado, x.nome, x.codigo))
        ]
    )
    if comparativo.empty or base.empty:
        return base
    extra = comparativo[["codigo", "periodo", "valor", "n_obs"]].rename(
        columns={"periodo": "ultimo_periodo", "valor": "ultimo_valor"}
    )
    return base.merge(extra, on="codigo", how="left")


def baixar_arquivo(url: str, destino: Path, timeout: int = 180) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists() and destino.stat().st_size > 0:
        return destino
    tmp = destino.with_suffix(destino.suffix + ".part")
    req = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(req, timeout=timeout) as resp, open(tmp, "wb") as fh:
        while True:
            chunk = resp.read(1024 * 256)
            if not chunk:
                break
            fh.write(chunk)
    tmp.replace(destino)
    return destino


def ler_csv_zip(zip_path: Path, chunksize: int | None = None):
    with zipfile.ZipFile(zip_path) as zf:
        nomes = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not nomes:
            raise FileNotFoundError(f"Nenhum CSV em {zip_path}")
        with zf.open(nomes[0]) as raw:
            buf = io.TextIOWrapper(raw, encoding="utf-8", errors="replace")
            if chunksize:
                yield from pd.read_csv(buf, dtype=str, chunksize=chunksize, low_memory=False)
            else:
                yield pd.read_csv(buf, dtype=str, low_memory=False)


def _aplicar_filtros_chunk(
    chunk: pd.DataFrame,
    dataset: str,
    col_freq: str | None,
    col_tempo: str | None,
    anos: int | None,
    recortes: list[str],
) -> pd.DataFrame:
    chunk, rec_f = filtrar_frequencia(chunk, col_freq)
    chunk, rec_a = filtrar_anos(chunk, col_tempo, anos)
    if rec_f and rec_f not in recortes:
        recortes.append(rec_f)
    if rec_a and rec_a not in recortes:
        recortes.append(rec_a)
    valor_col = detectar_coluna_valor(chunk.columns)
    if valor_col:
        chunk = chunk.loc[
            chunk[valor_col].notna() & (chunk[valor_col].astype(str).str.strip() != "")
        ]
    if not chunk.empty:
        chunk = chunk.copy()
        chunk.insert(0, "DATASET", dataset)
    return chunk


def carregar_dataset(
    dataset: str,
    cache_dir: Path,
    grande: bool,
    anos: int | None,
    chunksize: int = 250_000,
) -> tuple[pd.DataFrame, list[str]]:
    recortes: list[str] = []
    url = BIS_BULK_FILE.format(dataset=dataset)
    zip_path = cache_dir / f"{dataset}_csv_flat.zip"
    baixar_arquivo(url, zip_path)
    partes: list[pd.DataFrame] = []
    col_freq = col_tempo = None
    for i, chunk in enumerate(ler_csv_zip(zip_path, chunksize=chunksize if grande else None)):
        if i == 0:
            col_freq = detectar_coluna_freq(chunk.columns)
            col_tempo = detectar_coluna_tempo(chunk.columns)
        chunk = _aplicar_filtros_chunk(chunk, dataset, col_freq, col_tempo, anos, recortes)
        if not chunk.empty:
            partes.append(chunk)
    if not partes:
        return pd.DataFrame(), recortes
    df = pd.concat(partes, ignore_index=True)
    return df, recortes


def _gravar_partes_pais(chunk: pd.DataFrame, col_pais: str, dest: Path, seq: int) -> dict[str, Pais]:
    found = paises_de_serie(chunk[col_pais])
    tmp = chunk.copy()
    tmp["_codigo"] = tmp[col_pais].map(lambda v: partir_codigo_rotulo(v)[0])
    for codigo, grupo in tmp.groupby("_codigo", sort=False):
        if not codigo:
            continue
        pasta = dest / codigo
        pasta.mkdir(parents=True, exist_ok=True)
        grupo.drop(columns=["_codigo"]).to_parquet(pasta / f"p{seq:05d}.parquet", index=False)
    return found


def carregar_dataset_particionado(
    dataset: str,
    cache_dir: Path,
    dest: Path,
    anos: int | None,
    chunksize: int = 250_000,
) -> tuple[dict[str, Pais], str, list[str]]:
    recortes: list[str] = []
    url = BIS_BULK_FILE.format(dataset=dataset)
    zip_path = cache_dir / f"{dataset}_csv_flat.zip"
    baixar_arquivo(url, zip_path)
    paises: dict[str, Pais] = {}
    col_pais = None
    col_freq = col_tempo = None
    for i, chunk in enumerate(ler_csv_zip(zip_path, chunksize=chunksize)):
        if i == 0:
            col_pais = detectar_coluna_pais(chunk.columns)
            col_freq = detectar_coluna_freq(chunk.columns)
            col_tempo = detectar_coluna_tempo(chunk.columns)
            if not col_pais:
                raise RuntimeError(f"{dataset}: coluna de país não identificada ({list(chunk.columns)})")
        chunk = _aplicar_filtros_chunk(chunk, dataset, col_freq, col_tempo, anos, recortes)
        if chunk.empty:
            continue
        paises.update(_gravar_partes_pais(chunk, col_pais, dest, i))
        print(f"    {dataset} bloco {i + 1}: +{len(chunk):,} linhas", flush=True)
    if not col_pais:
        raise RuntimeError(f"{dataset}: arquivo vazio")
    return paises, col_pais, recortes


def ler_pais_particionado(dest: Path, codigo: str) -> pd.DataFrame:
    pasta = dest / codigo
    arquivos = sorted(pasta.glob("*.parquet")) if pasta.exists() else []
    if not arquivos:
        return pd.DataFrame()
    return pd.concat([pd.read_parquet(p) for p in arquivos], ignore_index=True)


def montar_capa(
    topico: dict,
    recortes: list[str],
    n_paises: int,
    n_agregados: int,
    n_linhas: int,
    gerado_em: datetime,
) -> pd.DataFrame:
    linhas = [
        ("Tema", topico["titulo"]),
        ("Tema (en)", topico["titulo_en"]),
        ("Portal", BIS_BULK_PAGE),
        ("Dataflows", ", ".join(topico["datasets"])),
        ("Arquivos-fonte", ", ".join(f"{d}_csv_flat.zip" for d in topico["datasets"])),
        ("Gerado em", gerado_em.strftime("%Y-%m-%d %H:%M UTC")),
        ("Países (abas)", str(n_paises)),
        ("Agregados (abas)", str(n_agregados)),
        ("Linhas de dados", f"{n_linhas:,}".replace(",", ".")),
        (
            "Organização",
            "Uma aba por país/jurisdição com as séries do tema. "
            "Agregados regionais (ex.: zona do euro, economias avançadas) também têm aba própria.",
        ),
        (
            "Recorte",
            "; ".join(recortes) if recortes else "conjunto completo (após excluir linhas sem valor)",
        ),
        (
            "Citação",
            "Bank for International Settlements, BIS Data Portal "
            f"({date.today().year}), {topico['titulo_en']}. {BIS_BULK_PAGE}",
        ),
        (
            "Termos",
            "https://www.bis.org/terms_conditions.htm — observe as condições de uso do BIS.",
        ),
    ]
    return pd.DataFrame(linhas, columns=["campo", "conteudo"])


def escrever_excel(
    caminho: Path,
    capa: pd.DataFrame,
    indice: pd.DataFrame,
    comparativo: pd.DataFrame,
    abas: list[tuple[Pais, pd.DataFrame]],
) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    usados: dict[str, int] = {}

    def nome_unico(base: str) -> str:
        n = sanitizar_aba(base)
        k = usados.get(n, 0)
        usados[n] = k + 1
        if k == 0:
            return n
        suf = f"_{k+1}"
        return sanitizar_aba(n[: 31 - len(suf)] + suf)

    with pd.ExcelWriter(caminho, engine="xlsxwriter") as writer:
        capa.to_excel(writer, sheet_name="Capa", index=False)
        indice.to_excel(writer, sheet_name="Indice", index=False)
        comparativo.to_excel(writer, sheet_name="Comparativo", index=False)
        book = writer.book
        header_fmt = book.add_format({"bold": True, "bg_color": "#1B4F72", "font_color": "white"})
        wrap = book.add_format({"text_wrap": True, "valign": "top"})
        for sheet_name in ("Capa", "Indice", "Comparativo"):
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 0, max(0, (capa if sheet_name == "Capa" else indice if sheet_name == "Indice" else comparativo).shape[1] - 1))
            ws.set_column(0, 0, 22)
            ws.set_column(1, 8, 28, wrap)
            ws.set_row(0, 18, header_fmt)
        for pais, bloco in abas:
            aba = nome_unico(pais.aba)
            if bloco.empty:
                bloco = pd.DataFrame({"aviso": [f"Sem observações para {pais.nome}"]})
            bloco.to_excel(writer, sheet_name=aba, index=False)
            ws = writer.sheets[aba]
            ws.freeze_panes(1, 0)
            ncols = max(0, bloco.shape[1] - 1)
            if len(bloco):
                ws.autofilter(0, 0, len(bloco), ncols)
            ws.set_column(0, ncols, 18)
            ws.set_row(0, 18, header_fmt)


def _slug_arquivo(topico: dict) -> str:
    slug = topico["titulo_en"].lower()
    return re.sub(r"[^a-z0-9]+", "_", slug).strip("_")


def _montar_abas(
    obter_df,
    col_pais: str,
    paises: dict[str, Pais],
    max_linhas_aba: int,
) -> tuple[pd.DataFrame, list[tuple[Pais, pd.DataFrame]], int]:
    """obter_df(codigo) -> DataFrame cru (ainda com coluna de país)."""
    amostras: list[pd.DataFrame] = []
    abas: list[tuple[Pais, pd.DataFrame]] = []
    n_linhas = 0
    ordenados = sorted(paises.values(), key=lambda p: (p.agregado, p.nome.lower(), p.codigo))
    for pais in ordenados:
        bruto = obter_df(pais.codigo)
        if bruto.empty:
            abas.append((pais, bruto))
            continue
        bruto = limpar_colunas(bruto)
        col = codigo_coluna(col_pais)
        if col not in bruto.columns:
            # já pode ter sido limpa com o nome curto
            col = next((c for c in bruto.columns if codigo_coluna(c) == codigo_coluna(col_pais)), col)
        bloco = preparar_aba_pais(bruto, col, pais, max_linhas_aba)
        n_linhas += len(bloco)
        abas.append((pais, bloco))
        # amostra para o comparativo (último período)
        if "TIME_PERIOD" in bruto.columns:
            amostras.append(bruto)
        else:
            amostras.append(bruto)
    if amostras:
        # comparativo só precisa do último período: reduz memória
        reduzidas = []
        for a in amostras:
            if "TIME_PERIOD" in a.columns and len(a) > 4000:
                ultimo = a["TIME_PERIOD"].astype(str).max()
                reduzidas.append(a.loc[a["TIME_PERIOD"].astype(str) == ultimo])
            else:
                reduzidas.append(a)
        painel = pd.concat(reduzidas, ignore_index=True, sort=False)
    else:
        painel = pd.DataFrame()
    col = codigo_coluna(col_pais)
    if not painel.empty and col not in painel.columns:
        col = next((c for c in painel.columns if c == col or codigo_coluna(c) == col), col)
    comparativo = comparativo_paises(painel, col, paises) if not painel.empty else pd.DataFrame()
    return comparativo, abas, n_linhas


def gerar_tema(
    topico: dict,
    cache_dir: Path,
    saida_dir: Path,
    max_linhas_aba: int,
    gerado_em: datetime | None = None,
) -> ResultadoTema:
    gerado_em = gerado_em or datetime.now(timezone.utc)
    recortes: list[str] = []
    grande = bool(topico.get("grande"))
    teto = min(max_linhas_aba, 15_000) if grande else max_linhas_aba

    if grande:
        tmp = Path(tempfile.mkdtemp(prefix=f"bis_{topico['id']}_"))
        try:
            paises: dict[str, Pais] = {}
            col_pais = None
            for dataset in topico["datasets"]:
                pmap, col, rec = carregar_dataset_particionado(
                    dataset,
                    cache_dir=cache_dir,
                    dest=tmp,
                    anos=topico.get("anos"),
                )
                recortes.extend(f"{dataset}: {r}" for r in rec)
                paises.update(pmap)
                col_pais = col_pais or col
            if not paises or not col_pais:
                raise RuntimeError(f"Sem dados para o tema {topico['id']}")

            def obter(codigo: str) -> pd.DataFrame:
                return ler_pais_particionado(tmp, codigo)

            comparativo, abas, n_linhas = _montar_abas(obter, col_pais, paises, teto)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    else:
        frames: list[pd.DataFrame] = []
        for dataset in topico["datasets"]:
            df, rec = carregar_dataset(
                dataset,
                cache_dir=cache_dir,
                grande=False,
                anos=topico.get("anos"),
            )
            recortes.extend(f"{dataset}: {r}" for r in rec)
            if not df.empty:
                frames.append(df)
        if not frames:
            raise RuntimeError(f"Sem dados para o tema {topico['id']}")
        dados = pd.concat(frames, ignore_index=True, sort=False)
        col_pais = detectar_coluna_pais(dados.columns)
        if not col_pais:
            raise RuntimeError(
                f"Tema {topico['id']}: coluna de país não identificada ({list(dados.columns)})"
            )
        paises = paises_de_serie(dados[col_pais])
        if not paises:
            raise RuntimeError(f"Tema {topico['id']}: nenhum país na coluna {col_pais}")

        def obter(codigo: str) -> pd.DataFrame:
            mask = dados[col_pais].map(lambda v: partir_codigo_rotulo(v)[0] == codigo)
            return dados.loc[mask].copy()

        comparativo, abas, n_linhas = _montar_abas(obter, col_pais, paises, teto)

    indice = indice_paises(paises, comparativo)
    n_paises = sum(1 for p in paises.values() if not p.agregado)
    n_agregados = sum(1 for p in paises.values() if p.agregado)
    capa = montar_capa(topico, recortes, n_paises, n_agregados, n_linhas, gerado_em)
    arquivo = saida_dir / f"BIS_{topico['id']}_{_slug_arquivo(topico)}.xlsx"
    escrever_excel(arquivo, capa, indice, comparativo, abas)
    return ResultadoTema(
        topico_id=topico["id"],
        arquivo=arquivo,
        n_paises=n_paises,
        n_agregados=n_agregados,
        n_linhas=n_linhas,
        recortes=recortes,
    )


def escrever_catalogo(resultados: list[ResultadoTema], saida_dir: Path, gerado_em: datetime) -> Path:
    rows = [
        {
            "arquivo": r.arquivo.name,
            "topico": r.topico_id,
            "paises": r.n_paises,
            "agregados": r.n_agregados,
            "linhas": r.n_linhas,
            "recorte": "; ".join(r.recortes),
        }
        for r in resultados
    ]
    df = pd.DataFrame(rows)
    path = saida_dir / "BIS_00_catalogo_temas.xlsx"
    capa = pd.DataFrame(
        [
            ("Fonte", BIS_BULK_PAGE),
            ("Gerado em", gerado_em.strftime("%Y-%m-%d %H:%M UTC")),
            ("Arquivos", str(len(resultados))),
            (
                "Descrição",
                "Um Excel por tema do portal BIS. Em cada arquivo, aba por país "
                "com as séries daquele tema.",
            ),
        ],
        columns=["campo", "conteudo"],
    )
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        capa.to_excel(writer, sheet_name="Capa", index=False)
        df.to_excel(writer, sheet_name="Temas", index=False)
    return path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--topics",
        default="",
        help="IDs separados por vírgula (ex.: CBPOL,CPI). Vazio = todos os 20 temas.",
    )
    p.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    p.add_argument("--saida", type=Path, default=DEFAULT_OUT)
    p.add_argument(
        "--max-linhas-aba",
        type=int,
        default=200_000,
        help="Teto de linhas por aba de país (as mais recentes são mantidas).",
    )
    return p.parse_args(argv)


def selecionar_topicos(filtro: str) -> list[dict]:
    if not filtro.strip():
        return list(TOPICOS)
    ids = {x.strip().upper() for x in filtro.split(",") if x.strip()}
    escolhidos = [t for t in TOPICOS if t["id"] in ids]
    faltando = ids - {t["id"] for t in escolhidos}
    if faltando:
        raise SystemExit(f"Tópicos desconhecidos: {sorted(faltando)}. Válidos: {[t['id'] for t in TOPICOS]}")
    return escolhidos


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    topicos = selecionar_topicos(args.topics)
    args.cache.mkdir(parents=True, exist_ok=True)
    args.saida.mkdir(parents=True, exist_ok=True)
    gerado_em = datetime.now(timezone.utc)
    resultados: list[ResultadoTema] = []
    for topico in topicos:
        print(f"→ {topico['id']}: {topico['titulo']}", flush=True)
        try:
            res = gerar_tema(
                topico,
                cache_dir=args.cache,
                saida_dir=args.saida,
                max_linhas_aba=args.max_linhas_aba,
                gerado_em=gerado_em,
            )
        except Exception as exc:  # noqa: BLE001 — relata e segue os demais temas
            print(f"  ERRO {topico['id']}: {exc}", file=sys.stderr, flush=True)
            continue
        resultados.append(res)
        print(
            f"  {res.arquivo.name}  países={res.n_paises} agregados={res.n_agregados} linhas={res.n_linhas}",
            flush=True,
        )
    if resultados:
        cat = escrever_catalogo(resultados, args.saida, gerado_em)
        print(f"Catálogo: {cat}", flush=True)
    print(f"Concluído: {len(resultados)}/{len(topicos)} temas", flush=True)
    return 0 if resultados else 1


if __name__ == "__main__":
    raise SystemExit(main())
