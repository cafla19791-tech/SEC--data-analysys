"""CPI oficial, juros básicos em dias úteis e ranking 1995–2025 (BIS).

Fontes:
  https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip
  https://data.bis.org/static/bulk/WS_CBPOL_csv_col.zip

Uso:
  python3 scripts/bis_cpi_juros_ranking.py
  python3 scripts/bis_cpi_juros_ranking.py --sem-download
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.bis_urls_paises import (  # noqa: E402
    PAISES_PT,
    nome_aba,
    nome_pais,
)

DATA_DIR = ROOT / "data" / "bis"
OUTPUT_DIR = ROOT / "output" / "bis_paises"
URL_CPI = "https://data.bis.org/static/bulk/WS_LONG_CPI_csv_col.zip"
URL_POL = "https://data.bis.org/static/bulk/WS_CBPOL_csv_col.zip"
INICIO = pd.Timestamp("1995-01-01")
ANO_INI_RANK = 1995
ANO_FIM_RANK = 2025
DIA_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
MES_RE = re.compile(r"^\d{4}-\d{2}$")
ANO_RE = re.compile(r"^\d{4}$")


def baixar(url: str, dest: Path, baixar: bool) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 0:
        return dest
    if not baixar:
        raise FileNotFoundError(dest)
    print(f"  baixando {url}", flush=True)
    resp = requests.get(url, timeout=300)
    resp.raise_for_status()
    dest.write_bytes(resp.content)
    return dest


def _abrir_csv_zip(zip_path: Path):
    zf = zipfile.ZipFile(zip_path)
    nome = next(n for n in zf.namelist() if n.lower().endswith(".csv"))
    raw = zf.open(nome)
    return zf, raw


def ler_csv_col(zip_path: Path, usecols=None) -> pd.DataFrame:
    import csv
    import io

    zf, raw = _abrir_csv_zip(zip_path)
    try:
        fh = io.TextIOWrapper(raw, encoding="utf-8", newline="")
        reader = csv.reader(fh)
        header = next(reader)
        if callable(usecols):
            keep = [i for i, c in enumerate(header) if usecols(c)]
        elif usecols is None:
            keep = list(range(len(header)))
        else:
            wanted = set(usecols)
            keep = [i for i, c in enumerate(header) if c in wanted]
        cols = [header[i] for i in keep]
        linhas = [[rec[i] if i < len(rec) else "" for i in keep] for rec in reader]
        return pd.DataFrame(linhas, columns=cols)
    finally:
        raw.close()
        zf.close()


def _formatos(wb) -> dict:
    return {
        "titulo": wb.add_format({"font_name": "Calibri", "font_size": 14, "bold": True, "font_color": "#1F4E79"}),
        "nota": wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True, "valign": "top"}),
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
        "num4": wb.add_format({"font_name": "Calibri", "font_size": 9, "num_format": "0.000000"}),
        "data": wb.add_format({"font_name": "Calibri", "font_size": 9, "num_format": "dd/mm/yyyy"}),
        "txt": wb.add_format({"font_name": "Calibri", "font_size": 9}),
        "int": wb.add_format({"font_name": "Calibri", "font_size": 9, "align": "center"}),
    }


def _notas(wb, linhas: list[str], fmt: dict) -> None:
    ws = wb.add_worksheet("Notas")
    ws.write(0, 0, linhas[0], fmt["titulo"])
    ws.set_column(0, 0, 120)
    for i, txt in enumerate(linhas[1:], start=1):
        ws.write(i, 0, txt, fmt["nota"])
        ws.set_row(i, 30)


def _cabecalho(ws, cabs: list[str], fmt, larguras: list[float] | None = None) -> None:
    for i, cab in enumerate(cabs):
        ws.write(0, i, cab, fmt["cab"])
        ws.set_column(i, i, larguras[i] if larguras else 16)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, 0, max(len(cabs) - 1, 0))
    ws.set_row(0, 24)


def fator_diario(taxa_aa: float, n_uteis_ano: int) -> float:
    if n_uteis_ano <= 0 or pd.isna(taxa_aa):
        return np.nan
    return float((1.0 + float(taxa_aa) / 100.0) ** (1.0 / n_uteis_ano))


def juros_real_fisher(nominal: float, inflacao: float) -> float:
    return (1.0 + nominal / 100.0) / (1.0 + inflacao / 100.0) * 100.0 - 100.0


def calendario_uteis(inicio: pd.Timestamp, fim: pd.Timestamp) -> pd.DatetimeIndex:
    dias = pd.date_range(inicio, fim, freq="D")
    return dias[dias.weekday < 5]


def montar_juros_uteis(serie_diaria: pd.Series, inicio: pd.Timestamp = INICIO) -> pd.DataFrame:
    """Expande a taxa oficial para dias úteis (seg–sex) e acumula no mês e no ano.

    A taxa BIS já é % a.a. O fator de um pregão é (1+i/100)^(1/N), em que N é
    o número de segunda–sexta daquele ano-calendário. Assim, se a taxa ficar
    constante o ano inteiro, o acumulado no ano coincide com a taxa oficial.
    Feriados nacionais não são excluídos (o BIS não publica o calendário).
    Pregões sem cotação herdam a última taxa vigente (forward-fill).
    """
    s = pd.to_numeric(serie_diaria, errors="coerce")
    s.index = pd.to_datetime(s.index)
    s = s.sort_index()
    s = s[s.index >= inicio]
    s = s.dropna()
    if s.empty:
        return pd.DataFrame(
            columns=[
                "data",
                "taxa_basica_aa",
                "taxa_eq_diaria",
                "taxa_acum_mes",
                "taxa_acum_ano",
            ]
        )
    uteis = calendario_uteis(max(inicio, s.index.min().normalize()), s.index.max().normalize())
    vig = s.reindex(s.index.union(uteis)).sort_index().ffill().reindex(uteis)
    vig = vig.dropna()
    if vig.empty:
        return montar_juros_uteis(pd.Series(dtype=float))
    df = pd.DataFrame({"data": vig.index, "taxa_basica_aa": vig.to_numpy(dtype=float)})
    df["ano"] = df["data"].dt.year
    df["mes"] = df["data"].dt.month
    n_ano = df.groupby("ano")["data"].transform("size")
    fator = (1.0 + df["taxa_basica_aa"] / 100.0) ** (1.0 / n_ano)
    df["taxa_eq_diaria"] = (fator - 1.0) * 100.0
    df["taxa_acum_mes"] = (fator.groupby([df["ano"], df["mes"]]).cumprod() - 1.0) * 100.0
    df["taxa_acum_ano"] = (fator.groupby(df["ano"]).cumprod() - 1.0) * 100.0
    return df[["data", "taxa_basica_aa", "taxa_eq_diaria", "taxa_acum_mes", "taxa_acum_ano"]]


def extrair_cpi(df: pd.DataFrame) -> dict[str, dict[str, pd.DataFrame]]:
    """{codigo: {'mensal': df, 'anual': df, 'nome': str}}."""
    periodos_m = [c for c in df.columns if MES_RE.match(str(c))]
    periodos_a = [c for c in df.columns if ANO_RE.match(str(c))]
    id_cols = ["FREQ", "REF_AREA", "Reference area", "UNIT_MEASURE"]
    nomes = {
        str(c): nome_pais(str(c), str(n))
        for c, n in df.drop_duplicates("REF_AREA")[["REF_AREA", "Reference area"]].itertuples(index=False)
    }

    def _bloco(periodos: list[str], freq: str) -> pd.DataFrame:
        if not periodos:
            return pd.DataFrame(columns=["REF_AREA", "periodo", "indice_2010", "var_12m"])
        sub = df.loc[df["FREQ"] == freq, id_cols + periodos]
        long = sub.melt(id_vars=id_cols, value_vars=periodos, var_name="periodo", value_name="valor")
        long["valor"] = pd.to_numeric(long["valor"], errors="coerce")
        long = long.dropna(subset=["valor"])
        if long.empty:
            return pd.DataFrame(columns=["REF_AREA", "periodo", "indice_2010", "var_12m"])
        idx = long.loc[long["UNIT_MEASURE"] == "628", ["REF_AREA", "periodo", "valor"]].rename(
            columns={"valor": "indice_2010"}
        )
        yoy = long.loc[long["UNIT_MEASURE"] == "771", ["REF_AREA", "periodo", "valor"]].rename(
            columns={"valor": "var_12m"}
        )
        return pd.merge(idx, yoy, on=["REF_AREA", "periodo"], how="outer")

    mensal = _bloco(periodos_m, "M")
    anual = _bloco(periodos_a, "A")
    out: dict[str, dict] = {}
    for codigo, nome in nomes.items():
        m = mensal.loc[mensal["REF_AREA"] == codigo, ["periodo", "indice_2010", "var_12m"]].sort_values("periodo")
        a = anual.loc[anual["REF_AREA"] == codigo, ["periodo", "indice_2010", "var_12m"]].sort_values("periodo")
        if m.empty and a.empty:
            continue
        out[codigo] = {"nome": nome, "mensal": m.reset_index(drop=True), "anual": a.reset_index(drop=True)}
    return out


def extrair_politica_diaria(df: pd.DataFrame) -> dict[str, tuple[str, pd.Series]]:
    dias = [c for c in df.columns if DIA_RE.match(str(c)) and str(c) >= "1995-01-01"]
    diario = df[df["FREQ"] == "D"].copy()
    out = {}
    for _, rec in diario.iterrows():
        codigo = str(rec["REF_AREA"])
        nome = nome_pais(codigo, str(rec["Reference area"]))
        s = pd.to_numeric(rec[dias], errors="coerce")
        s.index = pd.to_datetime(s.index)
        out[codigo] = (nome, s)
    return out


def gerar_cpi(paises: dict, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path), {"strings_to_urls": False})
    fmt = _formatos(wb)
    _notas(
        wb,
        [
            "Índices oficiais de inflação ao consumidor — BIS WS_LONG_CPI",
            "Uma aba por país/economia. Colunas mensais e anuais lado a lado.",
            "Índice: 2010 = 100 (unidade BIS 628). Variação em 12 meses: % a.a. (unidade BIS 771).",
            f"Fonte: {URL_CPI}",
            "A variação anual oficial é a variação do índice médio do ano (não necessariamente dez/dez).",
        ],
        fmt,
    )
    usados = {"Notas", "Indice"}
    ordem = sorted(paises.items(), key=lambda kv: (kv[0] != "BR", paises[kv[0]]["nome"].casefold(), kv[0]))
    indice = []
    ws_i = wb.add_worksheet("Indice")
    for codigo, info in ordem:
        aba = nome_aba(codigo, info["nome"], usados)
        ws = wb.add_worksheet(aba)
        mensal = info.get("mensal", pd.DataFrame())
        anual = info.get("anual", pd.DataFrame())
        cabs = [
            "Período mensal",
            "Índice mensal (2010=100)",
            "Inflação 12 meses %",
            "",
            "Ano",
            "Índice anual (2010=100)",
            "Inflação anual %",
        ]
        _cabecalho(ws, cabs, fmt, [16, 24, 20, 3, 10, 24, 18])
        n = max(len(mensal), len(anual))
        for i in range(n):
            if i < len(mensal):
                ws.write(i + 1, 0, mensal.iloc[i]["periodo"], fmt["txt"])
                v = mensal.iloc[i]["indice_2010"]
                w = mensal.iloc[i]["var_12m"]
                if pd.notna(v):
                    ws.write_number(i + 1, 1, float(v), fmt["num"])
                if pd.notna(w):
                    ws.write_number(i + 1, 2, float(w), fmt["num"])
            if i < len(anual):
                ws.write(i + 1, 4, anual.iloc[i]["periodo"], fmt["txt"])
                v = anual.iloc[i]["indice_2010"]
                w = anual.iloc[i]["var_12m"]
                if pd.notna(v):
                    ws.write_number(i + 1, 5, float(v), fmt["num"])
                if pd.notna(w):
                    ws.write_number(i + 1, 6, float(w), fmt["num"])
        if n:
            ws.autofilter(0, 0, n, 2)
        indice.append(
            {
                "Código": codigo,
                "País": info["nome"],
                "Aba": aba,
                "Meses": int(len(mensal)),
                "Anos": int(len(anual)),
                "Início mensal": "" if mensal.empty else mensal.iloc[0]["periodo"],
                "Fim mensal": "" if mensal.empty else mensal.iloc[-1]["periodo"],
            }
        )
    _cabecalho(ws_i, ["Código", "País", "Aba", "Meses", "Anos", "Início mensal", "Fim mensal"], fmt, [10, 28, 28, 10, 10, 14, 14])
    for i, rec in enumerate(indice, start=1):
        for j, k in enumerate(["Código", "País", "Aba", "Meses", "Anos", "Início mensal", "Fim mensal"]):
            val = rec[k]
            if isinstance(val, int):
                ws_i.write_number(i, j, val, fmt["int"])
            else:
                ws_i.write(i, j, val, fmt["txt"])
    wb.close()
    return path


def gerar_juros(paises: dict[str, tuple[str, pd.Series]], path: Path) -> dict[str, pd.DataFrame]:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path), {"strings_to_urls": False})
    fmt = _formatos(wb)
    _notas(
        wb,
        [
            "Taxas básicas de juros — BIS WS_CBPOL, dias úteis desde 01/01/1995",
            "Uma aba por país. Só segunda a sexta. Feriados nacionais não são excluídos.",
            "Taxa básica % a.a.: valor oficial do BIS (fim do dia), com forward-fill nos pregões sem cotação.",
            "Taxa equivalente diária: (1 + i/100)^(1/N) − 1, N = número de segunda–sexta do ano.",
            "Taxa mensal / anual: produtório dos fatores diários no mês ou no ano (acumulada até aquele dia).",
            "Se a taxa oficial ficar constante o ano inteiro, o acumulado no ano coincide com ela.",
            f"Fonte: {URL_POL}",
        ],
        fmt,
    )
    usados = {"Notas", "Indice"}
    ordem = sorted(paises.items(), key=lambda kv: (kv[0] != "BR", kv[1][0].casefold(), kv[0]))
    tabelas: dict[str, pd.DataFrame] = {}
    indice = []
    ws_i = wb.add_worksheet("Indice")
    cabs = [
        "Data",
        "Taxa básica % a.a.",
        "Taxa equivalente diária %",
        "Taxa mensal acumulada %",
        "Taxa anual acumulada %",
    ]
    for codigo, (nome, serie) in ordem:
        tab = montar_juros_uteis(serie, INICIO)
        tabelas[codigo] = tab
        aba = nome_aba(codigo, nome, usados)
        ws = wb.add_worksheet(aba)
        _cabecalho(ws, cabs, fmt, [12, 18, 24, 24, 24])
        dados = tab.itertuples(index=False, name=None)
        for i, rec in enumerate(dados, start=1):
            ws.write_datetime(i, 0, rec[0].to_pydatetime(), fmt["data"])
            ws.write_number(i, 1, float(rec[1]), fmt["num"])
            ws.write_number(i, 2, float(rec[2]), fmt["num4"])
            ws.write_number(i, 3, float(rec[3]), fmt["num"])
            ws.write_number(i, 4, float(rec[4]), fmt["num"])
        if not tab.empty:
            ws.autofilter(0, 0, len(tab), 4)
        indice.append(
            {
                "Código": codigo,
                "País": nome,
                "Aba": aba,
                "Pregões": int(len(tab)),
                "Início": "" if tab.empty else tab.iloc[0]["data"].strftime("%d/%m/%Y"),
                "Fim": "" if tab.empty else tab.iloc[-1]["data"].strftime("%d/%m/%Y"),
            }
        )
        print(f"  {codigo} {nome}: {len(tab)} pregões", flush=True)
    _cabecalho(ws_i, ["Código", "País", "Aba", "Pregões", "Início", "Fim"], fmt, [10, 28, 28, 12, 14, 14])
    for i, rec in enumerate(indice, start=1):
        for j, k in enumerate(["Código", "País", "Aba", "Pregões", "Início", "Fim"]):
            val = rec[k]
            if isinstance(val, int):
                ws_i.write_number(i, j, val, fmt["int"])
            else:
                ws_i.write(i, j, val, fmt["txt"])
    wb.close()
    return tabelas


def inflacao_anual(paises_cpi: dict) -> pd.DataFrame:
    recs = []
    for codigo, info in paises_cpi.items():
        anual = info.get("anual", pd.DataFrame())
        if anual.empty:
            continue
        for _, row in anual.iterrows():
            if not ANO_RE.match(str(row["periodo"])):
                continue
            recs.append(
                {
                    "codigo": codigo,
                    "pais": info["nome"],
                    "ano": int(row["periodo"]),
                    "inflacao": row["var_12m"],
                }
            )
    return pd.DataFrame(recs)


def juros_anual(tabelas: dict[str, pd.DataFrame], nomes: dict[str, str]) -> pd.DataFrame:
    recs = []
    for codigo, tab in tabelas.items():
        if tab.empty:
            continue
        for ano, g in tab.groupby(tab["data"].dt.year):
            recs.append(
                {
                    "codigo": codigo,
                    "pais": nomes[codigo],
                    "ano": int(ano),
                    "juros_nominais": float(g.iloc[-1]["taxa_acum_ano"]),
                    "taxa_fim_ano": float(g.iloc[-1]["taxa_basica_aa"]),
                }
            )
    return pd.DataFrame(recs)


def montar_ranking(juros: pd.DataFrame, infla: pd.DataFrame) -> pd.DataFrame:
    base = pd.merge(juros, infla, on=["codigo", "ano"], how="outer", suffixes=("", "_inf"))
    if "pais_inf" in base.columns:
        base["pais"] = base["pais"].where(base["pais"].notna(), base["pais_inf"])
        base = base.drop(columns=["pais_inf"])
    base["juros_reais"] = np.where(
        base["juros_nominais"].notna() & base["inflacao"].notna(),
        juros_real_fisher(base["juros_nominais"].to_numpy(), base["inflacao"].to_numpy()),
        np.nan,
    )
    return base


def _rank_desc(serie: pd.Series) -> pd.Series:
    return serie.rank(ascending=False, method="min")


def gerar_ranking(base: pd.DataFrame, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path), {"strings_to_urls": False})
    fmt = _formatos(wb)
    _notas(
        wb,
        [
            "Ranking anual 1995–2025 — juros nominais, juros reais e inflação",
            "Juros nominais: taxa básica acumulada no ano (produtório dos fatores diários em dias úteis).",
            "Inflação: variação oficial em 12 meses do índice anual BIS (WS_LONG_CPI, unidade 771).",
            "Juros reais: fórmula de Fisher, (1+i)/(1+π) − 1, com i = juros nominais acumulados e π = inflação oficial.",
            "Só entra no ranking o país com o indicador daquela coluna no ano. A zona do euro (XM) entra como economia.",
            f"Fontes: {URL_POL}  |  {URL_CPI}",
        ],
        fmt,
    )
    for ano in range(ANO_INI_RANK, ANO_FIM_RANK + 1):
        rec = base[base["ano"] == ano].copy()
        ws = wb.add_worksheet(str(ano))
        blocos = [
            (0, "Juros nominais % a.a.", "juros_nominais"),
            (4, "Inflação oficial %", "inflacao"),
            (8, "Juros reais (Fisher) %", "juros_reais"),
        ]
        ws.merge_range(0, 0, 0, 10, f"Rankings {ano}", fmt["titulo"])
        for col0, titulo, campo in blocos:
            sub = rec.dropna(subset=[campo]).sort_values(campo, ascending=False).reset_index(drop=True)
            ws.write(2, col0, titulo, fmt["cab"])
            ws.merge_range(2, col0, 2, col0 + 2, titulo, fmt["cab"])
            for j, cab in enumerate(["Posição", "País", titulo.split(" (")[0] + " %"]):
                ws.write(3, col0 + j, cab, fmt["cab"])
            for i, row in sub.iterrows():
                ws.write_number(i + 4, col0, i + 1, fmt["int"])
                ws.write(i + 4, col0 + 1, row["pais"], fmt["txt"])
                ws.write_number(i + 4, col0 + 2, float(row[campo]), fmt["num"])
            ws.set_column(col0, col0, 10)
            ws.set_column(col0 + 1, col0 + 1, 28)
            ws.set_column(col0 + 2, col0 + 2, 22)
        ws.set_column(3, 3, 3)
        ws.set_column(7, 7, 3)
        ws.freeze_panes(4, 0)
    wb.close()
    return path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--sem-download", action="store_true")
    args = parser.parse_args(argv)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    zip_cpi = baixar(URL_CPI, args.cache_dir / "WS_LONG_CPI_csv_col.zip", not args.sem_download)
    zip_pol = baixar(URL_POL, args.cache_dir / "WS_CBPOL_csv_col.zip", not args.sem_download)

    print("CPI — lendo", flush=True)
    cpi_bruto = ler_csv_col(zip_cpi)
    print(f"  {cpi_bruto.shape[0]} séries × {cpi_bruto.shape[1]} colunas", flush=True)
    paises_cpi = extrair_cpi(cpi_bruto)
    dest_cpi = args.output_dir / "bis_WS_LONG_CPI_inflacao.xlsx"
    print("  gravando planilha de inflação", flush=True)
    gerar_cpi(paises_cpi, dest_cpi)
    print(f"  {len(paises_cpi)} países → {dest_cpi}", flush=True)

    print("CBPOL — lendo diário", flush=True)
    def _cols_pol(nome: str) -> bool:
        return nome in {"FREQ", "REF_AREA", "Reference area"} or bool(DIA_RE.match(nome) and nome >= "1995-01-01")

    pol_bruto = ler_csv_col(zip_pol, usecols=_cols_pol)
    paises_pol = extrair_politica_diaria(pol_bruto)
    dest_j = args.output_dir / "bis_WS_CBPOL_juros_uteis.xlsx"
    print("  montando pregões e acumulados", flush=True)
    tabelas = gerar_juros(paises_pol, dest_j)
    print(f"  {len(tabelas)} países → {dest_j}", flush=True)

    nomes = {c: n for c, (n, _s) in paises_pol.items()}
    juros = juros_anual(tabelas, nomes)
    infla = inflacao_anual(paises_cpi)
    base = montar_ranking(juros, infla)
    dest_r = args.output_dir / "bis_ranking_juros_inflacao_1995_2025.xlsx"
    gerar_ranking(base, dest_r)
    print(f"Ranking → {dest_r}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
