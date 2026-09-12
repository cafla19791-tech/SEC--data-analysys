#!/usr/bin/env python3
"""Municípios de Minas Gerais em que Lula venceu o 2º turno de 2022, por mesorregião.

A composição das 12 mesorregiões é a do documento do governo de Minas
(lista IBGE de meso e microrregiões):
https://www.mg.gov.br/sites/default/files/paginas/arquivos/2016/ligminas_10_2_04_listamesomicro.pdf

O catálogo estruturado está em data/tse_catalog/mesorregioes_mg.csv
(Divisão Territorial Brasileira 2016 do IBGE — a mesma classificação do PDF).

Saída:
  output/tse_planilhas/lula_2022_mesorregioes_mg.xlsx

Uso:
  python3 scripts/lula_2022_mesorregioes_mg.py
  python lula_2022_mesorregioes_mg.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.discriminativo_interior_nordeste import normalizar_nome, pct
    from scripts.planilha_resultados_presidente import pasta_dados, pastas_saida
except ImportError:  # ContAgil
    from discriminativo_interior_nordeste import normalizar_nome, pct  # type: ignore
    from planilha_resultados_presidente import pasta_dados, pastas_saida  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[1]
CATALOGO_MESO = REPO_ROOT / "data" / "tse_catalog" / "mesorregioes_mg.csv"
FONTE_PDF = (
    "https://www.mg.gov.br/sites/default/files/paginas/arquivos/2016/"
    "ligminas_10_2_04_listamesomicro.pdf"
)

COLUNAS_MUN = [
    "codigo_meso",
    "nome_meso",
    "codigo_micro",
    "nome_micro",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "QT_SECOES",
    "QT_VOTOS_LULA",
    "QT_VOTOS_BOLSONARO",
    "QT_VOTOS_VALIDOS",
    "PCT_LULA",
    "PCT_BOLSONARO",
    "DIF_PP",
]


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconf = getattr(stream, "reconfigure", None)
        if reconf is None:
            continue
        try:
            reconf(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dados", type=Path, default=None)
    p.add_argument("--saida", type=Path, default=None)
    p.add_argument("--catalogo", type=Path, default=CATALOGO_MESO)
    return p.parse_args(argv)


def chave_nome(nome: object) -> str:
    texto = normalizar_nome(nome)
    for ch in ("-", "'", "´", "`"):
        texto = texto.replace(ch, " ")
    return " ".join(texto.split())


def resolver_catalogo(caminho: Path | None) -> Path:
    candidatos = [
        caminho,
        CATALOGO_MESO,
        Path.cwd() / "data" / "tse_catalog" / "mesorregioes_mg.csv",
        REPO_ROOT / "data" / "tse_catalog" / "mesorregioes_mg.csv",
    ]
    for cand in candidatos:
        if cand is not None and Path(cand).exists():
            return Path(cand)
    raise FileNotFoundError(
        "Falta data/tse_catalog/mesorregioes_mg.csv (mesorregiões IBGE de Minas Gerais)."
    )


def carregar_catalogo(caminho: Path) -> pd.DataFrame:
    cat = pd.read_csv(caminho)
    cat["codigo_tse"] = pd.to_numeric(cat["codigo_tse"], errors="coerce")
    cat["codigo_ibge"] = pd.to_numeric(cat["codigo_ibge"], errors="coerce")
    cat["codigo_meso"] = pd.to_numeric(cat["codigo_meso"], errors="coerce").astype("Int64")
    cat["codigo_micro"] = pd.to_numeric(cat["codigo_micro"], errors="coerce").astype("Int64")
    cat["nome_meso"] = cat["nome_meso"].astype(str).str.strip()
    cat["chave_nome"] = cat["nome_municipio"].map(chave_nome)
    return cat.dropna(subset=["codigo_tse", "codigo_meso"])


def _vazio_municipio() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "SG_UF",
            "CD_MUNICIPIO",
            "NM_MUNICIPIO",
            "QT_SECOES",
            "QT_VOTOS_LULA",
            "QT_VOTOS_BOLSONARO",
            "QT_VOTOS_VALIDOS",
            "PCT_LULA",
            "PCT_BOLSONARO",
            "DIF_PP",
            "VENCEDOR",
        ]
    )


def municipios_do_discriminativo(caminho: Path) -> pd.DataFrame:
    df = pd.read_csv(caminho)
    mg = df[df["SG_UF"].astype(str).str.upper() == "MG"].copy()
    if mg.empty:
        return _vazio_municipio()
    out = pd.DataFrame(
        {
            "SG_UF": "MG",
            "CD_MUNICIPIO": pd.to_numeric(mg["CD_MUNICIPIO"], errors="coerce"),
            "NM_MUNICIPIO": mg["NM_MUNICIPIO"],
            "QT_SECOES": pd.to_numeric(mg.get("QT_SECOES_2022"), errors="coerce"),
            "QT_VOTOS_LULA": pd.to_numeric(mg["QT_VOTOS_PT_2022"], errors="coerce"),
            "QT_VOTOS_BOLSONARO": pd.to_numeric(mg["QT_VOTOS_OPP_2022"], errors="coerce"),
            "QT_VOTOS_VALIDOS": pd.to_numeric(mg["QT_VOTOS_VALIDOS_2022"], errors="coerce"),
        }
    )
    return enriquecer_municipios(out)


def municipios_das_urnas(caminho: Path) -> pd.DataFrame:
    urnas = pd.read_csv(caminho, compression="gzip", low_memory=False)
    urnas["SG_UF"] = urnas["SG_UF"].astype(str).str.strip().str.upper()
    mg = urnas[urnas["SG_UF"] == "MG"].copy()
    for col in ("QT_VOTOS_LULA", "QT_VOTOS_BOLSONARO", "QT_VOTOS_VALIDOS", "CD_MUNICIPIO"):
        mg[col] = pd.to_numeric(mg[col], errors="coerce")
    g = mg.groupby(["CD_MUNICIPIO", "NM_MUNICIPIO"], as_index=False).agg(
        QT_SECOES=("CD_MUNICIPIO", "size"),
        QT_VOTOS_LULA=("QT_VOTOS_LULA", "sum"),
        QT_VOTOS_BOLSONARO=("QT_VOTOS_BOLSONARO", "sum"),
        QT_VOTOS_VALIDOS=("QT_VOTOS_VALIDOS", "sum"),
    )
    g["SG_UF"] = "MG"
    return enriquecer_municipios(g)


def enriquecer_municipios(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["PCT_LULA"] = [
        pct(l, v) for l, v in zip(out["QT_VOTOS_LULA"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_BOLSONARO"] = [
        pct(b, v) for b, v in zip(out["QT_VOTOS_BOLSONARO"], out["QT_VOTOS_VALIDOS"])
    ]
    out["DIF_PP"] = [
        None if a is None or b is None else round(float(a) - float(b), 2)
        for a, b in zip(out["PCT_LULA"], out["PCT_BOLSONARO"])
    ]
    out["VENCEDOR"] = [
        "Lula"
        if float(l or 0) > float(b or 0)
        else "Bolsonaro"
        if float(b or 0) > float(l or 0)
        else "Empate"
        for l, b in zip(out["QT_VOTOS_LULA"], out["QT_VOTOS_BOLSONARO"])
    ]
    return out


def carregar_municipios_mg(dados: Path) -> pd.DataFrame:
    disc = dados / "tse_planilhas" / "discriminativo_presidente_municipio_2t.csv"
    if not disc.exists():
        disc = REPO_ROOT / "output" / "tse_planilhas" / "discriminativo_presidente_municipio_2t.csv"
    if disc.exists():
        return municipios_do_discriminativo(disc)
    urnas = dados / "tse2022" / "urnas_2t_presidente.csv.gz"
    if not urnas.exists():
        urnas = REPO_ROOT / "output" / "tse2022" / "urnas_2t_presidente.csv.gz"
    if urnas.exists():
        return municipios_das_urnas(urnas)
    raise FileNotFoundError(
        "Falta o resultado municipal de 2022 2T "
        "(discriminativo_presidente_municipio_2t.csv ou urnas_2t_presidente.csv.gz)."
    )


def cruzar_mesorregiao(mun: pd.DataFrame, cat: pd.DataFrame) -> pd.DataFrame:
    base = mun.copy()
    base["CD_MUNICIPIO"] = pd.to_numeric(base["CD_MUNICIPIO"], errors="coerce")
    base["chave_nome"] = base["NM_MUNICIPIO"].map(chave_nome)
    cols_cat = [
        "codigo_tse",
        "codigo_ibge",
        "codigo_meso",
        "nome_meso",
        "codigo_micro",
        "nome_micro",
    ]
    out = base.merge(cat[cols_cat], left_on="CD_MUNICIPIO", right_on="codigo_tse", how="left")
    falta = out["nome_meso"].isna()
    if falta.any():
        por_nome = cat.drop_duplicates("chave_nome").set_index("chave_nome")
        for col in ("codigo_meso", "nome_meso", "codigo_micro", "nome_micro", "codigo_ibge"):
            out.loc[falta, col] = out.loc[falta, "chave_nome"].map(por_nome[col])
    return out


def municipios_lula(mun: pd.DataFrame, cat: pd.DataFrame) -> pd.DataFrame:
    cruzado = cruzar_mesorregiao(mun, cat)
    lula = cruzado[cruzado["VENCEDOR"] == "Lula"].copy()
    sem = lula[lula["nome_meso"].isna()]
    if not sem.empty:
        nomes = ", ".join(sem["NM_MUNICIPIO"].astype(str).head(8))
        raise ValueError(f"{len(sem)} município(s) sem mesorregião: {nomes}")
    return lula.sort_values(["codigo_meso", "nome_micro", "NM_MUNICIPIO"]).reset_index(drop=True)


def resumo_mesorregioes(lula: pd.DataFrame, cat: pd.DataFrame) -> pd.DataFrame:
    total_por_meso = (
        cat.groupby(["codigo_meso", "nome_meso"], as_index=False)
        .size()
        .rename(columns={"size": "Municípios na mesorregião"})
    )
    g = lula.groupby(["codigo_meso", "nome_meso"], as_index=False).agg(
        **{
            "Municípios com vitória de Lula": ("CD_MUNICIPIO", "nunique"),
            "Seções": ("QT_SECOES", "sum"),
            "Lula": ("QT_VOTOS_LULA", "sum"),
            "Bolsonaro": ("QT_VOTOS_BOLSONARO", "sum"),
            "Votos válidos": ("QT_VOTOS_VALIDOS", "sum"),
        }
    )
    out = total_por_meso.merge(g, on=["codigo_meso", "nome_meso"], how="left")
    for col in (
        "Municípios com vitória de Lula",
        "Seções",
        "Lula",
        "Bolsonaro",
        "Votos válidos",
    ):
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)
    out["% dos municípios da mesorregião"] = [
        pct(a, b)
        for a, b in zip(out["Municípios com vitória de Lula"], out["Municípios na mesorregião"])
    ]
    out["% Lula (válidos)"] = [
        pct(a, b) for a, b in zip(out["Lula"], out["Votos válidos"])
    ]
    out["% Bolsonaro (válidos)"] = [
        pct(a, b) for a, b in zip(out["Bolsonaro"], out["Votos válidos"])
    ]
    out["Diferença p.p. (Lula − Bolsonaro)"] = [
        None if a is None or b is None else round(float(a) - float(b), 2)
        for a, b in zip(out["% Lula (válidos)"], out["% Bolsonaro (válidos)"])
    ]
    out = out.sort_values("codigo_meso").reset_index(drop=True)
    out = out.rename(columns={"codigo_meso": "Código", "nome_meso": "Mesorregião"})
    return out


def _fmts(wb):
    header = {
        "bold": True,
        "bg_color": "#1F4E79",
        "font_color": "white",
        "border": 1,
        "valign": "vcenter",
        "align": "center",
        "text_wrap": True,
    }
    return {
        "title": wb.add_format(
            {"bold": True, "font_size": 13, "font_color": "white", "bg_color": "#1F4E79"}
        ),
        "header": wb.add_format(header),
        "label": wb.add_format(
            {
                "bold": True,
                "bg_color": "#1F4E79",
                "font_color": "white",
                "border": 1,
                "valign": "top",
            }
        ),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "text": wb.add_format({"border": 1}),
        "int": wb.add_format({"border": 1, "num_format": "#,##0"}),
        "num": wb.add_format({"border": 1, "num_format": "0.00"}),
        "pp": wb.add_format({"border": 1, "num_format": "+0.00;-0.00;0.00"}),
    }


def _escrever_df(ws, df: pd.DataFrame, fmts, titulo: str) -> None:
    ws.merge_range(0, 0, 0, max(len(df.columns) - 1, 0), titulo, fmts["title"])
    for c, col in enumerate(df.columns):
        ws.write(2, c, col, fmts["header"])
        ws.set_column(c, c, min(34, max(12, len(str(col)) + 2)))
    pp_cols = {c for c in df.columns if "Diferença" in c or c == "DIF_PP"}
    pct_cols = {c for c in df.columns if c.startswith("%") or c.startswith("PCT_")}
    for r in range(len(df)):
        row = df.iloc[r]
        for c, col in enumerate(df.columns):
            val = row[col]
            if val is None or (not isinstance(val, str) and pd.isna(val)):
                ws.write_blank(r + 3, c, None, fmts["text"])
            elif isinstance(val, bool):
                ws.write(r + 3, c, "S" if val else "N", fmts["text"])
            elif isinstance(val, (int,)) and not isinstance(val, bool):
                ws.write_number(r + 3, c, int(val), fmts["int"])
            elif isinstance(val, float):
                fmt = fmts["pp"] if col in pp_cols else fmts["num"] if col in pct_cols else fmts["num"]
                ws.write_number(r + 3, c, float(val), fmt)
            else:
                ws.write(r + 3, c, str(val), fmts["text"])
    if len(df):
        ws.autofilter(2, 0, 2 + len(df), len(df.columns) - 1)
    ws.freeze_panes(3, 2)
    ws.set_row(2, 30)


def _nome_aba(codigo: int, nome: str) -> str:
    limpo = (
        str(nome)
        .replace("/", "-")
        .replace("\\", "-")
        .replace("?", "")
        .replace("*", "")
        .replace("[", "")
        .replace("]", "")
        .replace(":", " ")
    )
    texto = f"{int(codigo):02d} {limpo}"
    return texto[:31].strip()


def gravar_xlsx(destino: Path, lula: pd.DataFrame, resumo: pd.DataFrame) -> Path:
    import xlsxwriter

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)
    n_lula = int(resumo["Municípios com vitória de Lula"].sum())
    n_mg = int(resumo["Municípios na mesorregião"].sum())

    leia = wb.add_worksheet("Leia-me")
    leia.set_column(0, 0, 28)
    leia.set_column(1, 1, 118)
    linhas = [
        ("Pleito", "Presidente 2022, 2º turno — Lula × Bolsonaro"),
        ("Universo", f"Municípios de Minas Gerais em que Lula venceu ({n_lula} de {n_mg})"),
        (
            "Mesorregiões",
            "12 mesorregiões do IBGE usadas pelo governo de Minas no documento "
            "ligminas_10_2_04_listamesomicro.pdf (MESO E MICRORREGIÕES DO IBGE).",
        ),
        ("Fonte da lista", FONTE_PDF),
        (
            "Catálogo",
            "data/tse_catalog/mesorregioes_mg.csv — DTB IBGE 2016, a mesma classificação do PDF.",
        ),
        (
            "Vitória",
            "Lula venceu o município se teve mais votos válidos que Bolsonaro.",
        ),
        ("Abas", "Resumo, Municipios e uma aba por mesorregião"),
    ]
    leia.write(0, 0, "Campo", fmts["header"])
    leia.write(0, 1, "Valor", fmts["header"])
    for i, (k, v) in enumerate(linhas, 1):
        leia.write(i, 0, k, fmts["label"])
        leia.write(i, 1, v, fmts["wrap"])
        leia.set_row(i, 28)

    _escrever_df(
        wb.add_worksheet("Resumo"),
        resumo,
        fmts,
        "Minas Gerais — municípios com vitória de Lula (2022 2T) por mesorregião",
    )
    mun_out = lula[COLUNAS_MUN].rename(
        columns={
            "codigo_meso": "Código meso",
            "nome_meso": "Mesorregião",
            "codigo_micro": "Código micro",
            "nome_micro": "Microrregião",
            "CD_MUNICIPIO": "Código TSE",
            "NM_MUNICIPIO": "Município",
            "QT_SECOES": "Seções",
            "QT_VOTOS_LULA": "Lula",
            "QT_VOTOS_BOLSONARO": "Bolsonaro",
            "QT_VOTOS_VALIDOS": "Votos válidos",
            "PCT_LULA": "% Lula",
            "PCT_BOLSONARO": "% Bolsonaro",
            "DIF_PP": "Diferença p.p.",
        }
    )
    _escrever_df(
        wb.add_worksheet("Municipios"),
        mun_out,
        fmts,
        "Municípios mineiros em que Lula venceu o 2º turno de 2022",
    )
    for codigo, bloco in lula.groupby("codigo_meso", sort=True):
        nome = str(bloco["nome_meso"].iloc[0])
        fatia = mun_out[mun_out["Mesorregião"] == nome]
        _escrever_df(
            wb.add_worksheet(_nome_aba(int(codigo), nome)),
            fatia,
            fmts,
            f"{nome} — {len(fatia)} município(s) com vitória de Lula",
        )
    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    cat = carregar_catalogo(resolver_catalogo(args.catalogo))
    mun = carregar_municipios_mg(dados)
    lula = municipios_lula(mun, cat)
    resumo = resumo_mesorregioes(lula, cat)
    destino = saida / "lula_2022_mesorregioes_mg.xlsx"
    gravar_xlsx(destino, lula, resumo)
    print(resumo.to_string(index=False))
    print(f"Municípios com Lula: {len(lula)}")
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
