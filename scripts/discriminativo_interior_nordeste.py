#!/usr/bin/env python3
"""Discriminativo Lula × Bolsonaro no interior do Nordeste (2022, 2º turno).

Interior = municípios das 9 UFs do Nordeste, excluídas as capitais.
Diferença em votos válidos e em votos totais (comparecimento = válidos
+ brancos + nulos).

Saída:
  output/tse_planilhas/discriminativo_interior_nordeste_lula_bolsonaro_2022.xlsx

Uso:
  python3 scripts/discriminativo_interior_nordeste.py
  python discriminativo_interior_nordeste.py
"""

from __future__ import annotations

import argparse
import sys
import unicodedata
from pathlib import Path

import pandas as pd

try:
    from scripts.planilha_resultados_presidente import (
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )
except ImportError:  # ContAgil
    from planilha_resultados_presidente import (  # type: ignore
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )

UFS_NORDESTE = ("AL", "BA", "CE", "MA", "PB", "PE", "PI", "RN", "SE")

# Capitais estaduais (nome TSE). Chave (UF, nome normalizado).
CAPITAIS_NE = {
    "AL": "MACEIO",
    "BA": "SALVADOR",
    "CE": "FORTALEZA",
    "MA": "SAO LUIS",
    "PB": "JOAO PESSOA",
    "PE": "RECIFE",
    "PI": "TERESINA",
    "RN": "NATAL",
    "SE": "ARACAJU",
}

COLUNAS_SAIDA = [
    ("SG_UF", "UF", 6, "text"),
    ("CD_MUNICIPIO", "Código TSE", 12, "int"),
    ("NM_MUNICIPIO", "Município", 32, "text"),
    ("QT_SECOES", "Seções", 10, "int"),
    ("QT_VOTOS_LULA", "Lula", 14, "int"),
    ("QT_VOTOS_BOLSONARO", "Bolsonaro", 14, "int"),
    ("QT_DIF_LULA_BOLSONARO", "Diferença (Lula − Bolsonaro)", 22, "int"),
    ("QT_VOTOS_VALIDOS", "Votos válidos", 14, "int"),
    ("PCT_LULA_VALIDOS", "% Lula (válidos)", 16, "pct"),
    ("PCT_BOLSONARO_VALIDOS", "% Bolsonaro (válidos)", 20, "pct"),
    ("DIF_PCT_VALIDOS", "Diferença p.p. (válidos)", 20, "pp"),
    ("QT_VOTOS_BRANCO", "Brancos", 12, "int"),
    ("QT_VOTOS_NULO", "Nulos", 10, "int"),
    ("QT_VOTOS_TOTAIS", "Votos totais", 14, "int"),
    ("PCT_LULA_TOTAIS", "% Lula (totais)", 16, "pct"),
    ("PCT_BOLSONARO_TOTAIS", "% Bolsonaro (totais)", 20, "pct"),
    ("DIF_PCT_TOTAIS", "Diferença p.p. (totais)", 20, "pp"),
    ("QT_APTOS", "Aptos", 12, "int"),
    ("QT_ABSTENCOES", "Abstenções", 12, "int"),
    ("VENCEDOR", "Vencedor", 12, "text"),
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
    return p.parse_args(argv)


def normalizar_nome(nome: object) -> str:
    texto = "" if nome is None or (isinstance(nome, float) and pd.isna(nome)) else str(nome)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return " ".join(texto.upper().split())


def eh_capital_nordeste(uf: object, nome: object) -> bool:
    sigla = "" if uf is None else str(uf).strip().upper()
    esperado = CAPITAIS_NE.get(sigla)
    return esperado is not None and normalizar_nome(nome) == esperado


def pct(parte: float, total: float) -> float | None:
    if total is None or pd.isna(total) or float(total) <= 0:
        return None
    return round(100.0 * float(parte) / float(total), 2)


def recorte_nordeste(df: pd.DataFrame) -> pd.DataFrame:
    base = df.copy()
    base["SG_UF"] = base["SG_UF"].astype(str).str.strip().str.upper()
    return base[base["SG_UF"].isin(UFS_NORDESTE)].copy()


def marcar_capital_interior(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["EH_CAPITAL"] = [
        eh_capital_nordeste(uf, nome)
        for uf, nome in zip(out["SG_UF"], out["NM_MUNICIPIO"])
    ]
    out["RECORTE"] = out["EH_CAPITAL"].map({True: "Capital", False: "Interior"})
    return out


def com_diferencas(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in (
        "QT_VOTOS_LULA",
        "QT_VOTOS_BOLSONARO",
        "QT_VOTOS_VALIDOS",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_COMPARECIMENTO",
        "QT_APTOS",
        "QT_ABSTENCOES",
    ):
        if col not in out.columns:
            out[col] = 0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)
    if "QT_COMPARECIMENTO" in out.columns:
        totais = out["QT_COMPARECIMENTO"]
        # Fallback se o comparecimento veio zerado mas há votos lançados.
        fallback = (
            out["QT_VOTOS_VALIDOS"] + out["QT_VOTOS_BRANCO"] + out["QT_VOTOS_NULO"]
        )
        totais = totais.where(totais > 0, fallback)
    else:
        totais = out["QT_VOTOS_VALIDOS"] + out["QT_VOTOS_BRANCO"] + out["QT_VOTOS_NULO"]
    out["QT_VOTOS_TOTAIS"] = totais
    out["QT_DIF_LULA_BOLSONARO"] = out["QT_VOTOS_LULA"] - out["QT_VOTOS_BOLSONARO"]
    out["PCT_LULA_VALIDOS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_LULA"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_BOLSONARO_VALIDOS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_BOLSONARO"], out["QT_VOTOS_VALIDOS"])
    ]
    out["DIF_PCT_VALIDOS"] = [
        pct(d, v)
        for d, v in zip(out["QT_DIF_LULA_BOLSONARO"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_LULA_TOTAIS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_LULA"], out["QT_VOTOS_TOTAIS"])
    ]
    out["PCT_BOLSONARO_TOTAIS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_BOLSONARO"], out["QT_VOTOS_TOTAIS"])
    ]
    out["DIF_PCT_TOTAIS"] = [
        pct(d, t)
        for d, t in zip(out["QT_DIF_LULA_BOLSONARO"], out["QT_VOTOS_TOTAIS"])
    ]
    if "VENCEDOR" not in out.columns:
        out["VENCEDOR"] = [
            "Lula"
            if lula > bolo
            else "Bolsonaro"
            if bolo > lula
            else "Empate"
            for lula, bolo in zip(out["QT_VOTOS_LULA"], out["QT_VOTOS_BOLSONARO"])
        ]
    else:
        out["VENCEDOR"] = (
            out["VENCEDOR"]
            .astype(str)
            .str.replace("LULA", "Lula", regex=False)
            .str.replace("BOLSONARO", "Bolsonaro", regex=False)
        )
    return out


def agregar_municipios(urnas: pd.DataFrame) -> pd.DataFrame:
    mun = agregar(
        urnas,
        ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"],
    )
    return com_diferencas(marcar_capital_interior(mun))


def agregar_ufs(municipios: pd.DataFrame) -> pd.DataFrame:
    soma = [
        "QT_SECOES",
        "QT_VOTOS_LULA",
        "QT_VOTOS_BOLSONARO",
        "QT_VOTOS_VALIDOS",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_VOTOS_TOTAIS",
        "QT_APTOS",
        "QT_ABSTENCOES",
        "QT_DIF_LULA_BOLSONARO",
    ]
    g = (
        municipios.groupby("SG_UF", dropna=False)[soma]
        .sum()
        .reset_index()
    )
    contagem = municipios.groupby("SG_UF", dropna=False)["CD_MUNICIPIO"].nunique()
    g["CD_MUNICIPIO"] = pd.NA
    def _rotulo_uf(uf: object) -> str:
        n = int(contagem.get(uf, 0))
        palavra = "município" if n == 1 else "municípios"
        return f"{n} {palavra} do interior"

    g["NM_MUNICIPIO"] = g["SG_UF"].map(_rotulo_uf)
    g = com_diferencas(g)
    return g.sort_values("SG_UF", kind="mergesort").reset_index(drop=True)


def linha_resumo(nome: str, recorte: pd.DataFrame) -> dict:
    base = {
        "Recorte": nome,
        "Municípios": int(recorte["CD_MUNICIPIO"].nunique())
        if "CD_MUNICIPIO" in recorte.columns
        else 0,
        "Seções": int(recorte["QT_SECOES"].sum()) if "QT_SECOES" in recorte.columns else 0,
        "Lula": int(recorte["QT_VOTOS_LULA"].sum()),
        "Bolsonaro": int(recorte["QT_VOTOS_BOLSONARO"].sum()),
        "Diferença (Lula − Bolsonaro)": int(recorte["QT_VOTOS_LULA"].sum())
        - int(recorte["QT_VOTOS_BOLSONARO"].sum()),
        "Votos válidos": int(recorte["QT_VOTOS_VALIDOS"].sum()),
        "Votos totais": int(recorte["QT_VOTOS_TOTAIS"].sum()),
    }
    base["% Lula (válidos)"] = pct(base["Lula"], base["Votos válidos"])
    base["% Bolsonaro (válidos)"] = pct(base["Bolsonaro"], base["Votos válidos"])
    base["Diferença p.p. (válidos)"] = pct(
        base["Diferença (Lula − Bolsonaro)"], base["Votos válidos"]
    )
    base["% Lula (totais)"] = pct(base["Lula"], base["Votos totais"])
    base["% Bolsonaro (totais)"] = pct(base["Bolsonaro"], base["Votos totais"])
    base["Diferença p.p. (totais)"] = pct(
        base["Diferença (Lula − Bolsonaro)"], base["Votos totais"]
    )
    return base


def montar_resumo(municipios: pd.DataFrame) -> pd.DataFrame:
    interior = municipios[~municipios["EH_CAPITAL"]]
    capitais = municipios[municipios["EH_CAPITAL"]]
    linhas = [
        linha_resumo("Interior do Nordeste", interior),
        linha_resumo("Capitais do Nordeste", capitais),
        linha_resumo("Nordeste (todos os municípios)", municipios),
    ]
    return pd.DataFrame(linhas)


def _fmt_livro(wb):
    return {
        "title": wb.add_format(
            {
                "bold": True,
                "font_size": 13,
                "font_color": "white",
                "bg_color": "#1F4E79",
                "valign": "vcenter",
            }
        ),
        "sub": wb.add_format(
            {
                "italic": True,
                "font_size": 10,
                "font_color": "#1F4E79",
                "valign": "vcenter",
            }
        ),
        "header": wb.add_format(
            {
                "bold": True,
                "bg_color": "#1F4E79",
                "font_color": "white",
                "border": 1,
                "valign": "vcenter",
                "align": "center",
                "text_wrap": True,
            }
        ),
        "text": wb.add_format({"border": 1, "valign": "vcenter"}),
        "int": wb.add_format(
            {"border": 1, "num_format": "#,##0", "valign": "vcenter"}
        ),
        "pct": wb.add_format(
            {"border": 1, "num_format": "0.00", "valign": "vcenter"}
        ),
        "pp": wb.add_format(
            {
                "border": 1,
                "num_format": "+0.00;-0.00;0.00",
                "valign": "vcenter",
            }
        ),
        "label": wb.add_format(
            {"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1}
        ),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
    }


def _escrever_tabela(ws, dados: pd.DataFrame, fmts, titulo: str, subtitulo: str) -> None:
    ncols = len(COLUNAS_SAIDA)
    ws.merge_range(0, 0, 0, ncols - 1, titulo, fmts["title"])
    ws.merge_range(1, 0, 1, ncols - 1, subtitulo, fmts["sub"])
    ws.set_row(0, 22)
    ws.set_row(1, 16)
    ws.set_row(2, 30)
    for c, (_, titulo_col, largura, _) in enumerate(COLUNAS_SAIDA):
        ws.write(2, c, titulo_col, fmts["header"])
        ws.set_column(c, c, largura)
    for r, rec in enumerate(dados.itertuples(index=False), 3):
        row = dados.iloc[r - 3]
        for c, (campo, _, _, tipo) in enumerate(COLUNAS_SAIDA):
            val = row[campo] if campo in row.index else None
            if val is None or pd.isna(val) or val == "":
                ws.write_blank(r, c, None, fmts["text"])
            elif tipo == "int":
                ws.write_number(r, c, int(val), fmts["int"])
            elif tipo in ("pct", "pp"):
                ws.write_number(r, c, float(val), fmts[tipo])
            else:
                ws.write_string(r, c, str(val), fmts["text"])
    last = 2 + len(dados)
    if len(dados):
        ws.autofilter(2, 0, last, ncols - 1)
    ws.freeze_panes(3, 3)


def gravar_xlsx(
    destino: Path,
    interior: pd.DataFrame,
    por_uf: pd.DataFrame,
    capitais: pd.DataFrame,
    resumo: pd.DataFrame,
    fonte: str,
) -> Path:
    import xlsxwriter

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmt_livro(wb)

    leia = wb.add_worksheet("Leia-me")
    leia.set_column(0, 0, 36)
    leia.set_column(1, 1, 92)
    linhas = [
        ("Pleito", "Presidente 2022, 2º turno — Lula × Bolsonaro"),
        ("Recorte", "Interior do Nordeste (AL BA CE MA PB PE PI RN SE, sem as capitais)"),
        (
            "Capitais excluídas",
            "Maceió, Salvador, Fortaleza, São Luís, João Pessoa, Recife, Teresina, Natal, Aracaju",
        ),
        ("Fonte", fonte),
        ("Municípios do interior", f"{len(interior):,}".replace(",", ".")),
        ("Critério de vitória", "Mais votos válidos no município"),
        (
            "Votos válidos",
            "Lula + Bolsonaro. % e diferença p.p. = (Lula − Bolsonaro) / válidos × 100",
        ),
        (
            "Votos totais",
            "Comparecimento (válidos + brancos + nulos). % e diferença p.p. = (Lula − Bolsonaro) / totais × 100",
        ),
        (
            "Sinal da diferença",
            "Positivo = vantagem de Lula; negativo = vantagem de Bolsonaro",
        ),
        ("Abas", "Municipios, Por_UF, Capitais, Resumo"),
    ]
    leia.write(0, 0, "Campo", fmts["header"])
    leia.write(0, 1, "Valor", fmts["header"])
    for i, (k, v) in enumerate(linhas, 1):
        leia.write(i, 0, k, fmts["label"])
        leia.write(i, 1, v, fmts["wrap"])
        leia.set_row(i, 18)
    leia.freeze_panes(1, 0)

    sub_int = (
        "Municípios do interior. Ordenados pela diferença em pontos percentuais dos votos válidos."
    )
    _escrever_tabela(
        wb.add_worksheet("Municipios"),
        interior.sort_values(
            ["DIF_PCT_VALIDOS", "NM_MUNICIPIO"],
            ascending=[False, True],
            kind="mergesort",
        ).reset_index(drop=True),
        fmts,
        f"Interior do Nordeste — {len(interior)} municípios — 2º turno 2022",
        sub_int,
    )
    _escrever_tabela(
        wb.add_worksheet("Por_UF"),
        por_uf,
        fmts,
        "Interior do Nordeste por UF — 2º turno 2022",
        "Soma só dos municípios do interior (capitais fora).",
    )
    _escrever_tabela(
        wb.add_worksheet("Capitais"),
        capitais.sort_values("SG_UF", kind="mergesort").reset_index(drop=True),
        fmts,
        "Capitais do Nordeste (excluídas do interior) — 2º turno 2022",
        "As 9 capitais, para comparação com o recorte do interior.",
    )

    res = wb.add_worksheet("Resumo")
    res.set_column(0, 0, 36)
    res.set_column(1, 20, 16)
    res.merge_range(
        0,
        0,
        0,
        len(resumo.columns) - 1,
        "Nordeste 2022 2º turno — interior × capitais × total",
        fmts["title"],
    )
    for c, col in enumerate(resumo.columns):
        res.write(2, c, col, fmts["header"])
    tipos_int = {
        "Municípios",
        "Seções",
        "Lula",
        "Bolsonaro",
        "Diferença (Lula − Bolsonaro)",
        "Votos válidos",
        "Votos totais",
    }
    tipos_pp = {"Diferença p.p. (válidos)", "Diferença p.p. (totais)"}
    tipos_pct = {
        "% Lula (válidos)",
        "% Bolsonaro (válidos)",
        "% Lula (totais)",
        "% Bolsonaro (totais)",
    }
    for r, row in enumerate(resumo.itertuples(index=False), 3):
        rec = resumo.iloc[r - 3]
        for c, col in enumerate(resumo.columns):
            val = rec[col]
            if col == "Recorte":
                res.write_string(r, c, str(val), fmts["text"])
            elif col in tipos_int:
                res.write_number(r, c, int(val), fmts["int"])
            elif col in tipos_pp:
                res.write_number(r, c, float(val), fmts["pp"])
            elif col in tipos_pct:
                res.write_number(r, c, float(val), fmts["pct"])
            else:
                res.write(r, c, val, fmts["text"])
    res.set_row(0, 22)
    res.set_row(2, 28)
    res.freeze_panes(3, 1)

    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    urnas, fonte = carregar_pleito(dados, 2022, 2)
    ne = recorte_nordeste(urnas)
    municipios = agregar_municipios(ne)
    n_capitais = int(municipios["EH_CAPITAL"].sum())
    if n_capitais != 9:
        raise RuntimeError(
            f"Esperadas 9 capitais do Nordeste; encontradas {n_capitais}: "
            + ", ".join(
                f"{r.SG_UF}-{r.NM_MUNICIPIO}"
                for r in municipios[municipios["EH_CAPITAL"]].itertuples()
            )
        )
    interior = municipios[~municipios["EH_CAPITAL"]].copy()
    capitais = municipios[municipios["EH_CAPITAL"]].copy()
    por_uf = agregar_ufs(interior)
    resumo = montar_resumo(municipios)
    destino = saida / "discriminativo_interior_nordeste_lula_bolsonaro_2022.xlsx"
    gravar_xlsx(destino, interior, por_uf, capitais, resumo, fonte)
    print(f"Municípios interior: {len(interior)}")
    print(f"Capitais: {len(capitais)}")
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
