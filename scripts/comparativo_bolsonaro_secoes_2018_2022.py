#!/usr/bin/env python3
"""Bolsonaro × Haddad (2018 2T) versus Bolsonaro × Lula (2022 2T), por seção.

Cruza cada seção eleitoral (UF + município TSE + zona + seção) nos dois
segundos turnos e compara a diferença percentual de Bolsonaro em relação
ao adversário do PT.

Saída:
  output/tse_planilhas/comparativo_bolsonaro_secoes_2018_2022.csv.gz
  output/tse_planilhas/comparativo_bolsonaro_secoes_2018_2022.xlsx

Uso:
  python3 scripts/comparativo_bolsonaro_secoes_2018_2022.py
  python comparativo_bolsonaro_secoes_2018_2022.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.discriminativo_interior_nordeste import pct
    from scripts.planilha_resultados_presidente import (
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )
except ImportError:  # ContAgil
    from discriminativo_interior_nordeste import pct  # type: ignore
    from planilha_resultados_presidente import (  # type: ignore
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )

REPO_ROOT = Path(__file__).resolve().parents[1]
CHAVES = ["SG_UF", "CD_MUNICIPIO", "NR_ZONA", "NR_SECAO"]
N_EXTREMOS = 500


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
    p.add_argument("--extremos", type=int, default=N_EXTREMOS)
    return p.parse_args(argv)


def _media(series: pd.Series) -> float | None:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return None
    return round(float(s.mean()), 2)


def _vencedor(bolso: float, adversario: float, nome_adv: str) -> str:
    if bolso > adversario:
        return "Bolsonaro"
    if adversario > bolso:
        return nome_adv
    return "Empate"


def preparar_ano(df: pd.DataFrame, *, ano: int, col_adv: str, nome_adv: str) -> pd.DataFrame:
    base = df.copy()
    base["SG_UF"] = base["SG_UF"].astype(str).str.strip().str.upper()
    for col in ("CD_MUNICIPIO", "NR_ZONA", "NR_SECAO"):
        base[col] = pd.to_numeric(base[col], errors="coerce")
    for col in ("QT_VOTOS_BOLSONARO", col_adv, "QT_VOTOS_VALIDOS"):
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)
    if "NM_MUNICIPIO" not in base.columns:
        base["NM_MUNICIPIO"] = ""
    base["NM_MUNICIPIO"] = base["NM_MUNICIPIO"].astype(str)
    out = pd.DataFrame(
        {
            "SG_UF": base["SG_UF"],
            "CD_MUNICIPIO": base["CD_MUNICIPIO"],
            "NM_MUNICIPIO": base["NM_MUNICIPIO"],
            "NR_ZONA": base["NR_ZONA"],
            "NR_SECAO": base["NR_SECAO"],
            "QT_VOTOS_BOLSONARO": base["QT_VOTOS_BOLSONARO"],
            "QT_VOTOS_ADV": base[col_adv],
            "QT_VOTOS_VALIDOS": base["QT_VOTOS_VALIDOS"],
        }
    )
    out["PCT_BOLSONARO"] = [
        pct(b, v) for b, v in zip(out["QT_VOTOS_BOLSONARO"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_ADV"] = [pct(a, v) for a, v in zip(out["QT_VOTOS_ADV"], out["QT_VOTOS_VALIDOS"])]
    out["DIF_BOLSO_ADV"] = [
        None if x is None or y is None else round(float(x) - float(y), 2)
        for x, y in zip(out["PCT_BOLSONARO"], out["PCT_ADV"])
    ]
    out["VENCEDOR"] = [
        _vencedor(float(b), float(a), nome_adv)
        for b, a in zip(out["QT_VOTOS_BOLSONARO"], out["QT_VOTOS_ADV"])
    ]
    return out.rename(
        columns={
            "NM_MUNICIPIO": f"NM_MUNICIPIO_{ano}",
            "QT_VOTOS_BOLSONARO": f"QT_VOTOS_BOLSONARO_{ano}",
            "QT_VOTOS_ADV": f"QT_VOTOS_{nome_adv.upper()}_{ano}",
            "QT_VOTOS_VALIDOS": f"QT_VOTOS_VALIDOS_{ano}",
            "PCT_BOLSONARO": f"PCT_BOLSONARO_{ano}",
            "PCT_ADV": f"PCT_{nome_adv.upper()}_{ano}",
            "DIF_BOLSO_ADV": f"DIF_BOLSO_{nome_adv.upper()}_{ano}",
            "VENCEDOR": f"VENCEDOR_{ano}",
        }
    )


def resolver_fonte(dados: Path, rels: list[Path]) -> Path:
    for rel in rels:
        for cand in (dados / rel, REPO_ROOT / "output" / rel):
            if cand.exists():
                return cand
    raise FileNotFoundError(" / ".join(str(r) for r in rels))


def carregar_anos(dados: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    f18 = resolver_fonte(
        dados,
        [
            Path("tse2018") / "secoes_2t_presidente.csv.gz",
            Path("tse2018") / "urnas_2t_presidente.csv.gz",
        ],
    )
    f22 = resolver_fonte(
        dados,
        [
            Path("tse2022") / "urnas_2t_presidente.csv.gz",
            Path("tse2022") / "secoes_2t_presidente.csv.gz",
        ],
    )
    a18 = preparar_ano(
        pd.read_csv(f18, compression="gzip", low_memory=False),
        ano=2018,
        col_adv="QT_VOTOS_HADDAD",
        nome_adv="Haddad",
    )
    a22 = preparar_ano(
        pd.read_csv(f22, compression="gzip", low_memory=False),
        ano=2022,
        col_adv="QT_VOTOS_LULA",
        nome_adv="Lula",
    )
    return a18, a22


def cruzar_secoes(a18: pd.DataFrame, a22: pd.DataFrame) -> pd.DataFrame:
    out = a18.merge(a22, on=CHAVES, how="outer")
    out["COMPARAVEL"] = (
        out["QT_VOTOS_VALIDOS_2018"].notna() & out["QT_VOTOS_VALIDOS_2022"].notna()
    ).map({True: "S", False: "N"})
    out["NM_MUNICIPIO"] = out["NM_MUNICIPIO_2022"].where(
        out["NM_MUNICIPIO_2022"].notna() & (out["NM_MUNICIPIO_2022"].astype(str) != ""),
        out["NM_MUNICIPIO_2018"],
    )
    out["REGIAO"] = out["SG_UF"].map(rotulo_regiao)
    out["DIF_PCT_BOLSONARO"] = [
        None if a is None or b is None or pd.isna(a) or pd.isna(b) else round(float(b) - float(a), 2)
        for a, b in zip(out["PCT_BOLSONARO_2018"], out["PCT_BOLSONARO_2022"])
    ]
    out["DIF_MARGEM_BOLSO"] = [
        None if a is None or b is None or pd.isna(a) or pd.isna(b) else round(float(b) - float(a), 2)
        for a, b in zip(out["DIF_BOLSO_HADDAD_2018"], out["DIF_BOLSO_LULA_2022"])
    ]
    def _lado(nome: object) -> str:
        if nome is None or (isinstance(nome, float) and pd.isna(nome)) or str(nome) == "":
            return ""
        texto = str(nome)
        if texto == "Bolsonaro":
            return "Bolsonaro"
        if texto == "Empate":
            return "Empate"
        return "PT"

    out["INVERTEU"] = [
        "S"
        if _lado(x) and _lado(y) and _lado(x) != _lado(y)
        else "N"
        if _lado(x) and _lado(y)
        else ""
        for x, y in zip(out["VENCEDOR_2018"], out["VENCEDOR_2022"])
    ]
    return out.sort_values(CHAVES).reset_index(drop=True)


def comparaveis(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["COMPARAVEL"] == "S"].copy()


def resumo_recorte(nome: str, bloco: pd.DataFrame) -> dict[str, object]:
    par = comparaveis(bloco)
    n18 = int(bloco["QT_VOTOS_VALIDOS_2018"].notna().sum())
    n22 = int(bloco["QT_VOTOS_VALIDOS_2022"].notna().sum())
    pond18 = None
    pond22 = None
    if not par.empty:
        v18 = float(par["QT_VOTOS_VALIDOS_2018"].sum())
        v22 = float(par["QT_VOTOS_VALIDOS_2022"].sum())
        pond18 = pct(
            par["QT_VOTOS_BOLSONARO_2018"].sum() - par["QT_VOTOS_HADDAD_2018"].sum(),
            v18,
        )
        pond22 = pct(
            par["QT_VOTOS_BOLSONARO_2022"].sum() - par["QT_VOTOS_LULA_2022"].sum(),
            v22,
        )
    r_bolso = None
    if len(par) >= 3:
        r = pd.to_numeric(par["PCT_BOLSONARO_2018"], errors="coerce").corr(
            pd.to_numeric(par["PCT_BOLSONARO_2022"], errors="coerce")
        )
        r_bolso = None if pd.isna(r) else round(float(r), 3)
    return {
        "Recorte": nome,
        "Seções 2018": n18,
        "Seções 2022": n22,
        "Seções comparáveis": len(par),
        "Só em 2018": n18 - len(par),
        "Só em 2022": n22 - len(par),
        "Bolsonaro − Haddad (média das seções, p.p.)": _media(par["DIF_BOLSO_HADDAD_2018"]),
        "Bolsonaro − Lula (média das seções, p.p.)": _media(par["DIF_BOLSO_LULA_2022"]),
        "Variação da margem (média, p.p.)": _media(par["DIF_MARGEM_BOLSO"]),
        "Bolsonaro − Haddad (ponderada, p.p.)": pond18,
        "Bolsonaro − Lula (ponderada, p.p.)": pond22,
        "Variação da margem (ponderada, p.p.)": (
            None if pond18 is None or pond22 is None else round(pond22 - pond18, 2)
        ),
        "% Bolsonaro 2018 (média)": _media(par["PCT_BOLSONARO_2018"]),
        "% Bolsonaro 2022 (média)": _media(par["PCT_BOLSONARO_2022"]),
        "Variação % Bolsonaro (média, p.p.)": _media(par["DIF_PCT_BOLSONARO"]),
        "Vitórias Bolsonaro 2018": int((par["VENCEDOR_2018"] == "Bolsonaro").sum()),
        "Vitórias Bolsonaro 2022": int((par["VENCEDOR_2022"] == "Bolsonaro").sum()),
        "Inverteram vencedor": int((par["INVERTEU"] == "S").sum()),
        "Pearson % Bolsonaro 2018×2022": r_bolso,
    }


def tabela_resumo(df: pd.DataFrame) -> pd.DataFrame:
    linhas = [resumo_recorte("Brasil", df)]
    for reg, g in df.groupby("REGIAO", dropna=False):
        linhas.append(resumo_recorte(str(reg), g))
    return pd.DataFrame(linhas)


def tabela_uf(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for uf, g in df.groupby("SG_UF", dropna=False):
        rec = resumo_recorte(str(uf), g)
        rec["UF"] = uf
        rec["Região"] = g["REGIAO"].iloc[0] if len(g) else ""
        linhas.append(rec)
    out = pd.DataFrame(linhas)
    cols = ["UF", "Região"] + [c for c in out.columns if c not in {"UF", "Região", "Recorte"}]
    return out[cols].sort_values("UF").reset_index(drop=True)


def tabela_municipio(df: pd.DataFrame) -> pd.DataFrame:
    par = comparaveis(df)
    if par.empty:
        return pd.DataFrame()
    g = par.groupby(["REGIAO", "SG_UF", "CD_MUNICIPIO"], dropna=False).agg(
        NM_MUNICIPIO=("NM_MUNICIPIO", "first"),
        QT_SECOES=("NR_SECAO", "size"),
        QT_VOTOS_BOLSONARO_2018=("QT_VOTOS_BOLSONARO_2018", "sum"),
        QT_VOTOS_HADDAD_2018=("QT_VOTOS_HADDAD_2018", "sum"),
        QT_VOTOS_VALIDOS_2018=("QT_VOTOS_VALIDOS_2018", "sum"),
        QT_VOTOS_BOLSONARO_2022=("QT_VOTOS_BOLSONARO_2022", "sum"),
        QT_VOTOS_LULA_2022=("QT_VOTOS_LULA_2022", "sum"),
        QT_VOTOS_VALIDOS_2022=("QT_VOTOS_VALIDOS_2022", "sum"),
        MEDIA_DIF_2018=("DIF_BOLSO_HADDAD_2018", "mean"),
        MEDIA_DIF_2022=("DIF_BOLSO_LULA_2022", "mean"),
        MEDIA_DIF_MARGEM=("DIF_MARGEM_BOLSO", "mean"),
        INVERTERAM=("INVERTEU", lambda s: int((s == "S").sum())),
    ).reset_index()
    g["DIF_BOLSO_HADDAD_2018"] = [
        pct(b - h, v)
        for b, h, v in zip(
            g["QT_VOTOS_BOLSONARO_2018"], g["QT_VOTOS_HADDAD_2018"], g["QT_VOTOS_VALIDOS_2018"]
        )
    ]
    g["DIF_BOLSO_LULA_2022"] = [
        pct(b - l, v)
        for b, l, v in zip(
            g["QT_VOTOS_BOLSONARO_2022"], g["QT_VOTOS_LULA_2022"], g["QT_VOTOS_VALIDOS_2022"]
        )
    ]
    g["DIF_MARGEM_BOLSO"] = [
        None if a is None or b is None else round(float(b) - float(a), 2)
        for a, b in zip(g["DIF_BOLSO_HADDAD_2018"], g["DIF_BOLSO_LULA_2022"])
    ]
    g["MEDIA_DIF_2018"] = g["MEDIA_DIF_2018"].round(2)
    g["MEDIA_DIF_2022"] = g["MEDIA_DIF_2022"].round(2)
    g["MEDIA_DIF_MARGEM"] = g["MEDIA_DIF_MARGEM"].round(2)
    return g.sort_values(["SG_UF", "NM_MUNICIPIO"]).reset_index(drop=True)


def tabela_cobertura(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for uf, g in df.groupby("SG_UF", dropna=False):
        n18 = int(g["QT_VOTOS_VALIDOS_2018"].notna().sum())
        n22 = int(g["QT_VOTOS_VALIDOS_2022"].notna().sum())
        npar = int((g["COMPARAVEL"] == "S").sum())
        linhas.append(
            {
                "UF": uf,
                "Região": g["REGIAO"].iloc[0],
                "Seções 2018": n18,
                "Seções 2022": n22,
                "Comparáveis": npar,
                "Só 2018": n18 - npar,
                "Só 2022": n22 - npar,
                "% comparáveis sobre 2018": pct(npar, n18),
                "% comparáveis sobre 2022": pct(npar, n22),
            }
        )
    return pd.DataFrame(linhas).sort_values("UF").reset_index(drop=True)


def extremos(df: pd.DataFrame, n: int, *, ganhos: bool) -> pd.DataFrame:
    par = comparaveis(df)
    par = par[par["DIF_MARGEM_BOLSO"].notna()]
    if par.empty:
        return par
    cols = [
        "REGIAO",
        "SG_UF",
        "NM_MUNICIPIO",
        "CD_MUNICIPIO",
        "NR_ZONA",
        "NR_SECAO",
        "DIF_BOLSO_HADDAD_2018",
        "DIF_BOLSO_LULA_2022",
        "DIF_MARGEM_BOLSO",
        "PCT_BOLSONARO_2018",
        "PCT_BOLSONARO_2022",
        "VENCEDOR_2018",
        "VENCEDOR_2022",
        "INVERTEU",
    ]
    ordem = par.sort_values("DIF_MARGEM_BOLSO", ascending=not ganhos)
    return ordem[cols].head(n).reset_index(drop=True)


COLUNAS_CSV = [
    "REGIAO",
    "SG_UF",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "COMPARAVEL",
    "QT_VOTOS_BOLSONARO_2018",
    "QT_VOTOS_HADDAD_2018",
    "QT_VOTOS_VALIDOS_2018",
    "PCT_BOLSONARO_2018",
    "PCT_HADDAD_2018",
    "DIF_BOLSO_HADDAD_2018",
    "VENCEDOR_2018",
    "QT_VOTOS_BOLSONARO_2022",
    "QT_VOTOS_LULA_2022",
    "QT_VOTOS_VALIDOS_2022",
    "PCT_BOLSONARO_2022",
    "PCT_LULA_2022",
    "DIF_BOLSO_LULA_2022",
    "VENCEDOR_2022",
    "DIF_PCT_BOLSONARO",
    "DIF_MARGEM_BOLSO",
    "INVERTEU",
]


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
            {"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1, "valign": "top"}
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
    pp_cols = {
        c
        for c in df.columns
        if "Variação" in c or "Diferença" in c or c.startswith("DIF_") or "margem" in c
    }
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
                ws.write_number(r + 3, c, float(val), fmts["pp"] if col in pp_cols else fmts["num"])
            else:
                ws.write(r + 3, c, str(val), fmts["text"])
    if len(df):
        ws.autofilter(2, 0, 2 + len(df), len(df.columns) - 1)
    ws.freeze_panes(3, 2)
    ws.set_row(2, 30)


def gravar_xlsx(
    destino: Path,
    df: pd.DataFrame,
    *,
    n_extremos: int,
) -> Path:
    import xlsxwriter

    destino.parent.mkdir(parents=True, exist_ok=True)
    resumo = tabela_resumo(df)
    ufs = tabela_uf(df)
    mun = tabela_municipio(df)
    cob = tabela_cobertura(df)
    br = resumo[resumo["Recorte"] == "Brasil"].iloc[0]
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)

    leia = wb.add_worksheet("Leia-me")
    leia.set_column(0, 0, 32)
    leia.set_column(1, 1, 118)
    linhas = [
        ("Pleitos", "2º turno 2018 (Bolsonaro × Haddad) e 2º turno 2022 (Bolsonaro × Lula)"),
        ("Chave da seção", "SG_UF + CD_MUNICIPIO + NR_ZONA + NR_SECAO"),
        (
            "Seções",
            f"2018: {br['Seções 2018']:,} · 2022: {br['Seções 2022']:,} · "
            f"comparáveis: {br['Seções comparáveis']:,} · "
            f"só 2018: {br['Só em 2018']:,} · só 2022: {br['Só em 2022']:,}".replace(",", "."),
        ),
        (
            "Diferença em relação ao adversário",
            "% Bolsonaro − % Haddad (2018) e % Bolsonaro − % Lula (2022), nos votos válidos. "
            "Positivo = vantagem de Bolsonaro naquela seção.",
        ),
        (
            "Variação da margem",
            "(Bolsonaro − Lula, 2022) − (Bolsonaro − Haddad, 2018). "
            "Negativo = Bolsonaro piorou frente ao PT na mesma seção.",
        ),
        (
            "Brasil — média das seções",
            f"2018: {br['Bolsonaro − Haddad (média das seções, p.p.)']} p.p. · "
            f"2022: {br['Bolsonaro − Lula (média das seções, p.p.)']} p.p. · "
            f"Δ {br['Variação da margem (média, p.p.)']} p.p.",
        ),
        (
            "Brasil — ponderada por votos",
            f"2018: {br['Bolsonaro − Haddad (ponderada, p.p.)']} p.p. · "
            f"2022: {br['Bolsonaro − Lula (ponderada, p.p.)']} p.p. · "
            f"Δ {br['Variação da margem (ponderada, p.p.)']} p.p.",
        ),
        (
            "Arquivo completo",
            "comparativo_bolsonaro_secoes_2018_2022.csv.gz — uma linha por seção "
            "(inclui seções que só existiram em um dos anos).",
        ),
        ("Abas", "Resumo, Por_UF, Por_Municipio, Cobertura, Maiores_ganhos, Maiores_perdas"),
    ]
    leia.write(0, 0, "Campo", fmts["header"])
    leia.write(0, 1, "Valor", fmts["header"])
    for i, (k, v) in enumerate(linhas, 1):
        leia.write(i, 0, k, fmts["label"])
        leia.write(i, 1, v, fmts["wrap"])
        leia.set_row(i, 30)

    _escrever_df(wb.add_worksheet("Resumo"), resumo, fmts, "Bolsonaro vs PT — média por seção, 2018 × 2022")
    _escrever_df(wb.add_worksheet("Por_UF"), ufs, fmts, "Mesmas médias por UF")
    _escrever_df(
        wb.add_worksheet("Por_Municipio"),
        mun,
        fmts,
        "Município — margem ponderada e média das seções comparáveis",
    )
    _escrever_df(wb.add_worksheet("Cobertura"), cob, fmts, "Quantas seções de cada UF existem nos dois anos")
    _escrever_df(
        wb.add_worksheet("Maiores_ganhos"),
        extremos(df, n_extremos, ganhos=True),
        fmts,
        f"{n_extremos} seções em que a margem de Bolsonaro mais subiu (2022 − 2018)",
    )
    _escrever_df(
        wb.add_worksheet("Maiores_perdas"),
        extremos(df, n_extremos, ganhos=False),
        fmts,
        f"{n_extremos} seções em que a margem de Bolsonaro mais caiu (2022 − 2018)",
    )
    wb.close()
    return destino


def gravar_csv_gz(df: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    cols = [c for c in COLUNAS_CSV if c in df.columns]
    df[cols].to_csv(destino, index=False, encoding="utf-8", compression="gzip")
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    a18, a22 = carregar_anos(dados)
    df = cruzar_secoes(a18, a22)
    csv = gravar_csv_gz(df, saida / "comparativo_bolsonaro_secoes_2018_2022.csv.gz")
    xlsx = gravar_xlsx(saida / "comparativo_bolsonaro_secoes_2018_2022.xlsx", df, n_extremos=int(args.extremos))
    resumo = tabela_resumo(df)
    print(resumo.to_string(index=False))
    print(f"CSV: {csv} ({csv.stat().st_size / 1_048_576:.1f} MB)")
    print(f"Workbook: {xlsx} ({xlsx.stat().st_size / 1_048_576:.1f} MB)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
