#!/usr/bin/env python3
"""Diferença Lula × Bolsonaro em UE2020 vs urnas anteriores — municípios até 50 mil hab.

2º turno 2022. Em cada município com população IBGE 2022 ≤ 50.000, apura a
margem média (pontos percentuais dos votos válidos) nas urnas favoráveis a
Lula e nas favoráveis a Bolsonaro, separando modelo 2020 e modelos anteriores.
Depois compara nova × antiga para cada candidato.

Saída:
  output/tse_planilhas/discriminativo_modelo_urna_ate_50mil.xlsx

Uso:
  python3 scripts/discriminativo_modelo_urna_ate_50mil.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.discriminativo_urnas_municipio import (
        GERACAO_ANTERIOR,
        GERACAO_UE2020,
        classificar_geracao,
        preparar_urnas,
        vencedor_votos,
    )
    from scripts.planilha_resultados_presidente import (
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )
except ImportError:  # ContAgil
    from discriminativo_urnas_municipio import (  # type: ignore
        GERACAO_ANTERIOR,
        GERACAO_UE2020,
        classificar_geracao,
        preparar_urnas,
        vencedor_votos,
    )
    from planilha_resultados_presidente import (  # type: ignore
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )

REPO_ROOT = Path(__file__).resolve().parents[1]
LIMITE_HAB = 50_000
CATALOGO_POP = REPO_ROOT / "data" / "tse_catalog" / "populacao_censo2022_tse.csv"


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
    p.add_argument("--limite-hab", type=int, default=LIMITE_HAB)
    p.add_argument("--catalogo-pop", type=Path, default=CATALOGO_POP)
    return p.parse_args(argv)


def pct(parte: float, total: float) -> float | None:
    if total is None or pd.isna(total) or float(total) <= 0:
        return None
    return round(100.0 * float(parte) / float(total), 2)


def resolver_catalogo(caminho: Path | None) -> Path:
    candidatos = [
        caminho,
        CATALOGO_POP,
        Path.cwd() / "data" / "tse_catalog" / "populacao_censo2022_tse.csv",
        Path(__file__).resolve().parents[1] / "data" / "tse_catalog" / "populacao_censo2022_tse.csv",
    ]
    for cand in candidatos:
        if cand is not None and Path(cand).exists():
            return Path(cand)
    raise FileNotFoundError(
        "Falta data/tse_catalog/populacao_censo2022_tse.csv (Censo 2022 × código TSE)."
    )


def carregar_populacao(caminho: Path) -> pd.DataFrame:
    pop = pd.read_csv(caminho)
    pop["codigo_tse"] = pd.to_numeric(pop["codigo_tse"], errors="coerce")
    pop["POP_2022"] = pd.to_numeric(pop["POP_2022"], errors="coerce")
    pop["uf"] = pop["uf"].astype(str).str.strip().str.upper()
    return pop.dropna(subset=["codigo_tse", "POP_2022"])


def enriquecer_urnas(df: pd.DataFrame) -> pd.DataFrame:
    base = preparar_urnas(df)
    base["SG_UF"] = base["SG_UF"].astype(str).str.strip().str.upper()
    base["CD_MUNICIPIO"] = pd.to_numeric(base["CD_MUNICIPIO"], errors="coerce")
    for col in ("QT_VOTOS_LULA", "QT_VOTOS_BOLSONARO", "QT_VOTOS_VALIDOS"):
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)
    validos = base["QT_VOTOS_VALIDOS"]
    base["DIF_PP"] = [
        pct(l - b, v)
        for l, b, v in zip(base["QT_VOTOS_LULA"], base["QT_VOTOS_BOLSONARO"], validos)
    ]
    base["MARGEM_VENCEDOR"] = [
        None
        if d is None
        else abs(d)
        if venc in ("Lula", "Bolsonaro")
        else None
        for d, venc in zip(base["DIF_PP"], base["VENCEDOR_URNA"])
    ]
    return base


def _media(series: pd.Series) -> float | None:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return None
    return round(float(s.mean()), 2)


def _bloco_favorecido(grupo: pd.DataFrame, candidato: str) -> dict[str, object]:
    fatia = grupo[grupo["VENCEDOR_URNA"] == candidato]
    n = int(len(fatia))
    lula = float(fatia["QT_VOTOS_LULA"].sum()) if n else 0.0
    bolo = float(fatia["QT_VOTOS_BOLSONARO"].sum()) if n else 0.0
    validos = float(fatia["QT_VOTOS_VALIDOS"].sum()) if n else 0.0
    if candidato == "Lula":
        margem_pond = pct(lula - bolo, validos)
    else:
        margem_pond = pct(bolo - lula, validos)
    return {
        "QT_URNAS": n,
        "MEDIA_MARGEM": _media(fatia["MARGEM_VENCEDOR"]),
        "MARGEM_PONDERADA": margem_pond,
        "QT_VOTOS_VALIDOS": int(validos),
    }


def apurar_municipios(urnas: pd.DataFrame, pop: pd.DataFrame, limite: int) -> pd.DataFrame:
    recorte = pop.loc[pop["POP_2022"] <= limite, ["codigo_tse", "POP_2022"]].rename(
        columns={"codigo_tse": "CD_MUNICIPIO"}
    )
    base = urnas.merge(recorte, on="CD_MUNICIPIO", how="inner")
    base = base[base["SG_UF"] != "ZZ"].copy()
    base = base[base["GERACAO"].isin((GERACAO_ANTERIOR, GERACAO_UE2020))].copy()
    base = base[base["QT_VOTOS_VALIDOS"] > 0].copy()

    linhas = []
    chaves = ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"]
    for ch, bloco in base.groupby(chaves, dropna=False):
        meta = dict(zip(chaves, ch))
        meta["POP_2022"] = int(bloco["POP_2022"].iloc[0])
        meta["REGIAO"] = rotulo_regiao(meta["SG_UF"])
        for geracao, sufixo in (
            (GERACAO_ANTERIOR, "PRE2020"),
            (GERACAO_UE2020, "UE2020"),
        ):
            fatia = bloco[bloco["GERACAO"] == geracao]
            meta[f"QT_URNAS_{sufixo}"] = int(len(fatia))
            fav_l = _bloco_favorecido(fatia, "Lula")
            fav_b = _bloco_favorecido(fatia, "Bolsonaro")
            meta[f"QT_URNAS_LULA_{sufixo}"] = fav_l["QT_URNAS"]
            meta[f"MARGEM_LULA_{sufixo}"] = fav_l["MEDIA_MARGEM"]
            meta[f"MARGEM_LULA_POND_{sufixo}"] = fav_l["MARGEM_PONDERADA"]
            meta[f"QT_URNAS_BOLSO_{sufixo}"] = fav_b["QT_URNAS"]
            meta[f"MARGEM_BOLSO_{sufixo}"] = fav_b["MEDIA_MARGEM"]
            meta[f"MARGEM_BOLSO_POND_{sufixo}"] = fav_b["MARGEM_PONDERADA"]
        tem_pre = meta["QT_URNAS_PRE2020"] > 0
        tem_ue = meta["QT_URNAS_UE2020"] > 0
        meta["COMPARAVEL"] = "S" if tem_pre and tem_ue else "N"
        meta["COMPARA_LULA"] = (
            "S"
            if meta["QT_URNAS_LULA_PRE2020"] > 0 and meta["QT_URNAS_LULA_UE2020"] > 0
            else "N"
        )
        meta["COMPARA_BOLSO"] = (
            "S"
            if meta["QT_URNAS_BOLSO_PRE2020"] > 0 and meta["QT_URNAS_BOLSO_UE2020"] > 0
            else "N"
        )
        if meta["COMPARA_LULA"] == "S":
            meta["DIF_MARGEM_LULA_NOVA_ANTIGA"] = round(
                float(meta["MARGEM_LULA_UE2020"]) - float(meta["MARGEM_LULA_PRE2020"]), 2
            )
        else:
            meta["DIF_MARGEM_LULA_NOVA_ANTIGA"] = None
        if meta["COMPARA_BOLSO"] == "S":
            meta["DIF_MARGEM_BOLSO_NOVA_ANTIGA"] = round(
                float(meta["MARGEM_BOLSO_UE2020"]) - float(meta["MARGEM_BOLSO_PRE2020"]), 2
            )
        else:
            meta["DIF_MARGEM_BOLSO_NOVA_ANTIGA"] = None
        linhas.append(meta)
    out = pd.DataFrame(linhas)
    return out.sort_values(["SG_UF", "NM_MUNICIPIO"]).reset_index(drop=True)


def media_col(df: pd.DataFrame, col: str) -> float | None:
    return _media(df[col])


def resumo_recorte(nome: str, mun: pd.DataFrame) -> dict[str, object]:
    lula = mun[mun["COMPARA_LULA"] == "S"]
    bolso = mun[mun["COMPARA_BOLSO"] == "S"]
    lula_pre = mun[mun["MARGEM_LULA_PRE2020"].notna()]
    lula_ue = mun[mun["MARGEM_LULA_UE2020"].notna()]
    bolso_pre = mun[mun["MARGEM_BOLSO_PRE2020"].notna()]
    bolso_ue = mun[mun["MARGEM_BOLSO_UE2020"].notna()]
    media_lula_pre = media_col(lula_pre, "MARGEM_LULA_PRE2020")
    media_lula_ue = media_col(lula_ue, "MARGEM_LULA_UE2020")
    media_bolso_pre = media_col(bolso_pre, "MARGEM_BOLSO_PRE2020")
    media_bolso_ue = media_col(bolso_ue, "MARGEM_BOLSO_UE2020")
    return {
        "Recorte": nome,
        "Municípios ≤50 mil no recorte": len(mun),
        "Com os dois modelos": int((mun["COMPARAVEL"] == "S").sum()),
        "Municípios com urna Lula antiga": len(lula_pre),
        "Municípios com urna Lula UE2020": len(lula_ue),
        "Margem média Lula — urnas antigas (todos)": media_lula_pre,
        "Margem média Lula — UE2020 (todos)": media_lula_ue,
        "Diferença Lula não pareada (nova − antiga)": (
            None
            if media_lula_pre is None or media_lula_ue is None
            else round(media_lula_ue - media_lula_pre, 2)
        ),
        "Comparam Lula nos dois modelos": len(lula),
        "Margem média Lula — urnas antigas (pareada)": media_col(lula, "MARGEM_LULA_PRE2020"),
        "Margem média Lula — UE2020 (pareada)": media_col(lula, "MARGEM_LULA_UE2020"),
        "Diferença Lula pareada (nova − antiga)": media_col(lula, "DIF_MARGEM_LULA_NOVA_ANTIGA"),
        "Municípios com urna Bolsonaro antiga": len(bolso_pre),
        "Municípios com urna Bolsonaro UE2020": len(bolso_ue),
        "Margem média Bolsonaro — urnas antigas (todos)": media_bolso_pre,
        "Margem média Bolsonaro — UE2020 (todos)": media_bolso_ue,
        "Diferença Bolsonaro não pareada (nova − antiga)": (
            None
            if media_bolso_pre is None or media_bolso_ue is None
            else round(media_bolso_ue - media_bolso_pre, 2)
        ),
        "Comparam Bolsonaro nos dois modelos": len(bolso),
        "Margem média Bolsonaro — urnas antigas (pareada)": media_col(
            bolso, "MARGEM_BOLSO_PRE2020"
        ),
        "Margem média Bolsonaro — UE2020 (pareada)": media_col(bolso, "MARGEM_BOLSO_UE2020"),
        "Diferença Bolsonaro pareada (nova − antiga)": media_col(
            bolso, "DIF_MARGEM_BOLSO_NOVA_ANTIGA"
        ),
    }


def tabela_resumo(mun: pd.DataFrame) -> pd.DataFrame:
    linhas = [resumo_recorte("Brasil", mun)]
    for reg, g in mun.groupby("REGIAO"):
        linhas.append(resumo_recorte(str(reg), g))
    return pd.DataFrame(linhas)


def tabela_uf(mun: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for uf, g in mun.groupby("SG_UF"):
        rec = resumo_recorte(str(uf), g)
        rec["UF"] = uf
        rec["Região"] = g["REGIAO"].iloc[0]
        linhas.append(rec)
    out = pd.DataFrame(linhas)
    cols = ["UF", "Região"] + [c for c in out.columns if c not in {"UF", "Região", "Recorte"}]
    return out[cols].sort_values("UF").reset_index(drop=True)


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
        ws.set_column(c, c, min(28, max(12, len(str(col)) + 2)))
    pp_cols = {c for c in df.columns if "Diferença" in c or c.startswith("DIF_")}
    for r, rec in enumerate(df.itertuples(index=False), 3):
        row = df.iloc[r - 3]
        for c, col in enumerate(df.columns):
            val = row[col]
            if val is None or (not isinstance(val, str) and pd.isna(val)):
                ws.write_blank(r, c, None, fmts["text"])
            elif isinstance(val, bool):
                ws.write(r, c, "S" if val else "N", fmts["text"])
            elif isinstance(val, (int,)) and not isinstance(val, bool):
                ws.write_number(r, c, int(val), fmts["int"])
            elif isinstance(val, float):
                ws.write_number(r, c, float(val), fmts["pp"] if col in pp_cols else fmts["num"])
            else:
                ws.write(r, c, str(val), fmts["text"])
    if len(df):
        ws.autofilter(2, 0, 2 + len(df), len(df.columns) - 1)
    ws.freeze_panes(3, 3 if "NM_MUNICIPIO" in df.columns else 1)
    ws.set_row(2, 30)


def gravar_xlsx(destino: Path, mun: pd.DataFrame, limite: int) -> Path:
    import xlsxwriter

    resumo = tabela_resumo(mun)
    ufs = tabela_uf(mun)
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)
    br = resumo[resumo["Recorte"] == "Brasil"].iloc[0]

    leia = wb.add_worksheet("Leia-me")
    leia.set_column(0, 0, 36)
    leia.set_column(1, 1, 110)
    linhas = [
        ("Pleito", "Presidente 2022, 2º turno — Lula × Bolsonaro"),
        ("Universo", f"Municípios com até {limite:,} habitantes no Censo IBGE 2022 (prévia)".replace(",", ".")),
        ("Cruzamento", "Código TSE × código IBGE (data/tse_catalog/populacao_censo2022_tse.csv)"),
        ("Urna nova", "Modelo 2020 (UE2020, NR_MODELO ≥ 2020)"),
        ("Urna antiga", "Modelos anteriores a 2020 (UE2009 a UE2015)"),
        (
            "Diferença / margem",
            "Em cada urna: |% Lula − % Bolsonaro| nos votos válidos. "
            "Urna favorável a Lula = Lula > Bolsonaro; favorável a Bolsonaro = o inverso. "
            "A média municipal é a média simples dessas margens.",
        ),
        (
            "Comparação nova × antiga",
            "Só entra município que tem pelo menos uma urna do candidato nos dois modelos. "
            "Diferença = margem média na UE2020 − margem média nas antigas. "
            "Positivo = vantagem maior nas urnas novas.",
        ),
        (
            "Brasil — Lula",
            f"Não pareada: antigas {br['Margem média Lula — urnas antigas (todos)']} p.p. × "
            f"UE2020 {br['Margem média Lula — UE2020 (todos)']} p.p. "
            f"(Δ {br['Diferença Lula não pareada (nova − antiga)']} p.p.). "
            f"Pareada ({br['Comparam Lula nos dois modelos']} mun.): "
            f"{br['Margem média Lula — urnas antigas (pareada)']} → "
            f"{br['Margem média Lula — UE2020 (pareada)']} "
            f"(Δ {br['Diferença Lula pareada (nova − antiga)']} p.p.).",
        ),
        (
            "Brasil — Bolsonaro",
            f"Não pareada: antigas {br['Margem média Bolsonaro — urnas antigas (todos)']} p.p. × "
            f"UE2020 {br['Margem média Bolsonaro — UE2020 (todos)']} p.p. "
            f"(Δ {br['Diferença Bolsonaro não pareada (nova − antiga)']} p.p.). "
            f"Pareada ({br['Comparam Bolsonaro nos dois modelos']} mun.): "
            f"{br['Margem média Bolsonaro — urnas antigas (pareada)']} → "
            f"{br['Margem média Bolsonaro — UE2020 (pareada)']} "
            f"(Δ {br['Diferença Bolsonaro pareada (nova − antiga)']} p.p.).",
        ),
        ("Abas", "Resumo, Por_UF, Municipios"),
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
        "Média das margens municipais — municípios até 50 mil habitantes",
    )
    _escrever_df(wb.add_worksheet("Por_UF"), ufs, fmts, "Mesmas médias por UF")
    cols_mun = [
        "REGIAO",
        "SG_UF",
        "CD_MUNICIPIO",
        "NM_MUNICIPIO",
        "POP_2022",
        "COMPARAVEL",
        "COMPARA_LULA",
        "COMPARA_BOLSO",
        "QT_URNAS_PRE2020",
        "QT_URNAS_LULA_PRE2020",
        "MARGEM_LULA_PRE2020",
        "QT_URNAS_BOLSO_PRE2020",
        "MARGEM_BOLSO_PRE2020",
        "QT_URNAS_UE2020",
        "QT_URNAS_LULA_UE2020",
        "MARGEM_LULA_UE2020",
        "QT_URNAS_BOLSO_UE2020",
        "MARGEM_BOLSO_UE2020",
        "DIF_MARGEM_LULA_NOVA_ANTIGA",
        "DIF_MARGEM_BOLSO_NOVA_ANTIGA",
    ]
    _escrever_df(
        wb.add_worksheet("Municipios"),
        mun[cols_mun],
        fmts,
        "Um município por linha — margem média nas urnas favoráveis a cada candidato",
    )
    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    limite = int(args.limite_hab)
    pop = carregar_populacao(resolver_catalogo(args.catalogo_pop))
    fonte = dados / "tse2022" / "urnas_2t_presidente.csv.gz"
    if not fonte.exists():
        fonte = REPO_ROOT / "output" / "tse2022" / "urnas_2t_presidente.csv.gz"
    urnas = enriquecer_urnas(pd.read_csv(fonte, compression="gzip", low_memory=False))
    mun = apurar_municipios(urnas, pop, limite)
    destino = saida / "discriminativo_modelo_urna_ate_50mil.xlsx"
    gravar_xlsx(destino, mun, limite)
    resumo = tabela_resumo(mun)
    print(resumo.to_string(index=False))
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
