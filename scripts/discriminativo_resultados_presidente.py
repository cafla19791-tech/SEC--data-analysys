#!/usr/bin/env python3
"""Discriminativo comparativo das eleições presidenciais 2014, 2018 e 2022.

Cruza o 2º turno (dois candidatos) alinhando o lado PT
(Dilma / Haddad / Lula) com a oposição (Aécio / Bolsonaro / Bolsonaro)
em Brasil, região, UF, município e zona. O 1º turno entra com os
principais candidatos por UF e região.

Saída:
  output/tse_planilhas/discriminativo_presidente_2014_2018_2022.xlsx
  output/tse_planilhas/discriminativo_presidente_uf_2t.csv
  output/tse_planilhas/discriminativo_presidente_municipio_2t.csv
  output/tse_planilhas/discriminativo_urnas_{ano}_{turno}t.csv.gz
  output/tse_planilhas/discriminativo_urnas_{ano}_{turno}t.xlsx

Uso:
  python3 scripts/discriminativo_resultados_presidente.py
  python discriminativo_resultados_presidente.py --somente-urna
  python discriminativo_resultados_presidente.py --somente-urna --sem-xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.planilha_resultados_presidente import (
        carregar_pleito,
        conferir_totais,
        detalhe_urna,
        escrever_xlsx,
        pasta_dados,
        pastas_saida,
    )
except ImportError:  # ContAgil
    from planilha_resultados_presidente import (  # type: ignore
        carregar_pleito,
        conferir_totais,
        detalhe_urna,
        escrever_xlsx,
        pasta_dados,
        pastas_saida,
    )

ANOS = (2014, 2018, 2022)

# 2º turno: mesmo recorte político (PT × oposição).
LADOS_2T: dict[int, dict[str, str]] = {
    2014: {"pt": "DILMA", "opp": "AECIO", "pt_nome": "Dilma", "opp_nome": "Aécio"},
    2018: {"pt": "HADDAD", "opp": "BOLSONARO", "pt_nome": "Haddad", "opp_nome": "Bolsonaro"},
    2022: {"pt": "LULA", "opp": "BOLSONARO", "pt_nome": "Lula", "opp_nome": "Bolsonaro"},
}

CANDIDATOS_1T: dict[int, tuple[str, ...]] = {
    2014: ("DILMA", "AECIO", "MARINA"),
    2018: ("BOLSONARO", "HADDAD", "CIRO", "AMOEDO"),
    2022: ("LULA", "BOLSONARO", "TEBET", "CIRO"),
}


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
    p.add_argument(
        "--somente-urna",
        action="store_true",
        help="Só gera o discriminativo por urna (um arquivo por ano/turno).",
    )
    p.add_argument(
        "--sem-xlsx",
        action="store_true",
        help="Não grava o XLSX por urna (só o CSV.gz).",
    )
    return p.parse_args(argv)


def _pct(parte: float, total: float) -> float | None:
    if total is None or pd.isna(total) or float(total) <= 0:
        return None
    return round(100.0 * float(parte) / float(total), 2)


def lado_vencedor(pt: float, opp: float) -> str:
    if pt > opp:
        return "PT"
    if opp > pt:
        return "OPP"
    return "EMPATE"


def nome_vencedor(ano: int, lado: str) -> str:
    info = LADOS_2T[ano]
    if lado == "PT":
        return info["pt_nome"]
    if lado == "OPP":
        return info["opp_nome"]
    return "Empate"


def inverteram(lado_a: object, lado_b: object) -> str:
    if lado_a in (None, "EMPATE") or lado_b in (None, "EMPATE"):
        return "N"
    if pd.isna(lado_a) or pd.isna(lado_b):
        return "N"
    return "S" if lado_a != lado_b else "N"


def preparar_2t(df: pd.DataFrame, ano: int) -> pd.DataFrame:
    info = LADOS_2T[ano]
    base = df.copy()
    pt = f"QT_VOTOS_{info['pt']}"
    opp = f"QT_VOTOS_{info['opp']}"
    if pt not in base.columns or opp not in base.columns:
        raise KeyError(f"{ano} 2T: faltam {pt} / {opp}")
    base["QT_VOTOS_PT"] = pd.to_numeric(base[pt], errors="coerce").fillna(0)
    base["QT_VOTOS_OPP"] = pd.to_numeric(base[opp], errors="coerce").fillna(0)
    if "QT_VOTOS_VALIDOS" in base.columns:
        base["QT_VOTOS_VALIDOS"] = pd.to_numeric(
            base["QT_VOTOS_VALIDOS"], errors="coerce"
        ).fillna(0)
    else:
        base["QT_VOTOS_VALIDOS"] = base["QT_VOTOS_PT"] + base["QT_VOTOS_OPP"]
    if "NR_SECAO" not in base.columns:
        base["NR_SECAO"] = range(1, len(base) + 1)
    return base


def chaves_cruzamento(chaves: list[str]) -> list[str]:
    """Nome do município não entra na chave: o código TSE é estável."""
    return [c for c in chaves if c != "NM_MUNICIPIO"]


def agregar_2t(df: pd.DataFrame, chaves: list[str], ano: int) -> pd.DataFrame:
    trabalho = df.copy()
    for col in chaves:
        if col == "NM_MUNICIPIO":
            trabalho[col] = trabalho[col].astype(str)
        elif col in ("CD_MUNICIPIO", "NR_ZONA"):
            trabalho[col] = pd.to_numeric(trabalho[col], errors="coerce")
        elif col == "SG_UF":
            trabalho[col] = trabalho[col].astype(str).str.strip().str.upper()
    grupo = chaves_cruzamento(chaves)
    agg: dict[str, tuple[str, str]] = {
        "QT_SECOES": ("NR_SECAO", "size"),
        "QT_VOTOS_PT": ("QT_VOTOS_PT", "sum"),
        "QT_VOTOS_OPP": ("QT_VOTOS_OPP", "sum"),
        "QT_VOTOS_VALIDOS": ("QT_VOTOS_VALIDOS", "sum"),
    }
    if "NM_MUNICIPIO" in chaves:
        agg["NM_MUNICIPIO"] = ("NM_MUNICIPIO", "first")
    g = trabalho.groupby(grupo, dropna=False).agg(**agg).reset_index()
    g["PCT_PT"] = [
        _pct(p, v) for p, v in zip(g["QT_VOTOS_PT"], g["QT_VOTOS_VALIDOS"])
    ]
    g["PCT_OPP"] = [
        _pct(o, v) for o, v in zip(g["QT_VOTOS_OPP"], g["QT_VOTOS_VALIDOS"])
    ]
    g["LADO"] = [
        lado_vencedor(p, o) for p, o in zip(g["QT_VOTOS_PT"], g["QT_VOTOS_OPP"])
    ]
    g["VENCEDOR"] = [nome_vencedor(ano, lado) for lado in g["LADO"]]
    g["CAND_PT"] = LADOS_2T[ano]["pt_nome"]
    g["CAND_OPP"] = LADOS_2T[ano]["opp_nome"]
    renome = {
        c: f"{c}_{ano}"
        for c in (
            "QT_SECOES",
            "QT_VOTOS_PT",
            "QT_VOTOS_OPP",
            "QT_VOTOS_VALIDOS",
            "PCT_PT",
            "PCT_OPP",
            "LADO",
            "VENCEDOR",
            "CAND_PT",
            "CAND_OPP",
        )
    }
    return g.rename(columns=renome)


def cruzar_anos(blocos: dict[int, pd.DataFrame], chaves: list[str]) -> pd.DataFrame:
    cruz = chaves_cruzamento(chaves)
    out: pd.DataFrame | None = None
    for ano in ANOS:
        parte = blocos[ano]
        if "NM_MUNICIPIO" in parte.columns:
            parte = parte.rename(columns={"NM_MUNICIPIO": f"NM_MUNICIPIO_{ano}"})
        out = parte if out is None else out.merge(parte, on=cruz, how="outer")
    assert out is not None
    nomes = [f"NM_MUNICIPIO_{ano}" for ano in (2022, 2018, 2014) if f"NM_MUNICIPIO_{ano}" in out.columns]
    if nomes:
        out["NM_MUNICIPIO"] = out[nomes[0]]
        for col in nomes[1:]:
            out["NM_MUNICIPIO"] = out["NM_MUNICIPIO"].where(
                out["NM_MUNICIPIO"].notna() & (out["NM_MUNICIPIO"].astype(str) != ""),
                out[col],
            )
        out = out.drop(columns=nomes)
    for ano in ANOS:
        if f"QT_SECOES_{ano}" in out.columns:
            out[f"QT_SECOES_{ano}"] = (
                pd.to_numeric(out[f"QT_SECOES_{ano}"], errors="coerce")
                .fillna(0)
                .astype(int)
            )
        for col in (f"QT_VOTOS_PT_{ano}", f"QT_VOTOS_OPP_{ano}", f"QT_VOTOS_VALIDOS_{ano}"):
            if col in out.columns:
                out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)
    out["COMPARAVEL"] = "N"
    tem = [f"QT_SECOES_{ano}" in out.columns for ano in ANOS]
    if all(tem):
        out.loc[
            (out["QT_SECOES_2014"] > 0)
            & (out["QT_SECOES_2018"] > 0)
            & (out["QT_SECOES_2022"] > 0),
            "COMPARAVEL",
        ] = "S"
    pares = ((2014, 2018), (2018, 2022), (2014, 2022))
    for a, b in pares:
        out[f"DIF_PCT_PT_{a}_{b}"] = [
            None
            if pd.isna(x) or pd.isna(y)
            else round(float(y) - float(x), 2)
            for x, y in zip(out[f"PCT_PT_{a}"], out[f"PCT_PT_{b}"])
        ]
        out[f"INVERTEU_{a}_{b}"] = [
            inverteram(la, lb)
            for la, lb in zip(out[f"LADO_{a}"], out[f"LADO_{b}"])
        ]
    if "NM_MUNICIPIO" in out.columns:
        out["NM_MUNICIPIO"] = out["NM_MUNICIPIO"].fillna("").astype(str)
    return out


def ordem_2t(chaves: list[str]) -> list[str]:
    meio: list[str] = []
    for ano in ANOS:
        meio.extend(
            [
                f"CAND_PT_{ano}",
                f"CAND_OPP_{ano}",
                f"QT_SECOES_{ano}",
                f"QT_VOTOS_PT_{ano}",
                f"QT_VOTOS_OPP_{ano}",
                f"QT_VOTOS_VALIDOS_{ano}",
                f"PCT_PT_{ano}",
                f"PCT_OPP_{ano}",
                f"VENCEDOR_{ano}",
                f"LADO_{ano}",
            ]
        )
    extras = [
        "DIF_PCT_PT_2014_2018",
        "DIF_PCT_PT_2018_2022",
        "DIF_PCT_PT_2014_2022",
        "INVERTEU_2014_2018",
        "INVERTEU_2018_2022",
        "INVERTEU_2014_2022",
        "COMPARAVEL",
    ]
    return chaves + meio + extras


def alinhar(df: pd.DataFrame, chaves: list[str]) -> pd.DataFrame:
    ordem = [c for c in ordem_2t(chaves) if c in df.columns]
    extras = [c for c in df.columns if c not in ordem]
    out = df.reindex(columns=ordem + extras)
    sort_cols = [c for c in chaves if c in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, kind="mergesort").reset_index(drop=True)
    return out


def agregar_1t(df: pd.DataFrame, chaves: list[str], ano: int) -> pd.DataFrame:
    cands = [c for c in CANDIDATOS_1T[ano] if f"QT_VOTOS_{c}" in df.columns]
    trabalho = df.copy()
    for col in chaves:
        if col == "SG_UF":
            trabalho[col] = trabalho[col].astype(str).str.strip().str.upper()
        elif col == "NM_MUNICIPIO":
            trabalho[col] = trabalho[col].astype(str)
    if "NR_SECAO" not in trabalho.columns:
        trabalho["NR_SECAO"] = 1
    for c in cands:
        trabalho[f"QT_VOTOS_{c}"] = pd.to_numeric(
            trabalho[f"QT_VOTOS_{c}"], errors="coerce"
        ).fillna(0)
    if "QT_VOTOS_VALIDOS" in trabalho.columns:
        trabalho["QT_VOTOS_VALIDOS"] = pd.to_numeric(
            trabalho["QT_VOTOS_VALIDOS"], errors="coerce"
        ).fillna(0)
    else:
        trabalho["QT_VOTOS_VALIDOS"] = sum(trabalho[f"QT_VOTOS_{c}"] for c in cands)
    agg = {"QT_SECOES": ("NR_SECAO", "size")}
    for c in cands:
        agg[c] = (f"QT_VOTOS_{c}", "sum")
    agg["VALIDOS"] = ("QT_VOTOS_VALIDOS", "sum")
    g = trabalho.groupby(chaves, dropna=False).agg(**agg).reset_index()
    for c in cands:
        g[f"PCT_{c}"] = [
            _pct(v, t) for v, t in zip(g[c], g["VALIDOS"])
        ]
    nomes = list(cands)

    def venc(row: pd.Series) -> str:
        vals = [float(row[c]) for c in nomes]
        if not vals or max(vals) <= 0:
            return ""
        if vals.count(max(vals)) > 1:
            return "Empate"
        return nomes[vals.index(max(vals))].title().replace("_", " ")

    g["VENCEDOR"] = g.apply(venc, axis=1)
    renome = {"QT_SECOES": f"QT_SECOES_{ano}", "VALIDOS": f"VALIDOS_{ano}", "VENCEDOR": f"VENCEDOR_{ano}"}
    for c in cands:
        renome[c] = f"{c}_{ano}"
        renome[f"PCT_{c}"] = f"PCT_{c}_{ano}"
    return g.rename(columns=renome)


def cruzar_1t(blocos: dict[int, pd.DataFrame], chaves: list[str]) -> pd.DataFrame:
    out: pd.DataFrame | None = None
    for ano in ANOS:
        out = blocos[ano] if out is None else out.merge(blocos[ano], on=chaves, how="outer")
    assert out is not None
    sort_cols = [c for c in chaves if c in out.columns]
    if sort_cols:
        out = out.sort_values(sort_cols, kind="mergesort").reset_index(drop=True)
    return out


def linha_brasil_2t(df: pd.DataFrame, ano: int) -> dict[str, object]:
    info = LADOS_2T[ano]
    pt = int(df["QT_VOTOS_PT"].sum())
    opp = int(df["QT_VOTOS_OPP"].sum())
    val = int(df["QT_VOTOS_VALIDOS"].sum())
    lado = lado_vencedor(pt, opp)
    return {
        "ANO": ano,
        "TURNO": 2,
        "CAND_PT": info["pt_nome"],
        "CAND_OPP": info["opp_nome"],
        "QT_SECOES": int(len(df)),
        "QT_VOTOS_PT": pt,
        "QT_VOTOS_OPP": opp,
        "QT_VOTOS_VALIDOS": val,
        "PCT_PT": _pct(pt, val),
        "PCT_OPP": _pct(opp, val),
        "VENCEDOR": nome_vencedor(ano, lado),
        "LADO": lado,
    }


def discriminar_urna(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    """Uma linha por urna/seção, com % e vencedor."""
    out = detalhe_urna(df)
    out["ANO_ELEICAO"] = ano
    out["NR_TURNO"] = turno
    if turno != 2:
        return out
    info = LADOS_2T[ano]
    pt_col = f"QT_VOTOS_{info['pt']}"
    opp_col = f"QT_VOTOS_{info['opp']}"
    out["CAND_PT"] = info["pt_nome"]
    out["CAND_OPP"] = info["opp_nome"]
    out["QT_VOTOS_PT"] = pd.to_numeric(out[pt_col], errors="coerce").fillna(0)
    out["QT_VOTOS_OPP"] = pd.to_numeric(out[opp_col], errors="coerce").fillna(0)
    if "QT_VOTOS_VALIDOS" in out.columns:
        validos = pd.to_numeric(out["QT_VOTOS_VALIDOS"], errors="coerce").fillna(0)
    else:
        validos = out["QT_VOTOS_PT"] + out["QT_VOTOS_OPP"]
        out["QT_VOTOS_VALIDOS"] = validos
    out["PCT_PT"] = [_pct(p, v) for p, v in zip(out["QT_VOTOS_PT"], validos)]
    out["PCT_OPP"] = [_pct(o, v) for o, v in zip(out["QT_VOTOS_OPP"], validos)]
    out["LADO"] = [
        lado_vencedor(p, o) for p, o in zip(out["QT_VOTOS_PT"], out["QT_VOTOS_OPP"])
    ]
    out["VENCEDOR"] = [nome_vencedor(ano, lado) for lado in out["LADO"]]
    frente = [
        "ANO_ELEICAO",
        "NR_TURNO",
        "REGIAO",
        "SG_UF",
        "CD_MUNICIPIO",
        "NM_MUNICIPIO",
        "NR_ZONA",
        "NR_SECAO",
        "NR_LOCAL_VOTACAO",
        "NR_URNA_EFETIVADA",
        "NR_MODELO",
        "DS_MODELO_URNA",
        "CAND_PT",
        "CAND_OPP",
        "QT_VOTOS_PT",
        "QT_VOTOS_OPP",
        "QT_VOTOS_VALIDOS",
        "PCT_PT",
        "PCT_OPP",
        "VENCEDOR",
        "LADO",
    ]
    resto = [c for c in out.columns if c not in frente]
    return out.reindex(columns=[c for c in frente + resto if c in out.columns])


def leia_me_urna(ano: int, turno: int, df: pd.DataFrame, fonte: str) -> list[tuple[str, str]]:
    linhas = [
        ("Eleição", f"Presidente {ano}, {turno}º turno — discriminativo por urna"),
        ("Fonte", fonte),
        ("Linhas", f"{len(df):,}".replace(",", ".")),
        (
            "Série preenchida",
            f"{int(df['NR_URNA_EFETIVADA'].notna().sum()):,}".replace(",", ".")
            if "NR_URNA_EFETIVADA" in df.columns
            else "não se aplica",
        ),
        ("Aba", "Urna (uma linha por urna/seção)"),
    ]
    if turno == 2:
        info = LADOS_2T[ano]
        linhas.append(
            (
                "Lados",
                f"PT = {info['pt_nome']} · oposição = {info['opp_nome']}",
            )
        )
        linhas.append(
            (
                "PT",
                f"{int(df['QT_VOTOS_PT'].fillna(0).sum()):,}".replace(",", "."),
            )
        )
        linhas.append(
            (
                "Oposição",
                f"{int(df['QT_VOTOS_OPP'].fillna(0).sum()):,}".replace(",", "."),
            )
        )
    return linhas


def gravar_discriminativo_urna(
    df: pd.DataFrame,
    ano: int,
    turno: int,
    saida: Path,
    fonte: str,
    *,
    xlsx: bool,
) -> list[Path]:
    tabela = discriminar_urna(df, ano, turno)
    stem = f"discriminativo_urnas_{ano}_{turno}t"
    gz = saida / f"{stem}.csv.gz"
    tabela.to_csv(gz, index=False, encoding="utf-8", compression="gzip")
    print(f"  {gz} ({gz.stat().st_size / 1_048_576:.1f} MB, {len(tabela):,} urnas)", flush=True)
    escritos = [gz]
    if xlsx:
        dest = saida / f"{stem}.xlsx"
        escrever_xlsx(dest, leia_me_urna(ano, turno, tabela, fonte), {"Urna": tabela})
        print(f"  {dest} ({dest.stat().st_size / 1_048_576:.1f} MB)", flush=True)
        escritos.append(dest)
    return escritos


def leia_me(avisos: list[str]) -> list[tuple[str, str]]:
    linhas = [
        ("Conteúdo", "Discriminativo Presidente 2014 × 2018 × 2022"),
        (
            "2º turno",
            "Lado PT = Dilma (2014), Haddad (2018), Lula (2022). "
            "Oposição = Aécio (2014), Bolsonaro (2018 e 2022).",
        ),
        (
            "Abas",
            "Brasil_2T, Regiao_2T, UF_2T, Municipio_2T, Zona_2T, "
            "Regiao_1T, UF_1T",
        ),
        (
            "COMPARAVEL",
            "S = o recorte existe nos três pleitos (seções > 0).",
        ),
        (
            "INVERTEU_A_B",
            "S = o lado vencedor (PT/OPP) mudou entre os anos A e B.",
        ),
        (
            "DIF_PCT_PT_A_B",
            "Pontos percentuais do lado PT no ano B menos o ano A.",
        ),
        ("Fonte 2014", "Boletins de Urna (NR_URNA_EFETIVADA)"),
        ("Fonte 2018", "votacao_secao nacional (série parcial)"),
        ("Fonte 2022", "Boletins de Urna (NR_URNA_EFETIVADA)"),
    ]
    for aviso in avisos:
        linhas.append(("Conferência TSE", aviso))
    return linhas


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = Path(args.dados) if args.dados else pasta_dados()
    saida = Path(args.saida) if args.saida else pastas_saida()
    saida.mkdir(parents=True, exist_ok=True)

    avisos: list[str] = []
    brutos_2t: dict[int, pd.DataFrame] = {}
    brutos_1t: dict[int, pd.DataFrame] = {}
    for ano in ANOS:
        df2, fonte2 = carregar_pleito(dados, ano, 2)
        print(f"  {ano} T2: {len(df2):,} ← {fonte2}", flush=True)
        for aviso in conferir_totais(df2, ano, 2):
            print(f"    {aviso}", flush=True)
            avisos.append(f"{ano} 2T {aviso}")
        gravar_discriminativo_urna(
            df2, ano, 2, saida, fonte2, xlsx=not args.sem_xlsx
        )
        brutos_2t[ano] = preparar_2t(df2, ano)
        del df2
        df1, fonte1 = carregar_pleito(dados, ano, 1)
        print(f"  {ano} T1: {len(df1):,} ← {fonte1}", flush=True)
        for aviso in conferir_totais(df1, ano, 1):
            print(f"    {aviso}", flush=True)
            avisos.append(f"{ano} 1T {aviso}")
        gravar_discriminativo_urna(
            df1, ano, 1, saida, fonte1, xlsx=not args.sem_xlsx
        )
        brutos_1t[ano] = df1
        del df1

    if args.somente_urna:
        print(f"Saída: {saida}")
        return 0

    brasil = pd.DataFrame([linha_brasil_2t(brutos_2t[ano], ano) for ano in ANOS])
    recortes_2t = {
        "Regiao": ["REGIAO"],
        "UF": ["REGIAO", "SG_UF"],
        "Municipio": ["REGIAO", "SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"],
        "Zona": ["REGIAO", "SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO", "NR_ZONA"],
    }
    abas: dict[str, pd.DataFrame] = {"Brasil_2T": brasil}
    gerados_csv: dict[str, pd.DataFrame] = {}
    for nome, chaves in recortes_2t.items():
        blocos = {ano: agregar_2t(brutos_2t[ano], chaves, ano) for ano in ANOS}
        tabela = alinhar(cruzar_anos(blocos, chaves), chaves)
        abas[f"{nome}_2T"] = tabela
        print(f"  {nome}_2T: {len(tabela):,} linhas", flush=True)
        if nome in ("UF", "Municipio"):
            gerados_csv[nome.lower()] = tabela

    for nome, chaves in (("Regiao", ["REGIAO"]), ("UF", ["REGIAO", "SG_UF"])):
        blocos = {ano: agregar_1t(brutos_1t[ano], chaves, ano) for ano in ANOS}
        tabela = cruzar_1t(blocos, chaves)
        abas[f"{nome}_1T"] = tabela
        print(f"  {nome}_1T: {len(tabela):,} linhas", flush=True)

    dest = saida / "discriminativo_presidente_2014_2018_2022.xlsx"
    escrever_xlsx(dest, leia_me(avisos), abas)
    print(f"  {dest} ({dest.stat().st_size / 1_048_576:.1f} MB)", flush=True)

    uf_csv = saida / "discriminativo_presidente_uf_2t.csv"
    mun_csv = saida / "discriminativo_presidente_municipio_2t.csv"
    gerados_csv["uf"].to_csv(uf_csv, index=False, encoding="utf-8")
    gerados_csv["municipio"].to_csv(mun_csv, index=False, encoding="utf-8")
    print(f"  {uf_csv}", flush=True)
    print(f"  {mun_csv}", flush=True)
    print(f"Saída: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
