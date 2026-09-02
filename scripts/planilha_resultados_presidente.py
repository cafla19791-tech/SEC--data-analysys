#!/usr/bin/env python3
"""Planilhas de Presidente por região, UF, município, zona e urna.

Gera um XLSX por ano/turno (abas Leia-me, Regiao, UF, Municipio, Zona, Urna)
e um consolidado só com os recortes agregados.

Fontes (já processadas dos BUs / votacao_secao oficiais):
  output/tse2022/urnas_{1t|2t}_presidente.csv.gz
  output/tse2018/secoes_{1t|2t}_presidente.csv.gz
  output/tse2014/... (quando existir)

Uso:
  python3 scripts/planilha_resultados_presidente.py
  python planilha_resultados_presidente.py --ano 2022
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from scripts.baixar_boletins_urna import CANDIDATOS, descobrir_winpython
except ImportError:
    from baixar_boletins_urna import CANDIDATOS, descobrir_winpython  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[1]

REGIAO_POR_UF = {
    "AC": "Norte",
    "AM": "Norte",
    "AP": "Norte",
    "PA": "Norte",
    "RO": "Norte",
    "RR": "Norte",
    "TO": "Norte",
    "AL": "Nordeste",
    "BA": "Nordeste",
    "CE": "Nordeste",
    "MA": "Nordeste",
    "PB": "Nordeste",
    "PE": "Nordeste",
    "PI": "Nordeste",
    "RN": "Nordeste",
    "SE": "Nordeste",
    "DF": "Centro-Oeste",
    "GO": "Centro-Oeste",
    "MS": "Centro-Oeste",
    "MT": "Centro-Oeste",
    "ES": "Sudeste",
    "MG": "Sudeste",
    "RJ": "Sudeste",
    "SP": "Sudeste",
    "PR": "Sul",
    "RS": "Sul",
    "SC": "Sul",
    "ZZ": "Exterior",
}

SKIP_VOTOS = {"OUTROS", "BRANCO", "NULO", "VALIDOS"}

# 2018 usa votacao_secao (nacional completo). 2022/2014 usam BU (com série).
FONTES = {
    (2022, 1): ("tse2022", ("urnas_1t_presidente.csv.gz",)),
    (2022, 2): ("tse2022", ("urnas_2t_presidente.csv.gz",)),
    (2018, 1): ("tse2018", ("secoes_1t_presidente.csv.gz",)),
    (2018, 2): ("tse2018", ("secoes_2t_presidente.csv.gz",)),
    (2014, 1): ("tse2014", ("urnas_1t_presidente.csv.gz", "secoes_1t_presidente.csv.gz")),
    (2014, 2): ("tse2014", ("urnas_2t_presidente.csv.gz", "secoes_2t_presidente.csv.gz")),
}

TOTAIS_OFICIAIS = {
    (2022, 1): {"LULA": 57_259_504, "BOLSONARO": 51_072_345, "TEBET": 4_915_423, "CIRO": 3_599_287},
    (2022, 2): {"LULA": 60_345_999, "BOLSONARO": 58_206_354},
    (2018, 1): {"BOLSONARO": 49_277_010, "HADDAD": 31_342_051, "CIRO": 13_344_371, "AMOEDO": 2_679_745},
    (2018, 2): {"BOLSONARO": 57_797_847, "HADDAD": 47_040_906},
    (2014, 1): {"DILMA": 43_267_668, "AECIO": 34_897_211, "MARINA": 22_176_619},
    (2014, 2): {"DILMA": 54_501_118, "AECIO": 51_041_155},
}

SERIE_PARCIAL = {
    (2018, 1): ("tse2018", "urnas_1t_presidente.csv.gz"),
    (2018, 2): ("tse2018", "urnas_2t_presidente.csv.gz"),
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


def pastas_saida(winpy: Path | None = None) -> Path:
    raiz = winpy if winpy is not None else descobrir_winpython()
    if raiz is not None:
        return raiz / "saida" / "tse_planilhas"
    return REPO_ROOT / "output" / "tse_planilhas"


def pasta_dados(winpy: Path | None = None) -> Path:
    raiz = winpy if winpy is not None else descobrir_winpython()
    if raiz is not None:
        return raiz / "saida"
    return REPO_ROOT / "output"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--ano", type=int, nargs="*", default=[2014, 2018, 2022])
    p.add_argument("--turno", type=int, nargs="*", default=[1, 2])
    p.add_argument("--saida", type=Path, default=None)
    p.add_argument("--dados", type=Path, default=None)
    p.add_argument("--sem-urna", action="store_true", help="Não grava a aba Urna (arquivo menor).")
    return p.parse_args(argv)


def rotulo_regiao(uf: object) -> str:
    if uf is None or (isinstance(uf, float) and pd.isna(uf)):
        return "sem_uf"
    return REGIAO_POR_UF.get(str(uf).strip().upper(), "sem_uf")


def colunas_candidatos(df: pd.DataFrame) -> list[str]:
    out: list[str] = []
    for col in df.columns:
        if not str(col).startswith("QT_VOTOS_"):
            continue
        sufixo = str(col)[len("QT_VOTOS_") :]
        if sufixo not in SKIP_VOTOS:
            out.append(str(col))
    return out


def ler_csv(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(caminho)
    if str(caminho).endswith(".gz"):
        return pd.read_csv(caminho, compression="gzip")
    return pd.read_csv(caminho)


def enriquecer_serie(detalhe: pd.DataFrame, serie: pd.DataFrame) -> pd.DataFrame:
    """Cruza seção × BU para preencher NR_URNA_EFETIVADA quando existir."""
    keys = ["SG_UF", "CD_MUNICIPIO", "NR_ZONA", "NR_SECAO"]
    if any(k not in detalhe.columns or k not in serie.columns for k in keys):
        return detalhe
    extra_cols = [
        c
        for c in ("NR_URNA_EFETIVADA", "NR_MODELO", "DS_MODELO_URNA")
        if c in serie.columns
    ]
    extra = serie[keys + extra_cols].copy()
    base = detalhe.copy()
    for col in ("CD_MUNICIPIO", "NR_ZONA", "NR_SECAO"):
        base[col] = pd.to_numeric(base[col], errors="coerce")
        extra[col] = pd.to_numeric(extra[col], errors="coerce")
    extra["SG_UF"] = extra["SG_UF"].astype(str)
    base["SG_UF"] = base["SG_UF"].astype(str)
    extra = extra.drop_duplicates(keys, keep="first")
    base = base.drop(columns=extra_cols, errors="ignore")
    return base.merge(extra, on=keys, how="left")


def preparar(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    base = df.copy()
    base["ANO_ELEICAO"] = ano
    base["NR_TURNO"] = turno
    base["SG_UF"] = base["SG_UF"].astype(str).str.strip().str.upper()
    base["REGIAO"] = base["SG_UF"].map(rotulo_regiao)
    for col in base.columns:
        if str(col).startswith("QT_") or col in (
            "NR_ZONA",
            "NR_SECAO",
            "CD_MUNICIPIO",
            "NR_URNA_EFETIVADA",
            "NR_MODELO",
        ):
            base[col] = pd.to_numeric(base[col], errors="coerce")
    cand = colunas_candidatos(base)
    if "QT_VOTOS_VALIDOS" not in base.columns and cand:
        base["QT_VOTOS_VALIDOS"] = base[cand].fillna(0).sum(axis=1)
        if "QT_VOTOS_OUTROS" in base.columns:
            base["QT_VOTOS_VALIDOS"] = base["QT_VOTOS_VALIDOS"] + base[
                "QT_VOTOS_OUTROS"
            ].fillna(0)
    return base


def vencedor_frame(df: pd.DataFrame, cand_cols: list[str]) -> pd.Series:
    if not cand_cols:
        return pd.Series("", index=df.index)
    arr = df[cand_cols].fillna(0).to_numpy(dtype=float)
    if arr.size == 0:
        return pd.Series("", index=df.index)
    idx = arr.argmax(axis=1)
    maxv = arr.max(axis=1)
    ordem = np.sort(arr, axis=1)
    empate = (maxv > 0) & (ordem[:, -1] == ordem[:, -2])
    nomes = [c[len("QT_VOTOS_") :] for c in cand_cols]
    out = pd.Series(
        [nomes[i] if m > 0 else "" for i, m in zip(idx, maxv)], index=df.index
    )
    out.loc[empate] = "Empate"
    return out


def com_percentuais(df: pd.DataFrame, cand_cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    validos = out["QT_VOTOS_VALIDOS"].replace(0, np.nan) if "QT_VOTOS_VALIDOS" in out.columns else None
    if validos is not None:
        for col in cand_cols:
            out["PCT_" + col[len("QT_VOTOS_") :]] = (out[col] / validos * 100).round(2)
    out["VENCEDOR"] = vencedor_frame(out, cand_cols)
    return out


def agregar(df: pd.DataFrame, chaves: list[str]) -> pd.DataFrame:
    cand = colunas_candidatos(df)
    soma = cand + [
        c
        for c in (
            "QT_VOTOS_OUTROS",
            "QT_VOTOS_BRANCO",
            "QT_VOTOS_NULO",
            "QT_VOTOS_VALIDOS",
            "QT_APTOS",
            "QT_COMPARECIMENTO",
            "QT_ABSTENCOES",
        )
        if c in df.columns
    ]
    g = df.groupby(chaves, dropna=False)
    agg: dict[str, tuple[str, str]] = {"QT_SECOES": ("NR_SECAO", "size")}
    if "NR_URNA_EFETIVADA" in df.columns:
        agg["QT_URNAS_COM_SERIE"] = ("NR_URNA_EFETIVADA", "nunique")
    for col in soma:
        agg[col] = (col, "sum")
    out = g.agg(**agg).reset_index()
    return com_percentuais(out, cand)


def detalhe_urna(df: pd.DataFrame) -> pd.DataFrame:
    cand = colunas_candidatos(df)
    out = com_percentuais(df, cand)
    ordem = [
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
        *cand,
        "QT_VOTOS_OUTROS",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_VOTOS_VALIDOS",
        *[f"PCT_{c[len('QT_VOTOS_'):]}" for c in cand],
        "VENCEDOR",
        "QT_APTOS",
        "QT_COMPARECIMENTO",
        "QT_ABSTENCOES",
    ]
    return out[[c for c in ordem if c in out.columns]]


def recortes(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    return {
        "Regiao": agregar(df, ["ANO_ELEICAO", "NR_TURNO", "REGIAO"]),
        "UF": agregar(df, ["ANO_ELEICAO", "NR_TURNO", "REGIAO", "SG_UF"]),
        "Municipio": agregar(
            df,
            ["ANO_ELEICAO", "NR_TURNO", "REGIAO", "SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"],
        ),
        "Zona": agregar(
            df,
            [
                "ANO_ELEICAO",
                "NR_TURNO",
                "REGIAO",
                "SG_UF",
                "CD_MUNICIPIO",
                "NM_MUNICIPIO",
                "NR_ZONA",
            ],
        ),
        "Urna": detalhe_urna(df),
    }


def _estilo(writer: pd.ExcelWriter, nome: str, df: pd.DataFrame) -> None:
    ws = writer.sheets[nome]
    ws.freeze_panes(1, 0)
    nlin, ncol = len(df), len(df.columns)
    if ncol:
        ws.autofilter(0, 0, max(nlin, 1), ncol - 1)
    header = writer.book.add_format(
        {
            "bold": True,
            "bg_color": "#1F4E79",
            "font_color": "FFFFFF",
            "text_wrap": True,
            "valign": "vcenter",
        }
    )
    for i, col in enumerate(df.columns):
        ws.write(0, i, col, header)
        ws.set_column(i, i, min(26, max(11, len(str(col)) + 2)))


def escrever_xlsx(
    destino: Path,
    leia: list[tuple[str, str]],
    abas: dict[str, pd.DataFrame],
) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destino, engine="xlsxwriter") as writer:
        pd.DataFrame(leia, columns=["Campo", "Valor"]).to_excel(
            writer, sheet_name="Leia-me", index=False
        )
        _estilo(writer, "Leia-me", pd.DataFrame(leia, columns=["Campo", "Valor"]))
        for nome, df in abas.items():
            aba = nome[:31]
            df.to_excel(writer, sheet_name=aba, index=False)
            _estilo(writer, aba, df)
    return destino


def leia_me(ano: int, turno: int, df: pd.DataFrame, fonte: str) -> list[tuple[str, str]]:
    cand = colunas_candidatos(df)
    linhas = [
        ("Eleição", f"Presidente {ano}, {turno}º turno"),
        ("Fonte", fonte),
        ("Seções / urnas", f"{len(df):,}".replace(",", ".")),
        (
            "Série preenchida",
            f"{int(df['NR_URNA_EFETIVADA'].notna().sum()):,}".replace(",", ".")
            if "NR_URNA_EFETIVADA" in df.columns
            else "não se aplica",
        ),
        ("Abas", "Regiao, UF, Municipio, Zona, Urna (uma linha por urna/seção)"),
        (
            "Regiões",
            "Norte, Nordeste, Centro-Oeste, Sudeste, Sul, Exterior (ZZ)",
        ),
    ]
    for col in cand:
        linhas.append(
            (
                col[len("QT_VOTOS_") :],
                f"{int(df[col].fillna(0).sum()):,}".replace(",", "."),
            )
        )
    if "QT_VOTOS_VALIDOS" in df.columns:
        linhas.append(
            (
                "Válidos",
                f"{int(df['QT_VOTOS_VALIDOS'].fillna(0).sum()):,}".replace(",", "."),
            )
        )
    for aviso in conferir_totais(df, ano, turno):
        linhas.append(("Conferência TSE", aviso))
    return linhas


def resolver_fonte(dados: Path, ano: int, turno: int) -> Path:
    pasta, nomes = FONTES[(ano, turno)]
    candidatos: list[Path] = []
    for nome in nomes:
        candidatos.extend(
            (
                dados / pasta / nome,
                dados / pasta / nome.replace(".csv.gz", ".csv"),
                REPO_ROOT / "output" / pasta / nome,
            )
        )
    for cand in candidatos:
        if cand.exists():
            return cand
    raise FileNotFoundError(
        f"Falta {pasta}/{nomes[0]}. No ContAgil: "
        f"python baixar_boletins_urna.py --somente-resultado-github --ano {ano} --turno {turno}"
    )


def conferir_totais(df: pd.DataFrame, ano: int, turno: int) -> list[str]:
    esperado = TOTAIS_OFICIAIS.get((ano, turno), {})
    linhas: list[str] = []
    for nome, oficial in esperado.items():
        col = f"QT_VOTOS_{nome}"
        if col not in df.columns:
            linhas.append(f"{nome}: coluna ausente (oficial {oficial:,})".replace(",", "."))
            continue
        obtido = int(df[col].fillna(0).sum())
        ok = "OK" if obtido == oficial else "DIFERE"
        linhas.append(
            f"{nome}: {obtido:,} ({ok}; TSE {oficial:,})".replace(",", ".")
        )
    return linhas


def carregar_pleito(dados: Path, ano: int, turno: int) -> tuple[pd.DataFrame, str]:
    path = resolver_fonte(dados, ano, turno)
    df = preparar(ler_csv(path), ano, turno)
    fonte = f"{path.name}"
    chave_serie = (ano, turno)
    if chave_serie in SERIE_PARCIAL:
        pasta, nome = SERIE_PARCIAL[chave_serie]
        serie_path = dados / pasta / nome
        if not serie_path.exists():
            serie_path = REPO_ROOT / "output" / pasta / nome
        if serie_path.exists():
            df = enriquecer_serie(df, ler_csv(serie_path))
            df["REGIAO"] = df["SG_UF"].map(rotulo_regiao)
            fonte += f" + série parcial {nome}"
    return df, fonte


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = Path(args.dados) if args.dados else pasta_dados()
    saida = Path(args.saida) if args.saida else pastas_saida()
    saida.mkdir(parents=True, exist_ok=True)

    agregados: dict[str, list[pd.DataFrame]] = {
        "Regiao": [],
        "UF": [],
        "Municipio": [],
        "Zona": [],
    }
    gerados: list[Path] = []

    for ano in args.ano:
        for turno in args.turno:
            if (ano, turno) not in FONTES:
                continue
            try:
                df, fonte = carregar_pleito(dados, ano, turno)
            except FileNotFoundError as exc:
                print(f"  pulando {ano} T{turno}: {exc}", flush=True)
                continue
            print(f"  {ano} T{turno}: {len(df):,} linhas ← {fonte}", flush=True)
            for aviso in conferir_totais(df, ano, turno):
                print(f"    {aviso}", flush=True)
            abas = recortes(df)
            if args.sem_urna:
                abas.pop("Urna", None)
            dest = saida / f"resultados_presidente_{ano}_{turno}t.xlsx"
            escrever_xlsx(dest, leia_me(ano, turno, df, fonte), abas)
            gerados.append(dest)
            print(f"    {dest} ({dest.stat().st_size / 1_048_576:.1f} MB)", flush=True)
            for nome in ("Regiao", "UF", "Municipio", "Zona"):
                agregados[nome].append(abas[nome])

    if any(agregados.values()):
        combo = saida / "resultados_agregados_regiao_uf_municipio_zona.xlsx"
        blocos = {
            nome: pd.concat(partes, ignore_index=True)
            for nome, partes in agregados.items()
            if partes
        }
        anos_ok = sorted({int(p["ANO_ELEICAO"].iloc[0]) for p in agregados["Regiao"] if len(p)})
        escrever_xlsx(
            combo,
            [
                (
                    "Conteúdo",
                    "Região, UF, município e zona — "
                    + (", ".join(str(a) for a in anos_ok) or "nenhum ano")
                    + ", 1º e 2º turnos",
                ),
                (
                    "2014",
                    "Sem microdados nesta rede (TSE 403; Archive.org sem captura). "
                    "O script gera as abas quando existirem output/tse2014/urnas_* ou secoes_*.",
                ),
                ("Urna", "Nas planilhas resultados_presidente_{ano}_{turno}t.xlsx (aba Urna)"),
            ],
            blocos,
        )
        gerados.append(combo)
        print(f"  consolidado {combo} ({combo.stat().st_size / 1_048_576:.1f} MB)", flush=True)

    if not gerados:
        raise RuntimeError("Nenhuma planilha gerada. Confira os CSVs em output/tseYYYY/.")
    print(f"Saída: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
