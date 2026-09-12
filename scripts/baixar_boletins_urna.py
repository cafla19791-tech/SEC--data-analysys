#!/usr/bin/env python3
"""
Baixa e consolida Boletins de Urna (Presidente) por série da urna.

Eleições: 2022 (1º e 2º turno), 2018 (1º e 2º) e 2014 (1º e 2º).
Uma linha por urna/seção, com NR_URNA_EFETIVADA, modelo e votos.

Fontes: Dados Abertos do TSE (buweb por UF). O CDN costuma devolver 403;
o script tenta o Internet Archive e, na RFB, o CSV já consolidado no GitHub.

Uso:
  python3 scripts/baixar_boletins_urna.py --ano 2022 --turno 1
  python3 scripts/baixar_boletins_urna.py --ano 2018 --turno 2 --ufs RR AC
  python baixar_boletins_urna.py --somente-resultado-github
"""

from __future__ import annotations

import argparse
import gzip
import io
import json
import sys
import zipfile
from pathlib import Path

import pandas as pd

try:
    from scripts.baixar_boletins_urna_2022 import (
        UFS,
        baixar_arquivo,
        carregar_faixas_modelo,
        classificar_modelo,
        descobrir_winpython,
        rotulo_modelo,
        urls_espelho,
        user_agent_para,
    )
except ImportError:  # ContAgil: script na pasta winpython
    from baixar_boletins_urna_2022 import (  # type: ignore
        UFS,
        baixar_arquivo,
        carregar_faixas_modelo,
        classificar_modelo,
        descobrir_winpython,
        rotulo_modelo,
        urls_espelho,
        user_agent_para,
    )

REPO_ROOT = Path(__file__).resolve().parents[1]
CATALOGO_PADRAO = REPO_ROOT / "data" / "tse_catalog" / "boletins_urna_urls.json"
RESULTADO_GITHUB_REF = "cursor/tse-boletins-urna-209b"
RESULTADO_GITHUB_BASE = (
    "https://raw.githubusercontent.com/cafla19791-tech/SEC--data-analysys/"
    f"{RESULTADO_GITHUB_REF}/output/"
)

# Números oficiais dos candidatos a Presidente (TSE).
CANDIDATOS: dict[tuple[int, int], dict[int, str]] = {
    (2022, 1): {
        13: "LULA",
        22: "BOLSONARO",
        15: "TEBET",
        12: "CIRO",
        44: "SORAYA",
        30: "FELIPE_DAVILA",
        14: "PADRE_KELMON",
        80: "LEO_PERICLES",
        21: "SOFIA_MANZANO",
        16: "VERA",
        27: "EYMAEL",
    },
    (2022, 2): {13: "LULA", 22: "BOLSONARO"},
    (2018, 1): {
        17: "BOLSONARO",
        13: "HADDAD",
        12: "CIRO",
        45: "ALCKMIN",
        18: "MARINA",
        50: "BOULOS",
        15: "MEIRELLES",
        51: "DACIOLO",
        19: "ALVARO_DIAS",
        30: "AMOEDO",
        16: "VERA",
        27: "EYMAEL",
        54: "JOAO_GOULART",
    },
    (2018, 2): {17: "BOLSONARO", 13: "HADDAD"},
    (2014, 1): {
        13: "DILMA",
        45: "AECIO",
        40: "MARINA",
        50: "LUCIANA_GENRO",
        20: "EVERALDO",
        43: "EDUARDO_JORGE",
        16: "ZE_MARIA",
        21: "MAURO_IASI",
        27: "EYMAEL",
        28: "LEVY_FIDELIX",
        29: "RUI_COSTA_PIMENTA",
    },
    (2014, 2): {13: "DILMA", 45: "AECIO"},
}

# Layout posicional do BUWEB 2014 (TXT sem cabeçalho, 30 campos).
BWEB_2014_COLS = (
    "DT_GERACAO",
    "HH_GERACAO",
    "CD_PLEITO",
    "CD_ELEICAO",
    "SG_UF",
    "CD_CARGO_PERGUNTA",
    "DS_CARGO_PERGUNTA",
    "NR_ZONA",
    "NR_SECAO",
    "NR_LOCAL_VOTACAO",
    "NR_PARTIDO",
    "NM_PARTIDO",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "DT_BU_RECEBIDO",
    "QT_APTOS",
    "QT_ABSTENCOES",
    "QT_COMPARECIMENTO",
    "CD_TIPO_ELEICAO",
    "CD_TIPO_URNA",
    "DS_TIPO_URNA",
    "NR_VOTAVEL",
    "NM_VOTAVEL",
    "QT_VOTOS",
    "DS_TIPO_VOTAVEL",
    "NR_URNA_EFETIVADA",
    "CD_CARGA_1_URNA_EFETIVADA",
    "CD_CARGA_2_URNA_EFETIVADA",
    "CD_FLASHCARD_URNA_EFETIVADA",
    "DS_CARGO_PERGUNTA_SECAO",
)

BWEB_COLS = (
    "NR_TURNO",
    "SG_UF",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "NR_LOCAL_VOTACAO",
    "CD_CARGO_PERGUNTA",
    "DS_CARGO_PERGUNTA",
    "CD_TIPO_VOTAVEL",
    "DS_TIPO_VOTAVEL",
    "NR_VOTAVEL",
    "NM_VOTAVEL",
    "QT_VOTOS",
    "NR_URNA_EFETIVADA",
    "QT_APTOS",
    "QT_COMPARECIMENTO",
    "QT_ABSTENCOES",
    "DT_ABERTURA",
    "DT_ENCERRAMENTO",
    "DS_TIPO_URNA",
)

CHAVES_URNA = (
    "SG_UF",
    "CD_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "NR_URNA_EFETIVADA",
)

WAYBACK_EXTRA = (
    "https://web.archive.org/web/20221108000702id_/",
    "https://web.archive.org/web/20221105id_/",
    "https://web.archive.org/web/202409id_/",
    "https://web.archive.org/web/202606id_/",
    "https://web.archive.org/web/2023id_/",
)

ARQUIVOS_GITHUB = {
    (2022, 1): "tse2022/urnas_1t_presidente.csv.gz",
    (2022, 2): "tse2022/urnas_2t_presidente.csv.gz",
    (2018, 1): "tse2018/urnas_1t_presidente.csv.gz",
    (2018, 2): "tse2018/urnas_2t_presidente.csv.gz",
    (2014, 1): "tse2014/urnas_1t_presidente.csv.gz",
    (2014, 2): "tse2014/urnas_2t_presidente.csv.gz",
}
ARQUIVOS_GITHUB_EXTRA = {
    (2018, 1): (
        "tse2018/secoes_1t_presidente.csv.gz",
        "tse_planilhas/discriminativo_urnas_2018_1t.csv.gz",
    ),
    (2018, 2): (
        "tse2018/secoes_2t_presidente.csv.gz",
        "tse_planilhas/discriminativo_urnas_2018_2t.csv.gz",
    ),
    (2014, 1): ("tse_planilhas/discriminativo_urnas_2014_1t.csv.gz",),
    (2014, 2): ("tse_planilhas/discriminativo_urnas_2014_2t.csv.gz",),
    (2022, 1): ("tse_planilhas/discriminativo_urnas_2022_1t.csv.gz",),
    (2022, 2): ("tse_planilhas/discriminativo_urnas_2022_2t.csv.gz",),
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


def carregar_catalogo(path: Path | None = None) -> list[dict]:
    destino = path or CATALOGO_PADRAO
    return json.loads(destino.read_text(encoding="utf-8"))


def recurso_catalogo(
    catalogo: list[dict], ano: int, turno: int, uf: str
) -> dict:
    uf = uf.upper()
    for item in catalogo:
        if item["ano"] == ano and item["turno"] == turno and item["uf"] == uf:
            return item
    raise KeyError(f"Sem URL no catálogo para {ano} T{turno} {uf}")


def urls_espelho_historico(url: str) -> list[str]:
    """TSE + vários timestamps do Archive.org (1º turno 2022 / 2018)."""
    base = urls_espelho(url)
    extras: list[str] = []
    if url.startswith("https://web.archive.org/") or "githubusercontent.com" in url:
        return base
    for prefixo in WAYBACK_EXTRA:
        candidata = f"{prefixo}{url}"
        if candidata not in base:
            extras.append(candidata)
    return base + extras


def pasta_saida_eleicao(raiz: Path, ano: int) -> Path:
    return raiz / f"tse{ano}"


def pastas_padrao(ano: int, winpy: Path | None = None) -> tuple[Path, Path]:
    raiz = winpy if winpy is not None else descobrir_winpython()
    if raiz is not None:
        return (
            raiz / "dados" / f"tse{ano}" / "raw",
            raiz / "saida" / f"tse{ano}",
        )
    return REPO_ROOT / "data" / f"tse{ano}" / "raw", REPO_ROOT / "output" / f"tse{ano}"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--ano", type=int, choices=(2014, 2018, 2022), default=2022)
    p.add_argument("--turno", type=int, choices=(1, 2), default=1)
    p.add_argument("--ufs", nargs="+", default=list(UFS))
    p.add_argument("--raw-dir", type=Path, default=None)
    p.add_argument("--saida", type=Path, default=None)
    p.add_argument("--catalogo", type=Path, default=None)
    p.add_argument("--modelo", type=Path, default=None)
    p.add_argument("--somente-baixar", action="store_true")
    p.add_argument("--somente-processar", action="store_true")
    p.add_argument("--somente-resultado-github", action="store_true")
    p.add_argument("--apagar-zip", action="store_true", help="Apaga o ZIP após processar.")
    p.add_argument("--timeout", type=int, default=600)
    p.add_argument("--tentativas", type=int, default=2)
    p.add_argument("--usar-curl", action="store_true")
    return p.parse_args(argv)


def normalizar_ufs(ufs: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for raw in ufs:
        uf = raw.strip().upper()
        if uf not in UFS:
            raise ValueError(f"UF inválida: {raw!r}")
        if uf not in seen:
            out.append(uf)
            seen.add(uf)
    return out


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Remove aspas e espaços dos nomes (2018 tem 'SG_ UF')."""
    out = df.copy()
    novos = []
    for col in out.columns:
        nome = str(col).replace('"', "").replace(" ", "").strip()
        novos.append(nome)
    out.columns = novos
    return out


def nome_arquivo_saida(ano: int, turno: int) -> str:
    return f"urnas_{turno}t_presidente.csv"


def coluna_candidato(rotulo: str) -> str:
    return f"QT_VOTOS_{rotulo}"


def filtrar_presidente(df: pd.DataFrame, turno: int) -> pd.DataFrame:
    trabalho = normalizar_colunas(df)
    if "NR_TURNO" in trabalho.columns:
        nturno = pd.to_numeric(trabalho["NR_TURNO"], errors="coerce")
        trabalho = trabalho.loc[nturno == turno]
    if "CD_CARGO_PERGUNTA" in trabalho.columns:
        cargo = pd.to_numeric(trabalho["CD_CARGO_PERGUNTA"], errors="coerce")
        por_codigo = cargo == 1
        if por_codigo.any():
            trabalho = trabalho.loc[por_codigo]
        elif "DS_CARGO_PERGUNTA" in trabalho.columns:
            trabalho = trabalho.loc[
                trabalho["DS_CARGO_PERGUNTA"]
                .astype(str)
                .str.contains("Presidente", case=False, na=False)
            ]
    elif "DS_CARGO_PERGUNTA" in trabalho.columns:
        trabalho = trabalho.loc[
            trabalho["DS_CARGO_PERGUNTA"]
            .astype(str)
            .str.contains("Presidente", case=False, na=False)
        ]
    return trabalho


def consolidar_urnas(
    df: pd.DataFrame,
    faixas: pd.DataFrame,
    *,
    ano: int,
    turno: int,
) -> pd.DataFrame:
    base = filtrar_presidente(df, turno)
    if base.empty:
        return pd.DataFrame()

    for col in (
        "QT_VOTOS",
        "QT_APTOS",
        "QT_COMPARECIMENTO",
        "QT_ABSTENCOES",
        "NR_URNA_EFETIVADA",
        "NR_VOTAVEL",
        "CD_TIPO_VOTAVEL",
    ):
        if col in base.columns:
            base[col] = pd.to_numeric(base[col], errors="coerce")

    tipo = base.get("DS_TIPO_VOTAVEL", pd.Series("", index=base.index)).astype(str)
    nr_vot = pd.to_numeric(base.get("NR_VOTAVEL"), errors="coerce")
    cd_tipo = pd.to_numeric(base.get("CD_TIPO_VOTAVEL"), errors="coerce")
    nm = base.get("NM_VOTAVEL", pd.Series("", index=base.index)).astype(str).str.upper()
    votos = pd.to_numeric(base["QT_VOTOS"], errors="coerce").fillna(0)

    mapa = CANDIDATOS[(ano, turno)]
    for numero, rotulo in mapa.items():
        col = coluna_candidato(rotulo)
        base[col] = 0
        mask = nr_vot == numero
        base.loc[mask, col] = votos.loc[mask]

    base["QT_VOTOS_BRANCO"] = 0
    base["QT_VOTOS_NULO"] = 0
    branco = (
        tipo.str.contains("BRANCO", case=False, na=False)
        | (cd_tipo == 2)
        | (nr_vot == 95)
        | nm.str.contains("BRANCO", na=False)
    )
    nulo = (
        tipo.str.contains("NULO", case=False, na=False)
        | (cd_tipo == 3)
        | (nr_vot == 96)
        | nm.str.contains("NULO", na=False)
    )
    base.loc[branco, "QT_VOTOS_BRANCO"] = votos.loc[branco]
    base.loc[nulo, "QT_VOTOS_NULO"] = votos.loc[nulo]

    conhecidos = set(mapa) | {95, 96}
    outros_mask = nr_vot.notna() & ~nr_vot.isin(conhecidos) & ~branco & ~nulo
    base["QT_VOTOS_OUTROS"] = 0
    base.loc[outros_mask, "QT_VOTOS_OUTROS"] = votos.loc[outros_mask]

    chaves = [c for c in CHAVES_URNA if c in base.columns]
    extras = [
        c
        for c in (
            "NM_MUNICIPIO",
            "NR_LOCAL_VOTACAO",
            "QT_APTOS",
            "QT_COMPARECIMENTO",
            "QT_ABSTENCOES",
            "DT_ABERTURA",
            "DT_ENCERRAMENTO",
            "DS_TIPO_URNA",
        )
        if c in base.columns
    ]
    agg = {
        **{c: "first" for c in extras},
        **{coluna_candidato(r): "sum" for r in mapa.values()},
        "QT_VOTOS_BRANCO": "sum",
        "QT_VOTOS_NULO": "sum",
        "QT_VOTOS_OUTROS": "sum",
    }
    agrupado = base.groupby(chaves, dropna=False).agg(agg).reset_index()

    if "NR_URNA_EFETIVADA" in agrupado.columns:
        agrupado["NR_MODELO"] = classificar_modelo(
            agrupado["NR_URNA_EFETIVADA"], faixas
        )
        agrupado["DS_MODELO_URNA"] = agrupado["NR_MODELO"].map(rotulo_modelo)

    cand_cols = [coluna_candidato(r) for r in mapa.values()]
    agrupado["QT_VOTOS_VALIDOS"] = agrupado[cand_cols].sum(axis=1) + agrupado[
        "QT_VOTOS_OUTROS"
    ]
    agrupado["ANO_ELEICAO"] = ano
    agrupado["NR_TURNO"] = turno
    return alinhar_colunas_urna(agrupado, ano, turno)


def colunas_detalhe_urna(ano: int, turno: int) -> list[str]:
    cand_cols = [coluna_candidato(r) for r in CANDIDATOS[(ano, turno)].values()]
    return [
        "ANO_ELEICAO",
        "NR_TURNO",
        "SG_UF",
        "CD_MUNICIPIO",
        "NM_MUNICIPIO",
        "NR_ZONA",
        "NR_SECAO",
        "NR_LOCAL_VOTACAO",
        "NR_URNA_EFETIVADA",
        "NR_MODELO",
        "DS_MODELO_URNA",
        "QT_APTOS",
        "QT_COMPARECIMENTO",
        "QT_ABSTENCOES",
        *cand_cols,
        "QT_VOTOS_OUTROS",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_VOTOS_VALIDOS",
        "DT_ABERTURA",
        "DT_ENCERRAMENTO",
        "DS_TIPO_URNA",
    ]


def alinhar_colunas_urna(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    """Garante a mesma ordem de colunas em todas as UFs (evita CSV deslocado)."""
    if df.empty:
        return df
    ordem = [c for c in colunas_detalhe_urna(ano, turno) if c in df.columns]
    extras = [c for c in df.columns if c not in ordem]
    return df.reindex(columns=ordem + extras)


def _parece_cabecalho(linha: str) -> bool:
    alto = linha.upper()
    return "SG_UF" in alto or "DT_GERACAO" in alto or "ANO_ELEICAO" in alto


def ler_bweb_2014(fonte: Path | io.BytesIO | io.StringIO) -> pd.DataFrame:
    """TXT/CSV do BUWEB 2014: com ou sem cabeçalho, 30 campos posicionais."""
    if isinstance(fonte, Path):
        texto = fonte.read_text(encoding="latin-1")
    else:
        raw = fonte.read()
        texto = raw.decode("latin-1") if isinstance(raw, (bytes, bytearray)) else raw
    linhas = [ln for ln in texto.splitlines() if ln.strip()]
    if not linhas:
        return pd.DataFrame(columns=BWEB_2014_COLS)
    if _parece_cabecalho(linhas[0]):
        df = pd.read_csv(
            io.StringIO("\n".join(linhas)),
            sep=";",
            encoding="latin-1",
            dtype=str,
            low_memory=False,
        )
        return normalizar_colunas(df)
    registros = []
    for linha in linhas:
        partes = [p.strip().strip('"') for p in linha.split(";")]
        if len(partes) < 26:
            continue
        if len(partes) > len(BWEB_2014_COLS):
            partes = partes[: len(BWEB_2014_COLS)]
        elif len(partes) < len(BWEB_2014_COLS):
            partes = partes + [""] * (len(BWEB_2014_COLS) - len(partes))
        registros.append(partes)
    return pd.DataFrame(registros, columns=BWEB_2014_COLS)


def arquivo_dentro_do_zip(zip_path: Path) -> bytes:
    with zipfile.ZipFile(zip_path) as zf:
        nomes = [
            n
            for n in zf.namelist()
            if n.lower().endswith((".csv", ".txt")) and "leiame" not in n.lower()
        ]
        if not nomes:
            raise FileNotFoundError(f"Nenhum CSV/TXT em {zip_path}")
        nomes.sort()
        return zf.read(nomes[-1])


def _cols_uteis(nome: str) -> bool:
    limpo = str(nome).replace('"', "").replace(" ", "").strip()
    return limpo in set(BWEB_COLS)


def ler_bweb_zip(zip_path: Path, ano: int) -> pd.DataFrame:
    raw = arquivo_dentro_do_zip(zip_path)
    if ano == 2014:
        return ler_bweb_2014(io.BytesIO(raw))
    df = pd.read_csv(
        io.BytesIO(raw),
        sep=";",
        encoding="latin-1",
        dtype=str,
        usecols=_cols_uteis,
        low_memory=False,
    )
    return normalizar_colunas(df)


def _nome_dados_no_zip(zf: zipfile.ZipFile) -> str:
    nomes = [
        n
        for n in zf.namelist()
        if n.lower().endswith((".csv", ".txt")) and "leiame" not in n.lower()
    ]
    if not nomes:
        raise FileNotFoundError("Nenhum CSV/TXT no ZIP")
    nomes.sort()
    return nomes[-1]


def _partes_linha_2014(linha: str) -> list[str] | None:
    partes = [p.strip().strip('"') for p in linha.split(";")]
    if len(partes) < 26:
        return None
    if len(partes) > len(BWEB_2014_COLS):
        partes = partes[: len(BWEB_2014_COLS)]
    elif len(partes) < len(BWEB_2014_COLS):
        partes = partes + [""] * (len(BWEB_2014_COLS) - len(partes))
    return partes


def processar_zip_bweb_2014(
    zip_path: Path, faixas: pd.DataFrame, *, turno: int, chunk: int = 80_000
) -> pd.DataFrame:
    """Processa BU 2014 em lotes para não carregar o TXT inteiro de SP."""
    if zip_path.suffix.lower() == ".txt":
        texto = zip_path.read_text(encoding="latin-1")
    else:
        raw = arquivo_dentro_do_zip(zip_path)
        texto = raw.decode("latin-1")
        del raw
    linhas = [ln for ln in texto.splitlines() if ln.strip()]
    del texto
    if linhas and _parece_cabecalho(linhas[0]):
        linhas = linhas[1:]
    lotes: list[pd.DataFrame] = []
    buf: list[list[str]] = []
    for linha in linhas:
        rec = _partes_linha_2014(linha)
        if rec is None:
            continue
        buf.append(rec)
        if len(buf) >= chunk:
            df = pd.DataFrame(buf, columns=BWEB_2014_COLS)
            lotes.append(consolidar_urnas(df, faixas, ano=2014, turno=turno))
            buf = []
    del linhas
    if buf:
        df = pd.DataFrame(buf, columns=BWEB_2014_COLS)
        lotes.append(consolidar_urnas(df, faixas, ano=2014, turno=turno))
    if not lotes:
        return pd.DataFrame()
    if len(lotes) == 1:
        return alinhar_colunas_urna(lotes[0], 2014, turno)
    junto = pd.concat(lotes, ignore_index=True)
    chaves = [c for c in CHAVES_URNA if c in junto.columns]
    soma = [c for c in junto.columns if str(c).startswith("QT_VOTOS_")]
    primeiro = [c for c in junto.columns if c not in chaves and c not in soma]
    agg = {c: "first" for c in primeiro}
    agg.update({c: "sum" for c in soma})
    out = junto.groupby(chaves, dropna=False).agg(agg).reset_index()
    return alinhar_colunas_urna(out, 2014, turno)


def processar_zip_bweb(
    zip_path: Path, faixas: pd.DataFrame, *, ano: int, turno: int
) -> pd.DataFrame:
    """Lê o ZIP em pedaços (1º turno de SP passa de 1 GB descompactado)."""
    if ano == 2014:
        return processar_zip_bweb_2014(zip_path, faixas, turno=turno)

    partes: list[pd.DataFrame] = []
    with zipfile.ZipFile(zip_path) as zf:
        nome = _nome_dados_no_zip(zf)
        with zf.open(nome) as fh:
            leitor = pd.read_csv(
                fh,
                sep=";",
                encoding="latin-1",
                dtype=str,
                usecols=_cols_uteis,
                chunksize=250_000,
                low_memory=False,
            )
            for chunk in leitor:
                filtrado = filtrar_presidente(chunk, turno)
                if not filtrado.empty:
                    partes.append(filtrado)
    if not partes:
        return pd.DataFrame()
    return consolidar_urnas(
        pd.concat(partes, ignore_index=True), faixas, ano=ano, turno=turno
    )


def resumo_por_uf(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    cand_cols = [
        coluna_candidato(r) for r in CANDIDATOS[(ano, turno)].values() if coluna_candidato(r) in df.columns
    ]
    agg = {
        "QT_URNAS": ("NR_URNA_EFETIVADA", "nunique"),
        "QT_SECOES": ("NR_SECAO", "size"),
        **{c: (c, "sum") for c in cand_cols},
        "QT_VOTOS_BRANCO": ("QT_VOTOS_BRANCO", "sum"),
        "QT_VOTOS_NULO": ("QT_VOTOS_NULO", "sum"),
        "QT_VOTOS_VALIDOS": ("QT_VOTOS_VALIDOS", "sum"),
    }
    return (
        df.groupby("SG_UF", dropna=False)
        .agg(**{k: v for k, v in agg.items() if v[0] in df.columns or k in ("QT_URNAS", "QT_SECOES")})
        .reset_index()
        .sort_values("SG_UF")
    )


def resumo_por_modelo(df: pd.DataFrame, ano: int, turno: int) -> pd.DataFrame:
    cand_cols = [
        coluna_candidato(r)
        for r in CANDIDATOS[(ano, turno)].values()
        if coluna_candidato(r) in df.columns
    ]
    if "DS_MODELO_URNA" not in df.columns:
        return pd.DataFrame()
    trabalho = df.copy()
    for c in cand_cols + ["QT_VOTOS_VALIDOS"]:
        if c in trabalho.columns:
            trabalho[c] = pd.to_numeric(trabalho[c], errors="coerce").fillna(0)
    g = (
        trabalho.groupby(["NR_MODELO", "DS_MODELO_URNA"], dropna=False)
        .agg(
            QT_URNAS=("NR_URNA_EFETIVADA", "nunique"),
            **{c: (c, "sum") for c in cand_cols},
            QT_VOTOS_VALIDOS=("QT_VOTOS_VALIDOS", "sum"),
        )
        .reset_index()
        .sort_values("NR_MODELO")
    )
    return g


def localizar_zip(raw_dir: Path, arquivo: str, uf: str, turno: int) -> Path | None:
    direto = raw_dir / arquivo
    if direto.exists():
        return direto
    padroes = (
        f"bweb_{turno}t_{uf}_*.zip",
        f"BWEB_{turno}t_{uf}_*.zip",
        f"*{turno}t_{uf}_*.zip",
        f"bweb_{turno}t_{uf}_*.txt",
        f"BWEB_{turno}t_{uf}_*.txt",
    )
    for padrao in padroes:
        achados = sorted(raw_dir.glob(padrao))
        if achados:
            return achados[0]
    return None


def baixar_zip_uf(
    recurso: dict,
    dest: Path,
    *,
    timeout: int,
    tentativas: int,
    usar_curl: bool,
) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size >= 1000 and zipfile.is_zipfile(dest):
        return dest
    erros: list[str] = []
    for url in urls_espelho_historico(recurso["url"]):
        try:
            print(f"    {url}", flush=True)
            baixar_arquivo(
                url,
                dest,
                timeout=timeout,
                tentativas=tentativas,
                usar_curl=usar_curl,
            )
            if dest.exists() and dest.stat().st_size >= 1000 and zipfile.is_zipfile(dest):
                return dest
            if dest.exists():
                dest.unlink(missing_ok=True)
            erros.append(f"{url}: arquivo inválido")
        except Exception as exc:
            erros.append(f"{url}: {exc}")
            if dest.exists():
                dest.unlink(missing_ok=True)
    raise RuntimeError(" | ".join(erros[:6]))


def gravar_csv_e_gz(df: pd.DataFrame, csv_path: Path) -> Path:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(csv_path, index=False, encoding="utf-8")
    gz_path = csv_path.with_suffix(csv_path.suffix + ".gz")
    with csv_path.open("rb") as src, gzip.open(gz_path, "wb") as dst:
        dst.write(src.read())
    return gz_path


def baixar_resultado_github(
    saida_raiz: Path, *, timeout: int = 180, pares: list[tuple[int, int]] | None = None
) -> list[Path]:
    """Baixa CSVs já consolidados (GitHub raw — funciona na RFB)."""
    saida_raiz.mkdir(parents=True, exist_ok=True)
    obtidos: list[Path] = []
    alvos = pares or list(ARQUIVOS_GITHUB)
    rels: list[str] = []
    for chave in alvos:
        rels.append(ARQUIVOS_GITHUB[chave])
        rels.extend(ARQUIVOS_GITHUB_EXTRA.get(chave, ()))
    for rel in rels:
        dest = saida_raiz / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        url = RESULTADO_GITHUB_BASE + rel
        print(f"  github {rel}", flush=True)
        try:
            baixar_arquivo(url, dest, timeout=timeout, tentativas=1, usar_curl=True)
        except Exception as exc:
            print(f"  aviso: {rel} ainda não está no GitHub ({exc})", flush=True)
            continue
        if dest.name.endswith(".gz"):
            csv = dest.with_suffix("")
            with gzip.open(dest, "rb") as src, csv.open("wb") as dst:
                dst.write(src.read())
            obtidos.append(csv)
        else:
            obtidos.append(dest)
    return obtidos


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    ano = args.ano
    turno = args.turno
    ufs = normalizar_ufs(args.ufs)
    raw_default, saida_default = pastas_padrao(ano)
    raw_dir = Path(args.raw_dir) if args.raw_dir else raw_default
    saida = Path(args.saida) if args.saida else saida_default
    raw_dir.mkdir(parents=True, exist_ok=True)
    saida.mkdir(parents=True, exist_ok=True)

    if args.somente_resultado_github:
        raiz = saida.parent if saida.name.startswith("tse") else saida
        caminhos = baixar_resultado_github(
            raiz, timeout=args.timeout, pares=[(ano, turno)]
        )
        for path in caminhos:
            print(f"Baixado: {path}", flush=True)
        return 0 if caminhos else 2

    catalogo = carregar_catalogo(args.catalogo)
    print(f"Eleição {ano} {turno}º turno — {len(ufs)} UFs", flush=True)
    print(f"ZIPs : {raw_dir}", flush=True)
    print(f"Saida: {saida}", flush=True)

    if not args.somente_processar:
        erros: list[str] = []
        for uf in ufs:
            rec = recurso_catalogo(catalogo, ano, turno, uf)
            dest = raw_dir / rec["arquivo"]
            print(f"  baixando {uf} {rec['arquivo']}", flush=True)
            try:
                baixar_zip_uf(
                    rec,
                    dest,
                    timeout=args.timeout,
                    tentativas=args.tentativas,
                    usar_curl=args.usar_curl,
                )
                print(f"  ok {uf} ({dest.stat().st_size:,} bytes)", flush=True)
            except Exception as exc:
                erros.append(f"{uf}: {exc}")
                print(f"  ERRO {uf}: {exc}", flush=True)
        if erros:
            print(
                "\nFalha em downloads. Na RFB use:\n"
                "  python baixar_boletins_urna.py --somente-resultado-github "
                f"--ano {ano} --turno {turno}\n",
                flush=True,
            )
            if args.somente_baixar:
                return 2

    if args.somente_baixar:
        return 0

    faixas = carregar_faixas_modelo(args.modelo)
    partes: list[pd.DataFrame] = []
    faltando: list[str] = []
    for uf in ufs:
        rec = recurso_catalogo(catalogo, ano, turno, uf)
        zip_path = localizar_zip(raw_dir, rec["arquivo"], uf, turno)
        if zip_path is None:
            faltando.append(uf)
            continue
        print(f"  processando {uf} ← {zip_path.name}", flush=True)
        partes.append(processar_zip_bweb(zip_path, faixas, ano=ano, turno=turno))
        if args.apagar_zip:
            zip_path.unlink(missing_ok=True)

    if faltando and not partes:
        raise FileNotFoundError(
            "ZIP ausente para: " + ", ".join(faltando) + f" em {raw_dir}"
        )
    if faltando:
        print(f"Aviso: sem ZIP para {', '.join(faltando)}", flush=True)

    tabela = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if tabela.empty:
        raise RuntimeError("Nenhuma urna processada.")
    tabela = alinhar_colunas_urna(tabela, ano, turno)

    detalhe = saida / nome_arquivo_saida(ano, turno)
    gz = gravar_csv_e_gz(tabela, detalhe)
    resumo_por_uf(tabela, ano, turno).to_csv(
        saida / f"resumo_{turno}t_por_uf.csv", index=False, encoding="utf-8"
    )
    por_modelo = resumo_por_modelo(tabela, ano, turno)
    if not por_modelo.empty:
        por_modelo.to_csv(
            saida / f"resumo_{turno}t_por_modelo.csv", index=False, encoding="utf-8"
        )

    print()
    print(f"Urnas/seções: {len(tabela):,}".replace(",", "."))
    if "NR_URNA_EFETIVADA" in tabela.columns:
        print(
            "Série preenchida: "
            f"{tabela['NR_URNA_EFETIVADA'].notna().sum():,}".replace(",", ".")
        )
    for rotulo in CANDIDATOS[(ano, turno)].values():
        col = coluna_candidato(rotulo)
        if col in tabela.columns:
            print(f"{rotulo}: {int(tabela[col].sum()):,}".replace(",", "."))
    print(f"Detalhe: {detalhe}")
    print(f"Gzip: {gz}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
