#!/usr/bin/env python3
"""Índices oficiais de inflação — Brasil e Estados Unidos (variação no ano).

Critério: variação dezembro/dezembro, o mesmo conceito da “inflação no ano”
do IPCA (IBGE) e da variação de 12 meses do CPI-U (BLS).

Brasil:
  IPCA   IBGE / Bacen SGS 13522 (acumulado em 12 meses, dezembro)
  INPC   IBGE / Bacen SGS 188   (produto das 12 variações mensais)
  IGP-M  FGV  / Bacen SGS 189   (produto das 12 variações mensais)
  IGP-DI FGV  / Bacen SGS 190   (produto das 12 variações mensais)

Estados Unidos (índice de preços → Dez_t / Dez_{t-1} − 1):
  CPI-U  BLS / FRED CPIAUCNS (NSA)
  PCE    BEA / FRED PCEPI
"""

from __future__ import annotations

import argparse
import io
import sys
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

ANO_INICIO_DEFAULT = 1990
ANO_FIM_DEFAULT = 2025
SGS_API = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados"
FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}"

SERIES_BR: list[dict] = [
    {
        "coluna": "ipca",
        "nome": "IPCA",
        "pais": "Brasil",
        "orgao": "IBGE",
        "codigo": 13522,
        "tipo": "percentual_12m",
    },
    {
        "coluna": "inpc",
        "nome": "INPC",
        "pais": "Brasil",
        "orgao": "IBGE",
        "codigo": 188,
        "tipo": "percentual_mensal",
    },
    {
        "coluna": "igp_m",
        "nome": "IGP-M",
        "pais": "Brasil",
        "orgao": "FGV",
        "codigo": 189,
        "tipo": "percentual_mensal",
    },
    {
        "coluna": "igp_di",
        "nome": "IGP-DI",
        "pais": "Brasil",
        "orgao": "FGV",
        "codigo": 190,
        "tipo": "percentual_mensal",
    },
]

SERIES_US: list[dict] = [
    {
        "coluna": "cpi_u",
        "nome": "CPI-U",
        "pais": "Estados Unidos",
        "orgao": "BLS",
        "codigo": "CPIAUCNS",
        "tipo": "indice_precos",
    },
    {
        "coluna": "pce",
        "nome": "PCE",
        "pais": "Estados Unidos",
        "orgao": "BEA",
        "codigo": "PCEPI",
        "tipo": "indice_precos",
    },
]

SERIES = SERIES_BR + SERIES_US
COLUNAS_ORDEM = [s["coluna"] for s in SERIES]


def anos_periodo(ano_inicio: int, ano_fim: int) -> list[int]:
    if ano_fim < ano_inicio:
        raise ValueError("ano_fim deve ser >= ano_inicio")
    return list(range(ano_inicio, ano_fim + 1))


def _get_json(http, url: str, params: dict, tentativas: int = 4):
    ultimo_erro: Exception | None = None
    for tentativa in range(tentativas):
        resp = http.get(url, params=params, timeout=120)
        if resp.status_code == 404:
            return []
        corpo = getattr(resp, "content", b"x") or b""
        if resp.status_code in {429, 500, 502, 503} or not corpo.strip():
            ultimo_erro = RuntimeError(f"HTTP {resp.status_code} vazio")
            time.sleep(1.5 * (tentativa + 1))
            continue
        resp.raise_for_status()
        try:
            dados = resp.json()
        except ValueError as exc:
            ultimo_erro = exc
            time.sleep(1.5 * (tentativa + 1))
            continue
        return dados or []
    if ultimo_erro:
        raise RuntimeError(f"Falha ao ler {url}: {ultimo_erro}") from ultimo_erro
    return []


def baixar_sgs(
    codigo: int,
    inicio: str = "01/01/1989",
    fim: str = "31/12/2025",
    session: requests.Session | None = None,
) -> pd.DataFrame:
    http = session or requests
    dados = _get_json(
        http,
        SGS_API.format(cod=codigo),
        {
            "formato": "json",
            "dataInicial": inicio,
            "dataFinal": fim,
        },
    )
    if not dados:
        return pd.DataFrame(columns=["data", "valor"])
    df = pd.DataFrame(dados)
    df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    return (
        df.dropna(subset=["data", "valor"])
        .drop_duplicates(subset=["data"])
        .sort_values("data")
        .reset_index(drop=True)[["data", "valor"]]
    )


def baixar_fred(
    series_id: str,
    session: requests.Session | None = None,
) -> pd.DataFrame:
    http = session or requests
    ultimo_erro: Exception | None = None
    for tentativa in range(4):
        resp = http.get(FRED_CSV.format(sid=series_id), timeout=120)
        if resp.status_code in {429, 500, 502, 503} or not (resp.content or b"").strip():
            ultimo_erro = RuntimeError(f"FRED HTTP {resp.status_code}")
            time.sleep(1.5 * (tentativa + 1))
            continue
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        if df.empty or len(df.columns) < 2:
            ultimo_erro = RuntimeError("FRED CSV vazio")
            time.sleep(1.5 * (tentativa + 1))
            continue
        df = df.rename(columns={df.columns[0]: "data", df.columns[1]: "valor"})
        df["data"] = pd.to_datetime(df["data"], errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
        return (
            df.dropna(subset=["data", "valor"])
            .drop_duplicates(subset=["data"])
            .sort_values("data")
            .reset_index(drop=True)[["data", "valor"]]
        )
    raise RuntimeError(f"Falha ao ler FRED {series_id}: {ultimo_erro}")


def variacao_ano_percentual_mensal(df: pd.DataFrame, ano: int) -> float | None:
    """Π (1 + r_m/100) − 1 com os 12 meses do ano."""
    meses = df.loc[df["data"].dt.year == ano].sort_values("data")
    if len(meses) != 12:
        return None
    fator = (1.0 + meses["valor"].astype(float) / 100.0).prod()
    return float(fator - 1.0)


def variacao_ano_indice(df: pd.DataFrame, ano: int) -> float | None:
    """I(dezembro do ano) / I(dezembro do ano anterior) − 1."""
    dez = df.loc[(df["data"].dt.year == ano) & (df["data"].dt.month == 12), "valor"]
    ant = df.loc[(df["data"].dt.year == ano - 1) & (df["data"].dt.month == 12), "valor"]
    if dez.empty or ant.empty:
        return None
    base = float(ant.iloc[0])
    if base == 0:
        return None
    return float(dez.iloc[0]) / base - 1.0


def variacao_ano_percentual_12m(df: pd.DataFrame, ano: int) -> float | None:
    """Valor de dezembro da série já acumulada em 12 meses (unidade: %)."""
    dez = df.loc[(df["data"].dt.year == ano) & (df["data"].dt.month == 12), "valor"]
    if dez.empty:
        return None
    return float(dez.iloc[0]) / 100.0


def serie_anual(df: pd.DataFrame, tipo: str, anos: list[int]) -> dict[int, float]:
    out: dict[int, float] = {}
    for ano in anos:
        if tipo == "percentual_mensal":
            valor = variacao_ano_percentual_mensal(df, ano)
        elif tipo == "percentual_12m":
            valor = variacao_ano_percentual_12m(df, ano)
        elif tipo == "indice_precos":
            valor = variacao_ano_indice(df, ano)
        else:
            raise ValueError(f"tipo desconhecido: {tipo}")
        if valor is not None:
            out[ano] = valor
    return out


def montar_tabela(
    series_mensais: dict[str, pd.DataFrame],
    anos: list[int],
    catalogo: list[dict] | None = None,
) -> pd.DataFrame:
    catalogo = catalogo if catalogo is not None else SERIES
    linhas = []
    for ano in anos:
        row: dict = {"ano": ano}
        for meta in catalogo:
            mensal = series_mensais.get(meta["coluna"])
            if mensal is None or mensal.empty:
                row[meta["coluna"]] = None
                continue
            anual = serie_anual(mensal, meta["tipo"], [ano])
            row[meta["coluna"]] = anual.get(ano)
        linhas.append(row)
    return pd.DataFrame(linhas)


def formatar_pct(valor: float | None, casas: int = 2) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    numero = float(valor) * 100.0
    texto = f"{numero:,.{casas}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".") + "%"


def markdown_tabela(
    df: pd.DataFrame,
    ano_inicio: int = ANO_INICIO_DEFAULT,
    ano_fim: int = ANO_FIM_DEFAULT,
) -> str:
    linhas = [
        f"# Índices oficiais de inflação — Brasil e Estados Unidos ({ano_inicio}–{ano_fim})",
        "",
        "Variação no ano: **dezembro/dezembro**.",
        "",
        "| Ano | IPCA (BR) | INPC (BR) | IGP-M (BR) | IGP-DI (BR) | CPI-U (EUA) | PCE (EUA) |",
        "|----:|----------:|----------:|-----------:|------------:|------------:|----------:|",
    ]
    for _, r in df.iterrows():
        linhas.append(
            f"| {int(r['ano'])} | "
            f"{formatar_pct(r.get('ipca'))} | "
            f"{formatar_pct(r.get('inpc'))} | "
            f"{formatar_pct(r.get('igp_m'))} | "
            f"{formatar_pct(r.get('igp_di'))} | "
            f"{formatar_pct(r.get('cpi_u'))} | "
            f"{formatar_pct(r.get('pce'))} |"
        )
    linhas.extend(
        [
            "",
            "## Fontes e conceito",
            "",
            "- **IPCA** — IBGE, índice oficial de inflação ao consumidor e meta "
            "do regime de metas (Bacen SGS 13522, acumulado em 12 meses de dezembro).",
            "- **INPC** — IBGE, índice oficial usado em reajustes de salários e "
            "benefícios (Bacen SGS 188).",
            "- **IGP-M** e **IGP-DI** — FGV, índices gerais de preços oficiais "
            "da Fundação (contratos, aluguéis, dívida); Bacen SGS 189 e 190.",
            "- **CPI-U** — BLS, *Consumer Price Index for All Urban Consumers*, "
            "índice oficial de inflação ao consumidor dos EUA (FRED `CPIAUCNS`, "
            "série sem ajuste sazonal).",
            "- **PCE** — BEA, *Personal Consumption Expenditures Price Index*, "
            "deflator oficial do consumo e medida preferida do Fed (FRED `PCEPI`).",
            "",
            "IPCA: valor oficial de dezembro da série acumulada em 12 meses.",
            "INPC, IGP-M e IGP-DI: produto das 12 variações mensais do ano, "
            "`Π (1 + r_m/100) − 1` (arredondamento de 0,01 p.p. possível).",
            "EUA: `índice de dezembro ÷ índice de dezembro do ano anterior − 1`.",
            "",
            "Os anos 1990–1994 no Brasil cobrem a hiperinflação e o Plano Real "
            "(1º de julho de 1994). A partir de 1995 a unidade é o real.",
            "",
        ]
    )
    return "\n".join(linhas)


def tabela_percentual(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in COLUNAS_ORDEM:
        if col in out.columns:
            out[col] = out[col].map(lambda v: None if v is None or pd.isna(v) else v * 100.0)
    return out


def gravar(df: pd.DataFrame, pasta: Path, stem: str = "indices_inflacao_eua_brasil_1990_2025") -> dict[str, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    pct = tabela_percentual(df)
    csv = pasta / f"{stem}.csv"
    xlsx = pasta / f"{stem}.xlsx"
    md = pasta / f"{stem}.md"
    pct.to_csv(csv, index=False, float_format="%.4f")
    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        pct.to_excel(writer, index=False, sheet_name="Variacao_dez_dez_%")
        fontes = pd.DataFrame(
            [
                {
                    "coluna": s["coluna"],
                    "indice": s["nome"],
                    "pais": s["pais"],
                    "orgao": s["orgao"],
                    "codigo": s["codigo"],
                    "tipo": s["tipo"],
                }
                for s in SERIES
            ]
        )
        fontes.to_excel(writer, index=False, sheet_name="Fontes")
    md.write_text(
        markdown_tabela(
            df,
            int(df["ano"].min()) if not df.empty else ANO_INICIO_DEFAULT,
            int(df["ano"].max()) if not df.empty else ANO_FIM_DEFAULT,
        ),
        encoding="utf-8",
    )
    return {"csv": csv, "xlsx": xlsx, "md": md}


def baixar_todas(
    session: requests.Session | None = None,
    inicio_br: str = "01/01/1989",
    fim_br: str = "31/12/2025",
) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    for meta in SERIES_BR:
        out[meta["coluna"]] = baixar_sgs(
            int(meta["codigo"]), inicio=inicio_br, fim=fim_br, session=session
        )
    for meta in SERIES_US:
        out[meta["coluna"]] = baixar_fred(str(meta["codigo"]), session=session)
    return out


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Índices oficiais de inflação EUA e Brasil (dez/dez)."
    )
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO_DEFAULT)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM_DEFAULT)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "output")
    args = parser.parse_args(argv)

    anos = anos_periodo(args.ano_inicio, args.ano_fim)
    series = baixar_todas()
    tabela = montar_tabela(series, anos)
    caminhos = gravar(tabela, args.output_dir)
    texto = markdown_tabela(tabela, args.ano_inicio, args.ano_fim)
    print(texto)
    for nome, path in caminhos.items():
        print(f"[OK] {nome}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
