#!/usr/bin/env python3
"""Saldos de fim de ano dos fatores condicionantes da base monetária (Bacen SGS).

Baixa as séries oficiais de *saldo em final de período* no SGS e extrai a
última observação de cada ano civil (dezembro, ou o último mês disponível
no ano corrente).

Unidade original do SGS: milhares de unidades monetárias correntes (R$ mil).
As tabelas de divulgação usam R$ milhões (valor SGS ÷ 1.000).

Fonte: Banco Central do Brasil — SGS / Portal de Dados Abertos.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.calcular_diretas_ipca_selic import _baixar_sgs  # noqa: E402

ANO_INICIO_DEFAULT = 1995
ANO_FIM_DEFAULT = 2026
SGS_API = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados"

# Ordem da Nota de Estatísticas Monetárias (fatores + resultado).
SERIES: list[dict] = [
    {
        "codigo": 1810,
        "coluna": "tesouro_conta_unica",
        "nome": "Tesouro Nacional — Conta única",
        "grupo": "fator",
    },
    {
        "codigo": 1809,
        "coluna": "titulos_publicos_total",
        "nome": "Operações com títulos públicos federais — Total",
        "grupo": "fator",
    },
    {
        "codigo": 29004,
        "coluna": "titulos_mercado_primario",
        "nome": "Títulos públicos — mercado primário",
        "grupo": "desdobramento",
    },
    {
        "codigo": 29006,
        "coluna": "titulos_mercado_secundario",
        "nome": "Títulos públicos — mercado secundário",
        "grupo": "desdobramento",
    },
    {
        "codigo": 1811,
        "coluna": "setor_externo",
        "nome": "Operações com o setor externo",
        "grupo": "fator",
    },
    {
        "codigo": 12487,
        "coluna": "derivativos_ajustes",
        "nome": "Operações com derivativos — ajustes",
        "grupo": "fator",
    },
    {
        "codigo": 12484,
        "coluna": "redesconto",
        "nome": "Redesconto do Banco Central",
        "grupo": "fator",
    },
    {
        "codigo": 28724,
        "coluna": "linhas_temporarias_liquidez",
        "nome": "Linhas temporárias especiais de liquidez",
        "grupo": "fator",
    },
    {
        "codigo": 1815,
        "coluna": "depositos_instituicoes_financeiras",
        "nome": "Depósitos de instituições financeiras",
        "grupo": "fator",
    },
    {
        "codigo": 1818,
        "coluna": "outras_operacoes",
        "nome": "Autoridade Monetária — Outras operações",
        "grupo": "fator",
    },
    {
        "codigo": 1788,
        "coluna": "base_monetaria_restrita",
        "nome": "Base monetária restrita (resultado)",
        "grupo": "resultado",
    },
]


def anos_periodo(ano_inicio: int, ano_fim: int) -> list[int]:
    if ano_fim < ano_inicio:
        raise ValueError("ano_fim deve ser >= ano_inicio")
    return list(range(ano_inicio, ano_fim + 1))


def baixar_serie(
    codigo: int,
    inicio: str = "01/01/1995",
    fim: str | None = None,
    baixar=_baixar_sgs,
) -> pd.DataFrame:
    """Baixa uma série SGS e devolve colunas data, valor (R$ mil)."""
    bruto = baixar(codigo, inicio=inicio, fim=fim)
    out = bruto.rename(columns={"mes": "data"}).copy()
    if "data" not in out.columns:
        raise RuntimeError(f"Série SGS {codigo} sem coluna de data")
    out["data"] = pd.to_datetime(out["data"], errors="coerce")
    out["valor"] = pd.to_numeric(out["valor"], errors="coerce")
    out = out.dropna(subset=["data", "valor"]).sort_values("data")
    out["codigo"] = codigo
    return out[["codigo", "data", "valor"]].reset_index(drop=True)


def saldo_fim_de_ano(df: pd.DataFrame, anos: list[int]) -> pd.DataFrame:
    """Última observação de cada ano civil (dezembro quando a série chega lá)."""
    if df.empty:
        return pd.DataFrame(
            columns=["codigo", "ano", "data", "valor_rs_mil", "fechamento"]
        )
    work = df.copy()
    work["ano"] = work["data"].dt.year
    work = work[work["ano"].isin(anos)]
    if work.empty:
        return pd.DataFrame(
            columns=["codigo", "ano", "data", "valor_rs_mil", "fechamento"]
        )
    last = (
        work.sort_values("data")
        .groupby("ano", as_index=False)
        .tail(1)
        .reset_index(drop=True)
    )
    last["fechamento"] = last["data"].map(_rotulo_fechamento)
    last = last.rename(columns={"valor": "valor_rs_mil"})
    return last[["codigo", "ano", "data", "valor_rs_mil", "fechamento"]]


def _rotulo_fechamento(ts: pd.Timestamp) -> str:
    if ts.month == 12:
        return "fim_de_ano"
    return f"ultimo_disponivel_{ts.strftime('%m/%Y')}"


def montar_tabelas(
    series_long: dict[int, pd.DataFrame],
    anos: list[int],
    catalogo: list[dict] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Devolve (tabela longa, tabela larga em R$ milhões)."""
    catalogo = catalogo if catalogo is not None else SERIES
    partes: list[pd.DataFrame] = []
    for meta in catalogo:
        codigo = int(meta["codigo"])
        bruto = series_long.get(codigo)
        if bruto is None or bruto.empty:
            continue
        anual = saldo_fim_de_ano(bruto, anos)
        if anual.empty:
            continue
        anual["serie"] = meta["nome"]
        anual["coluna"] = meta["coluna"]
        anual["grupo"] = meta["grupo"]
        anual["valor_rs_milhoes"] = anual["valor_rs_mil"] / 1_000.0
        partes.append(anual)
    if not partes:
        raise RuntimeError("Nenhuma série SGS retornou saldo no período")
    longo = pd.concat(partes, ignore_index=True)
    longo = longo.sort_values(["ano", "codigo"]).reset_index(drop=True)

    largos = []
    for meta in catalogo:
        sub = longo.loc[longo["codigo"] == meta["codigo"], ["ano", "valor_rs_milhoes"]]
        if sub.empty:
            continue
        pivot = sub.rename(columns={"valor_rs_milhoes": meta["coluna"]})
        largos.append(pivot)
    largo = largos[0]
    for extra in largos[1:]:
        largo = largo.merge(extra, on="ano", how="outer")
    largo = largo.sort_values("ano").reset_index(drop=True)

    refs = (
        longo.groupby("ano", as_index=False)["data"]
        .max()
        .rename(columns={"data": "data_saldo"})
    )
    status = (
        longo.groupby("ano")["fechamento"]
        .agg(lambda s: "fim_de_ano" if (s == "fim_de_ano").all() else "parcial")
        .reset_index()
    )
    largo = largo.merge(refs, on="ano").merge(status, on="ano")
    cols = ["ano", "data_saldo", "fechamento"] + [
        m["coluna"] for m in catalogo if m["coluna"] in largo.columns
    ]
    return longo, largo[cols]


def formatar_milhoes(valor: float | None) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def markdown_tabela(largo: pd.DataFrame, catalogo: list[dict] | None = None) -> str:
    catalogo = catalogo if catalogo is not None else SERIES
    linhas = [
        "# Saldos dos fatores condicionantes da base monetária",
        "",
        "Fonte: Banco Central do Brasil, SGS (saldo em final de período).",
        "Unidade: **R$ milhões** (série original em R$ mil ÷ 1.000).",
        "Cada célula é o **último saldo do ano civil** (dezembro, ou o último mês disponível).",
        "",
        "| Ano | Data do saldo | "
        + " | ".join(m["nome"] for m in catalogo if m["coluna"] in largo.columns)
        + " |",
        "|-----|---------------|"
        + "|".join(["---:"] * sum(1 for m in catalogo if m["coluna"] in largo.columns))
        + "|",
    ]
    for _, row in largo.iterrows():
        data = pd.Timestamp(row["data_saldo"]).strftime("%d/%m/%Y")
        vals = [
            formatar_milhoes(row[m["coluna"]])
            for m in catalogo
            if m["coluna"] in largo.columns
        ]
        linhas.append(f"| {int(row['ano'])} | {data} | " + " | ".join(vals) + " |")
    linhas.extend(
        [
            "",
            "## Códigos SGS",
            "",
            "| Código | Série | Grupo |",
            "|-------:|-------|-------|",
        ]
    )
    for m in catalogo:
        linhas.append(f"| {m['codigo']} | {m['nome']} | {m['grupo']} |")
    linhas.extend(
        [
            "",
            "API: `https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados`.",
            "O ano corrente pode ainda não ter dezembro; nesse caso o saldo é o último disponível.",
            "",
        ]
    )
    return "\n".join(linhas)


def gravar_saidas(
    longo: pd.DataFrame,
    largo: pd.DataFrame,
    pasta: Path,
    stem: str = "fatores_condicionantes_base_monetaria",
) -> dict[str, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    csv_longo = pasta / f"{stem}_longo.csv"
    csv_largo = pasta / f"{stem}_anual_r$_milhoes.csv"
    xlsx = pasta / f"{stem}_anual.xlsx"
    md = pasta / f"{stem}_anual.md"

    longo_out = longo.copy()
    longo_out["data"] = pd.to_datetime(longo_out["data"]).dt.strftime("%Y-%m-%d")
    longo_out.to_csv(csv_longo, index=False)

    largo_out = largo.copy()
    largo_out["data_saldo"] = pd.to_datetime(largo_out["data_saldo"]).dt.strftime(
        "%Y-%m-%d"
    )
    largo_out.to_csv(csv_largo, index=False)

    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        largo_out.to_excel(writer, sheet_name="Anual_R$_milhoes", index=False)
        longo_out.to_excel(writer, sheet_name="Longo_SGS", index=False)
        pd.DataFrame(SERIES).to_excel(writer, sheet_name="Codigos_SGS", index=False)

    md.write_text(markdown_tabela(largo), encoding="utf-8")
    return {
        "csv_longo": csv_longo,
        "csv_largo": csv_largo,
        "xlsx": xlsx,
        "md": md,
    }


def coletar(
    ano_inicio: int = ANO_INICIO_DEFAULT,
    ano_fim: int = ANO_FIM_DEFAULT,
    baixar=_baixar_sgs,
    catalogo: list[dict] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    catalogo = catalogo if catalogo is not None else SERIES
    anos = anos_periodo(ano_inicio, ano_fim)
    inicio = f"01/01/{ano_inicio}"
    fim = datetime.now().strftime("%d/%m/%Y")
    series_long: dict[int, pd.DataFrame] = {}
    for meta in catalogo:
        codigo = int(meta["codigo"])
        print(f"[INFO] SGS {codigo} — {meta['nome']}")
        try:
            series_long[codigo] = baixar_serie(
                codigo, inicio=inicio, fim=fim, baixar=baixar
            )
        except Exception as exc:  # séries curtas/novas podem falhar no bloco inicial
            print(f"[AVISO] SGS {codigo} indisponível: {exc}")
            series_long[codigo] = pd.DataFrame(columns=["codigo", "data", "valor"])
    return montar_tabelas(series_long, anos, catalogo=catalogo)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Saldos de fim de ano dos fatores condicionantes da base monetária "
            "(Bacen SGS)."
        )
    )
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO_DEFAULT)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM_DEFAULT)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "output",
        help="Pasta das tabelas geradas",
    )
    parser.add_argument(
        "--stem",
        default="fatores_condicionantes_base_monetaria",
        help="Prefixo dos arquivos de saída",
    )
    args = parser.parse_args(argv)

    longo, largo = coletar(args.ano_inicio, args.ano_fim)
    caminhos = gravar_saidas(longo, largo, args.output_dir, stem=args.stem)
    print()
    print(markdown_tabela(largo))
    print("[OK] Arquivos:")
    for nome, path in caminhos.items():
        print(f"  {nome}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
