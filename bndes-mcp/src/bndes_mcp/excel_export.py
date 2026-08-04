"""Exporta operacoes BNDES para Excel."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .excel_format import autosize_dataframe_sheet
from .providers import expand_subcreditos, flatten_operacao, summarize


OPERACAO_COLS = [
    "cliente",
    "documentoCliente",
    "numeroContrato",
    "numeroOperacao",
    "dataContratacao",
    "anoPosicao",
    "descricaoProjeto",
    "valorContratacao",
    "valorDesembolsado",
    "faixaValorContratacao",
    "produtoBndes",
    "instrumentoBndes",
    "tipoOperacao",
    "tope",
    "operacaoDireta",
    "liquidada",
    "uf",
    "municipio",
    "setorApoiado",
    "subsetorApoiado",
    "cnae",
    "porteCliente",
    "naturezaCliente",
    "fonteRecursos",
    "custoFinanceiro",
    "taxaJuros",
    "prazoCarencia",
    "prazoAmortizacao",
    "inovacao",
    "reembolsavel",
    "agenteFinanceiro",
    "cnpjAgenteFinanceiro",
    "tipoGarantia",
    "areaOperacional",
    "paisDestino",
    "modalidadeOperacional",
    "moeda",
    "mutuario",
    "categoria",
    "id",
]


def _engine() -> str:
    try:
        import xlsxwriter  # noqa: F401

        return "xlsxwriter"
    except ImportError:
        return "openpyxl"


def _write_df(writer: Any, sheet: str, df: pd.DataFrame, engine: str) -> None:
    df.to_excel(writer, sheet_name=sheet, index=False)
    autosize_dataframe_sheet(writer, sheet, df, engine=engine, max_width=55)


def build_frames(docs: list[dict[str, Any]]) -> dict[str, pd.DataFrame]:
    ops = [flatten_operacao(d) for d in docs]
    df_ops = pd.DataFrame(ops)
    for c in OPERACAO_COLS:
        if c not in df_ops.columns:
            df_ops[c] = None
    df_ops = df_ops[OPERACAO_COLS].sort_values(
        by=["dataContratacao", "valorContratacao"],
        ascending=[False, False],
        na_position="last",
    )

    subs = expand_subcreditos(docs)
    df_sub = pd.DataFrame(subs)
    preferred_sub = [
        "idOperacao",
        "numeroContratoPai",
        "origemLinha",
        "cliente",
        "documentoCliente",
        "numeroContrato",
        "dataContratacao",
        "descricaoProjeto",
        "valorContratacao",
        "valorDesembolsado",
        "produtoBndes",
        "instrumentoBndes",
        "taxaJuros",
        "custoFinanceiro",
        "prazoCarencia",
        "prazoAmortizacao",
        "fonteRecursos",
        "operacaoDireta",
        "liquidada",
        "uf",
        "municipio",
        "agenteFinanceiro",
        "inovacao",
        "setorApoiado",
        "subsetorApoiado",
        "paisDestino",
    ]
    cols = [c for c in preferred_sub if c in df_sub.columns] + [
        c for c in df_sub.columns if c not in preferred_sub
    ]
    df_sub = df_sub[cols]

    summary = summarize(docs)
    df_kpi = pd.DataFrame(
        [
            {"indicador": "Cliente", "valor": summary["cliente"]},
            {"indicador": "Operacoes (docs)", "valor": summary["num_operacoes"]},
            {"indicador": "Com valor contratacao", "valor": summary["com_valor_contratacao"]},
            {"indicador": "Sem valor contratacao", "valor": summary["sem_valor_contratacao"]},
            {"indicador": "Soma valor contratacao (R$)", "valor": summary["soma_valor_contratacao"]},
            {"indicador": "Soma valor desembolsado (R$)", "valor": summary["soma_valor_desembolsado"]},
            {"indicador": "Ano minimo", "valor": summary["ano_min"]},
            {"indicador": "Ano maximo", "valor": summary["ano_max"]},
            {
                "indicador": "Nota",
                "valor": (
                    "Operacoes pos-embarque / EXIM frequentemente nao publicam "
                    "valorContratacao no documento agregado."
                ),
            },
        ]
    )

    by_year = (
        df_ops.groupby("anoPosicao", dropna=False)
        .agg(
            operacoes=("id", "count"),
            valor_contratacao=("valorContratacao", "sum"),
            valor_desembolsado=("valorDesembolsado", "sum"),
        )
        .reset_index()
        .sort_values("anoPosicao")
    )

    # produto pode ter multiplos separados por |
    prod_rows = []
    for _, r in df_ops.iterrows():
        prods = [p.strip() for p in str(r.get("produtoBndes") or "N/D").split("|") if p.strip()]
        if not prods:
            prods = ["N/D"]
        for p in prods:
            prod_rows.append(
                {
                    "produtoBndes": p,
                    "valorContratacao": r.get("valorContratacao") or 0.0,
                    "valorDesembolsado": r.get("valorDesembolsado") or 0.0,
                    "id": r.get("id"),
                }
            )
    df_prod_raw = pd.DataFrame(prod_rows)
    by_prod = (
        df_prod_raw.groupby("produtoBndes", dropna=False)
        .agg(
            operacoes=("id", "nunique"),
            valor_contratacao=("valorContratacao", "sum"),
            valor_desembolsado=("valorDesembolsado", "sum"),
        )
        .reset_index()
        .sort_values("valor_contratacao", ascending=False)
    )

    by_status = (
        df_ops.groupby("liquidada", dropna=False)
        .agg(
            operacoes=("id", "count"),
            valor_contratacao=("valorContratacao", "sum"),
            valor_desembolsado=("valorDesembolsado", "sum"),
        )
        .reset_index()
        .sort_values("operacoes", ascending=False)
    )

    by_tope = (
        df_ops.groupby("tope", dropna=False)
        .agg(
            operacoes=("id", "count"),
            valor_contratacao=("valorContratacao", "sum"),
            valor_desembolsado=("valorDesembolsado", "sum"),
        )
        .reset_index()
        .sort_values("valor_contratacao", ascending=False)
    )

    return {
        "Resumo": df_kpi,
        "Por_Ano": by_year,
        "Por_Produto": by_prod,
        "Por_Situacao": by_status,
        "Por_Tope": by_tope,
        "Operacoes": df_ops,
        "Subcreditos": df_sub,
    }


def write_excel(docs: list[dict[str, Any]], path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    frames = build_frames(docs)
    engine = _engine()
    money_names = {
        "valorContratacao",
        "valorDesembolsado",
        "valor_contratacao",
        "valor_desembolsado",
    }
    with pd.ExcelWriter(path, engine=engine) as writer:
        for name, df in frames.items():
            sheet = name[:31]
            _write_df(writer, sheet, df, engine)
            if engine != "xlsxwriter":
                continue
            ws = writer.sheets[sheet]
            money = writer.book.add_format({"num_format": "#,##0.00"})
            for idx, col_name in enumerate(df.columns):
                label = str(col_name)
                if label in money_names or "valor" in label.lower():
                    # Resumo KPI mixes text/numbers in one column — skip.
                    if sheet == "Resumo":
                        continue
                    ws.set_column(idx, idx, 18, money)
    return path
