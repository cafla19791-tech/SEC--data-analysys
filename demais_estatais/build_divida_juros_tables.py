#!/usr/bin/env python3
"""Build annual tables for demais estatais (ex-Petrobras, ex-Eletrobras, ex-bancos).

Official BCB fiscal statistics do not publish a homogeneous gross-debt stock for
this aggregate. The stock measure is DLSP (net debt). Interest is NFSP nominal
interest on an accrual / below-the-line basis (net of interest on financial assets).
"""

from __future__ import annotations

import json
import urllib.request
from collections import defaultdict
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw" / "bcb_series"
RAW.mkdir(parents=True, exist_ok=True)

SERIES = {
    4474: ("dlsp_total", "DLSP total — empresas estatais (R$ mi)"),
    4475: ("dlsp_federal", "DLSP — estatais federais (R$ mi)"),
    4476: ("dlsp_estadual", "DLSP — estatais estaduais (R$ mi)"),
    4477: ("dlsp_municipal", "DLSP — estatais municipais (R$ mi)"),
    4509: ("dlsp_total_pct_pib", "DLSP total estatais (% PIB)"),
    4510: ("dlsp_federal_pct_pib", "DLSP estatais federais (% PIB)"),
    4612: ("juros_total", "NFSP juros nominais — estatais total (R$ mi/mês)"),
    4613: ("juros_federal", "NFSP juros nominais — estatais federais (R$ mi/mês)"),
    4614: ("juros_estadual", "NFSP juros nominais — estatais estaduais (R$ mi/mês)"),
    4615: ("juros_municipal", "NFSP juros nominais — estatais municipais (R$ mi/mês)"),
    4579: ("resultado_nominal_total", "NFSP resultado nominal — estatais total (R$ mi/mês)"),
    4580: ("resultado_nominal_federal", "NFSP resultado nominal — estatais federais (R$ mi/mês)"),
}


def fetch_sgs(code: int) -> list[dict]:
    url = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{code}/dados"
        f"?formato=json&dataInicial=01/01/2001&dataFinal=31/12/2025"
    )
    with urllib.request.urlopen(url, timeout=90) as r:
        data = json.loads(r.read().decode())
    path = RAW / f"{code}_{SERIES[code][0]}.json"
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def to_monthly(series: list[dict]) -> dict[tuple[int, int], float]:
    out = {}
    for row in series:
        d, m, y = row["data"].split("/")
        out[(int(y), int(m))] = float(row["valor"])
    return out


def december_stock(monthly: dict[tuple[int, int], float], year: int) -> float | None:
    return monthly.get((year, 12))


def annual_sum(monthly: dict[tuple[int, int], float], year: int) -> float | None:
    vals = [v for (y, m), v in monthly.items() if y == year]
    if not vals:
        return None
    # Full calendar year preferred; allow partial early years if present.
    if year >= 2002 and len(vals) < 12 and year != 2001:
        # 2001 starts mid-year for some series; for 2002+ require 12 months.
        if len(vals) < 12:
            return None
    return sum(vals)


def br(n: float | None, digits: int = 2) -> str:
    if n is None:
        return "n/d"
    s = f"{n:,.{digits}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def main() -> None:
    monthly = {}
    meta = {}
    for code, (key, title) in SERIES.items():
        data = fetch_sgs(code)
        monthly[key] = to_monthly(data)
        meta[key] = {"sgs": code, "title": title, "n_obs": len(data)}

    years = list(range(2002, 2026))
    rows = []
    for y in years:
        rows.append(
            {
                "ano": y,
                "dlsp_total_rs_mi": december_stock(monthly["dlsp_total"], y),
                "dlsp_federal_rs_mi": december_stock(monthly["dlsp_federal"], y),
                "dlsp_estadual_rs_mi": december_stock(monthly["dlsp_estadual"], y),
                "dlsp_municipal_rs_mi": december_stock(monthly["dlsp_municipal"], y),
                "dlsp_total_pct_pib": december_stock(monthly["dlsp_total_pct_pib"], y),
                "dlsp_federal_pct_pib": december_stock(monthly["dlsp_federal_pct_pib"], y),
                "juros_total_rs_mi": annual_sum(monthly["juros_total"], y),
                "juros_federal_rs_mi": annual_sum(monthly["juros_federal"], y),
                "juros_estadual_rs_mi": annual_sum(monthly["juros_estadual"], y),
                "juros_municipal_rs_mi": annual_sum(monthly["juros_municipal"], y),
                "resultado_nominal_total_rs_mi": annual_sum(
                    monthly["resultado_nominal_total"], y
                ),
                "resultado_nominal_federal_rs_mi": annual_sum(
                    monthly["resultado_nominal_federal"], y
                ),
            }
        )

    notes = [
        "Abrangência BCB: empresas estatais NÃO financeiras das três esferas, "
        "exceto Grupo Petrobras e Grupo Eletrobras; inclui Itaipu Binacional. "
        "Bancos públicos já ficam fora do conceito de setor público não financeiro.",
        "Estoque = Dívida LÍQUIDA (DLSP), não dívida bruta. O BCB não publica série "
        "homogênea de dívida bruta agregada para esse conjunto. Valor negativo = "
        "posição credora líquida (ativos financeiros > passivos).",
        "Juros = juros nominais das NFSP sem desvalorização cambial (competência / "
        "abaixo da linha), líquidos de juros sobre ativos financeiros — NÃO são "
        "iguais a 'juros pagos em caixa' das DFs societárias.",
        "Fontes: BCB SGS 4474/4475/4476/4477 (DLSP); 4509/4510 (% PIB); "
        "4612/4613/4614/4615 (juros); 4579/4580 (resultado nominal).",
        "Estoque anual = saldo de dezembro; fluxos anuais = soma dos 12 meses.",
        "Comparabilidade com tabelas Petrobras/Eletrobras é limitada: aquelas usam "
        "dívida bruta societária e juros pagos em caixa; aqui o conceito é fiscal BCB.",
    ]

    payload = {
        "title": "Demais estatais — DLSP e juros nominais NFSP (2002–2025)",
        "exclusions": ["Petrobras", "Eletrobras", "bancos públicos / setor financeiro"],
        "concept_stock": "DLSP (dívida líquida) — não dívida bruta",
        "concept_interest": "Juros nominais NFSP (competência, líquido de juros de ativos)",
        "currency": "BRL",
        "unit_stock": "R$ milhões (saldo dezembro)",
        "unit_flow": "R$ milhões (soma anual dos fluxos mensais)",
        "sources": meta,
        "notes": notes,
        "rows": rows,
    }

    json_path = OUT / "demais_estatais_divida_juros_2002_2025.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    df = pd.DataFrame(rows)

    # Friendly tables
    divida = pd.DataFrame(
        {
            "Ano": df["ano"],
            "DLSP total (R$ mi)": df["dlsp_total_rs_mi"],
            "DLSP federal (R$ mi)": df["dlsp_federal_rs_mi"],
            "DLSP estadual (R$ mi)": df["dlsp_estadual_rs_mi"],
            "DLSP municipal (R$ mi)": df["dlsp_municipal_rs_mi"],
            "DLSP total (% PIB)": df["dlsp_total_pct_pib"],
            "DLSP federal (% PIB)": df["dlsp_federal_pct_pib"],
            "Conceito": "Dívida líquida DLSP (não bruta)",
            "Fonte": "BCB SGS 4474/4475/4476/4477; 4509/4510",
        }
    )
    juros = pd.DataFrame(
        {
            "Ano": df["ano"],
            "Juros total (R$ mi)": df["juros_total_rs_mi"],
            "Juros federal (R$ mi)": df["juros_federal_rs_mi"],
            "Juros estadual (R$ mi)": df["juros_estadual_rs_mi"],
            "Juros municipal (R$ mi)": df["juros_municipal_rs_mi"],
            "Conceito": "Juros nominais NFSP (competência; líquido)",
            "Fonte": "BCB SGS 4612/4613/4614/4615",
        }
    )
    consolidado = pd.DataFrame(
        {
            "Ano": df["ano"],
            "DLSP total (R$ mi)": df["dlsp_total_rs_mi"],
            "DLSP federal (R$ mi)": df["dlsp_federal_rs_mi"],
            "Juros total (R$ mi)": df["juros_total_rs_mi"],
            "Juros federal (R$ mi)": df["juros_federal_rs_mi"],
            "Resultado nominal total (R$ mi)": df["resultado_nominal_total_rs_mi"],
            "Resultado nominal federal (R$ mi)": df["resultado_nominal_federal_rs_mi"],
        }
    )
    notas_df = pd.DataFrame({"Nota": notes})

    xlsx_path = OUT / "demais_estatais_divida_juros_2002_2025.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        divida.to_excel(writer, sheet_name="Divida_DLSP", index=False)
        juros.to_excel(writer, sheet_name="Juros_NFSP", index=False)
        consolidado.to_excel(writer, sheet_name="Consolidado", index=False)
        notas_df.to_excel(writer, sheet_name="Notas", index=False)

    # Markdown
    md = []
    md.append("# Demais estatais — evolução da dívida (DLSP) e dos juros (NFSP), 2002–2025\n")
    md.append(
        "Recorte: empresas estatais **não financeiras**, excluindo **Petrobras**, "
        "**Eletrobras** e **bancos/setor financeiro**. Fonte: estatísticas fiscais do BCB.\n"
    )
    md.append(
        "> **Atenção conceitual:** não existe série oficial homogênea de **dívida bruta** "
        "para esse agregado. O estoque publicado é a **dívida líquida (DLSP)**. "
        "Os juros são **nominais das NFSP** (competência, líquidos de juros de ativos), "
        "não juros pagos em caixa.\n"
    )
    md.append("## Dívida líquida — DLSP (saldo de dezembro, R$ milhões)\n")
    md.append(
        "| Ano | Total | Federal | Estadual | Municipal | Total % PIB |\n"
        "|---:|---:|---:|---:|---:|---:|"
    )
    for r in rows:
        md.append(
            f"| {r['ano']} | {br(r['dlsp_total_rs_mi'])} | {br(r['dlsp_federal_rs_mi'])} | "
            f"{br(r['dlsp_estadual_rs_mi'])} | {br(r['dlsp_municipal_rs_mi'])} | "
            f"{br(r['dlsp_total_pct_pib'])} |"
        )
    md.append("\n## Juros nominais NFSP (soma anual, R$ milhões)\n")
    md.append(
        "| Ano | Total | Federal | Estadual | Municipal |\n"
        "|---:|---:|---:|---:|---:|"
    )
    for r in rows:
        md.append(
            f"| {r['ano']} | {br(r['juros_total_rs_mi'])} | {br(r['juros_federal_rs_mi'])} | "
            f"{br(r['juros_estadual_rs_mi'])} | {br(r['juros_municipal_rs_mi'])} |"
        )
    md.append("\n## Notas\n")
    for n in notes:
        md.append(f"- {n}")
    md_path = OUT / "demais_estatais_divida_juros_2002_2025.md"
    md_path.write_text("\n".join(md) + "\n", encoding="utf-8")

    print(f"Wrote {json_path}")
    print(f"Wrote {xlsx_path}")
    print(f"Wrote {md_path}")
    print("\nPreview DLSP total / juros total:")
    for r in rows:
        print(
            f"{r['ano']}  DLSP={br(r['dlsp_total_rs_mi']):>12}  "
            f"fed={br(r['dlsp_federal_rs_mi']):>12}  "
            f"juros={br(r['juros_total_rs_mi']):>12}  "
            f"j_fed={br(r['juros_federal_rs_mi']):>12}"
        )


if __name__ == "__main__":
    main()
