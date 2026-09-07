#!/usr/bin/env python3
"""Discriminante de lucros líquidos — grandes bancos BR (2023–2026).

Gera planilha Excel com:
  - Matriz banco × ano (R$ bilhões)
  - 2026 como 1º semestre (parcial; ano ainda em curso em set/2026)
  - Participações, variações e ranking
  - Notas metodológicas e fontes

Critérios:
  - Itaú / Bradesco / Santander: lucro líquido recorrente / gerencial (padrão de mercado)
  - BTG Pactual: lucro líquido ajustado
  - Nubank (Nu Holdings): lucro líquido em USD convertido a R$ pela taxa média anual
    (PTAX média aproximada) — a companhia reporta em dólares
"""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"
ART_DIR = Path("/opt/cursor/artifacts")

# Lucros em R$ bilhões (exceto coluna auxiliar USD do Nubank)
# Fontes: releases oficiais / cobertura Valor, Folha, Estadão, InfoMoney, etc.
DADOS = {
    "Banco": [
        "Itaú Unibanco",
        "Bradesco",
        "Santander Brasil",
        "Nubank (Nu Holdings)",
        "BTG Pactual",
    ],
    # Lucro líquido recorrente/gerencial/ajustado (R$ bi)
    "2023": [35.618, 16.297, 9.386, 4.99, 10.419],
    "2024": [41.403, 19.554, 13.872, 10.62, 12.321],
    "2025": [46.800, 24.650, 15.615, 16.04, 16.700],
    # 2026 parcial = 1S26 (ano ainda em curso)
    "2026_1S": [24.689, 13.861, 6.850, 10.43, 10.000],
}

# Nubank original em US$ bi e FX médio usado na conversão
NUBANK_USD = {
    "2023": 1.00,
    "2024": 1.97,
    "2025": 2.87,
    "2026_1S": 1.9324,  # 0.8714 + 1.061
}
FX_MEDIO = {
    "2023": 4.99,  # PTAX média aprox. → 1.00 * 4.99 = 4.99
    "2024": 5.39,  # 1.97 * 5.39 ≈ 10.62
    "2025": 5.59,  # 2.87 * 5.59 ≈ 16.04 (Finsiders)
    "2026_1S": 5.40,  # 1.9324 * 5.40 ≈ 10.43 (aprox. 1S26)
}

METRICAS = {
    "Itaú Unibanco": "Lucro líquido recorrente gerencial",
    "Bradesco": "Lucro líquido recorrente",
    "Santander Brasil": "Lucro líquido gerencial",
    "Nubank (Nu Holdings)": "Lucro líquido (US$ → R$ pela PTAX média)",
    "BTG Pactual": "Lucro líquido ajustado",
}

FONTES = [
    "Itaú: releases 2023–2025 e 1S26 (recorrente gerencial).",
    "Bradesco: releases 2023–2025 e 1S26 (recorrente).",
    "Santander Brasil: Folha/Valor — gerencial 2023 R$ 9,4 bi; 2024 R$ 13,872 bi; "
    "2025 R$ 15,615 bi; 1S26 ≈ R$ 6,85 bi (1T R$ 3,788 + 2T ≈ R$ 3,0; "
    "coerente com total Itaú+Bradesco+Santander = R$ 45,4 bi no 1S26).",
    "Nubank: Nu Holdings — US$ 1,0 bi (2023), US$ 1,97 bi (2024), US$ 2,87 bi (2025), "
    "US$ 1,932 bi no 1S26 (US$ 871,4 mi + US$ 1,061 bi); convertido pela PTAX média.",
    "BTG Pactual: ajustado R$ 10,419 bi (2023), R$ 12,321 bi (2024), R$ 16,7 bi (2025); "
    "1S26 ≈ R$ 10,0 bi (ajustado).",
    "Referência de data: setembro/2026 — 2026 ainda sem fechamento anual.",
]


def thin_border() -> Border:
    s = Side(style="thin", color="B0B0B0")
    return Border(left=s, right=s, top=s, bottom=s)


def style_header(ws, row: int, cols: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border()


def autofit(ws, min_w: int = 12, max_w: int = 48) -> None:
    for col in ws.columns:
        letter = None
        width = min_w
        for cell in col:
            if hasattr(cell, "column_letter"):
                letter = cell.column_letter
            if cell.value is None:
                continue
            width = max(width, min(max_w, len(str(cell.value)) + 2))
        if letter:
            ws.column_dimensions[letter].width = width


def build_frame() -> pd.DataFrame:
    df = pd.DataFrame(DADOS)
    df = df.set_index("Banco")
    df["CAGR_2023_2025_%"] = (
        (df["2025"] / df["2023"]) ** (1 / 2) - 1
    ) * 100
    # Anualização simples do 1S26 (×2) — apenas referência, não guidance
    df["2026_anualizado_ref"] = df["2026_1S"] * 2
    df["Var_1S26_vs_metade_2025_%"] = (
        df["2026_1S"] / (df["2025"] / 2) - 1
    ) * 100
    return df


def add_notas(ws, start_row: int) -> None:
    ws.cell(start_row, 1, "Notas metodológicas").font = Font(bold=True, size=12)
    r = start_row + 1
    for nota in [
        "Valores em R$ bilhões.",
        "Coluna 2026_1S = lucro do 1º semestre de 2026 (parcial; ano em curso).",
        "2026_anualizado_ref = 1S26 × 2 (apenas referência mecânica; não é projeção oficial).",
        "Métricas diferem por instituição (recorrente/gerencial/ajustado) — padrão de mercado.",
        "Nubank: conversão US$→R$ pela PTAX média aproximada do período.",
    ] + FONTES:
        ws.cell(r, 1, f"• {nota}")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True)
        r += 1
    ws.row_dimensions[start_row].height = 18


def write_excel(df: pd.DataFrame, path: Path) -> None:
    wb = Workbook()

    # --- Aba 1: Discriminante principal ---
    ws = wb.active
    ws.title = "Discriminante"

    titulo = (
        "Discriminante — Lucros líquidos (R$ bi): "
        "Itaú, Bradesco, Santander Brasil, Nubank e BTG Pactual (2023–2026)"
    )
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws.merge_cells("A1:G1")

    base = df[["2023", "2024", "2025", "2026_1S"]].copy()
    base.loc["TOTAL"] = base.sum(numeric_only=True)

    start = 3
    headers = ["Banco", "2023", "2024", "2025", "2026 (1S parcial)", "Métrica"]
    for j, h in enumerate(headers, 1):
        ws.cell(start, j, h)
    style_header(ws, start, len(headers))

    fill_alt = PatternFill("solid", fgColor="D6EAF8")
    fill_tot = PatternFill("solid", fgColor="F4B183")
    for i, banco in enumerate(base.index):
        r = start + 1 + i
        ws.cell(r, 1, banco)
        for j, col in enumerate(["2023", "2024", "2025", "2026_1S"], 2):
            cell = ws.cell(r, j, round(float(base.loc[banco, col]), 3))
            cell.number_format = "#,##0.000"
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border()
        metrica = METRICAS.get(banco, "Soma")
        ws.cell(r, 6, metrica)
        for c in range(1, 7):
            ws.cell(r, c).border = thin_border()
            if banco == "TOTAL":
                ws.cell(r, c).fill = fill_tot
                ws.cell(r, c).font = Font(bold=True)
            elif i % 2 == 1:
                ws.cell(r, c).fill = fill_alt

    # Participação %
    part_row = start + len(base) + 3
    ws.cell(part_row, 1, "Participação no total do ano (%)").font = Font(
        bold=True, size=12, color="1F4E79"
    )
    part = df[["2023", "2024", "2025", "2026_1S"]].div(
        df[["2023", "2024", "2025", "2026_1S"]].sum(), axis=1
    ) * 100
    r0 = part_row + 1
    for j, h in enumerate(
        ["Banco", "2023 %", "2024 %", "2025 %", "2026 1S %"], 1
    ):
        ws.cell(r0, j, h)
    style_header(ws, r0, 5)
    for i, banco in enumerate(part.index):
        r = r0 + 1 + i
        ws.cell(r, 1, banco)
        for j, col in enumerate(["2023", "2024", "2025", "2026_1S"], 2):
            cell = ws.cell(r, j, round(float(part.loc[banco, col]), 1))
            cell.number_format = "0.0"
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border()

    # Variações / CAGR
    var_row = r0 + len(part) + 3
    ws.cell(var_row, 1, "Dinâmica (crescimento)").font = Font(
        bold=True, size=12, color="1F4E79"
    )
    dyn = df[
        [
            "2023",
            "2024",
            "2025",
            "2026_1S",
            "CAGR_2023_2025_%",
            "2026_anualizado_ref",
            "Var_1S26_vs_metade_2025_%",
        ]
    ].copy()
    dyn["Δ 2024/2023 %"] = (dyn["2024"] / dyn["2023"] - 1) * 100
    dyn["Δ 2025/2024 %"] = (dyn["2025"] / dyn["2024"] - 1) * 100

    r1 = var_row + 1
    dyn_headers = [
        "Banco",
        "Δ 2024/2023 %",
        "Δ 2025/2024 %",
        "CAGR 2023–2025 %",
        "2026 anualizado (ref.)",
        "1S26 vs metade 2025 %",
    ]
    for j, h in enumerate(dyn_headers, 1):
        ws.cell(r1, j, h)
    style_header(ws, r1, len(dyn_headers))
    for i, banco in enumerate(dyn.index):
        r = r1 + 1 + i
        ws.cell(r, 1, banco)
        vals = [
            dyn.loc[banco, "Δ 2024/2023 %"],
            dyn.loc[banco, "Δ 2025/2024 %"],
            dyn.loc[banco, "CAGR_2023_2025_%"],
            dyn.loc[banco, "2026_anualizado_ref"],
            dyn.loc[banco, "Var_1S26_vs_metade_2025_%"],
        ]
        for j, v in enumerate(vals, 2):
            cell = ws.cell(r, j, round(float(v), 1))
            cell.number_format = "#,##0.0"
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border()

    add_notas(ws, r1 + len(dyn) + 3)
    autofit(ws)

    # Gráfico de barras (dados da matriz)
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Lucro líquido (R$ bi) — discriminante 2023–2025 + 1S26"
    chart.y_axis.title = "R$ bilhões"
    chart.x_axis.title = None
    data = Reference(ws, min_col=2, min_row=start, max_col=5, max_row=start + 5)
    cats = Reference(ws, min_col=1, min_row=start + 1, max_row=start + 5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 18
    chart.height = 10
    ws.add_chart(chart, "H3")

    # --- Aba 2: Nubank USD ---
    ws2 = wb.create_sheet("Nubank_USD")
    ws2["A1"] = "Nubank (Nu Holdings) — lucro líquido reportado em US$ e conversão"
    ws2["A1"].font = Font(bold=True, size=12, color="1F4E79")
    ws2.merge_cells("A1:E1")
    headers2 = ["Período", "Lucro US$ bi", "PTAX média (R$/US$)", "Lucro R$ bi", "Obs."]
    for j, h in enumerate(headers2, 1):
        ws2.cell(3, j, h)
    style_header(ws2, 3, 5)
    rows_nu = [
        ("2023", NUBANK_USD["2023"], FX_MEDIO["2023"], DADOS["2023"][3], "Ano cheio"),
        ("2024", NUBANK_USD["2024"], FX_MEDIO["2024"], DADOS["2024"][3], "Ano cheio"),
        ("2025", NUBANK_USD["2025"], FX_MEDIO["2025"], DADOS["2025"][3], "Ano cheio"),
        (
            "2026 1S",
            NUBANK_USD["2026_1S"],
            FX_MEDIO["2026_1S"],
            DADOS["2026_1S"][3],
            "1T US$ 0,8714 + 2T US$ 1,061",
        ),
    ]
    for i, row in enumerate(rows_nu):
        r = 4 + i
        for j, v in enumerate(row, 1):
            cell = ws2.cell(r, j, v)
            cell.border = thin_border()
            if j in (2, 3, 4) and isinstance(v, float):
                cell.number_format = "0.000"
    autofit(ws2)

    # --- Aba 3: Ranking ---
    ws3 = wb.create_sheet("Ranking")
    ws3["A1"] = "Ranking por lucro líquido (R$ bi)"
    ws3["A1"].font = Font(bold=True, size=12, color="1F4E79")
    col = 1
    for periodo in ["2023", "2024", "2025", "2026_1S"]:
        ranking = df[periodo].sort_values(ascending=False)
        ws3.cell(3, col, periodo if periodo != "2026_1S" else "2026 1S")
        ws3.cell(3, col).font = Font(bold=True, color="FFFFFF")
        ws3.cell(3, col).fill = PatternFill("solid", fgColor="1F4E79")
        ws3.cell(3, col + 1, "R$ bi")
        ws3.cell(3, col + 1).font = Font(bold=True, color="FFFFFF")
        ws3.cell(3, col + 1).fill = PatternFill("solid", fgColor="1F4E79")
        for i, (banco, val) in enumerate(ranking.items(), 1):
            ws3.cell(3 + i, col, f"{i}º {banco}")
            cell = ws3.cell(3 + i, col + 1, round(float(val), 3))
            cell.number_format = "0.000"
        col += 3
    autofit(ws3)

    # --- Aba 4: série longa (tidy) ---
    ws4 = wb.create_sheet("Serie_tidy")
    tidy = (
        df[["2023", "2024", "2025", "2026_1S"]]
        .reset_index()
        .melt(id_vars="Banco", var_name="Periodo", value_name="Lucro_R$_bi")
    )
    for r_idx, row in enumerate(dataframe_to_rows(tidy, index=False, header=True), 1):
        for c_idx, v in enumerate(row, 1):
            ws4.cell(r_idx, c_idx, v)
            if r_idx == 1:
                ws4.cell(r_idx, c_idx).fill = PatternFill("solid", fgColor="1F4E79")
                ws4.cell(r_idx, c_idx).font = Font(color="FFFFFF", bold=True)
    autofit(ws4)

    # Linha chart na aba Discriminante a partir de série (opcional simplificado)
    line = LineChart()
    line.title = "Evolução do lucro líquido (R$ bi)"
    line.style = 10
    line.y_axis.title = "R$ bi"
    line.width = 18
    line.height = 10
    # reuse bar data for line overlay style — separate series
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "H20")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def save_png(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    plot_df = df[["2023", "2024", "2025", "2026_1S"]].copy()
    plot_df = plot_df.rename(columns={"2026_1S": "2026 (1S)"})

    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5), constrained_layout=True)
    fig.suptitle(
        "Discriminante — Lucros líquidos dos grandes bancos (R$ bi)",
        fontsize=13,
        fontweight="bold",
        color="#1F4E79",
    )

    plot_df.T.plot(kind="bar", ax=axes[0], width=0.8)
    axes[0].set_title("Por período")
    axes[0].set_ylabel("R$ bilhões")
    axes[0].tick_params(axis="x", rotation=0)
    axes[0].legend(fontsize=8, loc="upper left")
    axes[0].grid(axis="y", alpha=0.3)

    plot_df.plot(kind="barh", ax=axes[1], width=0.8)
    axes[1].set_title("Por banco")
    axes[1].set_xlabel("R$ bilhões")
    axes[1].legend(fontsize=8)
    axes[1].grid(axis="x", alpha=0.3)

    fig.savefig(path, dpi=140)
    plt.close(fig)


def main() -> int:
    df = build_frame()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ART_DIR.mkdir(parents=True, exist_ok=True)

    xlsx = OUT_DIR / "discriminante_lucros_bancos_2023_2026.xlsx"
    csv = OUT_DIR / "discriminante_lucros_bancos_2023_2026.csv"
    png = ART_DIR / "discriminante_lucros_bancos.png"
    png_out = OUT_DIR / "discriminante_lucros_bancos.png"

    write_excel(df, xlsx)
    df[["2023", "2024", "2025", "2026_1S", "CAGR_2023_2025_%", "2026_anualizado_ref"]].to_csv(
        csv, float_format="%.3f"
    )
    save_png(df, png)
    save_png(df, png_out)

    print("Arquivos gerados:")
    print(f"  - {xlsx}")
    print(f"  - {csv}")
    print(f"  - {png_out}")
    print()
    print(df[["2023", "2024", "2025", "2026_1S"]].round(3).to_string())
    print()
    print("Totais:")
    print(df[["2023", "2024", "2025", "2026_1S"]].sum().round(3).to_string())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
