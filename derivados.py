#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processa e consolida as três séries mensais de derivados de petróleo da ANP (2000-2026):
1. Volume de derivados produzidos no Brasil (barris) — producao-derivados-b.xlsx
2. Volume de derivados importados (barris) — importacoes-exportacoes-b.xlsx
3. Dispêndio com as importações de derivados (US$ FOB) — importacoes-exportacoes-b.xlsx

Calcula adicionalmente:
- Custo médio unitário do derivado importado (US$/barril)
- Custo médio ponderado nacional (adotando custo de produção brasileira a US$ 25/bbl e comparativo a US$ 20/bbl)
- Lucro bruto adicional se todo o volume importado tivesse sido refinado no Brasil

Compatível tanto com execução direta na pasta WinPython do ContAgil quanto no repositório.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

MONTHS_PT = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
]
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTHS_PT)}
MONTH_NUM["Marco"] = 3

DEFAULT_DOMESTIC_COST = 25.0


def year_key(v) -> int | None:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        y = int(v)
        if 1990 <= y <= 2100:
            return y
    if isinstance(v, str) and re.fullmatch(r"\d{4}", v.strip()):
        return int(v.strip())
    return None


def extract_monthly_matrix(ws, header_row: int) -> dict[tuple[int, int], float]:
    years = {}
    for c in range(3, 50):
        y = year_key(ws.cell(header_row, c).value)
        if y is not None:
            years[c] = y
    out: dict[tuple[int, int], float] = {}
    for r in range(header_row + 1, header_row + 20):
        label = ws.cell(r, 2).value
        if not isinstance(label, str):
            continue
        lab = label.strip()
        if lab not in MONTH_NUM:
            continue
        m = MONTH_NUM[lab]
        for c, y in years.items():
            v = ws.cell(r, c).value
            if v is None or v == "":
                continue
            if isinstance(v, str) and v.strip().lower() in {"n/d", "-", "nd"}:
                continue
            try:
                out[(y, m)] = float(v)
            except (ValueError, TypeError):
                continue
    return out


def resolve_file(candidates: list[Path], desc: str) -> Path:
    for c in candidates:
        if c.is_file():
            return c
    raise FileNotFoundError(
        f"Arquivo para {desc} não encontrado. Locais verificados:\n"
        + "\n".join(f" - {p}" for p in candidates)
    )


def format_br(n: float | None, digits: int = 2) -> str:
    if n is None:
        return "n/d"
    s = f"{n:,.{digits}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def run_pipeline(
    prod_path: Path,
    imp_path: Path,
    output_dir: Path,
    start_year: int = 2000,
    end_year: int = 2026,
    domestic_cost: float = DEFAULT_DOMESTIC_COST,
) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Lendo produção de derivados: {prod_path}")
    wb_prod = load_workbook(prod_path, data_only=True)
    ws_prod = wb_prod.active
    # Linha 36 é o cabeçalho 'ANO' de DERIVADOS TOTAL (b)
    prod_vol = extract_monthly_matrix(ws_prod, 36)
    wb_prod.close()

    print(f"Lendo importações de derivados: {imp_path}")
    wb_imp = load_workbook(imp_path, data_only=True)
    ws_imp = wb_imp["Plan1"]
    # Linha 329: Importação de derivados (barris)
    # Linha 390: Dispêndio com importação de derivados (US$ FOB)
    imp_vol = extract_monthly_matrix(ws_imp, 329)
    imp_usd = extract_monthly_matrix(ws_imp, 390)
    wb_imp.close()

    # Construir lista ordenada de meses
    all_keys = set(prod_vol.keys()) | set(imp_vol.keys()) | set(imp_usd.keys())
    valid_keys = sorted(
        [k for k in all_keys if start_year <= k[0] <= end_year],
        key=lambda x: (x[0], x[1]),
    )

    rows = []
    prod_matrix_rows = []
    imp_vol_matrix_rows = []
    imp_usd_matrix_rows = []

    for y, m in valid_keys:
        pv = prod_vol.get((y, m))
        iv = imp_vol.get((y, m))
        iu = imp_usd.get((y, m))

        # Ignorar meses futuros vazios (ex: jul/2026 a dez/2026)
        if pv is None and iv is None and iu is None:
            continue

        cost_imp = (iu / iv) if (iv and iu and iv > 0) else None
        tot_vol = (pv or 0.0) + (iv or 0.0)
        tot_cost_obs = (iu or 0.0) + ((pv or 0.0) * domestic_cost)
        wac = (tot_cost_obs / tot_vol) if tot_vol > 0 else None
        lucro_adicional = ((cost_imp - domestic_cost) * iv) if (cost_imp is not None and iv) else None

        rows.append(
            {
                "ano": y,
                "mes": m,
                "mes_nome": MONTHS_PT[m - 1],
                "periodo": f"{y}-{m:02d}",
                "volume_produzido_brasil_barris": pv,
                "volume_importado_barris": iv,
                "dispendio_importacao_usd": iu,
                "custo_medio_importado_usd_bbl": cost_imp,
                "custo_producao_brasil_usd_bbl": domestic_cost,
                "volume_total_proxy_consumo_barris": tot_vol,
                "custo_medio_ponderado_usd_bbl": wac,
                "lucro_bruto_adicional_usd": lucro_adicional,
            }
        )

    df = pd.DataFrame(rows)

    # 1. Matriz Produção Nacional (Planilha 1)
    df_pivot_prod = df.pivot(index="mes_nome", columns="ano", values="volume_produzido_brasil_barris")
    df_pivot_prod = df_pivot_prod.reindex(MONTHS_PT)
    df_pivot_prod.loc["Total do Ano"] = df_pivot_prod.sum(axis=0)

    # 2. Matriz Volume Importado (Planilha 2)
    df_pivot_imp_vol = df.pivot(index="mes_nome", columns="ano", values="volume_importado_barris")
    df_pivot_imp_vol = df_pivot_imp_vol.reindex(MONTHS_PT)
    df_pivot_imp_vol.loc["Total do Ano"] = df_pivot_imp_vol.sum(axis=0)

    # 3. Matriz Dispêndio de Importação (Planilha 3)
    df_pivot_imp_usd = df.pivot(index="mes_nome", columns="ano", values="dispendio_importacao_usd")
    df_pivot_imp_usd = df_pivot_imp_usd.reindex(MONTHS_PT)
    df_pivot_imp_usd.loc["Total do Ano"] = df_pivot_imp_usd.sum(axis=0)

    # Resumo Anual
    resumo_anual = (
        df.groupby("ano", as_index=False)
        .agg(
            volume_produzido_brasil_barris=("volume_produzido_brasil_barris", "sum"),
            volume_importado_barris=("volume_importado_barris", "sum"),
            dispendio_importacao_usd=("dispendio_importacao_usd", "sum"),
            volume_total_proxy_barris=("volume_total_proxy_consumo_barris", "sum"),
            lucro_bruto_adicional_usd=("lucro_bruto_adicional_usd", "sum"),
        )
        .assign(
            custo_medio_importado_usd_bbl=lambda x: x["dispendio_importacao_usd"] / x["volume_importado_barris"],
            custo_medio_ponderado_usd_bbl=lambda x: (
                x["dispendio_importacao_usd"] + x["volume_produzido_brasil_barris"] * domestic_cost
            )
            / x["volume_total_proxy_barris"],
        )
    )

    tot_prod = float(df["volume_produzido_brasil_barris"].dropna().sum())
    tot_imp = float(df["volume_importado_barris"].dropna().sum())
    tot_usd = float(df["dispendio_importacao_usd"].dropna().sum())
    tot_cons = tot_prod + tot_imp
    tot_lucro = float(df["lucro_bruto_adicional_usd"].dropna().sum())
    wac_geral = (tot_usd + tot_prod * domestic_cost) / tot_cons if tot_cons > 0 else 0.0
    pm_imp_geral = tot_usd / tot_imp if tot_imp > 0 else 0.0

    # Gravar Excel consolidado com as três planilhas solicitadas
    excel_path = output_dir / "DERIVADOS_ANP_2000_2026.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_pivot_prod.to_excel(writer, sheet_name="1_Producao_Brasil_barris")
        df_pivot_imp_vol.to_excel(writer, sheet_name="2_Importacao_Volume_barris")
        df_pivot_imp_usd.to_excel(writer, sheet_name="3_Dispendio_Importacao_USD")
        resumo_anual.to_excel(writer, sheet_name="Resumo_Anual", index=False)
        df.to_excel(writer, sheet_name="Serie_Mensal_Completa", index=False)

    # Gravar JSON
    json_path = output_dir / "DERIVADOS_ANP_2000_2026.json"
    summary = {
        "periodo": f"{start_year} a {end_year}",
        "meses_observados": len(df),
        "custo_producao_brasil_hipotese_usd_bbl": domestic_cost,
        "totais_acumulados": {
            "volume_produzido_brasil_barris": tot_prod,
            "volume_importado_barris": tot_imp,
            "volume_total_consumo_proxy_barris": tot_cons,
            "dispendio_importacao_usd": tot_usd,
            "preco_medio_importacao_usd_bbl": pm_imp_geral,
            "custo_medio_ponderado_usd_bbl": wac_geral,
            "lucro_bruto_adicional_se_tudo_produzido_brasil_usd": tot_lucro,
        },
        "resumo_anual": resumo_anual.to_dict(orient="records"),
    }
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    # Gravar CSV da série completa
    csv_path = output_dir / "DERIVADOS_ANP_2000_2026_mensal.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print(f"\n[SUCESSO] Processamento concluído!")
    print(f"Planilha Excel consolidada: {excel_path}")
    print(f"  - Aba 1: 1_Producao_Brasil_barris")
    print(f"  - Aba 2: 2_Importacao_Volume_barris")
    print(f"  - Aba 3: 3_Dispendio_Importacao_USD")
    print(f"  - Aba 4: Resumo_Anual")
    print(f"  - Aba 5: Serie_Mensal_Completa")
    print(f"Arquivo JSON: {json_path}")
    print(f"Arquivo CSV:  {csv_path}\n")

    print(f"--- TOTAIS DO PERÍODO ({start_year} - {end_year}) ---")
    print(f"Volume Produzido no Brasil: {format_br(tot_prod, 0)} barris")
    print(f"Volume Importado:           {format_br(tot_imp, 0)} barris")
    print(f"Volume Total Consumido:     {format_br(tot_cons, 0)} barris")
    print(f"Dispêndio com Importações:  US$ {format_br(tot_usd, 2)} (~US$ {tot_usd/1e9:.2f} bilhões)")
    print(f"Preço Médio de Importação:  US$ {format_br(pm_imp_geral, 2)}/barril")
    print(f"Custo Médio Ponderado:      US$ {format_br(wac_geral, 2)}/barril (considerando BR a US$ {domestic_cost:.2f}/b)")
    print(f"Lucro Adicional Estimado:   US$ {format_br(tot_lucro, 2)} (~US$ {tot_lucro/1e9:.2f} bilhões)\n")

    return summary


def main():
    parser = argparse.ArgumentParser(
        description="Executa o processamento de derivados de petróleo da ANP (2000-2026)."
    )
    parser.add_argument(
        "--producao",
        type=str,
        default=None,
        help="Caminho do arquivo producao-derivados-b.xlsx (ou .xls)",
    )
    parser.add_argument(
        "--importacoes",
        type=str,
        default=None,
        help="Caminho do arquivo importacoes-exportacoes-b.xlsx",
    )
    parser.add_argument(
        "--saida",
        type=str,
        default=None,
        help="Diretório onde salvar a planilha e os relatórios de saída",
    )
    parser.add_argument(
        "--ano-inicio",
        type=int,
        default=2000,
        help="Ano inicial da análise (padrão: 2000)",
    )
    parser.add_argument(
        "--ano-fim",
        type=int,
        default=2026,
        help="Ano final da análise (padrão: 2026)",
    )
    parser.add_argument(
        "--custo-domestico",
        type=float,
        default=DEFAULT_DOMESTIC_COST,
        help="Hipótese de custo de produção nacional em US$/barril (padrão: 25.0)",
    )

    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    winpy = Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython")

    # Localizar producao-derivados-b
    prod_candidates = []
    if args.producao:
        prod_candidates.append(Path(args.producao))
    prod_candidates.extend(
        [
            root / "data" / "raw" / "anp" / "producao-derivados-b.xlsx",
            root / "data" / "raw" / "anp" / "producao-derivados-b.xls",
            root / "producao-derivados-b.xlsx",
            root / "producao-derivados-b.xls",
            winpy / "dados" / "producao-derivados-b.xlsx",
            winpy / "producao-derivados-b.xlsx",
            Path("producao-derivados-b.xlsx"),
        ]
    )
    prod_file = resolve_file(prod_candidates, "produção de derivados (producao-derivados-b.xlsx)")

    # Localizar importacoes-exportacoes-b
    imp_candidates = []
    if args.importacoes:
        imp_candidates.append(Path(args.importacoes))
    imp_candidates.extend(
        [
            root / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xlsx",
            root / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xls",
            root / "importacoes-exportacoes-b.xlsx",
            root / "importacoes-exportacoes-b.xls",
            winpy / "dados" / "importacoes-exportacoes-b.xlsx",
            winpy / "importacoes-exportacoes-b.xlsx",
            Path("importacoes-exportacoes-b.xlsx"),
        ]
    )
    imp_file = resolve_file(imp_candidates, "importações e exportações (importacoes-exportacoes-b.xlsx)")

    # Definir diretório de saída
    if args.saida:
        out_dir = Path(args.saida)
    elif (root / "output").is_dir():
        out_dir = root / "output"
    else:
        out_dir = root / "saida"

    run_pipeline(
        prod_path=prod_file,
        imp_path=imp_file,
        output_dir=out_dir,
        start_year=args.ano_inicio,
        end_year=args.ano_fim,
        domestic_cost=args.custo_domestico,
    )


if __name__ == "__main__":
    main()
