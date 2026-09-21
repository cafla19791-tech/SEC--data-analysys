#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processa e consolida as três séries mensais de derivados de petróleo da ANP (2000-2026),
tanto no TOTAL NACIONAL quanto DISCRIMINADO POR DERIVADO (PRODUTO):

1. Volume de derivados produzidos no Brasil (barris) — producao-derivados-b.xlsx
2. Volume de derivados importados (barris) — importacoes-exportacoes-b.xlsx
3. Dispêndio com as importações de derivados (US$ FOB) — importacoes-exportacoes-b.xlsx

Calcula adicionalmente para cada derivado e para o total:
- Custo médio unitário de importação (US$/barril)
- Custo médio ponderado nacional (adotando custo de produção brasileira a US$ 25/bbl e comparativo a US$ 20/bbl)
- Lucro bruto adicional se todo o volume importado tivesse sido refinado no Brasil a US$ 25/bbl

Compatível tanto com execução direta na pasta WinPython do ContAgil quanto no repositório.
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

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
MONTH_ABBR = [
    "JAN",
    "FEV",
    "MAR",
    "ABR",
    "MAI",
    "JUN",
    "JUL",
    "AGO",
    "SET",
    "OUT",
    "NOV",
    "DEZ",
]
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTHS_PT)}
MONTH_NUM["Marco"] = 3

DEFAULT_DOMESTIC_COST = 25.0
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def clean_product_name(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    # Remove sufixos como (m3), (b), (bep) e espaços extras
    s = re.sub(r"\s*\((m3|b|bep)\)\s*", "", s, flags=re.I).strip().upper()
    return s


def parse_pivot_cache(path: Path, def_name: str, rec_name: str):
    """Lê diretamente o cache binário XML do Excel para extrair registros por produto."""
    with zipfile.ZipFile(path) as z:
        dt = ET.fromstring(z.read(def_name))
        fields = []
        for f in dt.findall(".//m:cacheField", NS):
            name = f.get("name")
            items = []
            si = f.find("m:sharedItems", NS)
            if si is not None:
                for child in si:
                    tag = child.tag.split("}")[-1]
                    if tag == "s":
                        items.append(child.get("v"))
                    elif tag == "n":
                        items.append(float(child.get("v")))
                    elif tag == "m":
                        items.append(None)
                    else:
                        items.append(child.get("v"))
            fields.append({"name": name, "items": items})

        rt = ET.fromstring(z.read(rec_name))
        records = []
        for r in rt.findall("m:r", NS):
            row = []
            for idx, c in enumerate(r):
                tag = c.tag.split("}")[-1]
                if tag == "x":
                    xi = int(c.get("v"))
                    items = fields[idx]["items"]
                    row.append(items[xi] if items else xi)
                elif tag == "n":
                    row.append(float(c.get("v")))
                elif tag == "s":
                    row.append(c.get("v"))
                elif tag == "m":
                    row.append(None)
                else:
                    row.append(c.get("v"))
            records.append(row)
        return fields, records


def extract_by_product_monthly(
    prod_path: Path, imp_path: Path, start_year: int = 2000, end_year: int = 2026
):
    """
    Extrai as séries mensais discriminadas por derivado (produto) para:
    - Produção Brasil (barris)
    - Importação Volume (barris)
    - Importação Dispêndio (US$ FOB)
    """
    # 1. Produção (Cache 1)
    f_p, r_p = parse_pivot_cache(
        prod_path,
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotCache/pivotCacheRecords1.xml",
    )
    # Fields: PRODUTO(0), ANO(1), ESTADO(2), REFINARIA(3), UNIDADE(4), JAN(5)..DEZ(16)
    prod_data = defaultdict(float)
    all_products = set()

    for r in r_p:
        prod_name = clean_product_name(r[0])
        all_products.add(prod_name)
        ano = int(r[1])
        if start_year <= ano <= end_year:
            for m_idx in range(12):
                val = r[5 + m_idx]
                if isinstance(val, (int, float)):
                    prod_data[(prod_name, ano, m_idx + 1)] += val

    # 2. Importação Volume (Cache 7)
    f_iv, r_iv = parse_pivot_cache(
        imp_path,
        "xl/pivotCache/pivotCacheDefinition7.xml",
        "xl/pivotCache/pivotCacheRecords7.xml",
    )
    # Fields: ANO(0), PRODUTO(1), MOVIMENTO(2), UNIDADE(3), JAN(4)..DEZ(15)
    imp_vol_data = defaultdict(float)
    for r in r_iv:
        ano = int(r[0])
        prod_name = clean_product_name(r[1])
        all_products.add(prod_name)
        if start_year <= ano <= end_year:
            for m_idx in range(12):
                val = r[4 + m_idx]
                if isinstance(val, (int, float)):
                    imp_vol_data[(prod_name, ano, m_idx + 1)] += val

    # 3. Importação Dispêndio US$ (Cache 15)
    f_iu, r_iu = parse_pivot_cache(
        imp_path,
        "xl/pivotCache/pivotCacheDefinition15.xml",
        "xl/pivotCache/pivotCacheRecords15.xml",
    )
    # Fields: ANO(0), PRODUTO(1), UNIDADE(2), JAN(3)..DEZ(14)
    imp_usd_data = defaultdict(float)
    for r in r_iu:
        ano = int(r[0])
        prod_name = clean_product_name(r[1])
        all_products.add(prod_name)
        if start_year <= ano <= end_year:
            for m_idx in range(12):
                val = r[3 + m_idx]
                if isinstance(val, (int, float)):
                    imp_usd_data[(prod_name, ano, m_idx + 1)] += val

    return sorted(all_products), prod_data, imp_vol_data, imp_usd_data


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

    print(f"Lendo dados de produção: {prod_path}")
    print(f"Lendo dados de importação: {imp_path}")

    products, prod_dict, imp_vol_dict, imp_usd_dict = extract_by_product_monthly(
        prod_path, imp_path, start_year=start_year, end_year=end_year
    )

    # Obter anos e meses válidos (2000-01 a 2026-06, por exemplo)
    active_months = []
    for y in range(start_year, end_year + 1):
        for m in range(1, 13):
            # Verificar se algum produto tem dado nesse mês/ano
            has_data = any(
                (p, y, m) in prod_dict or (p, y, m) in imp_vol_dict or (p, y, m) in imp_usd_dict
                for p in products
            )
            if has_data:
                active_months.append((y, m))

    # Construir linhas detalhadas: produto x ano x mês
    rows_detail = []
    rows_total = []

    # Mapas para pivôs anuais e mensais
    prod_annual = defaultdict(lambda: defaultdict(float))
    imp_vol_annual = defaultdict(lambda: defaultdict(float))
    imp_usd_annual = defaultdict(lambda: defaultdict(float))

    for y, m in active_months:
        tot_pv = 0.0
        tot_iv = 0.0
        tot_iu = 0.0

        for p in products:
            pv = prod_dict.get((p, y, m), 0.0)
            iv = imp_vol_dict.get((p, y, m), 0.0)
            iu = imp_usd_dict.get((p, y, m), 0.0)

            tot_pv += pv
            tot_iv += iv
            tot_iu += iu

            prod_annual[p][y] += pv
            imp_vol_annual[p][y] += iv
            imp_usd_annual[p][y] += iu

            cost_imp = (iu / iv) if iv > 0 else None
            tot_vol = pv + iv
            tot_cost_obs = iu + (pv * domestic_cost)
            wac = (tot_cost_obs / tot_vol) if tot_vol > 0 else None
            lucro_adic = ((cost_imp - domestic_cost) * iv) if (cost_imp is not None and iv > 0) else 0.0

            rows_detail.append(
                {
                    "produto": p,
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
                    "lucro_bruto_adicional_usd": lucro_adic,
                }
            )

        # Totais nacionais do mês
        cost_imp_tot = (tot_iu / tot_iv) if tot_iv > 0 else None
        vol_tot_nac = tot_pv + tot_iv
        cost_tot_nac = tot_iu + (tot_pv * domestic_cost)
        wac_tot = (cost_tot_nac / vol_tot_nac) if vol_tot_nac > 0 else None
        lucro_adic_tot = ((cost_imp_tot - domestic_cost) * tot_iv) if (cost_imp_tot is not None and tot_iv > 0) else 0.0

        rows_total.append(
            {
                "ano": y,
                "mes": m,
                "mes_nome": MONTHS_PT[m - 1],
                "periodo": f"{y}-{m:02d}",
                "volume_produzido_brasil_barris": tot_pv,
                "volume_importado_barris": tot_iv,
                "dispendio_importacao_usd": tot_iu,
                "custo_medio_importado_usd_bbl": cost_imp_tot,
                "custo_producao_brasil_usd_bbl": domestic_cost,
                "volume_total_proxy_consumo_barris": vol_tot_nac,
                "custo_medio_ponderado_usd_bbl": wac_tot,
                "lucro_bruto_adicional_usd": lucro_adic_tot,
            }
        )

    df_detail = pd.DataFrame(rows_detail)
    df_total = pd.DataFrame(rows_total)

    years_list = sorted(list(set(y for y, m in active_months)))

    # --- 1. MATRIZES POR PRODUTO X ANO ---
    # Tabela 1: Produção Nacional Anual por Derivado
    df_prod_by_product = pd.DataFrame(
        [
            {"Produto": p, **{y: prod_annual[p][y] for y in years_list}, "Total_Periodo": sum(prod_annual[p][y] for y in years_list)}
            for p in products
        ]
    )
    df_prod_by_product.loc[len(df_prod_by_product)] = {
        "Produto": "TOTAL DERIVADOS",
        **{y: df_prod_by_product[y].sum() for y in years_list},
        "Total_Periodo": df_prod_by_product["Total_Periodo"].sum(),
    }

    # Tabela 2: Volume Importado Anual por Derivado
    df_imp_vol_by_product = pd.DataFrame(
        [
            {"Produto": p, **{y: imp_vol_annual[p][y] for y in years_list}, "Total_Periodo": sum(imp_vol_annual[p][y] for y in years_list)}
            for p in products
        ]
    )
    df_imp_vol_by_product.loc[len(df_imp_vol_by_product)] = {
        "Produto": "TOTAL DERIVADOS",
        **{y: df_imp_vol_by_product[y].sum() for y in years_list},
        "Total_Periodo": df_imp_vol_by_product["Total_Periodo"].sum(),
    }

    # Tabela 3: Dispêndio de Importação Anual por Derivado
    df_imp_usd_by_product = pd.DataFrame(
        [
            {"Produto": p, **{y: imp_usd_annual[p][y] for y in years_list}, "Total_Periodo": sum(imp_usd_annual[p][y] for y in years_list)}
            for p in products
        ]
    )
    df_imp_usd_by_product.loc[len(df_imp_usd_by_product)] = {
        "Produto": "TOTAL DERIVADOS",
        **{y: df_imp_usd_by_product[y].sum() for y in years_list},
        "Total_Periodo": df_imp_usd_by_product["Total_Periodo"].sum(),
    }

    # Tabela 4: Preço Médio Importado por Derivado (US$/bbl)
    df_price_by_product = pd.DataFrame(
        [
            {
                "Produto": p,
                **{
                    y: (imp_usd_annual[p][y] / imp_vol_annual[p][y]) if imp_vol_annual[p][y] > 0 else None
                    for y in years_list
                },
                "Preco_Medio_Periodo": (
                    sum(imp_usd_annual[p][y] for y in years_list) / sum(imp_vol_annual[p][y] for y in years_list)
                    if sum(imp_vol_annual[p][y] for y in years_list) > 0
                    else None
                ),
            }
            for p in products
        ]
    )
    tot_usd_per = df_imp_usd_by_product.loc[df_imp_usd_by_product["Produto"] == "TOTAL DERIVADOS", "Total_Periodo"].values[0]
    tot_vol_per = df_imp_vol_by_product.loc[df_imp_vol_by_product["Produto"] == "TOTAL DERIVADOS", "Total_Periodo"].values[0]
    df_price_by_product.loc[len(df_price_by_product)] = {
        "Produto": "TOTAL DERIVADOS",
        **{
            y: (df_imp_usd_by_product.loc[df_imp_usd_by_product["Produto"] == "TOTAL DERIVADOS", y].values[0] /
                df_imp_vol_by_product.loc[df_imp_vol_by_product["Produto"] == "TOTAL DERIVADOS", y].values[0])
            for y in years_list
        },
        "Preco_Medio_Periodo": tot_usd_per / tot_vol_per,
    }

    # --- 2. RESUMO CONSOLIDADO POR DERIVADO NO PERÍODO ---
    resumo_derivados_totais = []
    for p in products:
        pv_tot = sum(prod_annual[p][y] for y in years_list)
        iv_tot = sum(imp_vol_annual[p][y] for y in years_list)
        iu_tot = sum(imp_usd_annual[p][y] for y in years_list)
        cons_tot = pv_tot + iv_tot
        pm_imp = (iu_tot / iv_tot) if iv_tot > 0 else 0.0
        wac = (iu_tot + pv_tot * domestic_cost) / cons_tot if cons_tot > 0 else 0.0
        lucro_adic = ((pm_imp - domestic_cost) * iv_tot) if iv_tot > 0 else 0.0

        resumo_derivados_totais.append(
            {
                "Produto": p,
                "Volume_Produzido_Brasil_barris": pv_tot,
                "Volume_Importado_barris": iv_tot,
                "Volume_Total_Consumido_barris": cons_tot,
                "Share_Producao_Nacional_pct": (pv_tot / cons_tot * 100) if cons_tot > 0 else 0.0,
                "Share_Importacao_pct": (iv_tot / cons_tot * 100) if cons_tot > 0 else 0.0,
                "Dispendio_Importacao_USD": iu_tot,
                "Preco_Medio_Importacao_USD_bbl": pm_imp,
                "Custo_Medio_Ponderado_USD_bbl": wac,
                "Lucro_Bruto_Adicional_se_tudo_BR_25_USD": lucro_adic,
            }
        )
    df_resumo_derivados = pd.DataFrame(resumo_derivados_totais).sort_values(
        by="Volume_Total_Consumido_barris", ascending=False
    )

    # --- 3. MATRIZES MENSAIS GERAIS (Total Brasil) ---
    df_pivot_prod_total = df_total.pivot(index="mes_nome", columns="ano", values="volume_produzido_brasil_barris").reindex(MONTHS_PT)
    df_pivot_prod_total.loc["Total do Ano"] = df_pivot_prod_total.sum(axis=0)

    df_pivot_imp_vol_total = df_total.pivot(index="mes_nome", columns="ano", values="volume_importado_barris").reindex(MONTHS_PT)
    df_pivot_imp_vol_total.loc["Total do Ano"] = df_pivot_imp_vol_total.sum(axis=0)

    df_pivot_imp_usd_total = df_total.pivot(index="mes_nome", columns="ano", values="dispendio_importacao_usd").reindex(MONTHS_PT)
    df_pivot_imp_usd_total.loc["Total do Ano"] = df_pivot_imp_usd_total.sum(axis=0)

    # Gravar Excel com abas discriminadas por derivado e consolidadas
    excel_path = output_dir / "DERIVADOS_ANP_2000_2026.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_resumo_derivados.to_excel(writer, sheet_name="Resumo_Por_Derivado", index=False)
        df_prod_by_product.to_excel(writer, sheet_name="Prod_Por_Derivado_Anual", index=False)
        df_imp_vol_by_product.to_excel(writer, sheet_name="ImpVol_Por_Derivado_Anual", index=False)
        df_imp_usd_by_product.to_excel(writer, sheet_name="DispUSD_Por_Derivado_Anual", index=False)
        df_price_by_product.to_excel(writer, sheet_name="PrecoImp_Por_Derivado", index=False)
        df_pivot_prod_total.to_excel(writer, sheet_name="1_Producao_Total_Mensal")
        df_pivot_imp_vol_total.to_excel(writer, sheet_name="2_Importacao_Vol_Mensal")
        df_pivot_imp_usd_total.to_excel(writer, sheet_name="3_Dispendio_USD_Mensal")
        df_total.to_excel(writer, sheet_name="Serie_Total_Mensal", index=False)
        df_detail.to_excel(writer, sheet_name="Detalhe_Mensal_Por_Produto", index=False)

    # Gravar CSVs
    csv_detail = output_dir / "DERIVADOS_ANP_2000_2026_por_produto_mensal.csv"
    df_detail.to_csv(csv_detail, index=False, encoding="utf-8-sig")

    csv_total = output_dir / "DERIVADOS_ANP_2000_2026_mensal.csv"
    df_total.to_csv(csv_total, index=False, encoding="utf-8-sig")

    csv_resumo = output_dir / "DERIVADOS_ANP_2000_2026_resumo_por_derivado.csv"
    df_resumo_derivados.to_csv(csv_resumo, index=False, encoding="utf-8-sig")

    # Gravar JSON
    tot_prod_val = float(df_total["volume_produzido_brasil_barris"].sum())
    tot_imp_val = float(df_total["volume_importado_barris"].sum())
    tot_usd_val = float(df_total["dispendio_importacao_usd"].sum())
    tot_cons_val = tot_prod_val + tot_imp_val
    tot_lucro_val = float(df_total["lucro_bruto_adicional_usd"].sum())

    json_path = output_dir / "DERIVADOS_ANP_2000_2026.json"
    summary = {
        "periodo": f"{start_year} a {end_year}",
        "meses_observados": len(active_months),
        "total_derivados_analisados": len(products),
        "lista_derivados": products,
        "totais_nacionais": {
            "volume_produzido_brasil_barris": tot_prod_val,
            "volume_importado_barris": tot_imp_val,
            "volume_total_consumo_proxy_barris": tot_cons_val,
            "dispendio_importacao_usd": tot_usd_val,
            "preco_medio_importacao_usd_bbl": tot_usd_val / tot_imp_val if tot_imp_val > 0 else 0.0,
            "custo_medio_ponderado_usd_bbl": (tot_usd_val + tot_prod_val * domestic_cost) / tot_cons_val,
            "lucro_bruto_adicional_se_tudo_produzido_brasil_usd": tot_lucro_val,
        },
        "resumo_por_derivado": df_resumo_derivados.to_dict(orient="records"),
    }
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n[SUCESSO] Processamento CONCLUÍDO (Geral e Por Derivado)!")
    print(f"Planilha Excel consolidada: {excel_path}")
    print(f"  - Aba 1: Resumo_Por_Derivado (Ranking e estatísticas de cada derivado)")
    print(f"  - Aba 2: Prod_Por_Derivado_Anual (Produção de cada derivado por ano 2000–2026)")
    print(f"  - Aba 3: ImpVol_Por_Derivado_Anual (Volume importado por derivado por ano)")
    print(f"  - Aba 4: DispUSD_Por_Derivado_Anual (Dispêndio em US$ por derivado por ano)")
    print(f"  - Aba 5: PrecoImp_Por_Derivado (Preço unitário US$/barril por derivado)")
    print(f"  - Abas 6 a 10: Matrizes mensais do total e detalhes por produto mês a mês")

    print(f"\n--- RESUMO CONSOLIDADO POR DERIVADO (2000–2026) ---")
    print(f"{'DERIVADO':26s} | {'PROD (mi b)':12s} | {'IMP (mi b)':11s} | {'DISP (US$ mi)':14s} | {'PREÇO IMP':10s} | {'SHARE BR':9s}")
    print("-" * 92)
    for row in resumo_derivados_totais:
        print(
            f"{row['Produto']:26s} | "
            f"{row['Volume_Produzido_Brasil_barris']/1e6:12.2f} | "
            f"{row['Volume_Importado_barris']/1e6:11.2f} | "
            f"{row['Dispendio_Importacao_USD']/1e6:14.2f} | "
            f"US$ {row['Preco_Medio_Importacao_USD_bbl']:6.2f} | "
            f"{row['Share_Producao_Nacional_pct']:7.1f}%"
        )
    print("-" * 92)
    print(
        f"{'TOTAL DERIVADOS':26s} | "
        f"{tot_prod_val/1e6:12.2f} | "
        f"{tot_imp_val/1e6:11.2f} | "
        f"{tot_usd_val/1e6:14.2f} | "
        f"US$ {tot_usd_val/tot_imp_val:6.2f} | "
        f"{tot_prod_val/tot_cons_val*100:7.1f}%\n"
    )

    return summary


def main():
    parser = argparse.ArgumentParser(
        description="Processa séries ANP de derivados de petróleo (Geral e Por Derivado, 2000–2026)."
    )
    parser.add_argument("--producao", type=str, default=None, help="Caminho do producao-derivados-b.xlsx")
    parser.add_argument("--importacoes", type=str, default=None, help="Caminho do importacoes-exportacoes-b.xlsx")
    parser.add_argument("--saida", type=str, default=None, help="Diretório de saída")
    parser.add_argument("--ano-inicio", type=int, default=2000, help="Ano inicial (padrão: 2000)")
    parser.add_argument("--ano-fim", type=int, default=2026, help="Ano final (padrão: 2026)")
    parser.add_argument(
        "--custo-domestico",
        type=float,
        default=DEFAULT_DOMESTIC_COST,
        help="Custo doméstico de referência em US$/barril (padrão: 25.0)",
    )

    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    winpy = Path(r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython")

    prod_candidates = []
    if args.producao:
        prod_candidates.append(Path(args.producao))
    prod_candidates.extend(
        [
            root / "data" / "raw" / "anp" / "producao-derivados-b.xlsx",
            root / "data" / "raw" / "anp" / "producao-derivados-b.xls",
            root / "producao-derivados-b.xlsx",
            winpy / "dados" / "producao-derivados-b.xlsx",
            winpy / "producao-derivados-b.xlsx",
        ]
    )
    prod_file = resolve_file(prod_candidates, "produção de derivados")

    imp_candidates = []
    if args.importacoes:
        imp_candidates.append(Path(args.importacoes))
    imp_candidates.extend(
        [
            root / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xlsx",
            root / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xls",
            root / "importacoes-exportacoes-b.xlsx",
            winpy / "dados" / "importacoes-exportacoes-b.xlsx",
            winpy / "importacoes-exportacoes-b.xlsx",
        ]
    )
    imp_file = resolve_file(imp_candidates, "importações e exportações de derivados")

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
