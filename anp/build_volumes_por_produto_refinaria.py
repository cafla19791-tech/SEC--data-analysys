#!/usr/bin/env python3
"""Extract annual volumes by derivative product (imports + national production)
and national production by refinery, 2011–2026, from ANP Excel pivot caches.
"""

from __future__ import annotations

import json
import re
import zipfile
from collections import defaultdict
from pathlib import Path

import pandas as pd
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent
IMP = ROOT / "data" / "raw" / "anp" / "importacoes-exportacoes-b.xlsx"
PROD = ROOT / "data" / "raw" / "anp" / "producao-derivados-b.xlsx"

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
YEARS = list(range(2011, 2027))


def clean_product(name: str) -> str:
    if not isinstance(name, str):
        return str(name)
    s = name.strip()
    s = re.sub(r"\s*\((m3|b|bep)\)\s*", "", s, flags=re.I)
    return s.strip()


def load_pivot_cache(path: Path, def_name: str, rec_name: str):
    with zipfile.ZipFile(path) as z:
        def_xml = etree.fromstring(z.read(def_name))
        fields = []
        for f in def_xml.findall(".//m:cacheField", NS):
            name = f.get("name")
            items = []
            shared = f.find("m:sharedItems", NS)
            if shared is not None:
                for i in list(shared):
                    tag = i.tag.split("}")[-1]
                    if tag == "s":
                        items.append(i.get("v"))
                    elif tag == "n":
                        items.append(float(i.get("v")))
                    elif tag == "d":
                        items.append(i.get("v"))
                    elif tag == "m":
                        items.append(None)
                    else:
                        items.append(i.get("v"))
            fields.append({"name": name, "items": items})

        rec = etree.fromstring(z.read(rec_name))
        rows = []
        for r in rec.findall("m:r", NS):
            vals = []
            for idx, c in enumerate(r):
                tag = c.tag.split("}")[-1]
                if tag == "x":
                    xi = int(c.get("v"))
                    items = fields[idx]["items"]
                    vals.append(items[xi] if items else xi)
                elif tag == "n":
                    vals.append(float(c.get("v")))
                elif tag == "s":
                    vals.append(c.get("v"))
                elif tag == "m":
                    vals.append(None)
                else:
                    vals.append(c.get("v"))
            rows.append(vals)
    return fields, rows


def field_index(fields, name: str) -> int:
    for i, f in enumerate(fields):
        if f["name"] == name:
            return i
    raise KeyError(name)


def br(n, digits=0):
    if n is None:
        return "n/d"
    s = f"{n:,.{digits}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def main() -> None:
    # --- Imports by product (barris) ---
    # Cache 7: ANO, PRODUTO, MOVIMENTO=IMPORTAÇÃO, UNIDADE=b, months..., TOTAL
    imp_fields, imp_rows = load_pivot_cache(
        IMP,
        "xl/pivotCache/pivotCacheDefinition7.xml",
        "xl/pivotCache/pivotCacheRecords7.xml",
    )
    yi = field_index(imp_fields, "ANO")
    pi = field_index(imp_fields, "PRODUTO")
    mi = field_index(imp_fields, "MOVIMENTO COMERCIAL")
    ui = field_index(imp_fields, "UNIDADE")
    month_idxs = [
        field_index(imp_fields, m)
        for m in (
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
        )
    ]

    imports_by_product_year: dict[tuple[str, int], float] = defaultdict(float)
    for r in imp_rows:
        if str(r[mi]).strip().upper() != "IMPORTAÇÃO":
            continue
        if str(r[ui]).strip().lower() not in {"b", "barris"}:
            continue
        y = int(r[yi])
        if y not in YEARS:
            continue
        prod = clean_product(r[pi])
        tot = 0.0
        for idx in month_idxs:
            v = r[idx]
            if v is None:
                continue
            tot += float(v)
        imports_by_product_year[(prod, y)] += tot

    # --- Production by product and refinery (barris) ---
    # Cache 1: refinarias; product labels say (m3) but UNIDADE='b' and totals match barrel sheet
    prod_fields, prod_rows = load_pivot_cache(
        PROD,
        "xl/pivotCache/pivotCacheDefinition1.xml",
        "xl/pivotCache/pivotCacheRecords1.xml",
    )
    pyi = field_index(prod_fields, "ANO")
    ppi = field_index(prod_fields, "PRODUTO")
    pri = field_index(prod_fields, "REFINARIA")
    pui = field_index(prod_fields, "UNIDADE")
    pti = field_index(prod_fields, "TOTAL")

    production_by_product_year: dict[tuple[str, int], float] = defaultdict(float)
    production_by_refinery_year: dict[tuple[str, int], float] = defaultdict(float)
    production_by_refinery_product_year: dict[tuple[str, str, int], float] = defaultdict(
        float
    )

    for r in prod_rows:
        if str(r[pui]).strip().lower() not in {"b", "barris"}:
            continue
        y = int(r[pyi])
        if y not in YEARS:
            continue
        prod = clean_product(r[ppi])
        ref = str(r[pri]).strip()
        tot = float(r[pti] or 0.0)
        production_by_product_year[(prod, y)] += tot
        production_by_refinery_year[(ref, y)] += tot
        production_by_refinery_product_year[(ref, prod, y)] += tot

    products_imp = sorted({p for p, _ in imports_by_product_year})
    products_prod = sorted({p for p, _ in production_by_product_year})
    products_all = sorted(set(products_imp) | set(products_prod))
    refineries = sorted({r for r, _ in production_by_refinery_year})

    # Long tables
    rows_imp = [
        {"produto": p, "ano": y, "volume_importado_barris": imports_by_product_year[(p, y)]}
        for p in products_imp
        for y in YEARS
        if (p, y) in imports_by_product_year
    ]
    # include zeros for missing year/product in import list for completeness
    for p in products_imp:
        for y in YEARS:
            if (p, y) not in imports_by_product_year:
                rows_imp.append({"produto": p, "ano": y, "volume_importado_barris": 0.0})
    rows_imp = sorted(rows_imp, key=lambda x: (x["produto"], x["ano"]))

    rows_prod_prod = []
    for p in products_prod:
        for y in YEARS:
            rows_prod_prod.append(
                {
                    "produto": p,
                    "ano": y,
                    "volume_produzido_barris": production_by_product_year.get((p, y), 0.0),
                }
            )

    rows_ref = []
    for ref in refineries:
        for y in YEARS:
            rows_ref.append(
                {
                    "refinaria": ref,
                    "ano": y,
                    "volume_produzido_barris": production_by_refinery_year.get((ref, y), 0.0),
                }
            )

    # Wide pivots
    def wide(rows, id_col, value_col):
        df = pd.DataFrame(rows)
        return (
            df.pivot_table(index=id_col, columns="ano", values=value_col, aggfunc="sum")
            .reindex(columns=YEARS)
            .reset_index()
        )

    wide_imp = wide(rows_imp, "produto", "volume_importado_barris")
    wide_prod = wide(rows_prod_prod, "produto", "volume_produzido_barris")
    wide_ref = wide(rows_ref, "refinaria", "volume_produzido_barris")

    # Combined product comparison (import vs production)
    compare_rows = []
    for p in products_all:
        for y in YEARS:
            compare_rows.append(
                {
                    "produto": p,
                    "ano": y,
                    "volume_importado_barris": imports_by_product_year.get((p, y), 0.0),
                    "volume_produzido_barris": production_by_product_year.get((p, y), 0.0),
                }
            )
    df_compare = pd.DataFrame(compare_rows)

    notes = [
        "Fonte importações: ANP importacoes-exportacoes-b.xlsx — pivot cache de derivados "
        "(movimento IMPORTAÇÃO, unidade barris).",
        "Fonte produção: ANP producao-derivados-b.xls — pivot cache por refinaria/produto "
        "(unidade 'b' = barris; rótulos de produto no cache ainda trazem '(m3)', mas os "
        "totais batem com a aba em barris).",
        "Período: anos-calendário 2011–2026. Para 2026, a ANP pode estar parcial (YTD).",
        "Produção por refinaria = soma de todos os produtos da refinaria no ano.",
        "Importações por produto: 14 categorias ANP (sem combustíveis para navios/aeronaves "
        "na série de importação; esses entram sobretudo na exportação).",
        "Produção por produto: 15 categorias ANP de refinarias (inclui OUTROS ENERGÉTICOS).",
        "Valores anuais = campo TOTAL do cache (soma dos 12 meses, quando disponíveis).",
    ]

    payload = {
        "title": "Volumes anuais de derivados por produto e por refinaria (2011–2026)",
        "unit": "barris",
        "period": {"start": 2011, "end": 2026},
        "notes": notes,
        "products_import": products_imp,
        "products_production": products_prod,
        "refineries": refineries,
        "imports_by_product": rows_imp,
        "production_by_product": rows_prod_prod,
        "production_by_refinery": rows_ref,
        "production_by_refinery_product": [
            {
                "refinaria": ref,
                "produto": prod,
                "ano": y,
                "volume_produzido_barris": v,
            }
            for (ref, prod, y), v in sorted(production_by_refinery_product_year.items())
            if y in YEARS
        ],
    }

    json_path = OUT / "derivados_volumes_produto_refinaria_2011_2026.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    xlsx_path = OUT / "derivados_volumes_produto_refinaria_2011_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        wide_imp.to_excel(writer, sheet_name="Import_por_produto", index=False)
        wide_prod.to_excel(writer, sheet_name="Producao_por_produto", index=False)
        wide_ref.to_excel(writer, sheet_name="Producao_por_refinaria", index=False)
        df_compare.to_excel(writer, sheet_name="Comparativo_produto_ano", index=False)
        pd.DataFrame(rows_ref).to_excel(writer, sheet_name="Refinaria_longo", index=False)
        pd.DataFrame(payload["production_by_refinery_product"]).to_excel(
            writer, sheet_name="Refinaria_produto_ano", index=False
        )
        pd.DataFrame({"Nota": notes}).to_excel(writer, sheet_name="Notas", index=False)

    # Markdown summary (totals by product across period + latest full-ish year 2025)
    md = [
        "# Volumes anuais de derivados por produto e por refinaria (2011–2026)\n",
        "Unidade: **barris**. Fontes ANP (pivot caches das planilhas oficiais).\n",
        "## Importação por produto — 2025 (barris)\n",
        "| Produto | Volume importado |",
        "|---|---:|",
    ]
    for p in products_imp:
        md.append(f"| {p} | {br(imports_by_product_year.get((p, 2025), 0.0))} |")

    md += [
        "\n## Produção nacional por produto — 2025 (barris)\n",
        "| Produto | Volume produzido |",
        "|---|---:|",
    ]
    for p in products_prod:
        md.append(f"| {p} | {br(production_by_product_year.get((p, 2025), 0.0))} |")

    md += [
        "\n## Produção por refinaria — 2025 (barris)\n",
        "| Refinaria | Volume produzido |",
        "|---|---:|",
    ]
    for ref, vol in sorted(
        ((r, production_by_refinery_year.get((r, 2025), 0.0)) for r in refineries),
        key=lambda x: -x[1],
    ):
        md.append(f"| {ref} | {br(vol)} |")

    md.append("\n## Notas\n")
    for n in notes:
        md.append(f"- {n}")

    md_path = OUT / "derivados_volumes_produto_refinaria_2011_2026.md"
    md_path.write_text("\n".join(md) + "\n", encoding="utf-8")

    print(f"Wrote {json_path}")
    print(f"Wrote {xlsx_path}")
    print(f"Wrote {md_path}")
    print("import products:", products_imp)
    print("prod products:", products_prod)
    print("refineries:", refineries)
    # sanity totals
    for y in (2011, 2024, 2025, 2026):
        i = sum(imports_by_product_year.get((p, y), 0) for p in products_imp)
        p = sum(production_by_product_year.get((pr, y), 0) for pr in products_prod)
        print(f"{y}: import_total={i:,.0f} prod_total={p:,.0f}")


if __name__ == "__main__":
    main()
