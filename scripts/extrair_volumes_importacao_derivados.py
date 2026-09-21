#!/usr/bin/env python3
"""Extrai volumes importados de derivados de petróleo por produto (ANP).

Fonte principal: planilha ANP em barris (Google Sheets / Excel), seção
"Importação de derivados de petróleo por produto - 2000-2026 (b)".

A planilha usa tabelas dinâmicas; os valores por produto ficam no pivot cache
(não só no total "(Tudo)" visível). Este script lê o cache correspondente.

Saídas em output/anp/:
  - volumes_importacao_derivados_barris_longo.csv
  - volumes_importacao_derivados_barris_resumo_anual.csv
  - volumes_importacao_derivados_barris.xlsx  (pivôs mês×ano por produto)
"""

from __future__ import annotations

import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

DEFAULT_XLSX = Path("data/anp/anp_importacoes_exportacoes_barris.xlsx")
DEFAULT_OUT_DIR = Path("output/anp")

MES_ORDEM = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MES_NOME = {
    "JAN": "Janeiro",
    "FEV": "Fevereiro",
    "MAR": "Março",
    "ABR": "Abril",
    "MAI": "Maio",
    "JUN": "Junho",
    "JUL": "Julho",
    "AGO": "Agosto",
    "SET": "Setembro",
    "OUT": "Outubro",
    "NOV": "Novembro",
    "DEZ": "Dezembro",
}


def _shared_item(node: ET.Element):
    tag = node.tag.split("}")[-1]
    if tag == "s":
        return node.get("v")
    if tag == "n":
        return float(node.get("v")) if node.get("v") is not None else None
    if tag == "m":
        return None
    if tag == "d":
        return node.get("v")
    return node.get("v")


def parse_pivot_cache(xlsx: Path, definition: str, records: str) -> pd.DataFrame:
    with zipfile.ZipFile(xlsx) as z:
        def_root = ET.fromstring(z.read(definition))
        fields: list[str] = []
        shared: list[list] = []
        for cf in def_root.findall("main:cacheFields/main:cacheField", NS):
            fields.append(cf.get("name") or "")
            items: list = []
            shared_items = cf.find("main:sharedItems", NS)
            if shared_items is not None:
                items = [_shared_item(child) for child in list(shared_items)]
            shared.append(items)

        rec_root = ET.fromstring(z.read(records))
        rows = []
        for r in rec_root.findall("main:r", NS):
            vals = []
            for i, child in enumerate(list(r)):
                tag = child.tag.split("}")[-1]
                if tag == "x":
                    idx = int(child.get("v"))
                    vals.append(shared[i][idx] if i < len(shared) and shared[i] else idx)
                elif tag == "n":
                    vals.append(float(child.get("v")))
                elif tag == "s":
                    vals.append(child.get("v"))
                elif tag == "m":
                    vals.append(None)
                else:
                    vals.append(child.get("v"))
            while len(vals) < len(fields):
                vals.append(None)
            rows.append(dict(zip(fields, vals)))
    return pd.DataFrame(rows)


def find_import_volumes_cache(xlsx: Path) -> tuple[str, str]:
    """Localiza o pivot cache da tabela em B328 (importação derivados, barris)."""
    with zipfile.ZipFile(xlsx) as z:
        # Preferência: pivotTable cujo location começa em B328
        for i in range(1, 30):
            path = f"xl/pivotTables/pivotTable{i}.xml"
            if path not in z.namelist():
                continue
            root = ET.fromstring(z.read(path))
            loc = root.find("main:location", NS)
            if loc is None:
                continue
            ref = loc.get("ref") or ""
            if not re.match(r"^B328:", ref):
                continue
            rels = ET.fromstring(z.read(f"xl/pivotTables/_rels/pivotTable{i}.xml.rels"))
            target = None
            for rel in rels:
                if rel.get("Type", "").endswith("/pivotCacheDefinition"):
                    target = rel.get("Target")
                    break
            if not target:
                raise ValueError(f"Pivot {i} sem cache definition")
            # Target is like ../pivotCache/pivotCacheDefinition12.xml
            def_name = Path(target).name
            def_path = f"xl/pivotCache/{def_name}"
            # records via rels of definition
            def_rels = ET.fromstring(z.read(f"xl/pivotCache/_rels/{def_name}.rels"))
            rec_target = None
            for rel in def_rels:
                if rel.get("Type", "").endswith("/pivotCacheRecords"):
                    rec_target = rel.get("Target")
                    break
            if not rec_target:
                raise ValueError(f"Cache {def_name} sem records")
            rec_path = f"xl/pivotCache/{Path(rec_target).name}"
            return def_path, rec_path

    # Fallback conhecido da planilha ANP atualizada em ago/2026
    return (
        "xl/pivotCache/pivotCacheDefinition12.xml",
        "xl/pivotCache/pivotCacheRecords12.xml",
    )


def limpar_produto(nome: str) -> str:
    nome = str(nome).strip()
    nome = re.sub(r"\s*\(b\)\s*$", "", nome, flags=re.IGNORECASE).strip()
    return nome


def carregar_volumes_barris(xlsx: Path) -> pd.DataFrame:
    definition, records = find_import_volumes_cache(xlsx)
    raw = parse_pivot_cache(xlsx, definition, records)

    required = {"ANO", "PRODUTO", *MES_ORDEM}
    missing = required - set(raw.columns)
    if missing:
        raise ValueError(f"Campos ausentes no pivot cache: {sorted(missing)}")

    if "MOVIMENTO COMERCIAL" in raw.columns:
        raw = raw[raw["MOVIMENTO COMERCIAL"].astype(str).str.upper().str.startswith("IMPORT")]

    long_rows = []
    for _, row in raw.iterrows():
        produto = limpar_produto(row["PRODUTO"])
        ano = int(float(row["ANO"]))
        for mes in MES_ORDEM:
            val = row.get(mes)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            val = float(val)
            if val == 0.0:
                # mantém zeros de meses já reportados (úteis para série)
                pass
            long_rows.append(
                {
                    "ano": ano,
                    "mes": mes,
                    "mes_nome": MES_NOME[mes],
                    "mes_ordem": MES_ORDEM.index(mes) + 1,
                    "produto": produto,
                    "volume_barris": val,
                    "unidade": "b",
                }
            )

    df = pd.DataFrame(long_rows)
    df = (
        df.groupby(["ano", "mes", "mes_nome", "mes_ordem", "produto", "unidade"], as_index=False)[
            "volume_barris"
        ]
        .sum()
        .sort_values(["produto", "ano", "mes_ordem"])
        .reset_index(drop=True)
    )
    return df


def pivot_mes_ano(df: pd.DataFrame, value_col: str = "volume_barris") -> pd.DataFrame:
    pivot = (
        df.pivot_table(index="mes", columns="ano", values=value_col, aggfunc="sum", fill_value=0.0)
        .reindex(MES_ORDEM)
    )
    pivot.index = [MES_NOME.get(m, m) for m in pivot.index]
    pivot.loc["Total do Ano"] = pivot.sum(axis=0)

    anos = [c for c in pivot.columns if isinstance(c, (int, float))]
    if len(anos) >= 2:
        a_ref, a_base = int(anos[-1]), int(anos[-2])
        ytd_ref = pivot.loc[[MES_NOME[m] for m in MES_ORDEM], a_ref].cumsum()
        ytd_base = pivot.loc[[MES_NOME[m] for m in MES_ORDEM], a_base].cumsum()
        var = ((ytd_ref / ytd_base.replace(0, pd.NA)) - 1.0) * 100.0
        col_name = f"VARIAÇÃO ACUM. {a_ref}/{a_base} (%)"
        pivot[col_name] = pd.concat([var, pd.Series({"Total do Ano": pd.NA})])

    pivot.index.name = "Mês"
    return pivot.reset_index()


def sheet_name(produto: str) -> str:
    name = f"Vol {produto}"[:31]
    for ch in "[]:*?/\\":
        name = name.replace(ch, "-")
    return name


def gerar_saidas(df: pd.DataFrame, out_dir: Path) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_longo = out_dir / "volumes_importacao_derivados_barris_longo.csv"
    df.to_csv(csv_longo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    resumo = (
        df.groupby(["ano", "produto"], as_index=False)["volume_barris"]
        .sum()
        .sort_values(["ano", "produto"])
    )
    csv_resumo = out_dir / "volumes_importacao_derivados_barris_resumo_anual.csv"
    resumo.to_csv(csv_resumo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    xlsx_path = out_dir / "volumes_importacao_derivados_barris.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Longo", index=False)
        resumo.to_excel(writer, sheet_name="Resumo anual", index=False)
        pivot_mes_ano(df).to_excel(writer, sheet_name="Volume TOTAL barris", index=False)
        for produto in sorted(df["produto"].unique()):
            sub = df[df["produto"] == produto]
            pivot_mes_ano(sub).to_excel(writer, sheet_name=sheet_name(produto), index=False)

    return {"csv_longo": csv_longo, "csv_resumo": csv_resumo, "xlsx": xlsx_path}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_XLSX)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    args = parser.parse_args()

    if not args.source.exists():
        raise SystemExit(f"Arquivo fonte não encontrado: {args.source}")

    df = carregar_volumes_barris(args.source)
    paths = gerar_saidas(df, args.out_dir)

    print(f"Linhas: {len(df)}")
    print(f"Produtos: {df['produto'].nunique()} -> {sorted(df['produto'].unique())}")
    print(f"Anos: {df['ano'].min()}–{df['ano'].max()}")
    print(f"Volume total (barris): {df['volume_barris'].sum():,.2f}")
    for label, path in paths.items():
        print(f"{label}: {path}")


if __name__ == "__main__":
    main()
