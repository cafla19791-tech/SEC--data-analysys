#!/usr/bin/env python3
"""Extrai volumes (barris) e dispêndios (US$ FOB) de importação de derivados ANP.

Foco: ÓLEO DIESEL, GASOLINA A, GLP e QUEROSENE DE AVIAÇÃO, mensal 2010–2026.

Fonte: planilha ANP em barris (Google Sheets / Excel). Lê os pivot caches de
importação de derivados (volumes em B328 e dispêndio em B389).
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

# Produtos solicitados (gasolina = GASOLINA A; diesel = ÓLEO DIESEL; QAV)
PRODUTOS_FOCO = ("ÓLEO DIESEL", "GASOLINA A", "GLP", "QUEROSENE DE AVIAÇÃO")
ANO_INICIO_PADRAO = 2010
ANO_FIM_PADRAO = 2026

# Pivôs conhecidos na planilha ANP (atualizada em 28/08/2026)
PIVOT_REF_VOLUME = "B328:"
PIVOT_REF_DISPENDIO = "B389:"
FALLBACK_VOLUME = (
    "xl/pivotCache/pivotCacheDefinition12.xml",
    "xl/pivotCache/pivotCacheRecords12.xml",
)
FALLBACK_DISPENDIO = (
    "xl/pivotCache/pivotCacheDefinition4.xml",
    "xl/pivotCache/pivotCacheRecords4.xml",
)


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


def _resolve_cache_from_pivot(z: zipfile.ZipFile, pivot_index: int) -> tuple[str, str]:
    rels = ET.fromstring(z.read(f"xl/pivotTables/_rels/pivotTable{pivot_index}.xml.rels"))
    target = None
    for rel in rels:
        if rel.get("Type", "").endswith("/pivotCacheDefinition"):
            target = rel.get("Target")
            break
    if not target:
        raise ValueError(f"Pivot {pivot_index} sem cache definition")
    def_name = Path(target).name
    def_path = f"xl/pivotCache/{def_name}"
    def_rels = ET.fromstring(z.read(f"xl/pivotCache/_rels/{def_name}.rels"))
    rec_target = None
    for rel in def_rels:
        if rel.get("Type", "").endswith("/pivotCacheRecords"):
            rec_target = rel.get("Target")
            break
    if not rec_target:
        raise ValueError(f"Cache {def_name} sem records")
    return def_path, f"xl/pivotCache/{Path(rec_target).name}"


def find_cache_by_ref(xlsx: Path, ref_prefix: str, fallback: tuple[str, str]) -> tuple[str, str]:
    with zipfile.ZipFile(xlsx) as z:
        for i in range(1, 30):
            path = f"xl/pivotTables/pivotTable{i}.xml"
            if path not in z.namelist():
                continue
            root = ET.fromstring(z.read(path))
            loc = root.find("main:location", NS)
            if loc is None:
                continue
            ref = loc.get("ref") or ""
            if not ref.startswith(ref_prefix):
                continue
            return _resolve_cache_from_pivot(z, i)
    return fallback


def limpar_produto(nome: str) -> str:
    nome = str(nome).strip()
    nome = re.sub(r"\s*\(b\)\s*$", "", nome, flags=re.IGNORECASE).strip()
    return nome


def cache_wide_to_long(
    raw: pd.DataFrame,
    value_col: str,
    produtos: tuple[str, ...] | None = None,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
) -> pd.DataFrame:
    required = {"ANO", "PRODUTO", *MES_ORDEM}
    missing = required - set(raw.columns)
    if missing:
        raise ValueError(f"Campos ausentes no pivot cache: {sorted(missing)}")

    df = raw.copy()
    if "MOVIMENTO COMERCIAL" in df.columns:
        df = df[df["MOVIMENTO COMERCIAL"].astype(str).str.upper().str.startswith("IMPORT")]

    df["produto"] = df["PRODUTO"].map(limpar_produto)
    df["ano"] = df["ANO"].map(lambda x: int(float(x)))

    if produtos:
        df = df[df["produto"].isin(produtos)]
    if ano_inicio is not None:
        df = df[df["ano"] >= ano_inicio]
    if ano_fim is not None:
        df = df[df["ano"] <= ano_fim]

    rows = []
    for _, row in df.iterrows():
        for mes in MES_ORDEM:
            val = row.get(mes)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                val = 0.0
            rows.append(
                {
                    "ano": int(row["ano"]),
                    "mes": mes,
                    "mes_nome": MES_NOME[mes],
                    "mes_ordem": MES_ORDEM.index(mes) + 1,
                    "produto": row["produto"],
                    value_col: float(val),
                }
            )

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return (
        out.groupby(["ano", "mes", "mes_nome", "mes_ordem", "produto"], as_index=False)[value_col]
        .sum()
        .sort_values(["produto", "ano", "mes_ordem"])
        .reset_index(drop=True)
    )


def carregar_volumes_barris(
    xlsx: Path,
    produtos: tuple[str, ...] | None = None,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
) -> pd.DataFrame:
    definition, records = find_cache_by_ref(xlsx, PIVOT_REF_VOLUME, FALLBACK_VOLUME)
    raw = parse_pivot_cache(xlsx, definition, records)
    df = cache_wide_to_long(raw, "volume_barris", produtos, ano_inicio, ano_fim)
    if not df.empty:
        df["unidade_volume"] = "b"
    return df


def carregar_dispendio_usd(
    xlsx: Path,
    produtos: tuple[str, ...] | None = None,
    ano_inicio: int | None = None,
    ano_fim: int | None = None,
) -> pd.DataFrame:
    definition, records = find_cache_by_ref(xlsx, PIVOT_REF_DISPENDIO, FALLBACK_DISPENDIO)
    raw = parse_pivot_cache(xlsx, definition, records)
    df = cache_wide_to_long(raw, "dispendio_usd_fob", produtos, ano_inicio, ano_fim)
    if not df.empty:
        df["unidade_dispendio"] = "US$ FOB"
    return df


def carregar_volumes_e_dispendios(
    xlsx: Path,
    produtos: tuple[str, ...] = PRODUTOS_FOCO,
    ano_inicio: int = ANO_INICIO_PADRAO,
    ano_fim: int = ANO_FIM_PADRAO,
) -> pd.DataFrame:
    vol = carregar_volumes_barris(xlsx, produtos, ano_inicio, ano_fim)
    disp = carregar_dispendio_usd(xlsx, produtos, ano_inicio, ano_fim)
    keys = ["ano", "mes", "mes_nome", "mes_ordem", "produto"]
    merged = vol.merge(disp[keys + ["dispendio_usd_fob"]], on=keys, how="outer")
    merged["volume_barris"] = merged["volume_barris"].fillna(0.0)
    merged["dispendio_usd_fob"] = merged["dispendio_usd_fob"].fillna(0.0)
    merged["unidade_volume"] = "b"
    merged["unidade_dispendio"] = "US$ FOB"
    # Garante grade completa mês×ano×produto no período
    anos = range(ano_inicio, ano_fim + 1)
    grade = pd.MultiIndex.from_product(
        [anos, MES_ORDEM, produtos], names=["ano", "mes", "produto"]
    ).to_frame(index=False)
    grade["mes_nome"] = grade["mes"].map(MES_NOME)
    grade["mes_ordem"] = grade["mes"].map({m: i + 1 for i, m in enumerate(MES_ORDEM)})
    merged = grade.merge(
        merged[keys + ["volume_barris", "dispendio_usd_fob", "unidade_volume", "unidade_dispendio"]],
        on=keys,
        how="left",
    )
    merged["volume_barris"] = merged["volume_barris"].fillna(0.0)
    merged["dispendio_usd_fob"] = merged["dispendio_usd_fob"].fillna(0.0)
    merged["unidade_volume"] = "b"
    merged["unidade_dispendio"] = "US$ FOB"
    return merged.sort_values(["produto", "ano", "mes_ordem"]).reset_index(drop=True)


def pivot_mes_ano(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
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


def sheet_name(prefix: str, produto: str) -> str:
    name = f"{prefix} {produto}"[:31]
    for ch in "[]:*?/\\":
        name = name.replace(ch, "-")
    return name


def gerar_saidas(df: pd.DataFrame, out_dir: Path) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)

    cols = [
        "ano",
        "mes",
        "mes_nome",
        "mes_ordem",
        "produto",
        "volume_barris",
        "dispendio_usd_fob",
        "unidade_volume",
        "unidade_dispendio",
    ]
    longo = df[cols]

    csv_longo = out_dir / "volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.csv"
    longo.to_csv(csv_longo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    resumo = (
        longo.groupby(["ano", "produto"], as_index=False)[["volume_barris", "dispendio_usd_fob"]]
        .sum()
        .sort_values(["ano", "produto"])
    )
    csv_resumo = out_dir / "volumes_dispendios_diesel_gasolina_glp_qav_resumo_anual.csv"
    resumo.to_csv(csv_resumo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    xlsx_path = out_dir / "volumes_dispendios_diesel_gasolina_glp_qav_2010_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        longo.to_excel(writer, sheet_name="Longo", index=False)
        resumo.to_excel(writer, sheet_name="Resumo anual", index=False)

        # Matriz produto × mês-ano (largo) para leitura rápida
        matriz = longo.copy()
        matriz["ano_mes"] = matriz["ano"].astype(str) + "-" + matriz["mes"]
        for metric, sheet in [
            ("volume_barris", "Matriz volume barris"),
            ("dispendio_usd_fob", "Matriz dispendio USD"),
        ]:
            wide = matriz.pivot_table(
                index="produto", columns="ano_mes", values=metric, aggfunc="sum", fill_value=0.0
            )
            # ordena colunas cronologicamente
            ordered = sorted(
                wide.columns,
                key=lambda c: (int(c.split("-")[0]), MES_ORDEM.index(c.split("-")[1])),
            )
            wide = wide.reindex(columns=ordered).reindex(index=list(PRODUTOS_FOCO))
            wide.to_excel(writer, sheet_name=sheet)

        for produto in PRODUTOS_FOCO:
            sub = longo[longo["produto"] == produto]
            pivot_mes_ano(sub, "volume_barris").to_excel(
                writer, sheet_name=sheet_name("Vol", produto), index=False
            )
            pivot_mes_ano(sub, "dispendio_usd_fob").to_excel(
                writer, sheet_name=sheet_name("Disp", produto), index=False
            )

    return {"csv_longo": csv_longo, "csv_resumo": csv_resumo, "xlsx": xlsx_path}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_XLSX)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO_PADRAO)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM_PADRAO)
    args = parser.parse_args()

    if not args.source.exists():
        raise SystemExit(f"Arquivo fonte não encontrado: {args.source}")

    df = carregar_volumes_e_dispendios(
        args.source,
        produtos=PRODUTOS_FOCO,
        ano_inicio=args.ano_inicio,
        ano_fim=args.ano_fim,
    )
    paths = gerar_saidas(df, args.out_dir)

    print(f"Linhas: {len(df)}")
    print(f"Produtos: {sorted(df['produto'].unique())}")
    print(f"Anos: {df['ano'].min()}–{df['ano'].max()}")
    print(f"Volume total (barris): {df['volume_barris'].sum():,.2f}")
    print(f"Dispêndio total (US$ FOB): {df['dispendio_usd_fob'].sum():,.2f}")
    for label, path in paths.items():
        print(f"{label}: {path}")


if __name__ == "__main__":
    main()
