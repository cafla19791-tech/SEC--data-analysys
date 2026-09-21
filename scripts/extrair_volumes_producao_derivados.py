#!/usr/bin/env python3
"""Extrai volumes de produção nacional de derivados (ANP), mensal 2010–2026.

Produtos: ÓLEO DIESEL, GASOLINA A, GLP e QUEROSENE DE AVIAÇÃO.

Fonte: planilha ANP "Produção Nacional de Derivados de Petróleo (barris)",
seção "Produção de derivados de petróleo por refinaria e produto".
Agrega todas as refinárias (visão BRASIL / REFINARIA = Tudo).

Unidade: barris (conforme título da planilha; rótulos internos do cache
podem trazer sufixo "(m3)", mas os totais batem com a série em barris).
"""

from __future__ import annotations

import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

DEFAULT_XLSX = Path("data/anp/anp_producao_nacional_derivados_barris.xlsx")
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

PRODUTOS_FOCO = ("ÓLEO DIESEL", "GASOLINA A", "GLP", "QUEROSENE DE AVIAÇÃO")
ANO_INICIO_PADRAO = 2010
ANO_FIM_PADRAO = 2026

# Pivot principal da produção por refinaria (visão BRASIL)
PIVOT_REF_PRODUCAO = "B35:"
FALLBACK_CACHE = (
    "xl/pivotCache/pivotCacheDefinition5.xml",
    "xl/pivotCache/pivotCacheRecords5.xml",
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


def find_producao_cache(xlsx: Path) -> tuple[str, str]:
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
            if not ref.startswith(PIVOT_REF_PRODUCAO):
                continue
            rels = ET.fromstring(z.read(f"xl/pivotTables/_rels/pivotTable{i}.xml.rels"))
            target = None
            for rel in rels:
                if rel.get("Type", "").endswith("/pivotCacheDefinition"):
                    target = rel.get("Target")
                    break
            if not target:
                raise ValueError(f"Pivot {i} sem cache definition")
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
    return FALLBACK_CACHE


def limpar_produto(nome: str) -> str:
    nome = str(nome).strip()
    nome = re.sub(r"\s*\((?:b|m3|bep)\)\s*$", "", nome, flags=re.IGNORECASE).strip()
    return nome


def carregar_producao_brasil(
    xlsx: Path,
    produtos: tuple[str, ...] = PRODUTOS_FOCO,
    ano_inicio: int = ANO_INICIO_PADRAO,
    ano_fim: int = ANO_FIM_PADRAO,
) -> pd.DataFrame:
    definition, records = find_producao_cache(xlsx)
    raw = parse_pivot_cache(xlsx, definition, records)
    required = {"ANO", "PRODUTO", *MES_ORDEM}
    missing = required - set(raw.columns)
    if missing:
        raise ValueError(f"Campos ausentes no pivot cache: {sorted(missing)}")

    raw = raw.copy()
    raw["produto"] = raw["PRODUTO"].map(limpar_produto)
    raw["ano"] = raw["ANO"].map(lambda x: int(float(x)))
    raw = raw[raw["produto"].isin(produtos)]
    raw = raw[(raw["ano"] >= ano_inicio) & (raw["ano"] <= ano_fim)]

    # agrega todas as refinárias / estados → Brasil
    rows = []
    for _, row in raw.iterrows():
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
                    "volume_barris": float(val),
                }
            )

    df = pd.DataFrame(rows)
    if df.empty:
        aggregated = df
    else:
        aggregated = (
            df.groupby(["ano", "mes", "mes_nome", "mes_ordem", "produto"], as_index=False)[
                "volume_barris"
            ]
            .sum()
        )

    # grade completa
    anos = range(ano_inicio, ano_fim + 1)
    grade = pd.MultiIndex.from_product(
        [anos, MES_ORDEM, produtos], names=["ano", "mes", "produto"]
    ).to_frame(index=False)
    grade["mes_nome"] = grade["mes"].map(MES_NOME)
    grade["mes_ordem"] = grade["mes"].map({m: i + 1 for i, m in enumerate(MES_ORDEM)})
    out = grade.merge(
        aggregated,
        on=["ano", "mes", "mes_nome", "mes_ordem", "produto"],
        how="left",
    )
    out["volume_barris"] = out["volume_barris"].fillna(0.0)
    out["unidade"] = "b"
    out["escopo"] = "Brasil - refinarias"
    return out.sort_values(["produto", "ano", "mes_ordem"]).reset_index(drop=True)


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
    cols = ["ano", "mes", "mes_nome", "mes_ordem", "produto", "volume_barris", "unidade", "escopo"]
    longo = df[cols]

    csv_longo = out_dir / "producao_diesel_gasolina_glp_qav_2010_2026.csv"
    longo.to_csv(csv_longo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    resumo = (
        longo.groupby(["ano", "produto"], as_index=False)["volume_barris"]
        .sum()
        .sort_values(["ano", "produto"])
    )
    csv_resumo = out_dir / "producao_diesel_gasolina_glp_qav_resumo_anual.csv"
    resumo.to_csv(csv_resumo, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    xlsx_path = out_dir / "producao_diesel_gasolina_glp_qav_2010_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
        longo.to_excel(writer, sheet_name="Longo", index=False)
        resumo.to_excel(writer, sheet_name="Resumo anual", index=False)

        matriz = longo.copy()
        matriz["ano_mes"] = matriz["ano"].astype(str) + "-" + matriz["mes"]
        wide = matriz.pivot_table(
            index="produto", columns="ano_mes", values="volume_barris", aggfunc="sum", fill_value=0.0
        )
        ordered = sorted(
            wide.columns,
            key=lambda c: (int(c.split("-")[0]), MES_ORDEM.index(c.split("-")[1])),
        )
        wide = wide.reindex(columns=ordered).reindex(index=list(PRODUTOS_FOCO))
        wide.to_excel(writer, sheet_name="Matriz volume barris")

        pivot_mes_ano(longo).to_excel(writer, sheet_name="Volume TOTAL barris", index=False)
        for produto in PRODUTOS_FOCO:
            sub = longo[longo["produto"] == produto]
            pivot_mes_ano(sub).to_excel(writer, sheet_name=sheet_name(produto), index=False)

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

    df = carregar_producao_brasil(
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
    for label, path in paths.items():
        print(f"{label}: {path}")


if __name__ == "__main__":
    main()
