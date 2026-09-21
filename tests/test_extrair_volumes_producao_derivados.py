from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd

from scripts.extrair_volumes_producao_derivados import (
    PRODUTOS_FOCO,
    carregar_producao_brasil,
    gerar_saidas,
    pivot_mes_ano,
)


def _build_mini_xlsx(path: Path) -> None:
    definition = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="3">
  <main:cacheFields count="18">
    <main:cacheField name="PRODUTO"><main:sharedItems count="2"><main:s v="ÓLEO DIESEL (m3)"/><main:s v="GLP (m3)"/></main:sharedItems></main:cacheField>
    <main:cacheField name="ANO"><main:sharedItems count="1"><main:n v="2015"/></main:sharedItems></main:cacheField>
    <main:cacheField name="ESTADO"><main:sharedItems count="1"><main:s v="RJ"/></main:sharedItems></main:cacheField>
    <main:cacheField name="REFINARIA"><main:sharedItems count="2"><main:s v="REDUC"/><main:s v="REPLAN"/></main:sharedItems></main:cacheField>
    <main:cacheField name="UNIDADE"/><main:cacheField name="JAN"/><main:cacheField name="FEV"/><main:cacheField name="MAR"/>
    <main:cacheField name="ABR"/><main:cacheField name="MAI"/><main:cacheField name="JUN"/>
    <main:cacheField name="JUL"/><main:cacheField name="AGO"/><main:cacheField name="SET"/>
    <main:cacheField name="OUT"/><main:cacheField name="NOV"/><main:cacheField name="DEZ"/>
    <main:cacheField name="TOTAL"/>
  </main:cacheFields>
</main:pivotCacheDefinition>
"""
    # two refineries for diesel + one for GLP → diesel should sum
    records = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheRecords xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <main:r>
    <main:x v="0"/><main:x v="0"/><main:x v="0"/><main:x v="0"/><main:m/>
    <main:n v="1000"/><main:n v="500"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="1500"/>
  </main:r>
  <main:r>
    <main:x v="0"/><main:x v="0"/><main:x v="0"/><main:x v="1"/><main:m/>
    <main:n v="2000"/><main:n v="700"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="2700"/>
  </main:r>
  <main:r>
    <main:x v="1"/><main:x v="0"/><main:x v="0"/><main:x v="0"/><main:m/>
    <main:n v="100"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="100"/>
  </main:r>
</main:pivotCacheRecords>
"""
    pivot = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotTableDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="prod" cacheId="1">
  <main:location ref="B35:AC48" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
</main:pivotTableDefinition>
"""
    rels_pt = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition5.xml"/>
</Relationships>
"""
    rels_cache = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords5.xml"/>
</Relationships>
"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>
"""
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("xl/pivotTables/pivotTable6.xml", pivot)
        z.writestr("xl/pivotTables/_rels/pivotTable6.xml.rels", rels_pt)
        z.writestr("xl/pivotCache/pivotCacheDefinition5.xml", definition)
        z.writestr("xl/pivotCache/pivotCacheRecords5.xml", records)
        z.writestr("xl/pivotCache/_rels/pivotCacheDefinition5.xml.rels", rels_cache)


def test_agrega_refinarias_brasil(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    df = carregar_producao_brasil(xlsx, ano_inicio=2015, ano_fim=2015)
    assert len(df) == 4 * 12
    diesel_jan = df[(df["produto"] == "ÓLEO DIESEL") & (df["mes"] == "JAN")]
    assert float(diesel_jan["volume_barris"].iloc[0]) == 3000.0  # 1000+2000
    glp_jan = df[(df["produto"] == "GLP") & (df["mes"] == "JAN")]
    assert float(glp_jan["volume_barris"].iloc[0]) == 100.0


def test_gerar_saidas_producao(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    out = tmp_path / "out"
    df = carregar_producao_brasil(xlsx, ano_inicio=2015, ano_fim=2015)
    paths = gerar_saidas(df, out)
    assert paths["xlsx"].exists()
    longo = pd.read_csv(paths["csv_longo"], sep=";", decimal=",", encoding="utf-8-sig")
    assert set(longo["produto"]) == set(PRODUTOS_FOCO)
    pivot = pivot_mes_ano(df[df["produto"] == "ÓLEO DIESEL"])
    assert float(pivot.loc[pivot["Mês"] == "Janeiro", 2015].iloc[0]) == 3000.0
