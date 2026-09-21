from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd

from scripts.extrair_volumes_importacao_derivados import (
    carregar_volumes_barris,
    gerar_saidas,
    pivot_mes_ano,
)


def _build_mini_xlsx(path: Path) -> None:
    """Cria um xlsx mínimo com pivot cache no formato ANP (definição 12 / records 12)."""
    # Usamos a API pública parse_pivot_cache indiretamente via carregar; para o teste
    # de integração, geramos DataFrame direto e só testamos pivô/saídas.
    # O parse do cache real é coberto no teste com fixture XML zip.
    definition = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <main:cacheFields count="16">
    <main:cacheField name="ANO"><main:sharedItems count="1"><main:n v="2025"/></main:sharedItems></main:cacheField>
    <main:cacheField name="PRODUTO"><main:sharedItems count="2"><main:s v="ÓLEO DIESEL (b)"/><main:s v="GLP (b)"/></main:sharedItems></main:cacheField>
    <main:cacheField name="MOVIMENTO COMERCIAL"><main:sharedItems count="1"><main:s v="IMPORTAÇÃO"/></main:sharedItems></main:cacheField>
    <main:cacheField name="UNIDADE"><main:sharedItems count="1"><main:s v="b"/></main:sharedItems></main:cacheField>
    <main:cacheField name="JAN"/><main:cacheField name="FEV"/><main:cacheField name="MAR"/>
    <main:cacheField name="ABR"/><main:cacheField name="MAI"/><main:cacheField name="JUN"/>
    <main:cacheField name="JUL"/><main:cacheField name="AGO"/><main:cacheField name="SET"/>
    <main:cacheField name="OUT"/><main:cacheField name="NOV"/><main:cacheField name="DEZ"/>
  </main:cacheFields>
</main:pivotCacheDefinition>
"""
    records = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheRecords xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <main:r>
    <main:x v="0"/><main:x v="0"/><main:x v="0"/><main:x v="0"/>
    <main:n v="1000"/><main:n v="2000"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
  </main:r>
  <main:r>
    <main:x v="0"/><main:x v="1"/><main:x v="0"/><main:x v="0"/>
    <main:n v="100"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
  </main:r>
</main:pivotCacheRecords>
"""
    pivot = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotTableDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="t" cacheId="1">
  <main:location ref="B328:AC341" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
</main:pivotTableDefinition>
"""
    rels_pt = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition12.xml"/>
</Relationships>
"""
    rels_cache = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords12.xml"/>
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
        z.writestr("xl/pivotTables/pivotTable16.xml", pivot)
        z.writestr("xl/pivotTables/_rels/pivotTable16.xml.rels", rels_pt)
        z.writestr("xl/pivotCache/pivotCacheDefinition12.xml", definition)
        z.writestr("xl/pivotCache/pivotCacheRecords12.xml", records)
        z.writestr("xl/pivotCache/_rels/pivotCacheDefinition12.xml.rels", rels_cache)


def test_parse_e_carregar_volumes(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    df = carregar_volumes_barris(xlsx)
    assert set(df["produto"]) == {"ÓLEO DIESEL", "GLP"}
    diesel_jan = df[(df["produto"] == "ÓLEO DIESEL") & (df["mes"] == "JAN")]
    assert float(diesel_jan["volume_barris"].iloc[0]) == 1000.0


def test_gerar_saidas(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    out = tmp_path / "out"
    df = carregar_volumes_barris(xlsx)
    paths = gerar_saidas(df, out)
    assert paths["xlsx"].exists()
    longo = pd.read_csv(paths["csv_longo"], sep=";", decimal=",", encoding="utf-8-sig")
    assert longo["volume_barris"].sum() == 1000 + 2000 + 100
    pivot = pivot_mes_ano(df[df["produto"] == "ÓLEO DIESEL"])
    assert float(pivot.loc[pivot["Mês"] == "Janeiro", 2025].iloc[0]) == 1000.0
