from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd

from scripts.extrair_volumes_importacao_derivados import (
    PRODUTOS_FOCO,
    carregar_volumes_e_dispendios,
    gerar_saidas,
    pivot_mes_ano,
)


def _build_mini_xlsx(path: Path) -> None:
    definition_vol = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <main:cacheFields count="16">
    <main:cacheField name="ANO"><main:sharedItems count="1"><main:n v="2015"/></main:sharedItems></main:cacheField>
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
    records_vol = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    definition_disp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="2">
  <main:cacheFields count="15">
    <main:cacheField name="ANO"><main:sharedItems count="1"><main:n v="2015"/></main:sharedItems></main:cacheField>
    <main:cacheField name="PRODUTO"><main:sharedItems count="2"><main:s v="ÓLEO DIESEL"/><main:s v="GLP"/></main:sharedItems></main:cacheField>
    <main:cacheField name="UNIDADE"><main:sharedItems count="1"><main:s v="US$ FOB"/></main:sharedItems></main:cacheField>
    <main:cacheField name="JAN"/><main:cacheField name="FEV"/><main:cacheField name="MAR"/>
    <main:cacheField name="ABR"/><main:cacheField name="MAI"/><main:cacheField name="JUN"/>
    <main:cacheField name="JUL"/><main:cacheField name="AGO"/><main:cacheField name="SET"/>
    <main:cacheField name="OUT"/><main:cacheField name="NOV"/><main:cacheField name="DEZ"/>
  </main:cacheFields>
</main:pivotCacheDefinition>
"""
    records_disp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotCacheRecords xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <main:r>
    <main:x v="0"/><main:x v="0"/><main:x v="0"/>
    <main:n v="50000"/><main:n v="70000"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
  </main:r>
  <main:r>
    <main:x v="0"/><main:x v="1"/><main:x v="0"/>
    <main:n v="9000"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
    <main:n v="0"/><main:n v="0"/><main:n v="0"/><main:n v="0"/>
  </main:r>
</main:pivotCacheRecords>
"""
    pivot_vol = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotTableDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="vol" cacheId="1">
  <main:location ref="B328:AC341" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
</main:pivotTableDefinition>
"""
    pivot_disp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<main:pivotTableDefinition xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="disp" cacheId="2">
  <main:location ref="B389:AC402" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
</main:pivotTableDefinition>
"""
    rels_vol = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition12.xml"/>
</Relationships>
"""
    rels_disp = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition4.xml"/>
</Relationships>
"""
    rels_c12 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords12.xml"/>
</Relationships>
"""
    rels_c4 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords4.xml"/>
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
        z.writestr("xl/pivotTables/pivotTable16.xml", pivot_vol)
        z.writestr("xl/pivotTables/_rels/pivotTable16.xml.rels", rels_vol)
        z.writestr("xl/pivotTables/pivotTable13.xml", pivot_disp)
        z.writestr("xl/pivotTables/_rels/pivotTable13.xml.rels", rels_disp)
        z.writestr("xl/pivotCache/pivotCacheDefinition12.xml", definition_vol)
        z.writestr("xl/pivotCache/pivotCacheRecords12.xml", records_vol)
        z.writestr("xl/pivotCache/_rels/pivotCacheDefinition12.xml.rels", rels_c12)
        z.writestr("xl/pivotCache/pivotCacheDefinition4.xml", definition_disp)
        z.writestr("xl/pivotCache/pivotCacheRecords4.xml", records_disp)
        z.writestr("xl/pivotCache/_rels/pivotCacheDefinition4.xml.rels", rels_c4)


def test_carregar_volumes_e_dispendios_filtrados(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    df = carregar_volumes_e_dispendios(xlsx, produtos=PRODUTOS_FOCO, ano_inicio=2015, ano_fim=2015)
    # grade completa: 4 produtos × 12 meses
    assert len(df) == 4 * 12
    diesel_jan = df[(df["produto"] == "ÓLEO DIESEL") & (df["mes"] == "JAN")]
    assert float(diesel_jan["volume_barris"].iloc[0]) == 1000.0
    assert float(diesel_jan["dispendio_usd_fob"].iloc[0]) == 50000.0
    glp_jan = df[(df["produto"] == "GLP") & (df["mes"] == "JAN")]
    assert float(glp_jan["volume_barris"].iloc[0]) == 100.0
    assert float(glp_jan["dispendio_usd_fob"].iloc[0]) == 9000.0
    # produtos sem dado no fixture ficam zerados
    gas = df[df["produto"] == "GASOLINA A"]
    assert float(gas["volume_barris"].sum()) == 0.0


def test_gerar_saidas(tmp_path: Path) -> None:
    xlsx = tmp_path / "mini.xlsx"
    _build_mini_xlsx(xlsx)
    out = tmp_path / "out"
    df = carregar_volumes_e_dispendios(xlsx, ano_inicio=2015, ano_fim=2015)
    paths = gerar_saidas(df, out)
    assert paths["xlsx"].exists()
    longo = pd.read_csv(paths["csv_longo"], sep=";", decimal=",", encoding="utf-8-sig")
    assert set(longo["produto"]) == set(PRODUTOS_FOCO)
    assert longo["ano"].min() == 2015
    pivot = pivot_mes_ano(df[df["produto"] == "ÓLEO DIESEL"], "volume_barris")
    assert float(pivot.loc[pivot["Mês"] == "Janeiro", 2015].iloc[0]) == 1000.0
