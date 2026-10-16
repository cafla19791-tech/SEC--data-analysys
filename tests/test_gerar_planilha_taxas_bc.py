"""Testes do gerador de planilha Meta Selic + taxas BIS por país."""

from __future__ import annotations

import json
from pathlib import Path

from scripts.gerar_planilha_taxas_bc import build_meta_selic, sheet_name


def test_sheet_name_limits_and_sanitizes() -> None:
    assert sheet_name("United States") == "United_States"
    assert sheet_name("Türkiye") == "Turkiye"
    assert len(sheet_name("A" * 50)) == 31
    assert "?" not in sheet_name("Foo?Bar")


def test_build_meta_selic_from_cache() -> None:
    cache = Path("data/selic/historico_copom.json")
    assert cache.exists(), "cache Copom esperado em data/selic/"
    rows = json.loads(cache.read_text(encoding="utf-8"))["conteudo"]
    df = build_meta_selic(rows)
    assert len(df) >= 200
    assert df["Inicio_Vigencia"].min().isoformat() == "1996-07-01"
    assert float(df.iloc[-1]["Meta_Selic"]) > 0
    assert set(["Meta_Selic", "Inicio_Vigencia", "Fim_Vigencia"]).issubset(df.columns)
