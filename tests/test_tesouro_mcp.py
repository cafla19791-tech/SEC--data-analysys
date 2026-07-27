"""Unit tests for tesouro-mcp providers (mocked)."""

from __future__ import annotations

import pytest

from tesouro_mcp import providers


def test_parse_month_formats():
    assert providers._parse_month("01/2024") == "01/2024"
    assert providers._parse_month("2024-01") == "01/2024"
    assert providers._parse_month("2024-01-15") == "01/2024"


def test_resolve_serie_alias_and_code():
    a = providers.resolve_serie("resultado_primario")
    assert a["codigo_serie"] == "10.04.1"
    assert a["tema"] == "10"

    b = providers.resolve_serie("10.01.1")
    assert b["codigo_serie"] == "10.01.1"
    assert b["tema"] == "10"


def test_resolve_serie_unknown():
    with pytest.raises(ValueError, match="desconhecida"):
        providers.resolve_serie("nao_existe_xyz")


def test_normalize_row():
    row = providers._normalize_row(
        {
            "data": "2024-03-01T00:00:00.000Z",
            "valor": 199226.89,
            "codigoTema": "10",
            "nomeSubtema": "Receitas",
            "codigoSerie": "10.01.1",
            "nomeSerie": "Receita Total",
        }
    )
    assert row["date"] == "2024-03-01"
    assert row["value"] == pytest.approx(199226.89)
    assert row["unit"] == "R$ milhoes"


def test_get_serie_mocked(monkeypatch):
    def fake_resultado_fiscal(**kwargs):
        assert kwargs["tema"] == "10"
        assert kwargs["codigo_serie"] == "10.04.1"
        return {
            "tema": "10",
            "count": 1,
            "series": [{"date": "2024-01-01", "value": -10.0}],
        }

    monkeypatch.setattr(providers, "get_resultado_fiscal", fake_resultado_fiscal)
    out = providers.get_serie("resultado_primario", data_inicio="2024-01", data_fim="2024-01")
    assert out["alias"] == "resultado_primario"
    assert out["count"] == 1


def test_paginate_follows_next(monkeypatch):
    calls = []

    def fake_json(url, **kwargs):
        calls.append(url)
        if "page=2" in url:
            return {"registros": [{"codigoSerie": "b"}], "next": None, "status": "ok"}
        return {
            "registros": [{"codigoSerie": "a"}],
            "next": "https://apiapex.tesouro.gov.br/aria//v1/series-temporais/custom/series?page=2",
            "status": "ok",
        }

    monkeypatch.setattr(providers, "_get_json", fake_json)
    rows = providers._paginate("https://example/series")
    assert len(rows) == 2
    assert "aria/v1" in calls[1]
    assert "aria//v1" not in calls[1]


def test_cli_help():
    from tesouro_mcp.cli import build_parser

    help_text = build_parser().format_help()
    assert "serie" in help_text
    assert "headline" in help_text
    assert "ckan-show" in help_text
