"""Unit tests for bcb-mcp providers (mocked)."""

from __future__ import annotations

from datetime import date

import pytest

from bcb_mcp import providers


def test_resolve_series_alias_and_code():
    a = providers.resolve_series("selic")
    assert a["code"] == 11
    assert a["alias"] == "selic"

    b = providers.resolve_series(433)
    assert b["code"] == 433
    assert b["alias"] == "ipca"

    c = providers.resolve_series("ptax")
    assert c["code"] == 1


def test_resolve_series_unknown():
    with pytest.raises(ValueError, match="desconhecida"):
        providers.resolve_series("nao_existe_xyz")


def test_parse_date_formats():
    assert providers._parse_date("2024-01-15") == date(2024, 1, 15)
    assert providers._parse_date("15/01/2024") == date(2024, 1, 15)


def test_date_chunks_splits_long_range():
    chunks = providers._date_chunks(date(2000, 1, 1), date(2025, 1, 1))
    assert len(chunks) >= 2
    assert chunks[0][0] == date(2000, 1, 1)
    assert chunks[-1][1] == date(2025, 1, 1)
    # contiguous
    for i in range(len(chunks) - 1):
        assert chunks[i][1] + __import__("datetime").timedelta(days=1) == chunks[i + 1][0]


def test_normalize_sgs_rows():
    rows = providers._normalize_sgs_rows(
        [
            {"data": "01/01/2024", "valor": "10,5"},
            {"data": "02/01/2024", "valor": "11.0"},
            {"data": "01/01/2024", "valor": "10,5"},
        ]
    )
    assert len(rows) == 2
    assert rows[0]["date"] == "2024-01-01"
    assert rows[0]["value"] == pytest.approx(10.5)


def test_get_sgs_series_last(monkeypatch):
    def fake_json(url, **kwargs):
        assert "ultimos/2" in url
        return [
            {"data": "10/01/2024", "valor": "11.15"},
            {"data": "11/01/2024", "valor": "11.20"},
        ]

    monkeypatch.setattr(providers, "_get_json", fake_json)
    out = providers.get_sgs_series("selic", last=2)
    assert out["code"] == 11
    assert out["count"] == 2
    assert out["series"][-1]["value"] == pytest.approx(11.20)


def test_get_sgs_series_range_chunks(monkeypatch):
    calls: list[str] = []

    def fake_json(url, **kwargs):
        calls.append(url)
        return [{"data": "01/01/2010", "valor": "1"}]

    monkeypatch.setattr(providers, "_get_json", fake_json)
    out = providers.get_sgs_series(
        11,
        date_from="2000-01-01",
        date_to="2020-01-01",
    )
    assert out["chunks"] >= 2
    assert len(calls) == out["chunks"]
    assert out["count"] >= 1


def test_cli_help():
    from bcb_mcp.cli import build_parser

    help_text = build_parser().format_help()
    assert "serie" in help_text
    assert "ptax" in help_text
    assert "catalog" in help_text
