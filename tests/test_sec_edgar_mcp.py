"""Unit tests for sec-edgar-mcp providers (mocked)."""

from __future__ import annotations

import pytest

from sec_edgar_mcp import providers


def test_pad_cik():
    assert providers.pad_cik(320193) == "0000320193"
    assert providers.pad_cik("320193") == "0000320193"
    with pytest.raises(ValueError):
        providers.pad_cik("ABC")


def test_lookup_ticker(monkeypatch):
    providers._ticker_map.cache_clear()

    def fake_map():
        return {
            "AAPL": {"cik": "0000320193", "ticker": "AAPL", "title": "Apple Inc."}
        }

    monkeypatch.setattr(providers, "_ticker_map", fake_map)
    out = providers.lookup_ticker("aapl")
    assert out["found"] is True
    assert out["cik"] == "0000320193"


def test_list_filings_filters(monkeypatch):
    fake = {
        "name": "Apple Inc.",
        "tickers": ["AAPL"],
        "filings": {
            "recent": {
                "form": ["10-K", "10-Q", "8-K"],
                "filingDate": ["2024-11-01", "2024-08-01", "2024-07-01"],
                "accessionNumber": [
                    "0000320193-24-000001",
                    "0000320193-24-000002",
                    "0000320193-24-000003",
                ],
                "primaryDocument": ["aapl-20240928.htm", "aapl-10q.htm", "aapl-8k.htm"],
                "primaryDocDescription": ["10-K", "10-Q", "8-K"],
            }
        },
    }

    def fake_submissions(_):
        return {
            "cik": "0000320193",
            "name": "Apple Inc.",
            "tickers": ["AAPL"],
            "_data": fake,
        }

    monkeypatch.setattr(providers, "get_submissions", fake_submissions)
    out = providers.list_filings("AAPL", form="10-K", limit=5)
    assert out["count"] == 1
    assert out["filings"][0]["form"] == "10-K"
    assert "Archives/edgar/data" in out["filings"][0]["documentUrl"]


def test_cli_help():
    from sec_edgar_mcp.cli import build_parser

    p = build_parser()
    assert "lookup" in p.format_help()
