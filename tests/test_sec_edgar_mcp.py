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
    help_text = p.format_help()
    assert "lookup" in help_text
    assert "debt" in help_text
    assert "filing-xbrl" in help_text


def test_parse_xbrl_instant_facts_plain_contexts():
    xml = """<?xml version="1.0"?>
    <xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance"
          xmlns:ifrs-full="http://xbrl.ifrs.org/taxonomy/2024-03-27/ifrs-full">
      <xbrli:context id="AsOf2024-12-31">
        <xbrli:entity><xbrli:identifier scheme="http://www.sec.gov/CIK">0001119639</xbrli:identifier></xbrli:entity>
        <xbrli:period><xbrli:instant>2024-12-31</xbrli:instant></xbrli:period>
      </xbrli:context>
      <xbrli:context id="AsOf2025-12-31">
        <xbrli:entity><xbrli:identifier scheme="http://www.sec.gov/CIK">0001119639</xbrli:identifier></xbrli:entity>
        <xbrli:period><xbrli:instant>2025-12-31</xbrli:instant></xbrli:period>
      </xbrli:context>
      <xbrli:context id="AsOf2025-12-31_dim">
        <xbrli:entity><xbrli:identifier scheme="http://www.sec.gov/CIK">0001119639</xbrli:identifier></xbrli:entity>
        <xbrli:period><xbrli:instant>2025-12-31</xbrli:instant></xbrli:period>
        <xbrli:scenario>
          <xbrldi:explicitMember xmlns:xbrldi="http://xbrl.org/2006/xbrldi"
            dimension="ifrs-full:MaturityAxis">ifrs-full:NotLaterThanOneYearMember</xbrldi:explicitMember>
        </xbrli:scenario>
      </xbrli:context>
      <ifrs-full:Borrowings contextRef="AsOf2024-12-31" unitRef="USD">23162000000</ifrs-full:Borrowings>
      <ifrs-full:Borrowings contextRef="AsOf2025-12-31" unitRef="USD">26441000000</ifrs-full:Borrowings>
      <ifrs-full:Borrowings contextRef="AsOf2025-12-31_dim" unitRef="USD">2186000000</ifrs-full:Borrowings>
      <ifrs-full:LeaseLiabilities contextRef="AsOf2025-12-31" unitRef="USD">43352000000</ifrs-full:LeaseLiabilities>
    </xbrl>
    """
    facts = providers.parse_xbrl_instant_facts(
        xml, ["Borrowings", "LeaseLiabilities"]
    )
    assert facts["Borrowings"]["2024-12-31"] == 23162000000
    assert facts["Borrowings"]["2025-12-31"] == 26441000000
    assert facts["LeaseLiabilities"]["2025-12-31"] == 43352000000


def test_get_total_debt_sums_and_fills(monkeypatch):
    def fake_concept(ticker, concept, taxonomy="auto", limit=20, annual_only=False):
        if concept == "Borrowings":
            return {
                "entityName": "PETROBRAS",
                "concept": "Borrowings",
                "taxonomy": "ifrs-full",
                "unit": "USD",
                "recent": [
                    {
                        "end": "2024-12-31",
                        "val": 23162000000,
                        "fp": "FY",
                        "form": "20-F",
                        "frame": "CY2024",
                    }
                ],
            }
        if concept == "LeaseLiabilities":
            return {
                "entityName": "PETROBRAS",
                "concept": "LeaseLiabilities",
                "taxonomy": "ifrs-full",
                "unit": "USD",
                "recent": [
                    {
                        "end": "2024-12-31",
                        "val": 37149000000,
                        "fp": "FY",
                        "form": "20-F",
                        "frame": "CY2024",
                    }
                ],
            }
        raise ValueError(concept)

    def fake_extract(ticker, concepts, form=None, accession=None):
        return {
            "accessionNumber": "0001292814-26-002168",
            "form": "20-F",
            "filingDate": "2026-04-09",
            "instanceUrl": "https://example/x.xml",
            "concepts": {
                "Borrowings": [
                    {"end": "2025-12-31", "year": 2025, "val": 26441000000}
                ],
                "LeaseLiabilities": [
                    {"end": "2025-12-31", "year": 2025, "val": 43352000000}
                ],
            },
        }

    monkeypatch.setattr(providers, "resolve_cik", lambda _: "0001119639")
    monkeypatch.setattr(providers, "get_concept", fake_concept)
    monkeypatch.setattr(providers, "extract_filing_concepts", fake_extract)

    out = providers.get_total_debt("PBR", year_from=2024, year_to=2025)
    assert out["count"] == 2
    by_year = {r["year"]: r for r in out["series"]}
    assert by_year[2024]["total_debt"] == pytest.approx(23162000000 + 37149000000)
    assert by_year[2025]["total_debt"] == pytest.approx(26441000000 + 43352000000)
    assert by_year[2025]["sources"]["borrowings"] == "filing-ixbrl"
    assert out["filing_fill"]["used"] is True
