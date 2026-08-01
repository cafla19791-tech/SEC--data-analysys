"""Unit tests for bis-mcp providers (mocked + local CSV fixture)."""

from __future__ import annotations

from pathlib import Path

import pytest

from bis_mcp import providers


FIXTURE_CSV = """FREQ,REF_AREA,TIME_PERIOD,OBS_VALUE,TITLE
M,BR,2024-01,11.75,Brazil
M,BR,2024-02,11.25,Brazil
M,BR,2024-03,10.75,Brazil
M,US,2024-01,5.5,United States
M,US,2024-02,5.5,United States
M,US,2024-03,5.5,United States
D,BR,2024-03-01,10.75,Brazil daily
"""

# Same shape as data.bis.org/static/bulk/WS_CBPOL_csv_flat.zip
FLAT_FIXTURE_CSV = """STRUCTURE,STRUCTURE_ID,ACTION,FREQ:Frequency,REF_AREA:Reference area,TIME_PERIOD:Time period or range,OBS_VALUE:Observation Value,TITLE:Title,OBS_STATUS:Observation Status
dataflow,BIS:WS_CBPOL(1.0),I,M: Monthly,BR: Brazil,2024-01,11.75,Brazil,A: Normal value
dataflow,BIS:WS_CBPOL(1.0),I,M: Monthly,BR: Brazil,2024-02,11.25,Brazil,A: Normal value
dataflow,BIS:WS_CBPOL(1.0),I,M: Monthly,BR: Brazil,2024-03,10.75,Brazil,A: Normal value
dataflow,BIS:WS_CBPOL(1.0),I,M: Monthly,US: United States,2024-03,5.5,United States,A: Normal value
"""


def test_resolve_area_aliases():
    a = providers.resolve_area("brasil")
    assert a["code"] == "BR"
    b = providers.resolve_area("selic")
    assert b["code"] == "BR"
    c = providers.resolve_area("euro")
    assert c["code"] == "XM"
    d = providers.resolve_area("US")
    assert d["code"] == "US"


def test_resolve_area_unknown():
    with pytest.raises(ValueError, match="desconhecida"):
        providers.resolve_area("nao_existe_xyz")


def test_normalize_freq():
    assert providers._normalize_freq("mensal") == "M"
    assert providers._normalize_freq("D") == "D"
    with pytest.raises(ValueError):
        providers._normalize_freq("W")


def test_rows_from_csv_text():
    rows = providers._rows_from_csv_text(FIXTURE_CSV)
    assert len(rows) == 7
    assert rows[0]["ref_area"] == "BR"
    assert rows[0]["value"] == pytest.approx(11.75)


def test_read_local_series(tmp_path: Path):
    csv_path = tmp_path / "WS_CBPOL_csv_flat.csv"
    csv_path.write_text(FIXTURE_CSV, encoding="utf-8")
    out = providers.read_local_series(
        "BR,US",
        csv_path=csv_path,
        freq="M",
        last=2,
    )
    assert out["source"] == "local_csv"
    assert out["count"] == 4  # 2 per country
    br = [r for r in out["series"] if r["ref_area"] == "BR"]
    assert br[-1]["time_period"] == "2024-03"
    assert br[-1]["value"] == pytest.approx(10.75)


def test_read_local_series_flat_headers(tmp_path: Path):
    csv_path = tmp_path / "WS_CBPOL_csv_flat.csv"
    csv_path.write_text(FLAT_FIXTURE_CSV, encoding="utf-8")
    out = providers.read_local_series("brasil", csv_path=csv_path, freq="M", last=2)
    assert out["count"] == 2
    assert out["series"][0]["ref_area"] == "BR"
    assert out["series"][-1]["value"] == pytest.approx(10.75)
    assert out["series"][-1]["obs_status"] == "A"


def test_get_policy_rates_sdmx(monkeypatch):
    sample = (
        "FREQ,REF_AREA,TIME_PERIOD,OBS_VALUE\n"
        "M,BR,2025-05,14.75\n"
        "M,BR,2025-06,15\n"
        "M,BR,2025-07,15\n"
    )

    def fake_text(url, **kwargs):
        assert "WS_CBPOL/M.BR" in url
        assert "format=csv" in url
        return sample

    monkeypatch.setattr(providers, "_get_text", fake_text)
    out = providers.get_policy_rates("brasil", last=2)
    assert out["source"] == "sdmx"
    assert out["count"] == 2
    assert out["series"][-1]["value"] == pytest.approx(15.0)


def test_compare_latest(monkeypatch):
    sample = (
        "FREQ,REF_AREA,TIME_PERIOD,OBS_VALUE\n"
        "M,BR,2025-07,15\n"
        "M,US,2025-07,4.5\n"
        "M,XM,2025-07,2.15\n"
    )
    monkeypatch.setattr(providers, "_get_text", lambda url, **kwargs: sample)
    out = providers.compare_latest("BR,US,XM")
    assert len(out["latest"]) == 3
    by_code = {x["code"]: x["value"] for x in out["latest"]}
    assert by_code["BR"] == pytest.approx(15.0)
    assert by_code["US"] == pytest.approx(4.5)


def test_download_flat_csv(tmp_path: Path, monkeypatch):
    import io
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("WS_CBPOL_csv_flat.csv", FIXTURE_CSV)
    blob = buf.getvalue()

    monkeypatch.setattr(providers, "_get_bytes", lambda url, **kwargs: blob)
    out = providers.download_flat_csv(tmp_path)
    path = Path(out["path"])
    assert path.exists()
    assert path.name == "WS_CBPOL_csv_flat.csv"
    assert "11.75" in path.read_text(encoding="utf-8")


def test_extract_areas_csv(tmp_path: Path):
    src = tmp_path / "WS_CBPOL_csv_flat.csv"
    src.write_text(FLAT_FIXTURE_CSV, encoding="utf-8")
    out = tmp_path / "slim.csv"
    result = providers.extract_areas_csv("BR", out, csv_path=src, freq="M")
    assert result["rows"] == 3
    text = out.read_text(encoding="utf-8")
    assert "REF_AREA" in text
    assert "11.75" in text
    assert "US" not in text.splitlines()[1]


def test_cli_help():
    from bis_mcp.cli import build_parser

    help_text = build_parser().format_help()
    assert "serie" in help_text
    assert "compare" in help_text
    assert "download" in help_text
    assert "catalog" in help_text
    assert "extract" in help_text
