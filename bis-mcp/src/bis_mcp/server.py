"""MCP server: BIS central bank policy rates (WS_CBPOL) for Cursor Cloud Agents."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import providers

HOST = os.getenv("MCP_HOST", "0.0.0.0")
PORT = int(os.getenv("MCP_PORT", "8000"))

mcp = FastMCP(
    "bis-mcp",
    instructions=(
        "Tools for BIS central bank policy rates (WS_CBPOL / "
        "WS_CBPOL_csv_flat). Prefer country aliases (brasil/selic, us/fed, "
        "euro) or ISO codes (BR, US, XM). Monthly frequency by default. "
        "Can also read a local ContAgil WS_CBPOL_csv_flat.csv."
    ),
    host=HOST,
    port=PORT,
    stateless_http=True,
    json_response=True,
)


def _ok(data: dict[str, Any]) -> str:
    return json.dumps(data, indent=2, default=str, ensure_ascii=True)


def _err(exc: Exception) -> str:
    return json.dumps({"error": str(exc)}, indent=2, ensure_ascii=True)


@mcp.tool()
def list_known_areas() -> str:
    """List local aliases for common BIS REF_AREA codes (BR/SELIC, US, euro...)."""
    try:
        return _ok(providers.list_known_areas())
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def get_policy_rates(
    areas: str,
    freq: str = "M",
    date_from: str = "",
    date_to: str = "",
    last: int = 0,
    prefer_local: bool = False,
    csv_path: str = "",
) -> str:
    """Fetch BIS central bank policy rates for one or more countries.

    areas: comma-separated aliases/codes (BR,US,XM or brasil,fed,euro).
    freq: M (monthly, default), D (daily), A (annual).
    date_from / date_to: YYYY-MM or YYYY-MM-DD.
    last: if > 0, return last N points per country.
    prefer_local: read ContAgil WS_CBPOL_csv_flat.csv instead of SDMX.
    """
    try:
        return _ok(
            providers.get_policy_rates(
                areas,
                freq=freq,
                date_from=date_from or None,
                date_to=date_to or None,
                last=last if last and last > 0 else None,
                prefer_local=prefer_local,
                csv_path=csv_path or None,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def compare_latest(
    areas: str = "BR,US,XM,GB,JP",
    freq: str = "M",
    prefer_local: bool = False,
    csv_path: str = "",
) -> str:
    """Latest policy rate snapshot for each country (default BR/US/Euro/UK/JP)."""
    try:
        return _ok(
            providers.compare_latest(
                areas,
                freq=freq,
                prefer_local=prefer_local,
                csv_path=csv_path or None,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def download_flat_csv(dest_dir: str = ".", overwrite: bool = True) -> str:
    """Download WS_CBPOL_csv_flat.zip from BIS and extract the CSV into dest_dir.

    ContAgil tip: dest_dir = the winpython folder so the file lands next to python.exe.
    """
    try:
        return _ok(providers.download_flat_csv(dest_dir, overwrite=overwrite))
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def local_csv_info(csv_path: str = "") -> str:
    """Locate WS_CBPOL_csv_flat.csv (CWD, BIS_CBPOL_CSV, or ContAgil winpython path)."""
    try:
        path = providers.find_local_flat_csv(csv_path or None)
        return _ok(
            {
                "found": path is not None,
                "path": str(path) if path else None,
                "expected_name": providers.FLAT_CSV_NAME,
            }
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


@mcp.tool()
def extract_areas_csv(
    areas: str,
    out_path: str,
    csv_path: str = "",
    freq: str = "M",
) -> str:
    """Extract selected countries from the ~450MB flat CSV into a slim ContAgil-friendly file."""
    try:
        return _ok(
            providers.extract_areas_csv(
                areas,
                out_path,
                csv_path=csv_path or None,
                freq=freq or None,
            )
        )
    except Exception as exc:  # noqa: BLE001
        return _err(exc)


def main() -> None:
    transport = os.getenv("MCP_TRANSPORT", "stdio").strip().lower()
    if transport in {"http", "streamable-http", "streamable_http"}:
        mcp.run(transport="streamable-http")
        return
    if transport == "sse":
        mcp.run(transport="sse")
        return
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
