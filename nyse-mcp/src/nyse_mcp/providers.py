"""Market-data providers (Yahoo Finance by default, Alpha Vantage optional)."""

from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Any

import httpx
import yfinance as yf


class ProviderError(RuntimeError):
    """Raised when a market-data provider fails."""


def _provider_name() -> str:
    return os.getenv("MARKET_DATA_PROVIDER", "yahoo").strip().lower()


def _alpha_vantage_key() -> str:
    key = os.getenv("ALPHA_VANTAGE_API_KEY", "").strip()
    if not key:
        raise ProviderError(
            "ALPHA_VANTAGE_API_KEY is required when MARKET_DATA_PROVIDER=alphavantage. "
            "Get a free key at https://www.alphavantage.co/support/#api-key"
        )
    return key


def get_quote(symbol: str) -> dict[str, Any]:
    symbol = symbol.upper().strip()
    if _provider_name() == "alphavantage":
        return _av_quote(symbol)
    return _yahoo_quote(symbol)


def get_history(
    symbol: str,
    period: str = "1mo",
    interval: str = "1d",
) -> dict[str, Any]:
    symbol = symbol.upper().strip()
    if _provider_name() == "alphavantage":
        return _av_history(symbol, interval=interval)
    return _yahoo_history(symbol, period=period, interval=interval)


def get_fundamentals(symbol: str) -> dict[str, Any]:
    symbol = symbol.upper().strip()
    if _provider_name() == "alphavantage":
        return _av_overview(symbol)
    return _yahoo_fundamentals(symbol)


def search_ticker(query: str) -> dict[str, Any]:
    query = query.strip()
    if not query:
        raise ProviderError("query must not be empty")

    # Yahoo search works without an API key and is good enough for a skeleton.
    url = "https://query2.finance.yahoo.com/v1/finance/search"
    params = {"q": query, "quotesCount": 8, "newsCount": 0}
    headers = {"User-Agent": "nyse-mcp/0.1"}

    with httpx.Client(timeout=20.0) as client:
        response = client.get(url, params=params, headers=headers)
        response.raise_for_status()
        payload = response.json()

    results = []
    for item in payload.get("quotes", []):
        results.append(
            {
                "symbol": item.get("symbol"),
                "shortname": item.get("shortname") or item.get("longname"),
                "exchange": item.get("exchange"),
                "quoteType": item.get("quoteType"),
                "typeDisp": item.get("typeDisp"),
            }
        )
    return {"query": query, "results": results, "provider": "yahoo"}


def market_status() -> dict[str, Any]:
    """Rough US equity session status using Yahoo market state for ^GSPC."""
    ticker = yf.Ticker("^GSPC")
    state = None

    # yfinance versions expose market state in different shapes.
    try:
        fast = ticker.fast_info
        state = getattr(fast, "market_state", None)
        if state is None and hasattr(fast, "get"):
            state = fast.get("marketState") or fast.get("market_state")
    except Exception:  # noqa: BLE001
        state = None

    if not state:
        try:
            state = (ticker.info or {}).get("marketState")
        except Exception:  # noqa: BLE001
            state = None

    now = datetime.now(timezone.utc)
    # Fallback heuristic for US regular session (14:30–21:00 UTC, Mon–Fri).
    if not state:
        weekday = now.weekday()
        minutes = now.hour * 60 + now.minute
        if weekday >= 5:
            state = "CLOSED"
        elif 14 * 60 + 30 <= minutes < 21 * 60:
            state = "REGULAR_HEURISTIC"
        else:
            state = "CLOSED_HEURISTIC"

    return {
        "benchmark": "^GSPC",
        "market_state": state,
        "checked_at_utc": now.isoformat(),
        "note": "Indicative status from Yahoo Finance / UTC heuristic; not an official NYSE feed.",
        "provider": "yahoo",
    }


def _yahoo_quote(symbol: str) -> dict[str, Any]:
    ticker = yf.Ticker(symbol)
    info = ticker.fast_info
    hist = ticker.history(period="5d", interval="1d")
    last_close = float(hist["Close"].iloc[-1]) if not hist.empty else None
    prev_close = float(hist["Close"].iloc[-2]) if len(hist) > 1 else None
    change = None
    change_pct = None
    if last_close is not None and prev_close not in (None, 0):
        change = last_close - prev_close
        change_pct = (change / prev_close) * 100

    return {
        "symbol": symbol,
        "price": _safe_float(getattr(info, "last_price", None) or info.get("lastPrice"))
        or last_close,
        "currency": getattr(info, "currency", None) or info.get("currency"),
        "previous_close": prev_close,
        "change": change,
        "change_percent": change_pct,
        "provider": "yahoo",
        "delayed": True,
    }


def _yahoo_history(symbol: str, period: str, interval: str) -> dict[str, Any]:
    ticker = yf.Ticker(symbol)
    hist = ticker.history(period=period, interval=interval)
    if hist.empty:
        raise ProviderError(f"No history found for {symbol}")

    rows = []
    for idx, row in hist.iterrows():
        rows.append(
            {
                "date": idx.isoformat(),
                "open": _safe_float(row.get("Open")),
                "high": _safe_float(row.get("High")),
                "low": _safe_float(row.get("Low")),
                "close": _safe_float(row.get("Close")),
                "volume": _safe_int(row.get("Volume")),
            }
        )
    return {
        "symbol": symbol,
        "period": period,
        "interval": interval,
        "bars": rows,
        "provider": "yahoo",
    }


def _yahoo_fundamentals(symbol: str) -> dict[str, Any]:
    ticker = yf.Ticker(symbol)
    info = ticker.info or {}
    return {
        "symbol": symbol,
        "name": info.get("longName") or info.get("shortName"),
        "exchange": info.get("exchange"),
        "sector": info.get("sector"),
        "industry": info.get("industry"),
        "market_cap": info.get("marketCap"),
        "trailing_pe": info.get("trailingPE"),
        "forward_pe": info.get("forwardPE"),
        "dividend_yield": info.get("dividendYield"),
        "beta": info.get("beta"),
        "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
        "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
        "provider": "yahoo",
    }


def _av_quote(symbol: str) -> dict[str, Any]:
    data = _av_get(
        {
            "function": "GLOBAL_QUOTE",
            "symbol": symbol,
        }
    )
    quote = data.get("Global Quote") or {}
    if not quote:
        raise ProviderError(f"Alpha Vantage returned no quote for {symbol}: {data}")
    price = _safe_float(quote.get("05. price"))
    prev = _safe_float(quote.get("08. previous close"))
    change = _safe_float(quote.get("09. change"))
    change_pct_raw = (quote.get("10. change percent") or "").replace("%", "")
    return {
        "symbol": symbol,
        "price": price,
        "previous_close": prev,
        "change": change,
        "change_percent": _safe_float(change_pct_raw),
        "provider": "alphavantage",
        "delayed": True,
    }


def _av_history(symbol: str, interval: str) -> dict[str, Any]:
    # Free Alpha Vantage daily endpoint is the simplest reliable skeleton path.
    function = "TIME_SERIES_DAILY"
    if interval in {"1m", "5m", "15m", "30m", "60m"}:
        function = "TIME_SERIES_INTRADAY"

    params: dict[str, str] = {"function": function, "symbol": symbol, "outputsize": "compact"}
    if function == "TIME_SERIES_INTRADAY":
        params["interval"] = "5min" if interval == "1m" else interval

    data = _av_get(params)
    series_key = next((k for k in data if "Time Series" in k), None)
    if not series_key:
        raise ProviderError(f"Alpha Vantage returned no series for {symbol}: {data}")

    rows = []
    for ts, bar in list(data[series_key].items())[:100]:
        rows.append(
            {
                "date": ts,
                "open": _safe_float(bar.get("1. open")),
                "high": _safe_float(bar.get("2. high")),
                "low": _safe_float(bar.get("3. low")),
                "close": _safe_float(bar.get("4. close")),
                "volume": _safe_int(bar.get("5. volume")),
            }
        )
    rows.reverse()
    return {
        "symbol": symbol,
        "interval": interval,
        "bars": rows,
        "provider": "alphavantage",
    }


def _av_overview(symbol: str) -> dict[str, Any]:
    data = _av_get({"function": "OVERVIEW", "symbol": symbol})
    if not data or "Symbol" not in data:
        raise ProviderError(f"Alpha Vantage returned no overview for {symbol}: {data}")
    return {
        "symbol": symbol,
        "name": data.get("Name"),
        "exchange": data.get("Exchange"),
        "sector": data.get("Sector"),
        "industry": data.get("Industry"),
        "market_cap": _safe_int(data.get("MarketCapitalization")),
        "trailing_pe": _safe_float(data.get("PERatio")),
        "forward_pe": _safe_float(data.get("ForwardPE")),
        "dividend_yield": _safe_float(data.get("DividendYield")),
        "beta": _safe_float(data.get("Beta")),
        "fifty_two_week_high": _safe_float(data.get("52WeekHigh")),
        "fifty_two_week_low": _safe_float(data.get("52WeekLow")),
        "provider": "alphavantage",
    }


def _av_get(params: dict[str, str]) -> dict[str, Any]:
    query = {"apikey": _alpha_vantage_key(), **params}
    with httpx.Client(timeout=30.0) as client:
        response = client.get("https://www.alphavantage.co/query", params=query)
        response.raise_for_status()
        data = response.json()
    if "Note" in data or "Information" in data:
        raise ProviderError(f"Alpha Vantage rate limit / info: {data}")
    if "Error Message" in data:
        raise ProviderError(data["Error Message"])
    return data


def _safe_float(value: Any) -> float | None:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_int(value: Any) -> int | None:
    try:
        if value is None or value == "":
            return None
        return int(float(value))
    except (TypeError, ValueError):
        return None
