"""Extração do lucro líquido anual da Petrobras (SEC EDGAR + câmbio BCB)."""

from __future__ import annotations

from collections import defaultdict
from typing import Any

from .fx import fetch_annual_avg_usd_brl
from .sec_client import PETROBRAS_CIK, SecClient

# Conceitos IFRS preferidos (Petrobras é FPI e reporta 20-F em IFRS).
IFRS_NET_INCOME_CONCEPTS = [
    "ProfitLossAttributableToOwnersOfParent",  # lucro atribuível aos acionistas
    "ProfitLoss",
]

US_GAAP_NET_INCOME_CONCEPTS = [
    "NetIncomeLoss",
    "ProfitLoss",
]

ANNUAL_FORMS = {"20-F", "20-F/A", "10-K", "10-K/A"}


def extract_annual_usd(
    facts: dict[str, Any],
    years: int = 10,
) -> list[dict[str, Any]]:
    """
    Extrai série anual de lucro líquido em USD a partir do CompanyFacts.

    Preferência: IFRS ProfitLossAttributableToOwnersOfParent em formulários 20-F FY.
    """
    taxonomies = facts.get("facts", {})

    for taxonomy, concepts in (
        ("ifrs-full", IFRS_NET_INCOME_CONCEPTS),
        ("us-gaap", US_GAAP_NET_INCOME_CONCEPTS),
    ):
        bucket = taxonomies.get(taxonomy, {})
        for concept in concepts:
            if concept not in bucket:
                continue
            units = bucket[concept].get("units", {})
            entries = units.get("USD") or []
            series = _select_annual_fy(entries, concept=concept, taxonomy=taxonomy)
            if series:
                return series[-years:]
    return []


def _select_annual_fy(
    entries: list[dict[str, Any]],
    concept: str,
    taxonomy: str,
) -> list[dict[str, Any]]:
    by_year: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for entry in entries:
        end = entry.get("end")
        if not end:
            continue
        form = entry.get("form")
        fp = entry.get("fp")
        frame = entry.get("frame")

        is_fy_form = form in ANNUAL_FORMS and fp == "FY"
        is_annual_frame = (
            isinstance(frame, str)
            and frame.startswith("CY")
            and "Q" not in frame
            and form in ANNUAL_FORMS | {None, ""}
        )
        if not (is_fy_form or is_annual_frame):
            continue

        # Evita acumulados de períodos intermediários com end != 31/12
        if not end.endswith("-12-31"):
            continue

        by_year[end[:4]].append(entry)

    series: list[dict[str, Any]] = []
    for year in sorted(by_year):
        candidates = by_year[year]
        preferred = [
            e
            for e in candidates
            if e.get("form") in ANNUAL_FORMS and e.get("fp") == "FY"
        ]
        pool = preferred or candidates
        # Mais recente filed prevalece (amendments / restatements)
        pick = sorted(pool, key=lambda e: e.get("filed", ""))[-1]
        series.append(
            {
                "year": int(year),
                "end": pick.get("end"),
                "net_income_usd": float(pick["val"]),
                "filed": pick.get("filed"),
                "form": pick.get("form"),
                "frame": pick.get("frame"),
                "concept": concept,
                "taxonomy": taxonomy,
            }
        )
    return series


def build_table(
    usd_series: list[dict[str, Any]],
    fx_by_year: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """Monta tabela com R$, USD e variação YoY."""
    rows: list[dict[str, Any]] = []
    prev_usd: float | None = None
    prev_brl: float | None = None

    for item in usd_series:
        year = item["year"]
        usd = item["net_income_usd"]
        fx_info = fx_by_year.get(str(year), {})
        fx = float(fx_info["avg_usd_brl"])
        brl = usd * fx

        yoy_usd = _yoy(usd, prev_usd)
        yoy_brl = _yoy(brl, prev_brl)

        rows.append(
            {
                "year": year,
                "net_income_usd": usd,
                "net_income_brl": brl,
                "usd_brl_avg": fx,
                "yoy_usd_pct": yoy_usd,
                "yoy_brl_pct": yoy_brl,
                "form": item.get("form"),
                "filed": item.get("filed"),
                "concept": item.get("concept"),
                "taxonomy": item.get("taxonomy"),
                "fx_source": fx_info.get("source"),
            }
        )
        prev_usd, prev_brl = usd, brl

    return rows


def _yoy(current: float, previous: float | None) -> float | None:
    if previous is None:
        return None
    if previous == 0:
        return None
    return (current - previous) / abs(previous)


def extract_net_income(
    years: int = 10,
    user_agent: str = "SEC-Data-Analysis cafla19791@gmail.com",
    facts_cache: str | None = "data/raw/petrobras_CIK0001119639_companyfacts.json",
    fx_cache: str | None = "data/usdbrl_annual_avg.json",
    refresh: bool = False,
    cik: int = PETROBRAS_CIK,
) -> dict[str, Any]:
    """Pipeline completo: SEC → USD → câmbio BCB → tabela YoY."""
    client = SecClient(user_agent=user_agent)
    facts = client.get_company_facts(
        cik=cik,
        cache_path=facts_cache,
        use_cache=not refresh,
    )
    usd_series = extract_annual_usd(facts, years=years)
    if not usd_series:
        raise RuntimeError("Nenhum lucro líquido anual encontrado nos fatos SEC.")

    year_list = [row["year"] for row in usd_series]
    fx = fetch_annual_avg_usd_brl(
        year_list,
        cache_path=fx_cache,
        use_cache=not refresh,
    )
    table = build_table(usd_series, fx)

    return {
        "ticker": "PBR",
        "cik": int(cik),
        "name": facts.get("entityName", "PETROBRAS - PETROLEO BRASILEIRO SA"),
        "metric": "Net Income / Lucro Líquido atribuível aos acionistas",
        "currency_notes": {
            "usd": "Valores em USD extraídos da SEC EDGAR CompanyFacts (IFRS, 20-F FY).",
            "brl": (
                "Valores em R$ estimados como USD × média anual da taxa de câmbio "
                "USD/BRL (BCB SGS série 1, venda). Podem diferir ligeiramente dos "
                "valores oficiais em reais publicados na CVM/DFP."
            ),
        },
        "years": table,
    }
