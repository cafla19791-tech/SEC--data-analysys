"""Excel: tabela de taxas basicas anuais (% a.a.) por pais e ano.

Usa a serie mensal BIS WS_CBPOL. Para cada pais/ano:
  - preferencia: observacao de dezembro (YYYY-12) = fim de ano;
  - se nao houver dezembro: ultimo mes disponivel daquele ano
    (ex.: 2026 com dados ate junho -> 2026-06).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from . import excel_diario, excel_format, providers

DEFAULT_YEAR_FROM = 2003
DEFAULT_YEAR_TO = 2026


def year_end_rate(
    points: list[tuple[str, float]],
    year: int,
) -> tuple[float | None, str | None]:
    """Retorna (taxa_%_aa, periodo_YYYY-MM) do fim do ano, ou (None, None)."""
    prefix = f"{year:04d}-"
    in_year = [(p, v) for p, v in points if str(p).startswith(prefix)]
    if not in_year:
        return None, None
    # Preferencia dezembro; senao ultimo mes do ano
    dec = f"{year:04d}-12"
    for p, v in in_year:
        if str(p)[:7] == dec:
            return float(v), str(p)[:7]
    p, v = max(in_year, key=lambda x: str(x[0])[:7])
    return float(v), str(p)[:7]


def build_annual_table(
    series: dict[str, list[tuple[str, float]]],
    names: dict[str, str],
    *,
    year_from: int = DEFAULT_YEAR_FROM,
    year_to: int = DEFAULT_YEAR_TO,
) -> tuple[list[dict[str, Any]], list[int]]:
    """Linhas ordenadas por nome do pais; colunas = anos."""
    years = list(range(year_from, year_to + 1))
    rows: list[dict[str, Any]] = []
    for code in series:
        name = names.get(code, excel_diario.AREA_NAMES.get(code, code))
        row: dict[str, Any] = {"Pais": name, "Codigo": code}
        points = series[code]
        for y in years:
            rate, _period = year_end_rate(points, y)
            row[str(y)] = rate
        # Inclui pais se tiver ao menos um ano no intervalo
        if any(row[str(y)] is not None for y in years):
            rows.append(row)
    rows.sort(key=lambda r: (str(r["Pais"]).lower(), str(r["Codigo"])))
    return rows, years


def build_year_ranking(
    table_rows: list[dict[str, Any]],
    year: int,
) -> list[dict[str, Any]]:
    """Ranking do ano em ordem crescente da taxa (menor taxa = 1o)."""
    ykey = str(year)
    entries: list[tuple[float, str, str]] = []
    for r in table_rows:
        rate = r.get(ykey)
        if rate is None:
            continue
        try:
            val = float(rate)
        except (TypeError, ValueError):
            continue
        entries.append((val, str(r["Pais"]), str(r["Codigo"])))

    # Crescente pela taxa; empate: nome do pais, depois codigo
    entries.sort(key=lambda t: (t[0], t[1].lower(), t[2]))
    ranked: list[dict[str, Any]] = []
    for i, (rate, pais, codigo) in enumerate(entries, start=1):
        ranked.append(
            {
                "Posicao": i,
                "Pais": pais,
                "Codigo": codigo,
                "Taxa (% a.a.)": rate,
            }
        )
    return ranked


def ranking_sheet_name(year: int) -> str:
    """Nome de aba Excel (<=31 chars)."""
    return f"R_{year}"[:31]


def gerar_excel_basica_anual(
    out_path: str | Path,
    *,
    csv_path: str | Path | None = None,
    areas: str | None = None,
    year_from: int = DEFAULT_YEAR_FROM,
    year_to: int = DEFAULT_YEAR_TO,
    prefer_local: bool = True,
) -> dict[str, Any]:
    """Gera workbook .xlsx: Pais x anos com taxa basica % a.a. de fim de ano."""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pandas e necessario. pip install pandas openpyxl") from exc

    if year_from > year_to:
        raise ValueError(f"year_from ({year_from}) > year_to ({year_to})")

    wanted: set[str] | None = None
    if areas:
        wanted = {
            providers.resolve_area(a)["code"]
            for a in re.split(r"[\s,;+|]+", areas)
            if a
        }

    local = providers.find_local_flat_csv(csv_path) if prefer_local or csv_path else None
    if local is None and prefer_local:
        dest = Path(out_path).resolve().parent
        providers.download_flat_csv(dest)
        local = providers.find_local_flat_csv(dest / providers.FLAT_CSV_NAME)

    date_from = f"{year_from:04d}-01"
    date_to = f"{year_to:04d}-12"

    if local is not None:
        series, names = excel_diario._load_freq_from_flat(
            local, wanted, date_from, date_to, freq="M"
        )
        source = f"local:{local}"
    else:
        if wanted is None:
            wanted = set(excel_diario.AREA_NAMES.keys())
        series, names = excel_diario._load_freq_from_sdmx(
            sorted(wanted), date_from, date_to, freq="M"
        )
        source = "sdmx"

    if not series:
        raise RuntimeError("Nenhuma serie mensal encontrada para os filtros informados.")

    table_rows, years = build_annual_table(
        series, names, year_from=year_from, year_to=year_to
    )
    if not table_rows:
        raise RuntimeError("Nenhum pais com observacao no intervalo de anos.")

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    cols = ["Pais", "Codigo"] + [str(y) for y in years]
    df = pd.DataFrame(table_rows)[cols]

    # Contagem / indice de rankings por ano
    cobertura = []
    rankings: dict[int, list[dict[str, Any]]] = {}
    for y in years:
        ranked = build_year_ranking(table_rows, y)
        rankings[y] = ranked
        n = len(ranked)
        cobertura.append(
            {
                "Ano": y,
                "Aba_ranking": ranking_sheet_name(y),
                "Paises_com_taxa": n,
                "Total_paises": len(table_rows),
                "Menor_taxa_%": ranked[0]["Taxa (% a.a.)"] if ranked else None,
                "Pais_menor": ranked[0]["Pais"] if ranked else None,
                "Maior_taxa_%": ranked[-1]["Taxa (% a.a.)"] if ranked else None,
                "Pais_maior": ranked[-1]["Pais"] if ranked else None,
            }
        )

    legenda = pd.DataFrame(
        [
            {"Item": "Fonte", "Valor": "BIS WS_CBPOL (frequencia mensal)"},
            {"Item": "Origem dados", "Valor": source},
            {
                "Item": "Conceito",
                "Valor": "Taxa basica de juros / policy rate em % a.a. (OBS_VALUE)",
            },
            {
                "Item": "Regra por ano",
                "Valor": (
                    "Taxa de dezembro (YYYY-12); se ausente, ultimo mes "
                    "disponivel daquele ano (ex.: ano corrente incompleto)"
                ),
            },
            {
                "Item": "Ranking",
                "Valor": (
                    "Abas R_AAAA: Posicao | Pais | Codigo | Taxa (% a.a.), "
                    "ordem crescente da taxa (1 = menor taxa)"
                ),
            },
            {
                "Item": "Periodo",
                "Valor": f"{year_from} a {year_to}",
            },
            {
                "Item": "Paises",
                "Valor": str(len(table_rows)),
            },
            {
                "Item": "Tabela",
                "Valor": "Aba 02_Taxas_basicas_anuais: Pais | Codigo | anos",
            },
        ]
    )

    engine = "xlsxwriter"
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        engine = "openpyxl"

    with pd.ExcelWriter(out, engine=engine) as writer:
        legenda.to_excel(writer, sheet_name="00_Legenda", index=False)
        excel_format.autosize_dataframe_sheet(
            writer,
            "00_Legenda",
            legenda,
            engine=engine,
            max_width=90,
            padding=4,
            center=True,
            print_layout=True,
        )

        cov_df = pd.DataFrame(cobertura)
        cov_df.to_excel(writer, sheet_name="01_Cobertura", index=False)
        excel_format.autosize_dataframe_sheet(
            writer,
            "01_Cobertura",
            cov_df,
            engine=engine,
            padding=4,
            center=True,
            print_layout=True,
        )

        df.to_excel(writer, sheet_name="02_Taxas_basicas_anuais", index=False)
        formats = None
        if engine == "xlsxwriter":
            # Pais/Codigo centrados; anos com ate 4 casas (ex.: 3.625)
            nfmts: list[str | None] = [None, None] + ["0.####"] * len(years)
            formats = excel_format.make_center_formats(writer.book, nfmts)
        excel_format.autosize_dataframe_sheet(
            writer,
            "02_Taxas_basicas_anuais",
            df,
            engine=engine,
            col_formats=formats,
            min_width=8,
            max_width=28,
            padding=3,
            center=True,
            print_layout=True,
            page_header_left="Taxas basicas anuais (% a.a.)",
        )

        rank_formats = None
        if engine == "xlsxwriter":
            rank_formats = excel_format.make_center_formats(
                writer.book, [None, None, None, "0.####"]
            )
        for y in years:
            ranked = rankings[y]
            if not ranked:
                continue
            aba = ranking_sheet_name(y)
            rdf = pd.DataFrame(ranked)[
                ["Posicao", "Pais", "Codigo", "Taxa (% a.a.)"]
            ]
            rdf.to_excel(writer, sheet_name=aba, index=False)
            excel_format.autosize_dataframe_sheet(
                writer,
                aba,
                rdf,
                engine=engine,
                col_formats=rank_formats,
                min_width=10,
                max_width=36,
                padding=3,
                center=True,
                print_layout=True,
                page_header_left=f"Ranking {y} (crescente)",
            )

    return {
        "path": str(out.resolve()),
        "source": source,
        "countries": len(table_rows),
        "years": years,
        "year_from": year_from,
        "year_to": year_to,
        "rankings": len([y for y in years if rankings[y]]),
        "bytes": out.stat().st_size,
        "rule": "december_or_last_month_of_year",
        "ranking": "ascending_rate",
    }
