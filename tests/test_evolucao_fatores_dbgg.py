"""Testes dos fatores da DBGG."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.evolucao_fatores_dbgg import (
    ANO_FIM,
    ANO_INICIO,
    agregar_anual,
    exportar_tabelas,
    gerar_relatorio,
    gerar_tabelas_png,
    linhas_fatores,
)


def _mensal(ano_ini: int, ano_fim: int, valor_fn, mes_fim: int = 12) -> pd.DataFrame:
    meses = pd.date_range(f"{ano_ini}-01-01", f"{ano_fim}-{mes_fim:02d}-01", freq="MS")
    return pd.DataFrame({"mes": meses, "valor": [valor_fn(m) for m in meses]})


def _vazio() -> pd.DataFrame:
    return pd.DataFrame(columns=["mes", "valor"])


def _series() -> dict[str, pd.DataFrame]:
    # DBGG 50% PIB em 2001, sobe 1 p.p. ao ano; estoque = % * PIB
    def pib(m: pd.Timestamp) -> float:
        return 1_000.0 * (1.08 ** (m.year - 2001))

    def dbgg_pib(m: pd.Timestamp) -> float:
        return 50.0 + (m.year - 2001)

    return {
        "dbgg_rs": _mensal(2001, 2026, lambda m: dbgg_pib(m) / 100.0 * pib(m) * 1000, mes_fim=6),
        "dbgg_pib": _mensal(2001, 2026, dbgg_pib, mes_fim=6),
        "dlsp_rs": _mensal(2001, 2026, lambda m: 0.6 * dbgg_pib(m) / 100.0 * pib(m) * 1000, mes_fim=6),
        "dlsp_pib": _mensal(2001, 2026, lambda m: 0.6 * dbgg_pib(m), mes_fim=6),
        "primario_cons": _mensal(2002, 2026, lambda m: -1.0 if m.year < 2015 else 1.5, mes_fim=6),
        "primario_gfbc": _mensal(1995, 2026, lambda m: -2.0, mes_fim=6),
        "juros_cons": _mensal(2002, 2026, lambda m: 6.0, mes_fim=6),
        "nfsp_cons": _mensal(2002, 2026, lambda m: 5.0 if m.year < 2015 else 7.5, mes_fim=6),
        "nfsp_gfbc": _mensal(1995, 2026, lambda m: 8.0, mes_fim=6),
        "ajuste_priv": _mensal(2001, 2026, lambda m: -10_000 - (m.year - 2001) * 500, mes_fim=6),
        "ajuste_patrim": _mensal(2001, 2026, lambda m: 5_000 + (m.year - 2001) * 200, mes_fim=6),
        "ajuste_camb_ext": _mensal(2001, 2026, lambda m: (m.year - 2001) * -1_000, mes_fim=6),
        "ajuste_met_int": _mensal(2001, 2026, lambda m: 100.0, mes_fim=6),
        "selic": _mensal(1995, 2026, lambda m: 40.0 if m.year < 2000 else 13.0, mes_fim=6),
        "usd": _mensal(1995, 2026, lambda m: 1.0 if m.year < 1999 else 5.0, mes_fim=6),
        "ext_gg": _mensal(2013, 2026, lambda m: 200_000.0, mes_fim=6),
    }


def test_agregar_preenche_1995_com_drivers_e_dbgg_em_2001() -> None:
    anual = agregar_anual(_series())
    assert anual["ano"].min() == ANO_INICIO
    assert anual["ano"].max() == ANO_FIM
    row_1995 = anual.loc[anual["ano"] == 1995].iloc[0]
    assert pd.isna(row_1995["dbgg_pib"])
    assert row_1995["primario"] == -2.0
    assert row_1995["nfsp"] == 8.0
    assert row_1995["juros"] == 10.0  # NFSP − primário
    row_2002 = anual.loc[anual["ano"] == 2002].iloc[0]
    assert row_2002["primario"] == -1.0
    assert row_2002["juros"] == 6.0
    row_2001 = anual.loc[anual["ano"] == 2001].iloc[0]
    assert row_2001["dbgg_pib"] == 50.0
    row_2026 = anual.loc[anual["ano"] == 2026].iloc[0]
    assert row_2026["dbgg_pib"] == 50.0 + (2026 - 2001)
    assert pd.notna(row_2026["efeito_pib"])
    assert row_2026["efeito_pib"] < 0  # PIB cresce → dilui a razão


def test_relatorio_e_tabelas(tmp_path: Path) -> None:
    anual = agregar_anual(_series())
    rel = gerar_relatorio(anual, tmp_path, "Jun/2026")
    texto = rel.read_text(encoding="utf-8")
    assert "<table" in texto
    assert "border-collapse:collapse" in texto
    assert "4502" in texto
    assert "10822" in texto
    assert "2026*" in texto
    linhas = linhas_fatores(anual, "Jun/2026")
    assert linhas[-1][0] == "2026*"
    saidas = exportar_tabelas(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in saidas)
    pngs = gerar_tabelas_png(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in pngs)
