"""Testes dos fatores condicionantes da base monetária."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.evolucao_fatores_base_monetaria import (
    ANO_FIM,
    ANO_INICIO,
    FATORES,
    agregar_anual,
    exportar_tabelas,
    gerar_relatorio,
    gerar_tabelas_png,
    linhas_var,
)


def _mensal(ano_ini: int, ano_fim: int, valor_fn, mes_fim: int = 12) -> pd.DataFrame:
    meses = pd.date_range(f"{ano_ini}-01-01", f"{ano_fim}-{mes_fim:02d}-01", freq="MS")
    return pd.DataFrame({"mes": meses, "valor": [valor_fn(m) for m in meses]})


def _series() -> dict[str, pd.DataFrame]:
    def milhares(base_bi: float, m: pd.Timestamp) -> float:
        return (base_bi + (m.year - 1995) * 10 + m.month) * 1_000_000.0

    out = {
        "base": _mensal(1995, 2026, lambda m: milhares(20, m), mes_fim=6),
        "tesouro": _mensal(1995, 2026, lambda m: milhares(8, m), mes_fim=6),
        "titulos": _mensal(1995, 2026, lambda m: milhares(5, m), mes_fim=6),
        "externo": _mensal(1995, 2026, lambda m: milhares(4, m), mes_fim=6),
        "depositos_if": _mensal(1995, 2026, lambda m: milhares(1, m), mes_fim=6),
        "outras": _mensal(1995, 2026, lambda m: milhares(2, m), mes_fim=6),
        "redesconto": _mensal(1995, 2024, lambda m: 0.0),
        "derivativos": _mensal(2002, 2026, lambda m: milhares(0.5, m), mes_fim=6),
        "linhas_temp": _mensal(2020, 2026, lambda m: milhares(0.2, m), mes_fim=6),
        "titulos_primario": _mensal(2010, 2026, lambda m: milhares(3, m), mes_fim=6),
        "titulos_secundario": _mensal(2010, 2026, lambda m: milhares(2, m), mes_fim=6),
    }
    return out


def test_agregar_usa_dezembro_e_varia_no_ultimo_mes() -> None:
    anual = agregar_anual(_series())
    assert anual["ano"].min() == ANO_INICIO
    assert anual["ano"].max() == ANO_FIM
    row_1995 = anual.loc[anual["ano"] == 1995].iloc[0]
    assert abs(row_1995["base"] - (20 + 12)) < 1e-6
    row_2026 = anual.loc[anual["ano"] == 2026].iloc[0]
    assert abs(row_2026["base"] - (20 + (2026 - 1995) * 10 + 6)) < 1e-6
    row_1996 = anual.loc[anual["ano"] == 1996].iloc[0]
    assert abs(row_1996["d_base"] - (row_1996["base"] - row_1995["base"])) < 1e-9
    assert abs(row_1996["residuo_var"] - (row_1996["d_base"] - row_1996["d_soma"])) < 1e-9
    assert set(FATORES)


def test_relatorio_e_tabelas(tmp_path: Path) -> None:
    anual = agregar_anual(_series())
    rel = gerar_relatorio(anual, tmp_path, "Jun/2026")
    texto = rel.read_text(encoding="utf-8")
    assert "<table" in texto
    assert "border-collapse:collapse" in texto
    assert "1810" in texto
    assert "2026*" in texto
    linhas = linhas_var(anual, "Jun/2026")
    assert linhas[-1][0] == "2026*"
    saidas = exportar_tabelas(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in saidas)
    pngs = gerar_tabelas_png(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in pngs)
