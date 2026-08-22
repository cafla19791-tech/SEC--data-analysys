"""Testes da série anual dos agregados monetários M1–M4."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.evolucao_agregados_monetarios import (
    ANO_FIM,
    ANO_INICIO,
    ANO_QUEBRA,
    agregar_anual,
    exportar_tabelas,
    fases_historicas,
    gerar_relatorio,
    gerar_tabelas_png,
    linhas_tabela_anual,
)


def _mensal(ano_ini: int, ano_fim: int, valor_fn, mes_fim: int = 12) -> pd.DataFrame:
    meses = pd.date_range(f"{ano_ini}-01-01", f"{ano_fim}-{mes_fim:02d}-01", freq="MS")
    return pd.DataFrame({"mes": meses, "valor": [valor_fn(m) for m in meses]})


def _blocos() -> dict[str, dict[str, pd.DataFrame]]:
    def m4(m: pd.Timestamp, base: float) -> float:
        return base + (m.year - 1995) * 10 + m.month

    antiga = {
        "m1": _mensal(1995, 2018, lambda m: 0.15 * m4(m, 100)),
        "m2": _mensal(1995, 2018, lambda m: 0.50 * m4(m, 100)),
        "m3": _mensal(1995, 2018, lambda m: 0.80 * m4(m, 100)),
        "m4": _mensal(1995, 2018, lambda m: m4(m, 100)),
    }
    nova = {
        "m1": _mensal(2002, 2026, lambda m: 0.12 * m4(m, 200), mes_fim=6),
        "m2": _mensal(2001, 2026, lambda m: 0.45 * m4(m, 200), mes_fim=6),
        "m3": _mensal(2001, 2026, lambda m: 0.85 * m4(m, 200), mes_fim=6),
        "m4": _mensal(2001, 2026, lambda m: m4(m, 200), mes_fim=6),
    }
    return {"antiga": antiga, "nova": nova}


def test_agregar_usa_dezembro_e_quebra_em_2001() -> None:
    anual = agregar_anual(_blocos())
    assert anual["ano"].min() == ANO_INICIO
    assert anual["ano"].max() == ANO_FIM
    row_1995 = anual.loc[anual["ano"] == 1995].iloc[0]
    assert row_1995["fonte"] == "antiga"
    assert row_1995["m4"] == 100 + 12
    row_2000 = anual.loc[anual["ano"] == 2000].iloc[0]
    assert row_2000["fonte"] == "antiga"
    row_2001 = anual.loc[anual["ano"] == 2001].iloc[0]
    assert row_2001["fonte"] == "nova"
    # M1 nova só existe a partir de 2002 → fallback na antiga
    assert row_2001["m1_fallback"] == "antiga"
    assert row_2001["m4"] == 200 + (2001 - 1995) * 10 + 12
    row_2026 = anual.loc[anual["ano"] == 2026].iloc[0]
    assert row_2026["fonte"] == "nova"
    assert row_2026["m4"] == 200 + (2026 - 1995) * 10 + 6
    assert abs(row_2026["share_m3"] - 85.0) < 1e-9


def test_fases_relatorio_e_tabelas(tmp_path: Path) -> None:
    anual = agregar_anual(_blocos())
    fases = fases_historicas(anual)
    assert fases[0]["periodo"] == "1995–1998"
    assert fases[-1]["periodo"] == "2022–2026"
    rel = gerar_relatorio(anual, tmp_path, "Jun/2026")
    texto = rel.read_text(encoding="utf-8")
    assert "<table" in texto
    assert "border-collapse:collapse" in texto
    assert "27791" in texto
    assert "1827" in texto
    assert "2026*" in texto
    linhas = linhas_tabela_anual(anual, "Jun/2026")
    assert linhas[0][1] == "antiga"
    assert linhas[-1][0] == "2026*"
    saidas = exportar_tabelas(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in saidas)
    pngs = gerar_tabelas_png(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in pngs)
    assert ANO_QUEBRA == 2001
