"""Testes da agregação anual de recursos livres e direcionados."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.evolucao_recursos_livres_direcionados import (
    ANO_FIM,
    ANO_INICIO,
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


def _series() -> dict[str, pd.DataFrame]:
    def total(m: pd.Timestamp) -> float:
        return 100.0 + (m.year - 2002) * 10 + m.month

    def livres(m: pd.Timestamp) -> float:
        return 0.6 * total(m)

    def direc(m: pd.Timestamp) -> float:
        return 0.4 * total(m)

    def pib(_m: pd.Timestamp) -> float:
        return 30.0

    return {
        "total": _mensal(2002, 2026, total, mes_fim=6),
        "livres": _mensal(2007, 2026, livres, mes_fim=6),
        "direcionados": _mensal(2007, 2026, direc, mes_fim=6),
        "credito_pib": _mensal(2002, 2026, pib, mes_fim=6),
        "livres_pib": _mensal(2007, 2026, lambda m: 18.0, mes_fim=6),
        "direcionados_pib": _mensal(2007, 2026, lambda m: 12.0, mes_fim=6),
    }


def test_agregar_usa_dezembro_exceto_ano_incompleto():
    anual = agregar_anual(_series())
    assert anual["ano"].min() == ANO_INICIO
    assert anual["ano"].max() == ANO_FIM
    dez_2002 = anual.loc[anual["ano"] == 2002].iloc[0]
    assert dez_2002["total"] == 100.0 + 12  # dezembro
    assert pd.isna(dez_2002["livres"])
    jun_2026 = anual.loc[anual["ano"] == 2026].iloc[0]
    assert jun_2026["total"] == 100.0 + (2026 - 2002) * 10 + 6
    assert abs(jun_2026["share_livres"] + jun_2026["share_dir"] - 100.0) < 1e-9


def test_fases_e_relatorio(tmp_path: Path):
    anual = agregar_anual(_series())
    fases = fases_historicas(anual)
    assert fases[0]["periodo"] == "2002–2006"
    assert fases[0]["livres_ini"] is None
    assert fases[-1]["periodo"] == "2022–2026"
    rel = gerar_relatorio(anual, tmp_path, "Jun/2026")
    texto = rel.read_text(encoding="utf-8")
    assert "<table" in texto
    assert "border-collapse:collapse" in texto
    assert "20542" in texto
    assert "20593" in texto
    assert "2026*" in texto
    saidas = exportar_tabelas(anual, tmp_path, "Jun/2026")
    assert all(p.exists() for p in saidas)
    assert saidas[2].suffix == ".xlsx"


def test_tabelas_png(tmp_path: Path):
    anual = agregar_anual(_series())
    paths = gerar_tabelas_png(anual, tmp_path, "Jun/2026")
    assert len(paths) == 2
    assert all(p.exists() and p.stat().st_size > 0 for p in paths)
    linhas = linhas_tabela_anual(anual, "Jun/2026")
    assert linhas[0][1] == "—"  # 2002 sem split
    assert linhas[-1][0] == "2026*"
