"""Testes da série consignado / CDC / cartão (2002–2016)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.evolucao_consignado_cdc_cartao import (
    ANO_FIM,
    ANO_INICIO,
    agregar_anual,
    exportar_tabelas,
    fases_historicas,
    gerar_relatorio,
    gerar_tabelas_png,
    linhas_tabela_anual,
)


def _mensal(ano_ini: int, valor_fn) -> pd.DataFrame:
    meses = pd.date_range(f"{ano_ini}-01-01", "2016-12-01", freq="MS")
    return pd.DataFrame({"mes": meses, "valor": [valor_fn(m) for m in meses]})


def _series() -> dict[str, pd.DataFrame]:
    return {
        "consignado": _mensal(2007, lambda m: 50 + (m.year - 2007) * 10 + m.month),
        "cdc": _mensal(2007, lambda m: 80 + (m.year - 2007) * 5 + m.month),
        "veiculos": _mensal(2007, lambda m: 70 + (m.year - 2007) * 4 + m.month),
        "cartao": _mensal(2007, lambda m: 20 + (m.year - 2007) * 8 + m.month),
        "pf_livres": _mensal(2007, lambda m: 200 + (m.year - 2007) * 30 + m.month),
    }


def test_agregar_preenche_2002_2006_e_usa_dezembro():
    anual = agregar_anual(_series())
    assert list(anual["ano"]) == list(range(ANO_INICIO, ANO_FIM + 1))
    assert anual.loc[anual["ano"] == 2002, "consignado"].isna().all()
    dez = anual.loc[anual["ano"] == 2007].iloc[0]
    assert dez["consignado"] == 50 + 12
    assert dez["cdc"] == 80 + 12
    assert abs(dez["share_consignado"] - 100 * dez["consignado"] / dez["pf_livres"]) < 1e-9


def test_relatorio_e_tabelas(tmp_path: Path):
    anual = agregar_anual(_series())
    fases = fases_historicas(anual)
    assert fases[0]["periodo"] == "2002–2006"
    assert fases[0]["cons_ini"] is None
    assert fases[-1]["periodo"] == "2015–2016"
    rel = gerar_relatorio(anual, tmp_path)
    texto = rel.read_text(encoding="utf-8")
    assert "<table" in texto
    assert "border-collapse:collapse" in texto
    assert "20579" in texto
    assert "20583" in texto
    assert "20590" in texto
    saidas = exportar_tabelas(anual, tmp_path)
    assert all(p.exists() for p in saidas)
    linhas = linhas_tabela_anual(anual)
    assert linhas[0][1] == "—"
    assert linhas[-1][0] == "2016"


def test_png(tmp_path: Path):
    anual = agregar_anual(_series())
    paths = gerar_tabelas_png(anual, tmp_path)
    assert all(p.exists() and p.stat().st_size > 0 for p in paths)
