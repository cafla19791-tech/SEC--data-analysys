"""Testes dos saldos/variações anuais dos fatores condicionantes."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.fatores_condicionantes_base_monetaria import (
    SERIES,
    agregados_anuais,
    baixar_sgs,
    formatar_milhoes,
    gravar_saidas,
    markdown_tabela,
    montar_tabelas,
    ultimo_dia_periodo,
)


def _serie(codigo: int, datas: list[str], valores: list[float]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "codigo": codigo,
            "data": pd.to_datetime(datas),
            "valor": valores,
        }
    )


def test_ultimo_dia_periodo_converte_dia_1_em_31_12():
    assert ultimo_dia_periodo(pd.Timestamp("1995-12-01")) == pd.Timestamp("1995-12-31")
    assert ultimo_dia_periodo(pd.Timestamp("2026-06-01")) == pd.Timestamp("2026-06-30")


def test_agregados_dezembro_e_soma_do_ano():
    df = _serie(
        1810,
        ["1995-01-01", "1995-12-01", "1996-06-01", "1996-12-01"],
        [10.0, 20.0, 5.0, 7.0],
    )
    out = agregados_anuais(df, [1995, 1996])
    assert list(out["ano"]) == [1995, 1996]
    assert list(out["valor_dezembro_rs_mil"]) == [20.0, 7.0]
    assert list(out["variacao_ano_rs_mil"]) == [30.0, 12.0]
    assert set(out["fechamento"]) == {"31/12"}
    assert list(pd.to_datetime(out["data"]).dt.day.unique()) == [31]


def test_agregados_ano_corrente_parcial():
    df = _serie(1788, ["2026-05-01", "2026-06-01"], [100.0, 110.0])
    out = agregados_anuais(df, [2026])
    assert out.iloc[0]["valor_dezembro_rs_mil"] == 110.0
    assert out.iloc[0]["variacao_ano_rs_mil"] == 210.0
    assert out.iloc[0]["fechamento"] == "ultimo_30/06/2026"
    assert pd.Timestamp(out.iloc[0]["data"]) == pd.Timestamp("2026-06-30")
    assert out.iloc[0]["n_meses"] == 2


def test_montar_tabelas_converte_e_fecha_identidade():
    catalogo = [
        next(s for s in SERIES if s["codigo"] == 1810),
        next(s for s in SERIES if s["codigo"] == 1809),
        next(s for s in SERIES if s["codigo"] == 1788),
    ]
    series = {
        1810: _serie(1810, ["1995-12-01", "1996-06-01", "1996-12-01"], [1000.0, 400.0, 600.0]),
        1809: _serie(1809, ["1995-12-01", "1996-06-01", "1996-12-01"], [2000.0, 100.0, 900.0]),
        1788: _serie(1788, ["1995-12-01", "1996-06-01", "1996-12-01"], [10_000.0, 10_500.0, 12_000.0]),
    }
    longo, dez, var = montar_tabelas(series, [1995, 1996], catalogo=catalogo)
    assert dez.loc[dez["ano"] == 1996, "tesouro_conta_unica"].iloc[0] == 0.6
    assert var.loc[var["ano"] == 1996, "tesouro_conta_unica"].iloc[0] == 1.0
    assert var.loc[var["ano"] == 1996, "titulos_publicos_total"].iloc[0] == 1.0
    assert var.loc[var["ano"] == 1996, "soma_fatores"].iloc[0] == 2.0
    assert var.loc[var["ano"] == 1996, "base_monetaria_restrita"].iloc[0] == 2.0
    assert var.loc[var["ano"] == 1996, "variacao_base"].iloc[0] == 2.0
    assert abs(var.loc[var["ano"] == 1996, "discrepancia"].iloc[0]) < 1e-9


def test_baixar_sgs_ignora_404(monkeypatch):
    class _Resp:
        def __init__(self, status, payload=None):
            self.status_code = status
            self._payload = payload or []
            self.content = b"[]" if status == 404 else b'[{"data":"01/12/2015"}]'

        def raise_for_status(self):
            if self.status_code >= 400:
                raise RuntimeError(f"http {self.status_code}")

        def json(self):
            return self._payload

    calls = []

    def fake_get(url, params=None, timeout=120):
        calls.append(params["dataInicial"])
        if params["dataInicial"].endswith("1995"):
            return _Resp(404)
        return _Resp(200, [{"data": "01/12/2015", "valor": "1500"}])

    session = type("S", (), {"get": staticmethod(fake_get)})()
    out = baixar_sgs(29004, inicio="01/01/1995", fim="31/12/2016", session=session)
    assert not out.empty
    assert out.iloc[0]["valor"] == 1500.0
    assert any(c.endswith("1995") for c in calls)


def test_gravar_saidas_e_markdown(tmp_path: Path):
    catalogo = [next(s for s in SERIES if s["codigo"] == 1810)]
    series = {1810: _serie(1810, ["1995-12-01"], [1_500_000.0])}
    longo, dez, var = montar_tabelas(series, [1995], catalogo=catalogo)
    caminhos = gravar_saidas(longo, dez, var, tmp_path, stem="teste_fatores")
    assert caminhos["xlsx"].exists()
    md = caminhos["md_dezembro"].read_text(encoding="utf-8")
    assert formatar_milhoes(1500.0) in md
    xl = pd.ExcelFile(caminhos["xlsx"])
    assert "Saldo_ultimo_dia_ano" in xl.sheet_names
    assert "Variacao_no_ano" in xl.sheet_names
    assert "1810" in markdown_tabela(dez, "t", "n", catalogo=catalogo)
