"""Testes dos fatores condicionantes da DBGG (1995–2026)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from scripts.fatores_condicionantes_dbgg import (
    ANO_FIM,
    ANO_INICIO,
    COL_DEMAIS,
    COL_ESTOQUE_PP,
    COL_JUROS,
    COL_METODO,
    COL_PIB,
    COL_PRIMARIO,
    COL_VARIACAO,
    FATORES,
    FMT_PP,
    MENOS,
    METODO_ANTIGO,
    METODO_ATUAL,
    VERMELHO_FUNDO,
    anos_relatorio,
    carregar_painel,
    efeito_pib,
    identidade_anual,
    linha_anual,
    nfsp_gg,
    processar,
    tabela_anual,
    tabela_discriminativo,
)


def _painel_sintetico() -> pd.DataFrame:
    """Painel mensal com NFSP, PIB e as duas metodologias da DBGG."""
    meses = pd.date_range("1998-01-01", "2009-07-01", freq="MS")
    n = len(meses)
    pib = []
    d_old = []
    d_new = []
    d_rs = []
    j_fed_12, j_em_12 = [], []
    p_fed_12, p_em_12 = [], []
    j_fed_m, j_est_m, j_mun_m = [], [], []
    p_fed_m, p_est_m, p_mun_m = [], [], []

    # PIB 12m: 1.000.000 em dez/2005 e cresce 10% a cada dezembro.
    # Interpolar linearmente entre dezembros.
    pib_dez = {
        1997: 500_000.0,
        1998: 550_000.0,
        1999: 600_000.0,
        2000: 660_000.0,
        2001: 726_000.0,
        2002: 798_600.0,
        2003: 878_460.0,
        2004: 966_306.0,
        2005: 1_000_000.0,
        2006: 1_100_000.0,
        2007: 1_210_000.0,
        2008: 1_331_000.0,
        2009: 1_464_100.0,
    }

    def _pib_mes(ts: pd.Timestamp) -> float:
        y0 = pib_dez[ts.year - 1]
        y1 = pib_dez[ts.year]
        return y0 + (y1 - y0) * ts.month / 12.0

    # Estoques oficiais só em dezembro (e jul/2009 para o ano incompleto).
    estoque_old = {2004: 70.0, 2005: 68.0, 2006: 64.0, 2007: 62.0}
    estoque_new = {2006: 55.0, 2007: 56.0, 2008: 54.0, 2009: 57.0}

    # NFSP 12m em dezembro: juros GG = 6, primário GG = −2 (superávit).
    # No mensal, 1/12 desse fluxo em R$ = (pp/100)*PIB_dez/12.
    for ts in meses:
        y = ts.year
        pib.append(_pib_mes(ts))
        d_old.append(estoque_old[y] if ts.month == 12 and y in estoque_old else float("nan"))
        if y in estoque_new and (ts.month == 12 or (y == 2009 and ts.month == 7)):
            d_new.append(estoque_new[y])
            d_rs.append(estoque_new[y] / 100.0 * _pib_mes(ts))
        else:
            d_new.append(float("nan"))
            d_rs.append(float("nan"))

        if ts.month == 12 and y >= 1999:
            j_fed_12.append(4.0)
            j_em_12.append(2.0)
            p_fed_12.append(-1.5)
            p_em_12.append(-0.5)
        else:
            j_fed_12.append(float("nan"))
            j_em_12.append(float("nan"))
            p_fed_12.append(float("nan"))
            p_em_12.append(float("nan"))

        # Fluxo mensal em R$: (6% e −2% do PIB de dezembro) / 12.
        pib_dez_ano = pib_dez[y]
        j_fed_m.append(0.04 * pib_dez_ano / 12.0 if y >= 1999 else float("nan"))
        j_est_m.append(0.015 * pib_dez_ano / 12.0)
        j_mun_m.append(0.005 * pib_dez_ano / 12.0)
        p_fed_m.append(-0.015 * pib_dez_ano / 12.0 if y >= 1999 else float("nan"))
        p_est_m.append(-0.004 * pib_dez_ano / 12.0)
        p_mun_m.append(-0.001 * pib_dez_ano / 12.0)

    assert len(pib) == n
    return pd.DataFrame(
        {
            13762: d_new,
            13761: d_rs,
            4537: d_old,
            4382: pib,
            5751: j_fed_12,
            5753: j_em_12,
            5784: p_fed_12,
            5786: p_em_12,
            4607: j_fed_m,
            4610: j_est_m,
            4611: j_mun_m,
            4640: p_fed_m,
            4643: p_est_m,
            4644: p_mun_m,
        },
        index=meses,
    )


def _gravar_csvs(pasta: Path, painel: pd.DataFrame) -> dict[int, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    arquivos = {}
    for codigo in painel.columns:
        dest = pasta / f"{codigo}.csv"
        pd.DataFrame({"mes": painel.index, "valor": painel[codigo].to_numpy()}).to_csv(
            dest, index=False
        )
        arquivos[int(codigo)] = dest
    return arquivos


def test_periodo_1995_2026():
    assert ANO_INICIO == 1995
    assert ANO_FIM == 2026


def test_efeito_pib_formula():
    # d=50, g=10% → −50 × 0,10/1,10 = −4,545...
    assert efeito_pib(50.0, 0.10) == pytest.approx(-50.0 * 0.10 / 1.10)
    assert pd.isna(efeito_pib(50.0, -1.0))
    assert pd.isna(efeito_pib(float("nan"), 0.1))


def test_nfsp_gg_ano_fechado_usa_12_meses():
    p = _painel_sintetico()
    juros, prim = nfsp_gg(p, 2006)
    assert juros == pytest.approx(6.0)
    assert prim == pytest.approx(-2.0)


def test_nfsp_gg_ano_incompleto_usa_fluxo_mensal():
    p = _painel_sintetico()
    juros, prim = nfsp_gg(p, 2009)
    # 7/12 de 6% e −2% do PIB de 2009, sobre o PIB de jul/2009 (não o de dez).
    pib_jul = p.loc[pd.Timestamp("2009-07-01"), 4382]
    pib_dez = 1_464_100.0
    assert juros == pytest.approx(100.0 * (0.06 * pib_dez * 7 / 12.0) / pib_jul)
    assert prim == pytest.approx(100.0 * (-0.02 * pib_dez * 7 / 12.0) / pib_jul)


def test_nfsp_gg_sem_federal_fica_vazio():
    p = _painel_sintetico()
    juros, prim = nfsp_gg(p, 1998)
    assert pd.isna(juros) and pd.isna(prim)


def test_identidade_e_sinal_do_superavit():
    p = _painel_sintetico()
    row = linha_anual(p, 2007)
    assert row[COL_METODO] == METODO_ATUAL
    assert row[COL_ESTOQUE_PP] == pytest.approx(56.0)
    assert row[COL_PRIMARIO] == pytest.approx(-2.0)  # superávit reduz a dívida
    assert row[COL_JUROS] == pytest.approx(6.0)
    # g = 1.210.000 / 1.100.000 − 1 (PIB de dez/07 vs dez/06)
    g = 1_210_000.0 / 1_100_000.0 - 1.0
    assert row[COL_PIB] == pytest.approx(efeito_pib(55.0, g))
    delta = 56.0 - 55.0
    assert row[COL_VARIACAO] == pytest.approx(delta)
    assert row[COL_DEMAIS] == pytest.approx(
        delta - row[COL_JUROS] - row[COL_PRIMARIO] - row[COL_PIB]
    )
    assert row[COL_VARIACAO] == pytest.approx(
        row[COL_JUROS] + row[COL_PRIMARIO] + row[COL_PIB] + row[COL_DEMAIS]
    )


def test_quebra_metodologica_2006_2007():
    p = _painel_sintetico()
    r06 = linha_anual(p, 2006)
    r07 = linha_anual(p, 2007)
    assert r06[COL_METODO] == METODO_ANTIGO
    assert r06[COL_ESTOQUE_PP] == pytest.approx(64.0)
    assert r06[COL_VARIACAO] == pytest.approx(64.0 - 68.0)
    assert r07[COL_METODO] == METODO_ATUAL
    assert r07[COL_ESTOQUE_PP] == pytest.approx(56.0)


def test_antes_de_2002_sem_estoque():
    p = _painel_sintetico()
    r99 = linha_anual(p, 1999)
    assert pd.isna(r99[COL_ESTOQUE_PP])
    assert pd.isna(r99[COL_VARIACAO])
    assert r99[COL_JUROS] == pytest.approx(6.0)
    assert r99[COL_PRIMARIO] == pytest.approx(-2.0)
    r95 = linha_anual(p, 1995)
    assert pd.isna(r95[COL_JUROS])
    assert r95[COL_METODO] == ""


def test_discriminativo_soma_algebrica_e_anos():
    p = _painel_sintetico()
    disc = tabela_discriminativo(p, list(range(1995, 2010)))
    assert list(disc["Ano"]) == list(range(1995, 2010))
    for nome in FATORES:
        assert nome in disc.columns
    fechados = disc[disc["Ano"].isin([2006, 2007, 2008])]
    for _, row in fechados.iterrows():
        assert pd.notna(row[COL_VARIACAO])
        soma = sum(float(row[n]) for n in FATORES)
        assert row[COL_VARIACAO] == pytest.approx(soma)
    assert bool(disc.loc[disc["Ano"] == 2009, "incompleto"].iloc[0])


def test_identidade_anual_residuo_zero():
    p = _painel_sintetico()
    disc = tabela_discriminativo(p, [2006, 2007, 2008])
    ident = identidade_anual(disc)
    assert (ident["residuo"].abs() < 1e-9).all()


def test_anos_relatorio_corta_no_ultimo():
    p = _painel_sintetico()
    anos = anos_relatorio(p)
    assert anos[0] == 1995
    assert anos[-1] == 2009


def test_carregar_painel_arquivos_nao_baixa(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    painel = _painel_sintetico()
    arquivos = _gravar_csvs(tmp_path / "in", painel)

    def _fail(*_a, **_k):
        raise AssertionError("carregar_painel com arquivos= nao deve chamar a rede")

    monkeypatch.setattr("scripts.fatores_condicionantes_dbgg.baixar_sgs", _fail)
    monkeypatch.setattr("scripts.fatores_condicionantes_dbgg._http_get", _fail)
    out = carregar_painel(tmp_path / "cache", arquivos=arquivos)
    assert 13762 in out.columns
    assert 4537 in out.columns
    assert out.index.min().date() == date(1998, 1, 1)
    # código ausente de arquivos: não inventa coluna e não baixa
    parcial = {13762: arquivos[13762]}
    out2 = carregar_painel(tmp_path / "cache2", arquivos=parcial)
    assert list(out2.columns) == [13762]


def test_planilha_discriminativo(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    painel = _painel_sintetico()
    arquivos = _gravar_csvs(tmp_path / "in", painel)

    def _fail(*_a, **_k):
        raise AssertionError("processar com arquivos= nao deve chamar a rede")

    monkeypatch.setattr("scripts.fatores_condicionantes_dbgg.baixar_sgs", _fail)
    saida = tmp_path / "dbgg.xlsx"
    path = processar(
        pasta_cache=tmp_path / "cache",
        saida=saida,
        usar_cache=False,
        arquivos=arquivos,
    )
    assert path.exists()
    nomes = pd.ExcelFile(path).sheet_names
    assert nomes[0] == "Metodologia"
    for aba in ("Discriminativo", "Anual", "Identidade", "Grafico"):
        assert aba in nomes
    from openpyxl import load_workbook

    wb = load_workbook(path)
    ws = wb["Discriminativo"]
    headers = [c.value for c in ws[4]]
    assert headers[0] == "Ano"
    assert headers[1:5] == list(FATORES)
    assert COL_VARIACAO in headers
    idx_var = headers.index(COL_VARIACAO)
    assert idx_var == 1 + len(FATORES)
    # 2007 é um ano com os quatro fatores → fórmula SOMA
    cel_var = None
    cel_prim = None
    idx_prim = headers.index(COL_PRIMARIO)
    for row in ws.iter_rows(min_row=5, max_col=idx_var + 1):
        if str(row[0].value) == "2007":
            cel_var = row[idx_var]
            cel_prim = row[idx_prim]
            break
    assert cel_var is not None
    assert str(cel_var.value).startswith("=SUM(")
    assert MENOS in FMT_PP
    assert cel_prim is not None
    assert cel_prim.value == pytest.approx(-2.0)
    assert cel_prim.font.bold is True
    assert VERMELHO_FUNDO in str(cel_prim.fill.fgColor.rgb or "")
    # 2009 incompleto
    rotulos = [ws.cell(r, 1).value for r in range(5, 5 + 20)]
    assert "2009*" in rotulos
    anual = tabela_anual(tabela_discriminativo(painel, [2007]))
    assert anual.loc[anual["Item"] == COL_JUROS, 2007].iloc[0] == pytest.approx(6.0)
