"""Testes da atualização IPCA da base de desembolsos 1995–2001."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from scripts.atualizar_desembolsos_ipca_1995_2001 import (
    atualizar_ipca,
    carregar_desembolsos,
    carregar_ipca_desde_1995,
    main,
    resumo_anual,
    total_periodo,
)


def _mini_base(path: Path) -> None:
    # header na linha 2 (como o arquivo oficial)
    rows = [
        ["DESEMBOLSOS DO SISTEMA BNDES"] + [None] * 13,
        [None] * 14,
        [
            "ANO",
            "MÊS",
            "FORMA DE APOIO",
            "PRODUTO",
            "PORTE DE EMPRESA",
            "REGIÃO",
            "UF",
            "MUNICÍPIO",
            "MUNICÍPIO - CÓDIGO",
            "SETOR CNAE",
            "SUBSETOR CNAE AGRUPADO",
            "SETOR BNDES",
            "SUBSETOR BNDES",
            "DESEMBOLSOS\n(R$)",
        ],
        [2001, "DEZEMBRO", "DIRETA", "X", "GRANDE", "SE", "SP", "SAO PAULO", 1, "A", "B", "C", "D", 1000.0],
        [1995, "JANEIRO", "INDIRETA", "Y", "MICRO", "NE", "BA", "SALVADOR", 2, "A", "B", "C", "D", 2000.0],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="DESEMBOLSOS_BASE DE DADOS", index=False, header=False)


def _ipca_xlsx(path: Path, taxa: float = 0.5) -> None:
    mes = pd.date_range("1995-01-01", "2026-06-01", freq="MS")
    pd.DataFrame({"Data": mes, "IPCA": [taxa] * len(mes)}).to_excel(path, index=False)


def test_carregar_e_atualizar(tmp_path: Path):
    excel = tmp_path / "base.xlsx"
    _mini_base(excel)
    df = carregar_desembolsos(excel)
    assert len(df) == 2
    assert df.loc[0, "mes_num"] == 12

    ipca_path = tmp_path / "ipca.xlsx"
    _ipca_xlsx(ipca_path, taxa=0.5)
    ipca = carregar_ipca_desde_1995(ipca_path)
    out = atualizar_ipca(df, ipca, data_ref=datetime(2026, 6, 30))
    assert out["desembolso_ipca"].iloc[0] > 1000.0
    # 1995-01 tem mais meses de IPCA que 2001-12 → fator maior
    jan95 = out.loc[out["ANO"] == 1995, "desembolso_ipca"].iloc[0]
    dez01 = out.loc[out["ANO"] == 2001, "desembolso_ipca"].iloc[0]
    assert jan95 / 2000.0 > dez01 / 1000.0


def test_resumo_e_total(tmp_path: Path):
    excel = tmp_path / "base.xlsx"
    _mini_base(excel)
    ipca_path = tmp_path / "ipca.xlsx"
    _ipca_xlsx(ipca_path, taxa=0.4)
    df = atualizar_ipca(
        carregar_desembolsos(excel),
        carregar_ipca_desde_1995(ipca_path),
        data_ref=datetime(2026, 6, 30),
    )
    resumo = resumo_anual(df)
    total = total_periodo(resumo, pd.Timestamp("2026-06-30"))
    assert set(resumo["ANO"]) == {1995, 2001}
    assert abs(total.iloc[0]["valor_corrente"] - 3000.0) < 1e-9
    assert total.iloc[0]["valor_atualizado_IPCA"] > 3000.0


def test_main(tmp_path: Path):
    excel = tmp_path / "base.xlsx"
    _mini_base(excel)
    ipca_path = tmp_path / "ipca.xlsx"
    _ipca_xlsx(ipca_path, taxa=0.4)
    saida = tmp_path / "out.xlsx"
    rc = main(
        ["--excel", str(excel), "--ipca", str(ipca_path), "--saida", str(saida), "--sem-detalhe"]
    )
    assert rc == 0
    assert saida.exists()
    resumo = pd.read_excel(saida, sheet_name="Resumo_Anual")
    anos = set(resumo["ANO"].astype(str))
    assert {"1995", "2001", "1995-2001"} <= anos
    assert (resumo["tipo"] == "TOTAL 1995-2001").any()
