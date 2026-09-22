"""Discriminativo de lucro líquido e ativo dos bancos (IFData / COSIF)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from scripts.discriminativo_lucros_ativos_bancos import (
    eh_banco,
    escrever_excel,
    escrever_markdown,
    extrair_contas,
    grafico_evolucao,
    matriz_por_ano,
    montar_discriminativo,
    montar_evolucao,
    normalizar_conta,
    parse_saldo,
)


def test_parse_saldo_ptbr_e_vazio():
    assert parse_saldo("1744567704888,5") == 1744567704888.5
    assert parse_saldo("-108855,28") == -108855.28
    assert parse_saldo("1.234,56") == 1234.56
    assert parse_saldo(10.5) == 10.5
    assert parse_saldo(None) is None
    assert parse_saldo("null") is None
    assert parse_saldo("") is None


def test_normalizar_conta_ignora_acento_e_formula():
    assert normalizar_conta("Ativo Total") == "ativo"
    assert normalizar_conta("Lucro Líquido") == "lucro"
    assert normalizar_conta("Lucro Líquido \n(j) = (g) + (h) + (i)") == "lucro"
    assert normalizar_conta("Patrimônio Líquido") is None


def test_eh_banco_recorte():
    assert eh_banco(
        {"CodInst": "C0010069", "Tcb": "B1", "NomeInstituicao": "ITAU", "SegmentoTb": None}
    )
    assert eh_banco(
        {
            "CodInst": "00000000",
            "Tcb": "B1",
            "NomeInstituicao": "BANCO DO BRASIL S.A.",
            "SegmentoTb": "Banco Múltiplo",
        }
    )
    assert not eh_banco(
        {
            "CodInst": "C0080099",
            "Tcb": "B1",
            "NomeInstituicao": "ITAU - PRUDENCIAL",
            "SegmentoTb": None,
        }
    )
    assert not eh_banco(
        {
            "CodInst": "00068987",
            "Tcb": "B3S",
            "NomeInstituicao": "COOPERATIVA DE CRÉDITO",
            "SegmentoTb": "Cooperativa de Crédito",
        }
    )
    assert not eh_banco(
        {"CodInst": "C008", "Tcb": "N1", "NomeInstituicao": "DTVM", "SegmentoTb": None}
    )


def test_extrair_contas_deduplica_linha_repetida():
    rows = [
        {"CodInst": "C0010069", "NomeColuna": "Ativo Total", "Saldo": "10,5"},
        {"CodInst": "C0010069", "NomeColuna": "Ativo Total", "Saldo": "10,5"},
        {"CodInst": "C0010069", "NomeColuna": "Lucro Líquido", "Saldo": "1,25"},
        {"CodInst": "C0010069", "NomeColuna": "Captações", "Saldo": "9"},
        {"CodInst": "X", "NomeColuna": "Ativo Total", "Saldo": "null"},
    ]
    contas = extrair_contas(rows)
    assert contas["C0010069"]["ativo"] == 10.5
    assert contas["C0010069"]["lucro"] == 1.25
    assert "X" not in contas


def _cad(nome: str, tcb: str = "B1") -> dict:
    return {
        "CodInst": nome,
        "Tcb": tcb,
        "NomeInstituicao": nome,
        "SegmentoTb": "Banco Múltiplo",
        "Uf": "SP",
        "Tc": 2,
        "Sr": "S1",
    }


def test_lucro_anual_soma_os_dois_semestres_cosif():
    por = {
        202406: {"ITAU": {"ativo": 100.0, "lucro": 19.0}, "COOP": {"ativo": 5.0, "lucro": 1.0}},
        202412: {"ITAU": {"ativo": 110.0, "lucro": 21.0}},
    }
    cad = {
        202406: {"ITAU": _cad("ITAU"), "COOP": _cad("COOP", tcb="B3S")},
        202412: {"ITAU": _cad("ITAU"), "COOP": _cad("COOP", tcb="B3S")},
    }
    disc = montar_discriminativo(por, cad, ano_min=2024, ano_max=2024)
    assert list(disc["cod_inst"]) == ["ITAU"]
    row = disc.iloc[0]
    assert row["ativo"] == 110.0
    assert row["data_base_ativo"] == 202412
    assert row["lucro_liquido"] == 40.0
    assert bool(row["exercicio_completo"]) is True


def test_ano_parcial_usa_so_primeiro_semestre():
    por = {202606: {"ITAU": {"ativo": 120.0, "lucro": 12.0}}}
    cad = {202606: {"ITAU": _cad("ITAU")}}
    disc = montar_discriminativo(por, cad, ano_min=2026, ano_max=2026)
    row = disc.iloc[0]
    assert row["ativo"] == 120.0
    assert row["data_base_ativo"] == 202606
    assert row["lucro_liquido"] == 12.0
    assert bool(row["exercicio_completo"]) is False
    assert bool(row["ano_com_dezembro"]) is False


def test_evolucao_variacao_e_indice():
    por = {
        200206: {"A": {"ativo": 50.0, "lucro": 1.0}},
        200212: {"A": {"ativo": 80.0, "lucro": 2.0}},
        200306: {"A": {"ativo": 90.0, "lucro": 1.5}},
        200312: {"A": {"ativo": 100.0, "lucro": 2.5}},
    }
    cad = {
        m: {"A": _cad("A")}
        for m in (200206, 200212, 200306, 200312)
    }
    disc = montar_discriminativo(por, cad, ano_min=2002, ano_max=2003)
    evo = montar_evolucao(disc)
    assert list(evo["ano"]) == [2002, 2003]
    assert evo.loc[0, "lucro_liquido"] == 3.0
    assert evo.loc[1, "ativo"] == 100.0
    assert abs(evo.loc[1, "var_ativo"] - (100 / 80 - 1)) < 1e-12
    assert abs(evo.loc[1, "var_lucro"] - (4 / 3 - 1)) < 1e-12
    assert evo.loc[0, "indice_ativo_ano_base"] == 100
    assert pd.isna(evo.loc[0, "var_lucro"])


def test_migracao_de_codigo_nao_duplica_o_ativo():
    ind = {
        "CodInst": "00360305",
        "Tcb": "B1",
        "NomeInstituicao": "CAIXA ECONOMICA FEDERAL",
        "SegmentoTb": "Caixa Econômica Federal",
        "CodConglomeradoFinanceiro": "C0051626",
        "Uf": "DF",
        "Tc": 1,
        "Sr": "S1",
    }
    cong = {
        "CodInst": "C0051626",
        "Tcb": "B1",
        "NomeInstituicao": "CAIXA ECONÔMICA FEDERAL",
        "SegmentoTb": None,
        "CodConglomeradoFinanceiro": "C0051626",
        "Uf": "DF",
        "Tc": 1,
        "Sr": "S1",
    }
    por = {
        202106: {"00360305": {"ativo": 1460.0, "lucro": 10.0}},
        202112: {"C0051626": {"ativo": 1449.0, "lucro": 6.0}},
    }
    cad = {
        202106: {"00360305": {**ind, "CodConglomeradoFinanceiro": None}},
        202112: {"00360305": ind, "C0051626": cong},
    }
    disc = montar_discriminativo(por, cad, ano_min=2021, ano_max=2021)
    assert list(disc["cod_inst"]) == ["C0051626"]
    assert disc.iloc[0]["ativo"] == 1449.0
    assert disc.iloc[0]["lucro_1s"] == 10.0
    assert disc.iloc[0]["lucro_2s"] == 6.0
    assert disc.iloc[0]["lucro_liquido"] == 16.0


def test_banco_encerrado_no_semestre_nao_repete_ativo_de_junho():
    por = {
        200806: {"C0010100": {"ativo": 170.0, "lucro": 2.0}},
        200812: {"C0010069": {"ativo": 630.0, "lucro": 4.0}},
    }
    cad = {
        200806: {"C0010100": _cad("UNIBANCO")},
        200812: {"C0010069": _cad("ITAU")},
    }
    disc = montar_discriminativo(por, cad, ano_min=2008, ano_max=2008)
    assert set(disc["cod_inst"]) == {"C0010100", "C0010069"}
    uni = disc.set_index("cod_inst").loc["C0010100"]
    assert pd.isna(uni["ativo"])
    assert uni["lucro_liquido"] == 2.0
    evo = montar_evolucao(disc)
    assert evo.iloc[0]["ativo"] == 630.0
    assert evo.iloc[0]["lucro_liquido"] == 6.0
    assert evo.iloc[0]["n_bancos"] == 1


def test_matriz_preserva_nome_em_bilhoes():
    por = {
        202412: {
            "A": {"ativo": 2e9, "lucro": 1e8},
            "B": {"ativo": 1e9, "lucro": 5e7},
        }
    }
    cad = {202412: {"A": _cad("ALFA"), "B": _cad("BETA")}}
    disc = montar_discriminativo(por, cad, ano_min=2024, ano_max=2024)
    matriz = matriz_por_ano(disc, "ativo", em_bilhoes=True)
    assert list(matriz.index) == ["A", "B"]
    assert matriz.loc["A", "nome"] == "ALFA"
    assert matriz.loc["A", 2024] == 2


def test_markdown_e_excel_contem_o_periodo(tmp_path: Path):
    por = {
        200206: {"A": {"ativo": 10.0, "lucro": 1.0}},
        200212: {"A": {"ativo": 20.0, "lucro": 2.0}},
        202606: {"A": {"ativo": 30.0, "lucro": 3.0}},
    }
    cad = {m: {"A": _cad("BANCO A")} for m in por}
    disc = montar_discriminativo(por, cad, ano_min=2002, ano_max=2026)
    evo = montar_evolucao(disc)
    md = tmp_path / "disc.md"
    escrever_markdown(evo, disc, md, "2026-09-22 00:00 UTC")
    texto = md.read_text(encoding="utf-8")
    assert "2002" in texto
    assert "2026" in texto
    assert "COSIF" in texto
    assert "†" in texto
    xlsx = tmp_path / "disc.xlsx"
    escrever_excel(evo, disc, xlsx)
    livro = pd.ExcelFile(xlsx)
    assert {"Nota", "Evolucao", "Discriminativo", "Ativo_R$bi", "Lucro_R$bi"} <= set(livro.sheet_names)
    png = tmp_path / "disc.png"
    grafico_evolucao(evo, png)
    assert png.stat().st_size > 1000
