"""Testes do resumo explicativo das atas do Copom."""

from __future__ import annotations

import pandas as pd

from scripts.resumo_atas_copom import (
    agregar_reunioes,
    classificar_decisao,
    contar_temas,
    extrair_voto,
    html_para_texto,
    identificador,
    identificar_recuos,
    numero_reuniao,
    resumo_anual,
    tabela_html,
    trecho_decisao,
)


def test_html_para_texto_remove_tags_e_quebra_paragrafo() -> None:
    html = "<p>O Copom decidiu elevar a Selic.</p><script>x</script><br/>IPCA."
    texto = html_para_texto(html)
    assert "O Copom decidiu elevar a Selic." in texto
    assert "IPCA." in texto
    assert "<p>" not in texto
    assert "script" not in texto.lower()


def test_numero_reuniao_e_identificador() -> None:
    assert numero_reuniao("280ª reunião do Copom") == 280
    assert numero_reuniao("55ª Reunião") == 55
    assert numero_reuniao("sem número") is None
    assert identificador("/publicacoes/atascopom/05082026") == "05082026"


def test_classificar_decisao() -> None:
    assert classificar_decisao(0.50) == "alta"
    assert classificar_decisao(-0.25) == "corte"
    assert classificar_decisao(0.00) == "manutenção"
    assert classificar_decisao(0.02) == "manutenção"
    assert classificar_decisao(None) == "—"


def test_extrair_voto() -> None:
    assert extrair_voto("A decisão foi tomada por unanimidade.") == "unânime"
    assert extrair_voto("O Comitê então decidiu, unanimemente, pela elevação.") == "unânime"
    assert extrair_voto("O Copom decidiu, por 5 votos a 4, elevar a taxa.") == "5 a 4"
    assert extrair_voto("por sete votos a favor e dois contra, o Copom decidiu") == "7 a 2"
    assert extrair_voto("sem menção a voto") == "—"
    split = (
        "Votaram por uma redução de 0,50 ponto percentual os seguintes membros do Comitê: "
        "Roberto de Oliveira Campos Neto (presidente), Ailton de Aquino Santos, "
        "Carolina de Assis Barros, Gabriel Muricca Galípolo e Otávio Ribeiro Damaso. "
        "Votaram por uma redução de 0,25 ponto percentual os seguintes membros: "
        "Diogo Abry Guillen, Fernanda Magalhães Rumenos Guardado, "
        "Maurício Costa de Moura e Renato Dias de Brito Gomes."
    )
    assert extrair_voto(split) == "5 a 4"
    unica = (
        "Votaram por essa decisão os seguintes membros do Comitê: "
        "Gabriel Muricca Galípolo (presidente), Ailton de Aquino Santos, "
        "Gilneu Francisco Astolfi Vivan, Izabela Moreira Correa, "
        "Nilton José Schneider David, Paulo Picchetti e Rodrigo Alves Teixeira."
    )
    assert extrair_voto(unica) == "unânime"


def test_trecho_decisao_prioriza_secao() -> None:
    texto = "Introdução longa. Decisão de política monetária: o Copom manteve a Selic."
    trecho = trecho_decisao(texto, limite=80)
    assert trecho.startswith("Decisão de política monetária")


def test_contar_temas_inflacao_e_pandemia() -> None:
    texto = "A inflação do IPCA e a pandemia de covid-19."
    temas = contar_temas(texto)
    assert temas["inflacao"] >= 2
    assert temas["pandemia"] >= 1


def test_tabela_html_tem_grade_continua() -> None:
    html = tabela_html(["Ano", "Selic"], [["2001", "15,25"]])
    assert "border-collapse:collapse" in html
    assert "border:1px solid" in html
    assert "<th " in html
    assert "15,25" in html


def test_agregar_reunioes_e_resumo_anual() -> None:
    atas = pd.DataFrame(
        {
            "reuniao": [55, 56, 57],
            "data": pd.to_datetime(["2001-01-17", "2001-02-14", "2001-03-21"]),
            "texto": [
                "O Copom decidiu por unanimidade.",
                "Decisão de política monetária: alta. por 7 votos a 1.",
                "O Copom manteve a taxa.",
            ],
        }
    )
    selic = pd.DataFrame(
        {
            "data": pd.to_datetime(
                [
                    "2001-01-16",
                    "2001-01-17",
                    "2001-01-18",
                    "2001-02-14",
                    "2001-02-15",
                    "2001-03-21",
                    "2001-03-22",
                ]
            ),
            "selic": [15.75, 15.75, 15.25, 15.25, 15.75, 15.75, 15.75],
        }
    )
    reunioes = agregar_reunioes(atas, selic)
    assert reunioes.loc[0, "decisao"] == "corte"
    assert reunioes.loc[1, "decisao"] == "alta"
    assert reunioes.loc[2, "decisao"] == "manutenção"
    assert float(reunioes.loc[0, "selic"]) == 15.25
    assert reunioes.loc[0, "voto"] == "unânime"
    assert reunioes.loc[1, "voto"] == "7 a 1"
    anual = resumo_anual(reunioes)
    assert int(anual.iloc[0]["ano"]) == 2001
    assert int(anual.iloc[0]["n"]) == 3
    assert int(anual.iloc[0]["altas"]) == 1
    assert int(anual.iloc[0]["manutencoes"]) == 1


def test_identificar_recuos_corte_e_volta() -> None:
    reunioes = pd.DataFrame(
        {
            "reuniao": [0, 1, 2, 3, 4, 5, 6],
            "data": pd.to_datetime(
                [
                    "2003-12-01",
                    "2004-01-01",
                    "2004-03-01",
                    "2004-05-01",
                    "2004-10-01",
                    "2004-12-01",
                    "2005-03-01",
                ]
            ),
            "selic_antes": [18.0, 18.0, 17.0, 16.0, 16.0, 16.5, 17.5],
            "selic": [18.0, 17.0, 16.0, 16.0, 16.5, 17.5, 18.5],
            "decisao": ["manutenção", "corte", "corte", "manutenção", "alta", "alta", "alta"],
            "ano": [2003, 2004, 2004, 2004, 2004, 2004, 2005],
        }
    )
    recuos = identificar_recuos(reunioes, ano_ini=2003, ano_fim=2015)
    assert len(recuos) == 1
    assert recuos[0]["patamar"] == 18.0
    assert recuos[0]["piso"] == 16.0
    assert recuos[0]["pico"] == 18.5
    assert recuos[0]["retorno"] == "retornou e superou"
    assert recuos[0]["n_cortes"] == 2
    assert recuos[0]["meses_ate_reverter"] > 4
