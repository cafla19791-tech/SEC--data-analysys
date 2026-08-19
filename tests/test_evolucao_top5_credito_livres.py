"""Testes da carteira das 5 maiores IFs e do proxy de recursos livres."""

from __future__ import annotations

import pandas as pd

from scripts.evolucao_top5_credito_livres import (
    CANON_ORDEM,
    PF_LIVRES,
    PJ_LIVRES,
    agregar_big5,
    cabecalhos_carteira,
    linhas_carteira,
    nome_canonico,
    proxy_livres,
)


def test_canonico_une_fusoes_e_ignora_bndes() -> None:
    assert nome_canonico("Banco do Brasil S.A.") == "Banco do Brasil"
    assert nome_canonico("Itaú Unibanco S.A.") == "Itaú Unibanco"
    assert nome_canonico("Unibanco - União de Bancos Brasileiros S.A.") == "Itaú Unibanco"
    assert nome_canonico("Banco Bradesco S.A.") == "Bradesco"
    assert nome_canonico("Caixa Econômica Federal") == "Caixa"
    assert nome_canonico("Banco Santander (Brasil) S.A.") == "Santander"
    assert nome_canonico("ABN AMRO Bank") == "Santander"
    assert nome_canonico("BNDES") is None
    assert nome_canonico("Banco Nacional de Desenvolvimento Econômico e Social") is None
    assert nome_canonico("Nubank") is None


def test_agregar_big5_soma_fusoes_e_nao_duplica_if() -> None:
    cad = pd.DataFrame(
        [
            {"CodInst": "C001", "NomeInstituicao": "Itaú"},
            {"CodInst": "C002", "NomeInstituicao": "Unibanco"},
            {"CodInst": "C003", "NomeInstituicao": "Caixa Econômica Federal"},
            {"CodInst": "00360305", "NomeInstituicao": "Caixa Econômica Federal"},
            {"CodInst": "C004", "NomeInstituicao": "BNDES"},
        ]
    )
    cred = pd.DataFrame(
        [
            {"CodInst": "C001", "Saldo": 100.0},
            {"CodInst": "C002", "Saldo": 80.0},
            {"CodInst": "C003", "Saldo": 120.0},
            {"CodInst": "00360305", "Saldo": 90.0},
            {"CodInst": "C004", "Saldo": 999.0},
        ]
    )
    cons = agregar_big5(cred, cad)
    assert cons.loc[cons["banco"] == "Itaú Unibanco", "carteira"].iloc[0] == 180.0
    assert cons.loc[cons["banco"] == "Caixa", "carteira"].iloc[0] == 120.0
    assert "BNDES" not in set(cons["banco"])


def test_proxy_livres_nao_duplica_conglomerado_e_exclui_direcionados() -> None:
    cad = pd.DataFrame(
        [
            {"CodInst": "X", "NomeInstituicao": "Itaú Unibanco S.A."},
            {"CodInst": "Y", "NomeInstituicao": "Itaú Unibanco S.A."},
        ]
    )
    modal = pd.DataFrame(
        [
            {"CodInst": "X", "Grupo": "Cartão de Crédito", "Saldo": 10.0},
            {"CodInst": "Y", "Grupo": "Cartão de Crédito", "Saldo": 12.0},
            {"CodInst": "X", "Grupo": "Empréstimo com Consignação em Folha", "Saldo": 20.0},
            {"CodInst": "X", "Grupo": "Habitação", "Saldo": 999.0},
            {"CodInst": "X", "Grupo": "Capital de Giro", "Saldo": 30.0},
            {"CodInst": "X", "Grupo": "Rural e agroindustrial", "Saldo": 888.0},
        ]
    )
    livres = proxy_livres(modal, cad, PF_LIVRES | PJ_LIVRES)
    # cartão usa o maior saldo do grupo (12), não a soma 10+12
    assert livres.loc[livres["banco"] == "Itaú Unibanco", "livres"].iloc[0] == 62.0


def test_linhas_carteira_formato_brasileiro() -> None:
    painel = pd.DataFrame(
        [
            {
                "ano": 2024,
                "cart_Banco do Brasil": 100.0,
                "cart_Itaú Unibanco": 90.0,
                "cart_Bradesco": 80.0,
                "cart_Caixa": 70.0,
                "cart_Santander": 60.0,
                "cart_top5": 400.0,
                "liv_Banco do Brasil": None,
                "liv_top5": None,
            },
            {
                "ano": 2026,
                "cart_Banco do Brasil": 110.4,
                "cart_Itaú Unibanco": 91.0,
                "cart_Bradesco": 82.0,
                "cart_Caixa": 71.0,
                "cart_Santander": 61.0,
                "cart_top5": 415.4,
                "liv_Banco do Brasil": 40.0,
                "liv_Itaú Unibanco": 50.0,
                "liv_Bradesco": 30.0,
                "liv_Caixa": 10.0,
                "liv_Santander": 20.0,
                "liv_top5": 150.0,
            },
        ]
    )
    assert cabecalhos_carteira()[0] == "Ano"
    assert cabecalhos_carteira()[1:] == [*CANON_ORDEM, "Soma das 5"]
    linhas = linhas_carteira(painel, "cart")
    assert linhas[0][0] == "2024"
    assert linhas[0][1] == "100,0"
    assert linhas[1][0] == "2026*"
    assert linhas[1][-1] == "415,4"
    liv = linhas_carteira(painel, "liv")
    assert liv[0][1] == "—"
    assert liv[1][-1] == "150,0"
    so_liv = linhas_carteira(painel, "liv", so_com_valor=True)
    assert [linha[0] for linha in so_liv] == ["2026*"]
