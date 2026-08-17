"""Tabelas extraídas do Relatório TCU — Contas do Governo da República 2010.

Fonte oficial:
https://sites.tcu.gov.br/recursos/contas-do-governo-2010/CG%202010%20Relat%C3%B3rio.pdf

Valores nominais conforme o PDF (Relator Ministro Aroldo Cedraz, Brasília, 2011).
Unidades: R$ milhões, salvo indicação em `unidade`.
"""

from __future__ import annotations

from datetime import date, timedelta

FONTE_URL = (
    "https://sites.tcu.gov.br/recursos/contas-do-governo-2010/"
    "CG%202010%20Relat%C3%B3rio.pdf"
)
FONTE_TITULO = (
    "Relatório e Parecer Prévio sobre as Contas do Governo da República "
    "— Exercício de 2010"
)
FONTE_ORGAO = "Tribunal de Contas da União"
RELATOR = "Ministro Aroldo Cedraz"
PRESIDENTE_TCU = "Ministro Ubiratan Aguiar (edição) / Ministro Benjamin Zymler (capa 2011)"
ANO_PUBLICACAO = 2011
EXERCICIO = 2010

# PIB usado no quadro da DLSP/DBGG (Nota de Política Fiscal / Bacen)
PIB_2009_R_MILHOES = 3_185_125.0
PIB_2010_R_MILHOES = 3_657_366.0
# PIB do quadro PAC (RFB) — R$ bilhões
PIB_PAC_2010_R_BILHOES = 3_644.0


def indicadores() -> list[dict]:
    return [
        {
            "bloco": "Crédito",
            "indicador": "Saldo de crédito do SFN (PF+PJ)",
            "valor_2009": None,
            "valor_2010": 1_700_000.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 46.4,
            "var_pct": 22.9,
            "fonte": "item 2.3 Política Creditícia",
            "nota": "Maior patamar da série histórica até 2010",
        },
        {
            "bloco": "Crédito",
            "indicador": "Repasses do BNDES (recursos direcionados)",
            "valor_2009": 124_900.0,
            "valor_2010": 179_800.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": 43.9,
            "fonte": "item 2.3 Política Creditícia",
            "nota": "Destaque da expansão de 31,1% dos recursos direcionados",
        },
        {
            "bloco": "Crédito",
            "indicador": "Crédito habitacional (recursos direcionados)",
            "valor_2009": 87_400.0,
            "valor_2010": 131_400.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": 50.4,
            "fonte": "item 2.3 Política Creditícia",
            "nota": "",
        },
        {
            "bloco": "Crédito",
            "indicador": "Recursos livres",
            "valor_2009": None,
            "valor_2010": 1_100_000.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": 16.9,
            "fonte": "item 2.3 Política Creditícia",
            "nota": "",
        },
        {
            "bloco": "Crédito",
            "indicador": "Recursos direcionados",
            "valor_2009": None,
            "valor_2010": 589_800.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": 28.3,
            "fonte": "item 2.3 Política Creditícia",
            "nota": "",
        },
        {
            "bloco": "Dívida",
            "indicador": "DBGG",
            "valor_2009": 1_973_424.0,
            "valor_2010": 2_011_522.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 55.00,
            "var_pct": None,
            "fonte": "quadro Dívida Líquida e Bruta do Governo Geral",
            "nota": "2009 = 62,02% do PIB",
        },
        {
            "bloco": "Dívida",
            "indicador": "DLSP",
            "valor_2009": 1_362_711.0,
            "valor_2010": 1_475_820.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 40.35,
            "var_pct": None,
            "fonte": "quadro Dívida Líquida e Bruta do Governo Geral",
            "nota": "2009 = 42,78% do PIB; variação −2,43 p.p. do PIB",
        },
        {
            "bloco": "Dívida",
            "indicador": "DPF estoque em mercado",
            "valor_2009": 1_497_400.0,
            "valor_2010": 1_694_000.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.4 PAF / Relatório Anual da Dívida 2010",
            "nota": "PAF 2010: intervalo 1.600–1.730",
        },
        {
            "bloco": "Dívida",
            "indicador": "Juros nominais / déficit nominal",
            "valor_2009": 171_011.0,
            "valor_2010": 195_369.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 5.34,
            "var_pct": None,
            "fonte": "Fatores condicionantes da DLSP",
            "nota": "2009 = 5,37% do PIB",
        },
        {
            "bloco": "Dívida",
            "indicador": "Resultado primário (superávit)",
            "valor_2009": 64_769.0,
            "valor_2010": 101_696.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 2.78,
            "var_pct": None,
            "fonte": "Fatores condicionantes da DLSP",
            "nota": "2009 = 2,03% do PIB; valores com sinal positivo = superávit",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Créditos da União junto ao BNDES (ativo da DLSP)",
            "valor_2009": 129_237.0,
            "valor_2010": 236_723.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 6.47,
            "var_pct": None,
            "fonte": "quadro Dívida Líquida e Bruta do Governo Geral",
            "nota": "2009 = 4,06% do PIB; aumento de R$ 107,5 bi",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Créditos a instituições financeiras oficiais",
            "valor_2009": 144_787.0,
            "valor_2010": 256_602.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 7.02,
            "var_pct": None,
            "fonte": "quadro Dívida Líquida e Bruta do Governo Geral",
            "nota": "2009 = 4,55% do PIB",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Custo fiscal anual estimado Selic − TJLP sobre estoque BNDES",
            "valor_2009": None,
            "valor_2010": 14_200.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.3",
            "nota": "Captação Tesouro 10–12% a.a. vs aplicação 4–6% a.a. sobre R$ 236,7 bi",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Subsídios Lei 11.948/2009 (1ºs R$ 100 bi) — biênio 2009/2010",
            "valor_2009": None,
            "valor_2010": 3_100.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.3",
            "nota": "Cálculo do TCU para as transferências iniciais da Lei 11.948/2009",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Emissão direta ao BNDES sem contrapartida financeira (PAF)",
            "valor_2009": None,
            "valor_2010": 24_800.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.4",
            "nota": "Parte das emissões diretas de R$ 89,9 bi (Petrobras 42,9 + BNDES 24,8 + Petros 16,3)",
        },
        {
            "bloco": "Tesouro–BNDES",
            "indicador": "Carteira Lei 11.948/2009 aplicada em obras do PAC",
            "valor_2009": None,
            "valor_2010": 34_700.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 4.1 Atuação do Tesouro Nacional",
            "nota": "De R$ 163,7 bi da carteira; TC 022.684/2010-7",
        },
        {
            "bloco": "Renúncia",
            "indicador": "Renúncia federal total (trib. + prev. + fin./créd.)",
            "valor_2009": None,
            "valor_2010": 143_970.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 3.3.4",
            "nota": "Projetado: 105,8 trib. + 19,2 trib.-prev. + 18,9 fin./créd.",
        },
        {
            "bloco": "Renúncia",
            "indicador": "Benefícios tributários",
            "valor_2009": 89_524.56,
            "valor_2010": 105_843.31,
            "unidade": "R$ milhões",
            "pct_pib_2010": 2.88,
            "var_pct": 18.2,
            "fonte": "item 3.3.4 / RFB",
            "nota": "2010 projetado; 2009 estimado",
        },
        {
            "bloco": "Renúncia",
            "indicador": "Benefícios tributários-previdenciários",
            "valor_2009": 17_044.3,
            "valor_2010": 19_246.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 3.3.4 / RFB",
            "nota": "2010 projetado",
        },
        {
            "bloco": "Renúncia",
            "indicador": "Benefícios financeiros e creditícios",
            "valor_2009": 16_901.39,
            "valor_2010": 18_877.65,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": 11.69,
            "fonte": "item 3.3.4 / SPE-MF",
            "nota": "Negativos excluídos (Portaria MF 379/2006)",
        },
        {
            "bloco": "PAC",
            "indicador": "Desonerações tributárias do PAC",
            "valor_2009": 20_000.0,
            "valor_2010": 23_318.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": 0.6,
            "var_pct": None,
            "fonte": "item 4.1.2 / RFB",
            "nota": "Acumulado 2007–2010 = R$ 63,4 bi; 68% em 2009–2010",
        },
        {
            "bloco": "PAC",
            "indicador": "Subsídios creditícios PAC — desembolsos 2010",
            "valor_2009": None,
            "valor_2010": 4_444.24,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 4.1.2 / SPE-MF",
            "nota": "Custo de oportunidade: NTN-F 2017 = 12,32% a.a. em 2010",
        },
        {
            "bloco": "PAC",
            "indicador": "Subsídios creditícios PAC — contratações 2010",
            "valor_2009": None,
            "valor_2010": 4_933.82,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 4.1.2 / SPE-MF",
            "nota": "BNDES = R$ 4,8 bi (97,5% do total)",
        },
        {
            "bloco": "PAC",
            "indicador": "Desembolsos PAC (BNDES, BB, CEF, Basa, BNB)",
            "valor_2009": None,
            "valor_2010": 21_700.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 4.1.2 / SPE-MF",
            "nota": "Energia = R$ 15,94 bi",
        },
        {
            "bloco": "PAC",
            "indicador": "Contratações PAC (BNDES, BB, CEF, Basa, BNB)",
            "valor_2009": None,
            "valor_2010": 17_400.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 4.1.2 / SPE-MF",
            "nota": "Energia = R$ 10,47 bi; logística = 33,7%",
        },
        {
            "bloco": "BNDES DRE",
            "indicador": "Resultado com TVM (consolidado BNDES)",
            "valor_2009": 5_200.0,
            "valor_2010": 8_400.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.3 / DFs BNDES",
            "nota": "Diferencial entre custo pago ao Tesouro e remuneração dos títulos recebidos",
        },
        {
            "bloco": "BNDES DRE",
            "indicador": "Lucro líquido (consolidado BNDES)",
            "valor_2009": 6_700.0,
            "valor_2010": 9_900.0,
            "unidade": "R$ milhões",
            "pct_pib_2010": None,
            "var_pct": None,
            "fonte": "item 2.5.3 / DFs BNDES",
            "nota": "",
        },
    ]


def creditos_dlsp() -> list[dict]:
    """Ativos/passivos selecionados do quadro DLSP/DBGG (R$ milhões)."""
    return [
        {"item": "Dívida Líquida do Setor Público", "v2009": 1_362_711, "pib2009": 42.78, "v2010": 1_475_820, "pib2010": 40.35, "sinal": "passivo"},
        {"item": "Dívida líquida do governo geral", "v2009": 1_378_129, "pib2009": 43.33, "v2010": 1_495_285, "pib2010": 40.88, "sinal": "passivo"},
        {"item": "Dívida bruta do governo geral (DBGG)", "v2009": 1_973_424, "pib2009": 62.02, "v2010": 2_011_522, "pib2010": 55.00, "sinal": "passivo"},
        {"item": "Dívida interna", "v2009": 1_861_984, "pib2009": 58.46, "v2010": 1_902_125, "pib2010": 52.01, "sinal": "passivo"},
        {"item": "Dívida mobiliária do Tesouro Nacional", "v2009": 1_369_262, "pib2009": 42.99, "v2010": 1_569_450, "pib2010": 42.91, "sinal": "passivo"},
        {"item": "Dívida mobiliária em mercado", "v2009": 1_381_841, "pib2009": 43.38, "v2010": 1_590_719, "pib2010": 43.49, "sinal": "passivo"},
        {"item": "Operações compromissadas do BCB", "v2009": 454_710, "pib2009": 14.28, "v2010": 288_666, "pib2010": 7.89, "sinal": "passivo"},
        {"item": "Dívida externa", "v2009": 111_440, "pib2009": 3.56, "v2010": 109_397, "pib2010": 2.99, "sinal": "passivo"},
        {"item": "Créditos do governo geral", "v2009": 830_612, "pib2009": 26.08, "v2010": 979_408, "pib2010": 26.78, "sinal": "ativo"},
        {"item": "Créditos internos", "v2009": 830_612, "pib2009": 26.08, "v2010": 979_100, "pib2010": 26.77, "sinal": "ativo"},
        {"item": "Disponibilidades do governo geral", "v2009": 445_177, "pib2009": 13.98, "v2010": 451_320, "pib2010": 12.34, "sinal": "ativo"},
        {"item": "Disponibilidades do governo federal no BCB", "v2009": 406_354, "pib2009": 12.76, "v2010": 404_516, "pib2010": 11.06, "sinal": "ativo"},
        {"item": "Créditos concedidos a inst. financeiras oficiais", "v2009": 144_787, "pib2009": 4.55, "v2010": 256_602, "pib2010": 7.02, "sinal": "ativo"},
        {"item": "Instrumentos híbridos de capital e dívida", "v2009": 15_550, "pib2009": 0.49, "v2010": 19_879, "pib2010": 0.54, "sinal": "ativo"},
        {"item": "Créditos junto ao BNDES", "v2009": 129_237, "pib2009": 4.06, "v2010": 236_723, "pib2010": 6.47, "sinal": "ativo"},
        {"item": "Aplicações em fundos e programas", "v2009": 73_851, "pib2009": 2.32, "v2010": 95_910, "pib2010": 2.62, "sinal": "ativo"},
        {"item": "Créditos junto às estatais", "v2009": 16_518, "pib2009": 0.52, "v2010": 15_274, "pib2010": 0.42, "sinal": "ativo"},
        {"item": "Recursos do FAT na rede bancária", "v2009": 140_030, "pib2009": 4.40, "v2010": 146_360, "pib2010": 4.00, "sinal": "ativo"},
        {"item": "Títulos livres na carteira do Bacen", "v2009": 183_105, "pib2009": 5.75, "v2010": 414_537, "pib2010": 11.33, "sinal": "passivo"},
        {"item": "Equalização cambial", "v2009": 52_212, "pib2009": 1.64, "v2010": 48_634, "pib2010": 1.33, "sinal": "passivo"},
        {"item": "Dívida líquida do Banco Central", "v2009": -39_189, "pib2009": -1.23, "v2010": -43_401, "pib2010": -1.19, "sinal": "liquido"},
        {"item": "Dívida líquida das empresas estatais", "v2009": 23_771, "pib2009": 0.75, "v2010": 23_937, "pib2010": 0.65, "sinal": "passivo"},
    ]


def autorizacoes_legais() -> list[dict]:
    return [
        {
            "norma": "Lei 11.948/2009, art. 1º",
            "objeto": "Crédito da União ao BNDES",
            "limite_r_bi": 100.0,
            "observacao": "Autorização inicial; remuneração: até 30% pelo custo de captação externo e o restante pela TJLP (redação Lei 12.096/2009)",
        },
        {
            "norma": "Lei 12.249/2010, art. 44",
            "objeto": "Crédito da União ao BNDES (ampliação)",
            "limite_r_bi": 180.0,
            "observacao": "Acréscimo de R$ 80 bi sobre a Lei 11.948/2009, com superávit financeiro de 2008",
        },
        {
            "norma": "Lei 12.249/2010, arts. 34 a 36",
            "objeto": "Crédito aos agentes do Fundo da Marinha Mercante",
            "limite_r_bi": 15.0,
            "observacao": "Financiamentos aprovados pelo Conselho Diretor do FMM",
        },
        {
            "norma": "Lei 12.249/2010, art. 63",
            "objeto": "Crédito da União ao BNB",
            "limite_r_bi": 1.0,
            "observacao": "",
        },
        {
            "norma": "Lei 12.249/2010, art. 64",
            "objeto": "Instrumento híbrido de capital e dívida — BNB",
            "limite_r_bi": 1.0,
            "observacao": "Enquadramento de operações de crédito com o BNB",
        },
        {
            "norma": "MP 470/2009",
            "objeto": "Crédito à CEF",
            "limite_r_bi": 4.0,
            "observacao": "",
        },
        {
            "norma": "MP 505/2010 (Lei 12.397/2011)",
            "objeto": "Crédito ao BNDES para capitalização da Petrobras",
            "limite_r_bi": 30.0,
            "observacao": "Remuneração do Tesouro pela TJLP (§ 3º do art. 1º)",
        },
        {
            "norma": "Lei 12.276/2010",
            "objeto": "Subscrição de ações da Petrobras com títulos da DPMF",
            "limite_r_bi": 42.9,
            "observacao": "Integralização com títulos da dívida pública mobiliária federal",
        },
        {
            "norma": "Lei 10.179/2001",
            "objeto": "Permuta de títulos com a Petros",
            "limite_r_bi": 16.3,
            "observacao": "Emissão direta em 2010",
        },
        {
            "norma": "MP 2.196/2001",
            "objeto": "Permuta de títulos com a CEF",
            "limite_r_bi": 2.9,
            "observacao": "",
        },
        {
            "norma": "Lei 9.818/2001",
            "objeto": "Permuta de títulos com o FGE",
            "limite_r_bi": 2.7,
            "observacao": "",
        },
        {
            "norma": "MP 526/2011 (4/3/2011)",
            "objeto": "Novo crédito ao BNDES (títulos da DPMF) e à Finep",
            "limite_r_bi": 84.0,
            "observacao": "R$ 83 bi BNDES + R$ 1 bi Finep — posterior ao exercício de 2010; TCU registra o aumento da alavancagem",
        },
        {
            "norma": "Lei 12.017/2009",
            "objeto": "Exclui a Petrobras do cálculo do superávit primário",
            "limite_r_bi": None,
            "observacao": "Altera a cobertura do indicador fiscal",
        },
        {
            "norma": "Lei 12.377/2010",
            "objeto": "Exclui a Eletrobras do cálculo do superávit primário",
            "limite_r_bi": None,
            "observacao": "Altera a cobertura do indicador fiscal",
        },
    ]


def dpf_indicadores() -> list[dict]:
    anos = [2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010]
    estoque = [965.8, 1013.9, 1157.1, 1237.0, 1333.8, 1397.3, 1497.4, 1694.0]
    prazo = [39, 35.3, 33.3, 35.5, 39.2, 42.0, 42.4, 42.0]
    venc_12m = [30.7, 39.3, 36.3, 32.4, 28.2, 25.4, 23.6, 23.9]
    prefixado = [9.5, 16.1, 23.6, 31.9, 35.1, 29.9, 32.2, 36.6]
    precos = [10.3, 11.9, 13.1, 19.9, 24.1, 26.6, 26.7, 26.6]
    selic = [46.5, 45.7, 43.9, 33.4, 30.7, 32.4, 33.4, 30.8]
    cambio = [32.4, 24.2, 17.6, 12.7, 8.2, 9.7, 6.6, 5.1]
    tr = [1.4, 2.1, 1.8, 2.0, 1.9, 1.4, 1.1, 0.8]
    rows = []
    for i, ano in enumerate(anos):
        rows.append(
            {
                "ano": ano,
                "estoque_mercado_r_bi": estoque[i],
                "prazo_medio_meses": prazo[i],
                "pct_vencimento_12m": venc_12m[i],
                "pct_prefixado": prefixado[i],
                "pct_indice_precos": precos[i],
                "pct_selic": selic[i],
                "pct_cambio": cambio[i],
                "pct_tr_outros": tr[i],
            }
        )
    return rows


def renuncia_regional() -> list[dict]:
    return [
        {"regiao": "Norte", "tributarios": 19.51, "trib_prev": 0.49, "fin_cred": 2.17, "total": 22.16, "participacao_pct": 15.4, "per_capita_r": 1397.14},
        {"regiao": "Nordeste", "tributarios": 12.86, "trib_prev": 1.89, "fin_cred": 6.20, "total": 20.95, "participacao_pct": 14.5, "per_capita_r": 394.61},
        {"regiao": "Centro-Oeste", "tributarios": 6.21, "trib_prev": 1.62, "fin_cred": 1.80, "total": 9.63, "participacao_pct": 6.7, "per_capita_r": None},
        {"regiao": "Sudeste", "tributarios": 52.34, "trib_prev": 10.70, "fin_cred": 4.32, "total": 67.36, "participacao_pct": 46.8, "per_capita_r": None},
        {"regiao": "Sul", "tributarios": 14.92, "trib_prev": 4.55, "fin_cred": 1.31, "total": 20.79, "participacao_pct": 14.4, "per_capita_r": None},
        {"regiao": "Total", "tributarios": 105.84, "trib_prev": 19.25, "fin_cred": 18.88, "total": 143.97, "participacao_pct": 100.0, "per_capita_r": 754.75},
    ]


def renuncia_tributaria() -> list[dict]:
    return [
        {"tributo": "Imposto sobre Importação", "y2006": 1868.28, "y2007": 1805.28, "y2008": 2631.09, "y2009": 2031.13, "y2010": 2431.96, "var_2006_2010_pct": 30.2, "var_2009_2010_pct": 19.7},
        {"tributo": "IR — Pessoa Física", "y2006": 7957.58, "y2007": 10533.10, "y2008": 11833.43, "y2009": 12847.99, "y2010": 14738.82, "var_2006_2010_pct": 85.2, "var_2009_2010_pct": 14.7},
        {"tributo": "IR — Pessoa Jurídica", "y2006": 17986.68, "y2007": 19265.91, "y2008": 21146.34, "y2009": 22077.58, "y2010": 26025.51, "var_2006_2010_pct": 44.7, "var_2009_2010_pct": 17.9},
        {"tributo": "IR — Retido na Fonte", "y2006": 345.14, "y2007": 367.63, "y2008": 422.20, "y2009": 388.98, "y2010": 595.12, "var_2006_2010_pct": 72.4, "var_2009_2010_pct": 53.0},
        {"tributo": "IR — total", "y2006": 26289.41, "y2007": 30166.64, "y2008": 33401.97, "y2009": 35314.56, "y2010": 41359.44, "var_2006_2010_pct": 57.3, "var_2009_2010_pct": 17.1},
        {"tributo": "IPI — operações internas", "y2006": 9723.03, "y2007": 12365.47, "y2008": 13651.27, "y2009": 12698.66, "y2010": 14732.26, "var_2006_2010_pct": 51.5, "var_2009_2010_pct": 16.0},
        {"tributo": "IPI — vinculado à importação", "y2006": 1554.22, "y2007": 1349.82, "y2008": 1823.38, "y2009": 1434.99, "y2010": 1934.25, "var_2006_2010_pct": 24.5, "var_2009_2010_pct": 34.8},
        {"tributo": "IPI — total", "y2006": 11277.26, "y2007": 13715.30, "y2008": 15474.65, "y2009": 14133.65, "y2010": 16666.51, "var_2006_2010_pct": 47.8, "var_2009_2010_pct": 17.9},
        {"tributo": "IOF", "y2006": 261.17, "y2007": 391.04, "y2008": 651.43, "y2009": 859.96, "y2010": 971.05, "var_2006_2010_pct": 271.8, "var_2009_2010_pct": 12.9},
        {"tributo": "ITR", "y2006": 23.10, "y2007": 25.29, "y2008": 24.39, "y2009": 25.59, "y2010": 26.88, "var_2006_2010_pct": 16.3, "var_2009_2010_pct": 5.0},
        {"tributo": "PIS/Pasep", "y2006": 3748.04, "y2007": 4284.43, "y2008": 5022.25, "y2009": 5270.10, "y2010": 6353.01, "var_2006_2010_pct": 69.5, "var_2009_2010_pct": 20.5},
        {"tributo": "CSLL", "y2006": 3498.90, "y2007": 4178.75, "y2008": 4234.07, "y2009": 4887.79, "y2010": 5562.74, "var_2006_2010_pct": 59.0, "var_2009_2010_pct": 13.8},
        {"tributo": "Cofins", "y2006": 18431.36, "y2007": 21177.91, "y2008": 25131.62, "y2009": 27001.50, "y2010": 32357.18, "var_2006_2010_pct": 75.6, "var_2009_2010_pct": 19.8},
        {"tributo": "Cide", "y2006": 0.00, "y2007": 0.02, "y2008": 0.34, "y2009": 0.28, "y2010": 114.54, "var_2006_2010_pct": None, "var_2009_2010_pct": 41342.8},
        {"tributo": "Total", "y2006": 65397.52, "y2007": 75744.66, "y2008": 86571.82, "y2009": 89524.56, "y2010": 105843.31, "var_2006_2010_pct": 61.8, "var_2009_2010_pct": 18.2},
    ]


def renuncia_projetada_vs_estimada() -> list[dict]:
    return [
        {"ano": 2005, "projetado": 31288.20, "estimado": 41010.69, "variacao_pct": -23.71},
        {"ano": 2006, "projetado": 42499.55, "estimado": 65397.52, "variacao_pct": -35.01},
        {"ano": 2007, "projetado": 52739.77, "estimado": 75744.66, "variacao_pct": -30.37},
        {"ano": 2008, "projetado": 76055.96, "estimado": 86571.82, "variacao_pct": -12.15},
        {"ano": 2009, "projetado": 86245.02, "estimado": 89524.56, "variacao_pct": -3.66},
    ]


def carga_vs_renuncia_pib() -> list[dict]:
    return [
        {"ano": 2006, "carga_trib_federal_pct_pib": 23.59, "renuncia_trib_pct_pib": 2.80, "nota": "estimado"},
        {"ano": 2007, "carga_trib_federal_pct_pib": 23.76, "renuncia_trib_pct_pib": 2.96, "nota": "estimado"},
        {"ano": 2008, "carga_trib_federal_pct_pib": 23.76, "renuncia_trib_pct_pib": 2.88, "nota": "estimado"},
        {"ano": 2009, "carga_trib_federal_pct_pib": 22.80, "renuncia_trib_pct_pib": 2.81, "nota": "estimado"},
        {"ano": 2010, "carga_trib_federal_pct_pib": 23.18, "renuncia_trib_pct_pib": 2.88, "nota": "projetado"},
    ]


def renuncia_previdenciaria() -> list[dict]:
    return [
        {"item": "Simples — contribuição patronal diferenciada", "y2006": 6143.2, "y2007": 6880.3, "y2008": 7965.3, "y2009": 8723.3, "y2010": 9850.2, "var_2006_2010_pct": 60.3},
        {"item": "Entidades filantrópicas — isenção patronal", "y2006": 3831.8, "y2007": 4409.8, "y2008": 4983.5, "y2009": 5703.3, "y2010": 6440.0, "var_2006_2010_pct": 68.1},
        {"item": "Exportação produto rural — isenção sobre receitas", "y2006": 1853.0, "y2007": 2225.6, "y2008": 2577.8, "y2009": 2557.3, "y2010": 2887.7, "var_2006_2010_pct": 55.8},
        {"item": "Empregados — redução de alíquotas (CPMF)", "y2006": 461.7, "y2007": 528.0, "y2008": 0.0, "y2009": 0.0, "y2010": 0.0, "var_2006_2010_pct": None},
        {"item": "TIC — redução patronal e terceiros", "y2006": 0.0, "y2007": 0.0, "y2008": 31.3, "y2009": 60.3, "y2010": 68.1, "var_2006_2010_pct": None},
        {"item": "Total", "y2006": 12289.7, "y2007": 14043.7, "y2008": 15558.0, "y2009": 17044.3, "y2010": 19246.0, "var_2006_2010_pct": 56.6},
    ]


def beneficios_fin_cred() -> list[dict]:
    """Demonstrativo SPE/MF 2009–2010 (R$ milhões). Totais oficiais do TCU."""
    return [
        {"grupo": "Agropecuários", "item": "AGF e estoques estratégicos", "y2009": 2411.39, "y2010": 1115.62, "var_pct": -53.74},
        {"grupo": "Agropecuários", "item": "Garantia e sustentação de preços", "y2009": 925.65, "y2010": 1359.02, "var_pct": 46.82},
        {"grupo": "Agropecuários", "item": "Custeio agropecuário", "y2009": 197.96, "y2010": 488.84, "var_pct": 146.94},
        {"grupo": "Agropecuários", "item": "EGF", "y2009": 11.51, "y2010": 88.68, "var_pct": 670.66},
        {"grupo": "Agropecuários", "item": "PRONAF/PGPAF — equalização", "y2009": 767.87, "y2010": 323.86, "var_pct": -57.82},
        {"grupo": "Agropecuários", "item": "PRONAF/PGPAF — financiamento", "y2009": 136.76, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Agropecuários", "item": "Recoop — equalização (investimento)", "y2009": 0.95, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Agropecuários", "item": "Recoop — financiamento", "y2009": 13.63, "y2010": 10.11, "var_pct": -25.80},
        {"grupo": "Agropecuários", "item": "PESA", "y2009": 247.41, "y2010": 297.58, "var_pct": 20.28},
        {"grupo": "Agropecuários", "item": "Subvenção ao prêmio do seguro rural (Lei 10.823/2003)", "y2009": 178.51, "y2010": 198.28, "var_pct": 11.08},
        {"grupo": "Agropecuários", "item": "Funcafé — equalização", "y2009": 52.95, "y2010": 82.72, "var_pct": 56.21},
        {"grupo": "Agropecuários", "item": "Funcafé — financiamento", "y2009": 460.41, "y2010": 329.89, "var_pct": -28.35},
        {"grupo": "Agropecuários", "item": "Lavoura cacaueira baiana — equalização", "y2009": 0.84, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Agropecuários", "item": "Lavoura cacaueira baiana — financiamento", "y2009": 0.77, "y2010": 0.83, "var_pct": 7.39},
        {"grupo": "Setor produtivo", "item": "FND", "y2009": 395.51, "y2010": 0.81, "var_pct": -99.79},
        {"grupo": "Setor produtivo", "item": "Fundos constitucionais FNE, FNO e FCO", "y2009": 5364.76, "y2010": 6192.91, "var_pct": 15.44},
        {"grupo": "Setor produtivo", "item": "Investimentos Centro-Oeste (equalização FAT)", "y2009": 1.10, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Setor produtivo", "item": "Fundo da Marinha Mercante", "y2009": 3.67, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Setor produtivo", "item": "Proer", "y2009": 2925.89, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Setor produtivo", "item": "FGPC", "y2009": 0.00, "y2010": 6.88, "var_pct": None},
        {"grupo": "Setor produtivo", "item": "FGE", "y2009": 0.00, "y2010": 779.26, "var_pct": None},
        {"grupo": "Setor produtivo", "item": "Revitaliza", "y2009": 58.67, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Setor produtivo", "item": "Proex — equalização", "y2009": 394.30, "y2010": 231.73, "var_pct": -41.23},
        {"grupo": "Setor produtivo", "item": "FDNE", "y2009": 4.87, "y2010": 117.95, "var_pct": 2324.12},
        {"grupo": "Setor produtivo", "item": "FDA", "y2009": 58.24, "y2010": 39.88, "var_pct": -31.52},
        {"grupo": "Setor produtivo", "item": "FNDCT", "y2009": 220.63, "y2010": 169.18, "var_pct": -23.32},
        {"grupo": "Programas sociais", "item": "FRD", "y2009": 6.56, "y2010": 9.30, "var_pct": 41.81},
        {"grupo": "Programas sociais", "item": "FCVS", "y2009": 693.59, "y2010": 6497.73, "var_pct": 836.83},
        {"grupo": "Programas sociais", "item": "PSH — subsídio habitacional", "y2009": 193.32, "y2010": 0.00, "var_pct": -100.00},
        {"grupo": "Programas sociais", "item": "Fundo de Terras / Banco da Terra", "y2009": 135.91, "y2010": 319.69, "var_pct": 135.23},
        {"grupo": "Programas sociais", "item": "FIES", "y2009": 589.00, "y2010": 986.18, "var_pct": 67.43},
        {"grupo": "Programas sociais", "item": "Subvenção energia elétrica baixa renda (Lei 10.604/2002)", "y2009": 1719.97, "y2010": 1674.82, "var_pct": -2.63},
        {"grupo": "Programas sociais", "item": "Subvenção óleo diesel pesca (Lei 9.445/1997)", "y2009": 21.79, "y2010": 19.79, "var_pct": -9.17},
        {"grupo": "Total SPE", "item": "Total (metodologia SPE: negativos excluídos)", "y2009": 16901.39, "y2010": 18877.65, "var_pct": 11.69},
    ]


def pac_desoneracoes() -> list[dict]:
    return [
        {"medida": "Reajuste da tabela do Imposto de Renda", "periodo": "Pré-PAC", "projecao_2010": 13796.0},
        {"medida": "Prorrogação da depreciação acelerada", "periodo": "Pré-PAC", "projecao_2010": 196.0},
        {"medida": "Prorrogação da cumulatividade PIS/Cofins — construção civil", "periodo": "Pré-PAC", "projecao_2010": 2253.0},
        {"medida": "Lei Geral das Pequenas e Médias Empresas", "periodo": "Pré-PAC", "projecao_2010": 4500.0},
        {"medida": "Fundos de investimento em infraestrutura", "periodo": "Pós-PAC", "projecao_2010": 0.0},
        {"medida": "Desoneração de obras de infraestrutura (Reidi)", "periodo": "Pós-PAC", "projecao_2010": 370.0},
        {"medida": "Recuperação acelerada dos créditos do PIS e Cofins", "periodo": "Pós-PAC", "projecao_2010": 1943.0},
        {"medida": "Padis — semicondutores", "periodo": "Pós-PAC", "projecao_2010": 0.0},
        {"medida": "PATVD — TV digital", "periodo": "Pós-PAC", "projecao_2010": 0.0},
        {"medida": "Ampliação do benefício a microcomputadores", "periodo": "Pós-PAC", "projecao_2010": 200.0},
        {"medida": "Desoneração na compra de perfis de aço", "periodo": "Pós-PAC", "projecao_2010": 60.0},
        {"medida": "Total", "periodo": "", "projecao_2010": 23318.0},
    ]


def pac_desoneracoes_serie() -> list[dict]:
    return [
        {"ano": 2007, "desoneracoes_pac_r_bi": 7.0, "gastos_tributarios_r_bi": 76.0, "arrecadacao_rfb_r_bi": 432.0, "pib_r_bi": 2598.0, "a_sobre_b_pct": 9.8, "a_sobre_c_pct": 1.7, "a_sobre_d_pct": 0.3},
        {"ano": 2008, "desoneracoes_pac_r_bi": 13.0, "gastos_tributarios_r_bi": 87.0, "arrecadacao_rfb_r_bi": 480.0, "pib_r_bi": 3005.0, "a_sobre_b_pct": 14.6, "a_sobre_c_pct": 2.6, "a_sobre_d_pct": 0.4},
        {"ano": 2009, "desoneracoes_pac_r_bi": 20.0, "gastos_tributarios_r_bi": 90.0, "arrecadacao_rfb_r_bi": 471.0, "pib_r_bi": 3185.0, "a_sobre_b_pct": 22.4, "a_sobre_c_pct": 4.3, "a_sobre_d_pct": 0.6},
        {"ano": 2010, "desoneracoes_pac_r_bi": 23.0, "gastos_tributarios_r_bi": 106.0, "arrecadacao_rfb_r_bi": 545.0, "pib_r_bi": 3644.0, "a_sobre_b_pct": 22.0, "a_sobre_c_pct": 4.3, "a_sobre_d_pct": 0.6},
    ]


def pac_subsidios_eixo() -> list[dict]:
    return [
        {"eixo": "Infraestrutura logística", "desembolsos": 1103.10, "part_desemb_pct": 24.8, "contratacoes": 1906.35, "part_contr_pct": 38.6},
        {"eixo": "Infraestrutura em energia", "desembolsos": 3067.38, "part_desemb_pct": 69.0, "contratacoes": 2837.15, "part_contr_pct": 57.5},
        {"eixo": "Infraestrutura social e urbana", "desembolsos": 268.50, "part_desemb_pct": 6.0, "contratacoes": 188.12, "part_contr_pct": 3.8},
        {"eixo": "Infraestrutura administração pública", "desembolsos": 5.26, "part_desemb_pct": 0.1, "contratacoes": 2.21, "part_contr_pct": 0.0},
        {"eixo": "Total", "desembolsos": 4444.24, "part_desemb_pct": 100.0, "contratacoes": 4933.82, "part_contr_pct": 100.0},
    ]


def superavit_financeiro() -> list[dict]:
    return [
        {"item": "Ativo financeiro", "y2010": 549_411_832_267, "y2009": 531_258_765_890, "y2008": 994_523_972_226, "y2007": 890_496_864_188},
        {"item": "Passivo financeiro", "y2010": 191_560_330_261, "y2009": 165_906_685_728, "y2008": 751_820_866_839, "y2007": 596_944_965_879},
        {"item": "Superávit financeiro", "y2010": 357_851_502_006, "y2009": 365_352_080_162, "y2008": 242_703_105_387, "y2007": 293_551_898_309},
        {"item": "Variação do SF (exercício atual − anterior)", "y2010": -7_500_578_156, "y2009": 122_648_974_775, "y2008": -50_848_792_922, "y2007": None},
    ]


def fatores_dlsp() -> list[dict]:
    return [
        {"item": "DLSP — saldo", "y2008": 1_168_238, "pib2008": 38.53, "y2009": 1_362_711, "pib2009": 42.78, "y2010": 1_475_820, "pib2010": 40.35},
        {"item": "Variação acumulada no ano", "y2008": -43_524, "pib2008": -7.00, "y2009": 194_472, "pib2009": 4.25, "y2010": 113_109, "pib2010": -2.43},
        {"item": "NFSP", "y2008": 61_927, "pib2008": 2.04, "y2009": 106_242, "pib2009": 3.34, "y2010": 93_673, "pib2010": 2.56},
        {"item": "Primário (superávit = negativo na NFSP)", "y2008": -103_584, "pib2008": -3.42, "y2009": -64_769, "pib2009": -2.03, "y2010": -101_696, "pib2010": -2.78},
        {"item": "Juros nominais", "y2008": 165_511, "pib2008": 5.46, "y2009": 171_011, "pib2009": 5.37, "y2010": 195_369, "pib2010": 5.34},
        {"item": "Ajuste cambial", "y2008": -78_426, "pib2008": -2.59, "y2009": 80_886, "pib2009": 2.54, "y2010": 17_677, "pib2010": 0.48},
        {"item": "Reconhecimento de dívidas", "y2008": 135, "pib2008": 0.00, "y2009": -345, "pib2009": -0.01, "y2010": 2969, "pib2010": 0.08},
        {"item": "Privatizações", "y2008": -767, "pib2008": -0.03, "y2009": -3217, "pib2009": -0.10, "y2010": -2742, "pib2010": -0.07},
        {"item": "Efeito crescimento do PIB", "y2008": None, "pib2008": -5.56, "y2009": None, "pib2009": -1.85, "y2010": None, "pib2010": -5.52},
        {"item": "PIB 12 meses (valores correntes)", "y2008": 3_031_864, "pib2008": None, "y2009": 3_185_125, "pib2009": None, "y2010": 3_657_366, "pib2010": None},
    ]


def fatores_base_monetaria() -> list[dict]:
    """Quadro p. 35 — Fatores Condicionantes da Base Monetária – 2002 a 2010.

    O título do TCU cita 2002–2010 e o texto fala em “últimos sete exercícios”,
    mas a tabela impressa começa em 2003. Valores em R$ milhões.
    Convenção Bacen/TCU: (+) expansão da base / (−) retração.
    Identidade: tesouro + titulos + setor_externo + demais = var_base.
    """
    return [
        {"ano": 2003, "tesouro_nacional": -1_064, "titulos_publicos": 11_181, "setor_externo": 643, "demais_operacoes": -10_843, "var_base": -83},
        {"ano": 2004, "tesouro_nacional": -42_140, "titulos_publicos": 52_111, "setor_externo": 12_599, "demais_operacoes": -8_468, "var_base": 14_102},
        {"ano": 2005, "tesouro_nacional": -43_008, "titulos_publicos": 2_808, "setor_externo": 52_395, "demais_operacoes": 319, "var_base": 12_514},
        {"ano": 2006, "tesouro_nacional": -59_511, "titulos_publicos": -687, "setor_externo": 74_369, "demais_operacoes": 5_683, "var_base": 19_854},
        {"ano": 2007, "tesouro_nacional": -55_600, "titulos_publicos": -73_974, "setor_externo": 155_390, "demais_operacoes": -300, "var_base": 25_516},
        {"ano": 2008, "tesouro_nacional": -74_312, "titulos_publicos": 34_059, "setor_externo": -12_124, "demais_operacoes": 53_311, "var_base": 933},
        {"ano": 2009, "tesouro_nacional": -52_312, "titulos_publicos": 11_281, "setor_externo": 62_937, "demais_operacoes": -3_383, "var_base": 18_523},
        {"ano": 2010, "tesouro_nacional": -51_204, "titulos_publicos": 249_513, "setor_externo": 75_553, "demais_operacoes": -233_082, "var_base": 40_780},
    ]


def fatores_base_monetaria_detalhe_2009_2010() -> list[dict]:
    """Quadro p. 34 — detalhe 2009–2010 (R$ milhões)."""
    return [
        {
            "ano": 2009,
            "tesouro_nacional": -52_312,
            "titulos_publicos": 11_281,
            "setor_externo": 62_937,
            "depositos_inst_financ": -3_425,
            "derivativos_ajustes": -3_199,
            "outras_contas_ajustes": 3_242,
            "var_base": 18_523,
        },
        {
            "ano": 2010,
            "tesouro_nacional": -51_204,
            "titulos_publicos": 249_513,
            "setor_externo": 75_553,
            "depositos_inst_financ": -236_911,
            "derivativos_ajustes": -1,
            "outras_contas_ajustes": 3_830,
            "var_base": 40_780,
        },
    ]


def selic_copom_decisoes() -> list[dict]:
    """Mudanças da meta Selic (Copom / Bacen SGS 432), dez/2002 a mar/2011.

    `reuniao` é a data da reunião do Copom. `vigencia` é o primeiro dia da
    nova meta na série diária SGS 432 (em geral o dia útil seguinte).
    Inclui a taxa herdada de 18/12/2002 (25,00%), vigente em 1º/1/2003.

    O gráfico da p. 33 do TCU cobre jan/2006–mar/2011; a série abaixo
    cobre também 2003–2005, para cotejar com o quadro da p. 35.
    """
    return [
        {"reuniao": date(2002, 12, 18), "vigencia": date(2002, 12, 19), "selic": 25.00, "delta_pp": 3.00},
        {"reuniao": date(2003, 1, 22), "vigencia": date(2003, 1, 23), "selic": 25.50, "delta_pp": 0.50},
        {"reuniao": date(2003, 2, 19), "vigencia": date(2003, 2, 20), "selic": 26.50, "delta_pp": 1.00},
        {"reuniao": date(2003, 6, 18), "vigencia": date(2003, 6, 19), "selic": 26.00, "delta_pp": -0.50},
        {"reuniao": date(2003, 7, 23), "vigencia": date(2003, 7, 24), "selic": 24.50, "delta_pp": -1.50},
        {"reuniao": date(2003, 8, 20), "vigencia": date(2003, 8, 21), "selic": 22.00, "delta_pp": -2.50},
        {"reuniao": date(2003, 9, 17), "vigencia": date(2003, 9, 18), "selic": 20.00, "delta_pp": -2.00},
        {"reuniao": date(2003, 10, 22), "vigencia": date(2003, 10, 23), "selic": 19.00, "delta_pp": -1.00},
        {"reuniao": date(2003, 11, 19), "vigencia": date(2003, 11, 20), "selic": 17.50, "delta_pp": -1.50},
        {"reuniao": date(2003, 12, 17), "vigencia": date(2003, 12, 18), "selic": 16.50, "delta_pp": -1.00},
        {"reuniao": date(2004, 4, 14), "vigencia": date(2004, 4, 15), "selic": 16.00, "delta_pp": -0.50},
        {"reuniao": date(2004, 9, 15), "vigencia": date(2004, 9, 16), "selic": 16.25, "delta_pp": 0.25},
        {"reuniao": date(2004, 10, 20), "vigencia": date(2004, 10, 21), "selic": 16.75, "delta_pp": 0.50},
        {"reuniao": date(2004, 11, 17), "vigencia": date(2004, 11, 18), "selic": 17.25, "delta_pp": 0.50},
        {"reuniao": date(2004, 12, 15), "vigencia": date(2004, 12, 16), "selic": 17.75, "delta_pp": 0.50},
        {"reuniao": date(2005, 1, 19), "vigencia": date(2005, 1, 20), "selic": 18.25, "delta_pp": 0.50},
        {"reuniao": date(2005, 2, 16), "vigencia": date(2005, 2, 17), "selic": 18.75, "delta_pp": 0.50},
        {"reuniao": date(2005, 3, 16), "vigencia": date(2005, 3, 17), "selic": 19.25, "delta_pp": 0.50},
        {"reuniao": date(2005, 4, 20), "vigencia": date(2005, 4, 21), "selic": 19.50, "delta_pp": 0.25},
        {"reuniao": date(2005, 5, 18), "vigencia": date(2005, 5, 19), "selic": 19.75, "delta_pp": 0.25},
        {"reuniao": date(2005, 9, 14), "vigencia": date(2005, 9, 15), "selic": 19.50, "delta_pp": -0.25},
        {"reuniao": date(2005, 10, 19), "vigencia": date(2005, 10, 20), "selic": 19.00, "delta_pp": -0.50},
        {"reuniao": date(2005, 11, 23), "vigencia": date(2005, 11, 24), "selic": 18.50, "delta_pp": -0.50},
        {"reuniao": date(2005, 12, 14), "vigencia": date(2005, 12, 15), "selic": 18.00, "delta_pp": -0.50},
        {"reuniao": date(2006, 1, 18), "vigencia": date(2006, 1, 19), "selic": 17.25, "delta_pp": -0.75},
        {"reuniao": date(2006, 3, 8), "vigencia": date(2006, 3, 9), "selic": 16.50, "delta_pp": -0.75},
        {"reuniao": date(2006, 4, 19), "vigencia": date(2006, 4, 20), "selic": 15.75, "delta_pp": -0.75},
        {"reuniao": date(2006, 5, 31), "vigencia": date(2006, 6, 1), "selic": 15.25, "delta_pp": -0.50},
        {"reuniao": date(2006, 7, 19), "vigencia": date(2006, 7, 20), "selic": 14.75, "delta_pp": -0.50},
        {"reuniao": date(2006, 8, 30), "vigencia": date(2006, 8, 31), "selic": 14.25, "delta_pp": -0.50},
        {"reuniao": date(2006, 10, 18), "vigencia": date(2006, 10, 19), "selic": 13.75, "delta_pp": -0.50},
        {"reuniao": date(2006, 11, 29), "vigencia": date(2006, 11, 30), "selic": 13.25, "delta_pp": -0.50},
        {"reuniao": date(2007, 1, 24), "vigencia": date(2007, 1, 25), "selic": 13.00, "delta_pp": -0.25},
        {"reuniao": date(2007, 3, 7), "vigencia": date(2007, 3, 8), "selic": 12.75, "delta_pp": -0.25},
        {"reuniao": date(2007, 4, 18), "vigencia": date(2007, 4, 19), "selic": 12.50, "delta_pp": -0.25},
        {"reuniao": date(2007, 6, 6), "vigencia": date(2007, 6, 7), "selic": 12.00, "delta_pp": -0.50},
        {"reuniao": date(2007, 7, 18), "vigencia": date(2007, 7, 19), "selic": 11.50, "delta_pp": -0.50},
        {"reuniao": date(2007, 9, 5), "vigencia": date(2007, 9, 6), "selic": 11.25, "delta_pp": -0.25},
        {"reuniao": date(2008, 4, 16), "vigencia": date(2008, 4, 17), "selic": 11.75, "delta_pp": 0.50},
        {"reuniao": date(2008, 6, 4), "vigencia": date(2008, 6, 5), "selic": 12.25, "delta_pp": 0.50},
        {"reuniao": date(2008, 7, 23), "vigencia": date(2008, 7, 24), "selic": 13.00, "delta_pp": 0.75},
        {"reuniao": date(2008, 9, 10), "vigencia": date(2008, 9, 11), "selic": 13.75, "delta_pp": 0.75},
        {"reuniao": date(2009, 1, 21), "vigencia": date(2009, 1, 22), "selic": 12.75, "delta_pp": -1.00},
        {"reuniao": date(2009, 3, 11), "vigencia": date(2009, 3, 12), "selic": 11.25, "delta_pp": -1.50},
        {"reuniao": date(2009, 4, 29), "vigencia": date(2009, 4, 30), "selic": 10.25, "delta_pp": -1.00},
        {"reuniao": date(2009, 6, 10), "vigencia": date(2009, 6, 11), "selic": 9.25, "delta_pp": -1.00},
        {"reuniao": date(2009, 7, 22), "vigencia": date(2009, 7, 23), "selic": 8.75, "delta_pp": -0.50},
        {"reuniao": date(2010, 4, 28), "vigencia": date(2010, 4, 29), "selic": 9.50, "delta_pp": 0.75},
        {"reuniao": date(2010, 6, 9), "vigencia": date(2010, 6, 10), "selic": 10.25, "delta_pp": 0.75},
        {"reuniao": date(2010, 7, 21), "vigencia": date(2010, 7, 22), "selic": 10.75, "delta_pp": 0.50},
        {"reuniao": date(2011, 1, 19), "vigencia": date(2011, 1, 20), "selic": 11.25, "delta_pp": 0.50},
        {"reuniao": date(2011, 3, 2), "vigencia": date(2011, 3, 3), "selic": 11.75, "delta_pp": 0.50},
    ]


def selic_na_data(dia: date, decisoes: list[dict] | None = None) -> float:
    """Meta Selic vigente em `dia` (SGS 432)."""
    serie = decisoes if decisoes is not None else selic_copom_decisoes()
    meta = None
    for dec in serie:
        if dec["vigencia"] <= dia:
            meta = float(dec["selic"])
        else:
            break
    if meta is None:
        raise ValueError(f"Sem meta Selic para {dia.isoformat()}")
    return meta


def selic_anual(ano_ini: int = 2003, ano_fim: int = 2010) -> list[dict]:
    """Selic no 1º e no último dia do ano, média ponderada por dias e extremos."""
    serie = selic_copom_decisoes()
    rows = []
    for ano in range(ano_ini, ano_fim + 1):
        d0 = date(ano, 1, 1)
        d1 = date(ano, 12, 31)
        ini = selic_na_data(d0, serie)
        fim = selic_na_data(d1, serie)
        total = 0.0
        n = 0
        mx = ini
        mn = ini
        dia = d0
        while dia <= d1:
            v = selic_na_data(dia, serie)
            total += v
            n += 1
            if v > mx:
                mx = v
            if v < mn:
                mn = v
            dia += timedelta(days=1)
        delta = round(fim - ini, 2)
        if delta < -0.01:
            sentido = "queda"
        elif delta > 0.01:
            sentido = "alta"
        else:
            sentido = "estável"
        rows.append(
            {
                "ano": ano,
                "selic_ini": ini,
                "selic_fim": fim,
                "selic_media": round(total / n, 2),
                "selic_max": mx,
                "selic_min": mn,
                "delta_pp": delta,
                "sentido": sentido,
            }
        )
    return rows


def tcu_vs_oficial_selic_p33() -> list[dict]:
    """Datas que o TCU lê no gráfico da p. 33 versus o calendário oficial do Copom."""
    return [
        {
            "evento": "Piso de 8,75% a.a.",
            "leitura_tcu_p33": "vigente até a reunião de 22/7/2009",
            "oficial_copom": "corte para 8,75% na reunião de 22/7/2009, vigente em 23/7/2009; piso até 28/4/2010",
        },
        {
            "evento": "Selic em 10,25% a.a.",
            "leitura_tcu_p33": "aumentou para 10,25% apenas em 1º/9/2010",
            "oficial_copom": "10,25% vigente desde 10/6/2010 (reunião de 9/6); em 1º/9/2010 a meta já era 10,75%",
        },
        {
            "evento": "Alta de 0,50 pp para 11,25%",
            "leitura_tcu_p33": "reunião de 19/1/2011",
            "oficial_copom": "reunião de 19/1/2011, vigente em 20/1/2011",
        },
        {
            "evento": "Janela do gráfico da p. 33",
            "leitura_tcu_p33": "janeiro de 2006 a março de 2011",
            "oficial_copom": "o quadro da p. 35 começa em 2003; o cotejamento usa a série oficial 2003–2010 e marca a janela do gráfico",
        },
    ]


def reservas_internacionais_liquidez() -> list[dict]:
    """Estoque em dezembro — reservas internacionais, conceito de liquidez.

    Fonte do gráfico TCU p. 43 (Bacen, Indicadores Econômicos).
    Série oficial: Bacen SGS 3546, US$ milhões. PTAX venda no último dia
    útil do ano: SGS 1.
    """
    return [
        {"ano": 2002, "reservas_usd_mi": 37_823, "ptax_fim": 3.5333},
        {"ano": 2003, "reservas_usd_mi": 49_296, "ptax_fim": 2.8892},
        {"ano": 2004, "reservas_usd_mi": 52_935, "ptax_fim": 2.6544},
        {"ano": 2005, "reservas_usd_mi": 53_799, "ptax_fim": 2.3407},
        {"ano": 2006, "reservas_usd_mi": 85_839, "ptax_fim": 2.1380},
        {"ano": 2007, "reservas_usd_mi": 180_334, "ptax_fim": 1.7713},
        {"ano": 2008, "reservas_usd_mi": 193_783, "ptax_fim": 2.3370},
        {"ano": 2009, "reservas_usd_mi": 238_520, "ptax_fim": 1.7412},
        {"ano": 2010, "reservas_usd_mi": 288_575, "ptax_fim": 1.6662},
    ]


def agregados_monetarios_dezembro() -> list[dict]:
    """Saldos de fim de período, dezembro de 2002 a 2010 (R$ milhões).

    M1–M4: Bacen SGS 27791, 27810, 27813 e 27815 (metodologia da Nota de
    Estatísticas Monetárias de ago/2018, série retroagida a dez/2001).
    Unidade original: R$ mil; valores convertidos para R$ milhões.

    Definição vigente: M1 = PMPP + depósitos à vista; M2 = M1 + poupança
    e títulos privados de instituições depositárias; M3 = M2 + quotas de
    fundos depositários e operações compromissadas; M4 = M3 + títulos
    públicos federais em poder do público.

    Base monetária: SGS 1782 (fecha com as variações anuais da p. 35).
    M1 restrito contemporâneo (PMPP + depósitos à vista, série 1785):
    o M1 que o TCU/Bacen usavam em 2010, antes da revisão de 2018.
    """
    return [
        {"ano": 2002, "base": 73_302, "m1_restrito": 69_901, "m1": 110_334, "m2": 400_342, "m3": 691_108, "m4": 800_036},
        {"ano": 2003, "base": 73_219, "m1_restrito": 70_802, "m1": 112_766, "m2": 416_248, "m3": 841_738, "m4": 953_805},
        {"ano": 2004, "base": 88_733, "m1_restrito": 87_344, "m1": 129_660, "m2": 495_043, "m3": 990_548, "m4": 1_104_058},
        {"ano": 2005, "base": 101_247, "m1_restrito": 98_306, "m1": 146_745, "m2": 580_696, "m3": 1_166_453, "m4": 1_298_801},
        {"ano": 2006, "base": 121_102, "m1_restrito": 118_304, "m1": 176_890, "m2": 657_427, "m3": 1_365_958, "m4": 1_506_072},
        {"ano": 2007, "base": 146_617, "m1_restrito": 143_642, "m1": 235_075, "m2": 779_566, "m3": 1_600_006, "m4": 1_827_748},
        {"ano": 2008, "base": 147_550, "m1_restrito": 145_742, "m1": 227_167, "m2": 1_086_785, "m3": 1_894_809, "m4": 2_165_027},
        {"ano": 2009, "base": 166_073, "m1_restrito": 167_400, "m1": 254_714, "m2": 1_185_866, "m3": 2_196_530, "m4": 2_497_479},
        {"ano": 2010, "base": 206_853, "m1_restrito": 197_388, "m1": 287_739, "m2": 1_387_912, "m3": 2_681_421, "m4": 2_976_783},
    ]


def operacoes_compromissadas_prazos() -> list[dict]:
    """Quadro p. 36 — prazos de vencimento das compromissadas (R$ milhões)."""
    return [
        {"periodo": "dez/06", "curtissimo": 5_800, "ate_3_meses": None, "acima_3_meses": 54_231, "total": 60_030},
        {"periodo": "jun/07", "curtissimo": 10_198, "ate_3_meses": 48_767, "acima_3_meses": 77_795, "total": 136_760},
        {"periodo": "dez/07", "curtissimo": -1_460, "ate_3_meses": 82_781, "acima_3_meses": 84_493, "total": 165_813},
        {"periodo": "jun/08", "curtissimo": 20_348, "ate_3_meses": 138_195, "acima_3_meses": 74_586, "total": 233_129},
        {"periodo": "dez/08", "curtissimo": 75_834, "ate_3_meses": 180_666, "acima_3_meses": 43_990, "total": 300_491},
        {"periodo": "jun/09", "curtissimo": 56_029, "ate_3_meses": 272_136, "acima_3_meses": 57_458, "total": 385_624},
        {"periodo": "dez/09", "curtissimo": 31_846, "ate_3_meses": 316_634, "acima_3_meses": 79_394, "total": 427_874},
        {"periodo": "jun/10", "curtissimo": 25_853, "ate_3_meses": 231_049, "acima_3_meses": 93_827, "total": 350_729},
        {"periodo": "dez/10", "curtissimo": None, "ate_3_meses": 116_509, "acima_3_meses": 142_739, "total": 259_248},
    ]
