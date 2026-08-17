"""Séries oficiais para o debate Selic × balanço de pagamentos, 1995–2016.

Fontes: Bacen SGS 22710 (balança comercial, US$ milhões), 22701 (transações
correntes, US$ milhões), 3546 (reservas, liquidez), 4390 (Selic % a.m.),
433 (IPCA % a.m.). Acumulados 2003–2016 = produto dos 168 meses.
"""

from __future__ import annotations

# SGS 22710 — saldo anual da balança comercial (bens), US$ milhões
BALANCA_COMERCIAL = [
    {"ano": 1995, "usd_mi": -4_481},
    {"ano": 1996, "usd_mi": -6_601},
    {"ano": 1997, "usd_mi": -7_590},
    {"ano": 1998, "usd_mi": -7_596},
    {"ano": 1999, "usd_mi": -2_661},
    {"ano": 2000, "usd_mi": -2_169},
    {"ano": 2001, "usd_mi": 1_224},
    {"ano": 2002, "usd_mi": 11_667},
    {"ano": 2003, "usd_mi": 23_341},
    {"ano": 2004, "usd_mi": 32_106},
    {"ano": 2005, "usd_mi": 43_542},
    {"ano": 2006, "usd_mi": 44_519},
    {"ano": 2007, "usd_mi": 37_689},
    {"ano": 2008, "usd_mi": 22_040},
    {"ano": 2009, "usd_mi": 22_776},
    {"ano": 2010, "usd_mi": 16_462},
    {"ano": 2011, "usd_mi": 25_123},
    {"ano": 2012, "usd_mi": 14_157},
    {"ano": 2013, "usd_mi": -2_390},
    {"ano": 2014, "usd_mi": -9_148},
    {"ano": 2015, "usd_mi": 15_022},
    {"ano": 2016, "usd_mi": 41_544},
]

# SGS 22701 — transações correntes, US$ milhões
TRANSACOES_CORRENTES = [
    {"ano": 1999, "usd_mi": -26_784},
    {"ano": 2000, "usd_mi": -26_531},
    {"ano": 2001, "usd_mi": -24_890},
    {"ano": 2002, "usd_mi": -9_407},
    {"ano": 2003, "usd_mi": 2_193},
    {"ano": 2004, "usd_mi": 8_959},
    {"ano": 2005, "usd_mi": 11_679},
    {"ano": 2006, "usd_mi": 10_774},
    {"ano": 2007, "usd_mi": -2_754},
    {"ano": 2008, "usd_mi": -35_602},
    {"ano": 2009, "usd_mi": -29_328},
    {"ano": 2010, "usd_mi": -86_718},
    {"ano": 2011, "usd_mi": -83_576},
    {"ano": 2012, "usd_mi": -92_678},
    {"ano": 2013, "usd_mi": -88_384},
    {"ano": 2014, "usd_mi": -110_494},
    {"ano": 2015, "usd_mi": -63_409},
    {"ano": 2016, "usd_mi": -30_529},
]

# SGS 3546 — reservas, conceito de liquidez, dezembro, US$ milhões
RESERVAS_DEZEMBRO = [
    {"ano": 1999, "usd_mi": 36_342},
    {"ano": 2000, "usd_mi": 33_011},
    {"ano": 2001, "usd_mi": 35_866},
    {"ano": 2002, "usd_mi": 37_823},
    {"ano": 2003, "usd_mi": 49_296},
    {"ano": 2004, "usd_mi": 52_935},
    {"ano": 2005, "usd_mi": 53_799},
    {"ano": 2006, "usd_mi": 85_839},
    {"ano": 2007, "usd_mi": 180_334},
    {"ano": 2008, "usd_mi": 193_783},
    {"ano": 2009, "usd_mi": 238_520},
    {"ano": 2010, "usd_mi": 288_575},
    {"ano": 2011, "usd_mi": 352_012},
    {"ano": 2012, "usd_mi": 373_147},
    {"ano": 2013, "usd_mi": 358_808},
    {"ano": 2014, "usd_mi": 363_551},
    {"ano": 2015, "usd_mi": 356_464},
    {"ano": 2016, "usd_mi": 365_016},
]

# Selic efetiva anual (SGS 4390, produto dos 12 meses) e IPCA (SGS 433)
# Fatores acumulados desde jan/2003.
SELIC_IPCA_ANUAL = [
    {"ano": 2003, "selic_pct": 23.33, "ipca_pct": 9.30, "selic_acum_pct": 23.33, "ipca_acum_pct": 9.30},
    {"ano": 2004, "selic_pct": 16.24, "ipca_pct": 7.60, "selic_acum_pct": 43.36, "ipca_acum_pct": 17.61},
    {"ano": 2005, "selic_pct": 19.04, "ipca_pct": 5.69, "selic_acum_pct": 70.66, "ipca_acum_pct": 24.30},
    {"ano": 2006, "selic_pct": 15.08, "ipca_pct": 3.14, "selic_acum_pct": 96.40, "ipca_acum_pct": 28.20},
    {"ano": 2007, "selic_pct": 11.84, "ipca_pct": 4.46, "selic_acum_pct": 119.66, "ipca_acum_pct": 33.92},
    {"ano": 2008, "selic_pct": 12.48, "ipca_pct": 5.90, "selic_acum_pct": 147.08, "ipca_acum_pct": 41.82},
    {"ano": 2009, "selic_pct": 9.92, "ipca_pct": 4.31, "selic_acum_pct": 171.60, "ipca_acum_pct": 47.94},
    {"ano": 2010, "selic_pct": 9.78, "ipca_pct": 5.91, "selic_acum_pct": 198.17, "ipca_acum_pct": 56.68},
    {"ano": 2011, "selic_pct": 11.61, "ipca_pct": 6.50, "selic_acum_pct": 232.81, "ipca_acum_pct": 66.87},
    {"ano": 2012, "selic_pct": 8.48, "ipca_pct": 5.84, "selic_acum_pct": 261.03, "ipca_acum_pct": 76.61},
    {"ano": 2013, "selic_pct": 8.21, "ipca_pct": 5.91, "selic_acum_pct": 290.69, "ipca_acum_pct": 87.05},
    {"ano": 2014, "selic_pct": 10.91, "ipca_pct": 6.41, "selic_acum_pct": 333.31, "ipca_acum_pct": 99.04},
    {"ano": 2015, "selic_pct": 13.29, "ipca_pct": 10.67, "selic_acum_pct": 390.88, "ipca_acum_pct": 120.28},
    {"ano": 2016, "selic_pct": 14.03, "ipca_pct": 6.29, "selic_acum_pct": 459.74, "ipca_acum_pct": 134.13},
]

# Comparação Banco Mundial 2003–2016 (produto das taxas anuais)
# FR.INR.LEND = taxa de empréstimo (não é a taxa básica).
# FP.CPI.TOTL.ZG = inflação ao consumidor.
COMPARACAO_BM = [
    {"pais": "Brasil", "iso": "BRA", "emprestimo_acum_pct": 18_717.1, "emprestimo_geo_pct": 45.37, "cpi_acum_pct": 147.2, "cpi_geo_pct": 6.68},
    {"pais": "China", "iso": "CHN", "emprestimo_acum_pct": 116.3, "emprestimo_geo_pct": 5.67, "cpi_acum_pct": 44.8, "cpi_geo_pct": 2.68},
    {"pais": "Índia", "iso": "IND", "emprestimo_acum_pct": 323.6, "emprestimo_geo_pct": 10.86, "cpi_acum_pct": 161.9, "cpi_geo_pct": 7.12},
    {"pais": "Romênia", "iso": "ROU", "emprestimo_acum_pct": 532.5, "emprestimo_geo_pct": 14.08, "cpi_acum_pct": 113.3, "cpi_geo_pct": 5.56},
]


def balanca_comercial() -> list[dict]:
    return list(BALANCA_COMERCIAL)


def transacoes_correntes() -> list[dict]:
    return list(TRANSACOES_CORRENTES)


def reservas_dezembro() -> list[dict]:
    return list(RESERVAS_DEZEMBRO)


def selic_ipca_anual() -> list[dict]:
    return list(SELIC_IPCA_ANUAL)


def comparacao_banco_mundial() -> list[dict]:
    return list(COMPARACAO_BM)
