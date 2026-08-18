"""Séries oficiais da simulação Conta Única × fatores da DBGG, 2003–2015.

Unidades: R$ milhões (correntes), salvo indicação.

Fontes
------
1. Indiretas BNDES — portal de dados abertos, *Desembolsos Mensais*
   (resource 179950b8-…), agregado ``forma_de_apoio = INDIRETA``.
   Não usar o CSV “indiretas e produto”: ele zera Finame/Automático.

2. Participações acionárias — BNDESPAR, *Desembolsos via Renda Variável*
   (resource d446ae3e-…), ``tipo_de_ativo = PARTICIPAÇÃO ACIONÁRIA``.
   A base começa em 2007; 2003–2006 ficam sem valor.

3. Renúncias — âncora IFI/RFB 2015 (Nota Técnica nº 17, jun/2018, Tabela 3,
   bases efetivas) e totais TCU/IFI. A série anual 2003–2014 é reconstruída
   por participação constante no PIB de 2015 (ver ``renuncia_reconstruida``).

4. DBGG — Bacen SGS 13761 (R$ milhões) e 13762 (% PIB), metodologia 2008,
   dezembro. Começa em 2006.

5. PIB 12 meses — SGS 4382, dezembro (denominador da DBGG).

6. Selic efetiva anual — SGS 4390, produto dos 12 meses
   (mesmos fatores de ``selic_bp_2003_2016_dados``).

7. Captações Tesouro→BNDES — página oficial do BNDES
   “Recursos financeiros captados junto ao Tesouro Nacional” (R$ bilhões).
"""

from __future__ import annotations

ANOS = list(range(2003, 2016))

# Desembolsos mensais BNDES, forma INDIRETA, R$ milhões
INDIRETAS_R_MI = {
    2003: 14_546.0,
    2004: 19_854.0,
    2005: 21_009.0,
    2006: 25_307.0,
    2007: 32_810.0,
    2008: 42_779.0,
    2009: 50_409.0,
    2010: 81_538.0,
    2011: 69_353.0,
    2012: 67_109.0,
    2013: 97_222.0,
    2014: 85_493.0,
    2015: 50_946.0,
}

# Participações acionárias (não inclui debêntures nem cotas de fundo)
PARTICIPACOES_R_MI = {
    2003: None,
    2004: None,
    2005: None,
    2006: None,
    2007: 1_962.548,
    2008: 7_901.455,
    2009: 2_746.931,
    2010: 25_429.091,
    2011: 727.355,
    2012: 1_603.942,
    2013: 2_083.154,
    2014: 1_292.122,
    2015: 666.136,
}

# IFI NT 17 / RFB bases efetivas 2015, R$ milhões
RENUNCIA_2015_EFETIVA = {
    "desenvolvimento_regional": 5_899.1,
    "zfm_alc": 23_231.9,
    "imunes_isentas": 19_505.1,
    "inovacao_lei_bem": 3_392.0,
    "total_gt": 270_054.3,
}

# DGT PLOA 2015 (projeção, não base efetiva) — só para cotejo
RENUNCIA_2015_PLOA = {
    "desenvolvimento_regional": 7_274.459,
    "zfm_alc": 27_811.719,
    "imunes_isentas": 22_322.614,
    "inovacao": 3_403.039,
}

# TCU CG 2010 — isenção patronal de entidades filantrópicas (subconjunto)
FILANTROPICAS_PREV_R_MI = {
    2006: 3_831.8,
    2007: 4_409.8,
    2008: 4_983.5,
    2009: 5_703.3,
    2010: 6_440.0,
}

# SGS 4382 — PIB acumulado 12 meses, dezembro, R$ milhões
PIB_DEZ_R_MI = {
    2003: 1_717_950.4,
    2004: 1_957_751.2,
    2005: 2_170_584.5,
    2006: 2_409_449.9,
    2007: 2_720_262.9,
    2008: 3_109_803.1,
    2009: 3_333_039.4,
    2010: 3_885_847.0,
    2011: 4_376_382.0,
    2012: 4_814_760.0,
    2013: 5_331_619.0,
    2014: 5_778_953.0,
    2015: 5_995_787.0,
}

# SGS 13761 — DBGG, metodologia 2008, dezembro, R$ milhões
DBGG_DEZ_R_MI = {
    2006: 1_336_644.90,
    2007: 1_542_851.83,
    2008: 1_740_887.81,
    2009: 1_973_423.68,
    2010: 2_011_521.66,
    2011: 2_243_603.72,
    2012: 2_583_946.35,
    2013: 2_747_996.71,
    2014: 3_252_448.55,
    2015: 3_927_523.06,
}

# SGS 13762 — DBGG/PIB, metodologia 2008, dezembro
DBGG_PCT_PIB = {
    2006: 55.48,
    2007: 56.72,
    2008: 55.98,
    2009: 59.21,
    2010: 51.77,
    2011: 51.27,
    2012: 53.67,
    2013: 51.54,
    2014: 56.28,
    2015: 65.50,
}

# SGS 4390 — Selic efetiva no ano (%)
SELIC_EFETIVA_PCT = {
    2003: 23.33,
    2004: 16.24,
    2005: 19.04,
    2006: 15.08,
    2007: 11.84,
    2008: 12.48,
    2009: 9.92,
    2010: 9.78,
    2011: 11.61,
    2012: 8.48,
    2013: 8.21,
    2014: 10.91,
    2015: 13.29,
}

# BNDES — captações junto ao Tesouro, R$ milhões (página oficial)
TESOURO_CAPTACAO_R_MI = {
    2003: 0.0,
    2004: 0.0,
    2005: 0.0,
    2006: 0.0,
    2007: 0.0,
    2008: 22_500.0,
    2009: 105_000.0,
    2010: 107_100.0,  # 82,4 + 24,7 (capitalização Petrobras)
    2011: 50_200.0,
    2012: 55_000.0,
    2013: 41_000.0,
    2014: 60_000.0,
    2015: 0.0,
}

TESOURO_DEVOLUCAO_R_MI = {
    2009: 4_080.0,
    2010: 10_440.0,
    2011: 14_560.0,
    2012: 13_310.0,
    2013: 14_740.0,
    2014: 6_200.0,
    2015: 24_670.0,
}

# Lei 10.973/2004 entra em 2004; o gasto tributário mensurável da inovação
# (Lei do Bem, 11.196/2005) começa a aparecer no DGT a partir de 2006.
INOVACAO_INICIO = 2006


def _share_2015(chave: str) -> float:
    return RENUNCIA_2015_EFETIVA[chave] / PIB_DEZ_R_MI[2015]


def renuncia_reconstruida(ano: int) -> dict:
    """Reconstrói as três famílias pedidas, em R$ milhões.

    Calibração: bases efetivas RFB/IFI de 2015, aplicadas como fração
    constante do PIB (SGS 4382) em cada ano. Inovação = 0 até 2005.
    Desenvolvimento regional na acepção ampla = função RFB + ZFM/ALC.
    """
    pib = PIB_DEZ_R_MI[ano]
    regional = _share_2015("desenvolvimento_regional") * pib
    zfm = _share_2015("zfm_alc") * pib
    imunes = _share_2015("imunes_isentas") * pib
    if ano < INOVACAO_INICIO:
        inovacao = 0.0
    elif ano == 2015:
        regional = RENUNCIA_2015_EFETIVA["desenvolvimento_regional"]
        zfm = RENUNCIA_2015_EFETIVA["zfm_alc"]
        imunes = RENUNCIA_2015_EFETIVA["imunes_isentas"]
        inovacao = RENUNCIA_2015_EFETIVA["inovacao_lei_bem"]
    else:
        inovacao = _share_2015("inovacao_lei_bem") * pib
    return {
        "ano": ano,
        "desenvolvimento_regional": regional,
        "zfm_alc": zfm,
        "regional_ampla": regional + zfm,
        "inovacao": inovacao,
        "imunes_isentas": imunes,
        "renuncia_pedida": regional + zfm + inovacao + imunes,
    }


def indiretas() -> list[dict]:
    return [{"ano": a, "indiretas_r_mi": INDIRETAS_R_MI[a]} for a in ANOS]


def participacoes() -> list[dict]:
    return [{"ano": a, "participacoes_r_mi": PARTICIPACOES_R_MI[a]} for a in ANOS]


def tesouro_bndes() -> list[dict]:
    return [
        {
            "ano": a,
            "captacao_r_mi": TESOURO_CAPTACAO_R_MI[a],
            "devolucao_r_mi": TESOURO_DEVOLUCAO_R_MI.get(a, 0.0),
        }
        for a in ANOS
    ]


def dbgg_oficial() -> list[dict]:
    rows = []
    for a in ANOS:
        rows.append(
            {
                "ano": a,
                "dbgg_r_mi": DBGG_DEZ_R_MI.get(a),
                "dbgg_pct_pib": DBGG_PCT_PIB.get(a),
                "pib_r_mi": PIB_DEZ_R_MI[a],
                "selic_pct": SELIC_EFETIVA_PCT[a],
            }
        )
    return rows
