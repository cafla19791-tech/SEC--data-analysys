#!/usr/bin/env python3
"""Extrai e organiza o Relatório TCU das Contas do Governo 2010.

Gera:
  - output/TCU_CG_2010.xlsx
  - output/TCU_CG_2010_RELATORIO.md

Valores de fluxo/estoque de 2010 são atualizados pelo IPCA até 30/06/2026
(mesmo critério ContAgil / OSU deste repositório):

    valor_ipca = valor_nominal × fator_IPCA(jun/2026) / fator_IPCA(dez/2010)

Uso:
  python3 scripts/build_tcu_cg_2010.py
  python3 scripts/build_tcu_cg_2010.py --ipca data/ipca_sgs433.xlsx
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.calcular_diretas_ipca_selic import (  # noqa: E402
    DATA_REF_DEFAULT,
    carregar_ipca,
    fator_ipca_entre,
)
from scripts import tcu_cg_2010_dados as D  # noqa: E402
from scripts.analisar_base_monetaria_tcu import analisar  # noqa: E402
from scripts.cotejar_selic_base_tcu import cotejar  # noqa: E402
from scripts.analisar_reservas_agregados_tcu import analisar as analisar_reservas  # noqa: E402

MES_BASE_DEFAULT = datetime(2010, 12, 1)
OUTPUT_XLSX = ROOT / "output" / "TCU_CG_2010.xlsx"
OUTPUT_MD = ROOT / "output" / "TCU_CG_2010_RELATORIO.md"


def fator_dez2010_ref(
    ipca: pd.DataFrame,
    data_ref: datetime = DATA_REF_DEFAULT,
    mes_base: datetime = MES_BASE_DEFAULT,
) -> float:
    return float(fator_ipca_entre(ipca, pd.Timestamp(mes_base), pd.Timestamp(data_ref)))


def _br_bi(valor_milhoes: float | None, casas: int = 2) -> str:
    if valor_milhoes is None or (isinstance(valor_milhoes, float) and pd.isna(valor_milhoes)):
        return "—"
    return f"R$ {valor_milhoes / 1000.0:,.{casas}f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def _br_num(valor: float | None, casas: int = 2) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def df_indicadores(fator: float) -> pd.DataFrame:
    rows = []
    for r in D.indicadores():
        v2010 = r["valor_2010"]
        v2009 = r["valor_2009"]
        rows.append(
            {
                "bloco": r["bloco"],
                "indicador": r["indicador"],
                "valor_2009_r_mi": v2009,
                "valor_2010_r_mi": v2010,
                "valor_2010_ipca_jun2026_r_mi": None if v2010 is None else v2010 * fator,
                "unidade": r["unidade"],
                "pct_pib_2010": r["pct_pib_2010"],
                "var_pct": r["var_pct"],
                "fonte": r["fonte"],
                "nota": r["nota"],
            }
        )
    return pd.DataFrame(rows)


def df_creditos(fator: float) -> pd.DataFrame:
    rows = []
    for r in D.creditos_dlsp():
        rows.append(
            {
                "item": r["item"],
                "sinal": r["sinal"],
                "valor_2009_r_mi": r["v2009"],
                "pct_pib_2009": r["pib2009"],
                "valor_2010_r_mi": r["v2010"],
                "pct_pib_2010": r["pib2010"],
                "var_r_mi": r["v2010"] - r["v2009"],
                "valor_2010_ipca_jun2026_r_mi": r["v2010"] * fator,
            }
        )
    return pd.DataFrame(rows)


def df_autorizacoes() -> pd.DataFrame:
    return pd.DataFrame(D.autorizacoes_legais())


def df_dpf() -> pd.DataFrame:
    return pd.DataFrame(D.dpf_indicadores())


def df_renuncia_regional(fator: float) -> pd.DataFrame:
    out = pd.DataFrame(D.renuncia_regional())
    for col in ("tributarios", "trib_prev", "fin_cred", "total"):
        out[f"{col}_ipca_jun2026"] = out[col] * fator
    return out


def df_renuncia_trib() -> pd.DataFrame:
    return pd.DataFrame(D.renuncia_tributaria())


def df_renuncia_prev() -> pd.DataFrame:
    return pd.DataFrame(D.renuncia_previdenciaria())


def df_beneficios(fator: float) -> pd.DataFrame:
    out = pd.DataFrame(D.beneficios_fin_cred())
    out["y2010_ipca_jun2026"] = out["y2010"] * fator
    return out


def df_pac_deson(fator: float) -> pd.DataFrame:
    out = pd.DataFrame(D.pac_desoneracoes())
    out["projecao_2010_ipca_jun2026"] = out["projecao_2010"] * fator
    return out


def df_pac_eixo(fator: float) -> pd.DataFrame:
    out = pd.DataFrame(D.pac_subsidios_eixo())
    out["desembolsos_ipca_jun2026"] = out["desembolsos"] * fator
    out["contratacoes_ipca_jun2026"] = out["contratacoes"] * fator
    return out


def df_resumo_ipca(fator: float, data_ref: datetime) -> pd.DataFrame:
    destaques = [
        ("Créditos União → BNDES (estoque dez/2010)", 236_723.0, "estoque"),
        ("Custo fiscal anual Selic−TJLP (estimativa TCU)", 14_200.0, "fluxo"),
        ("Subsídios Lei 11.948/2009 — biênio 2009/2010", 3_100.0, "fluxo"),
        ("Emissão direta ao BNDES em 2010 (PAF)", 24_800.0, "fluxo"),
        ("Carteira Lei 11.948 aplicada ao PAC", 34_700.0, "estoque_carteira"),
        ("Repasses BNDES (saldo direcionado)", 179_800.0, "estoque"),
        ("Renúncia federal total 2010", 143_970.0, "fluxo"),
        ("Benefícios tributários 2010", 105_843.31, "fluxo"),
        ("Benefícios financeiros e creditícios 2010", 18_877.65, "fluxo"),
        ("Desonerações tributárias do PAC 2010", 23_318.0, "fluxo"),
        ("Subsídios creditícios PAC — desembolsos", 4_444.24, "fluxo"),
        ("Subsídios creditícios PAC — contratações (BNDES 97,5%)", 4_933.82, "fluxo"),
        ("DBGG", 2_011_522.0, "estoque"),
        ("DLSP", 1_475_820.0, "estoque"),
        ("Juros nominais / déficit nominal", 195_369.0, "fluxo"),
        ("Superávit primário", 101_696.0, "fluxo"),
    ]
    rows = []
    for nome, nominal, tipo in destaques:
        rows.append(
            {
                "indicador": nome,
                "tipo": tipo,
                "nominal_2010_r_mi": nominal,
                "nominal_2010_r_bi": nominal / 1000.0,
                "fator_ipca": fator,
                "ipca_jun2026_r_mi": nominal * fator,
                "ipca_jun2026_r_bi": nominal * fator / 1000.0,
                "data_base": "2010-12",
                "data_ref": data_ref.strftime("%Y-%m-%d"),
            }
        )
    return pd.DataFrame(rows)


def escrever_excel(
    destino: Path,
    fator: float,
    data_ref: datetime,
    extra_sheets: dict | None = None,
) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    fonte = pd.DataFrame(
        [
            {"campo": "Documento", "valor": D.FONTE_TITULO},
            {"campo": "Órgão", "valor": D.FONTE_ORGAO},
            {"campo": "Relator", "valor": D.RELATOR},
            {"campo": "Exercício", "valor": D.EXERCICIO},
            {"campo": "Publicação", "valor": D.ANO_PUBLICACAO},
            {"campo": "URL", "valor": D.FONTE_URL},
            {"campo": "PIB 2009 (R$ milhões)", "valor": D.PIB_2009_R_MILHOES},
            {"campo": "PIB 2010 quadro DLSP (R$ milhões)", "valor": D.PIB_2010_R_MILHOES},
            {"campo": "PIB 2010 quadro PAC (R$ bilhões)", "valor": D.PIB_PAC_2010_R_BILHOES},
            {"campo": "Fator IPCA dez/2010 → " + data_ref.strftime("%b/%Y"), "valor": fator},
            {"campo": "Data de referência IPCA", "valor": data_ref.strftime("%Y-%m-%d")},
            {
                "campo": "Metodologia IPCA",
                "valor": "valor × fator(ref) / fator(dez/2010); Bacen SGS 433",
            },
        ]
    )
    abas = {
        "Fonte": fonte,
        "Indicadores": df_indicadores(fator),
        "Creditos_DLSP": df_creditos(fator),
        "Autorizacoes_Legais": df_autorizacoes(),
        "DPF": df_dpf(),
        "Fatores_DLSP": pd.DataFrame(D.fatores_dlsp()),
        "Superavit_Financeiro": pd.DataFrame(D.superavit_financeiro()),
        "Renuncia_Regional": df_renuncia_regional(fator),
        "Renuncia_Tributaria": df_renuncia_trib(),
        "Renuncia_Projetada": pd.DataFrame(D.renuncia_projetada_vs_estimada()),
        "Carga_vs_Renuncia_PIB": pd.DataFrame(D.carga_vs_renuncia_pib()),
        "Renuncia_Previdenciaria": df_renuncia_prev(),
        "Beneficios_Fin_Cred": df_beneficios(fator),
        "PAC_Desoneracoes": df_pac_deson(fator),
        "PAC_Desoneracoes_Serie": pd.DataFrame(D.pac_desoneracoes_serie()),
        "PAC_Subsidios_Eixo": df_pac_eixo(fator),
        "Resumo_IPCA": df_resumo_ipca(fator, data_ref),
    }
    if extra_sheets:
        abas.update(extra_sheets)
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        for nome, frame in abas.items():
            frame.to_excel(writer, sheet_name=nome[:31], index=False)
    return destino


def escrever_markdown(destino: Path, fator: float, data_ref: datetime) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    resumo = df_resumo_ipca(fator, data_ref)
    cred = df_creditos(fator)
    eixo = df_pac_eixo(fator)
    bndes_2010 = 236_723.0
    bndes_2009 = 129_237.0
    custo = 14_200.0
    ref = data_ref.strftime("%d/%m/%Y")

    linhas_resumo = [
        "| Indicador | Nominal 2010 | IPCA "
        + data_ref.strftime("%b/%Y")
        + " | Tipo |",
        "|---|---:|---:|---|",
    ]
    for _, r in resumo.iterrows():
        linhas_resumo.append(
            f"| {r['indicador']} | {_br_bi(r['nominal_2010_r_mi'])} | "
            f"{_br_bi(r['ipca_jun2026_r_mi'])} | {r['tipo']} |"
        )

    md = f"""# TCU — Contas do Governo da República 2010

**Fonte:** [{D.FONTE_TITULO}]({D.FONTE_URL})
**Órgão:** {D.FONTE_ORGAO} · Relator {D.RELATOR} · publicado em {D.ANO_PUBLICACAO}
**Atualização:** IPCA Bacen SGS 433, de dez/2010 até {ref} (fator **{_br_num(fator, 4)}**)

Este recorte organiza as tabelas do relatório que importam para a análise de
crédito oficial, BNDES e subsídios — o mesmo recorte do restante deste
repositório (operações 2009–2010, spread Selic − taxa do contrato, impacto
fiscal a 30/06/2026).

## Destaques

O TCU registra, no exercício de 2010, a expansão do crédito direcionado e o
aumento dos créditos da União junto ao BNDES, com custo fiscal explícito
decorrente do diferencial entre a Selic (custo de captação do Tesouro) e a
TJLP (remuneração recebida do banco).

- Créditos da União no BNDES sobem de {_br_bi(bndes_2009)} (4,06% do PIB) para
  **{_br_bi(bndes_2010)}** (6,47% do PIB) — alta de {_br_bi(bndes_2010 - bndes_2009)}.
- Sobre esse estoque, o TCU estima custo fiscal anual de **{_br_bi(custo)}**,
  com captação do Tesouro entre 10% e 12% a.a. e aplicação no BNDES entre 4% e
  6% a.a.
- Só as transferências iniciais de {_br_bi(100_000)} da Lei 11.948/2009 geram
  **{_br_bi(3_100)}** de subsídios no biênio 2009/2010.
- Repasses do BNDES no SFN crescem 43,9%, de {_br_bi(124_900)} para
  **{_br_bi(179_800)}**.
- A Lei 12.249/2010 amplia o teto de crédito ao BNDES de {_br_bi(100_000)}
  para **{_br_bi(180_000)}**. A MP 505/2010 acrescenta {_br_bi(30_000)} para a
  capitalização da Petrobras. A MP 526/2011 (já em 2011) autoriza mais
  {_br_bi(83_000)}.
- Do crédito da Lei 11.948/2009, o BNDES informa carteira de {_br_bi(163_700)},
  dos quais **{_br_bi(34_700)}** em obras do PAC (TC 022.684/2010-7).
- Renúncia federal projetada: **{_br_bi(143_970)}** (tributária {_br_bi(105_843.31)}
  + previdenciária {_br_bi(19_246)} + financeira/creditícia {_br_bi(18_877.65)}).
- Subsídios creditícios do PAC em 2010: {_br_bi(4_444.24)} sobre desembolsos e
  {_br_bi(4_933.82)} sobre contratações; **97,5%** das contratações são do BNDES
  ({_br_bi(4_800)}). Custo de oportunidade da SPE: NTN-F 2017 a **12,32%** a.a.

Em reais de {ref}:

{chr(10).join(linhas_resumo)}

## Tesouro, Bacen e BNDES

O relatório (item 2.5) trata as emissões diretas de títulos a bancos oficiais
como alavancagem do Governo Central que **não entra por completo** nos
indicadores de endividamento líquido: o Tesouro emite e, no mesmo valor, passa
a ter um crédito contra o BNDES. A DLSP quase não se move; a DBGG e o custo
fiscal, sim.

Dos {_br_bi(181_200)} de emissões diretas em 2010, apenas {_br_bi(90_500)}
aumentam o estoque da dívida (Relatório Anual da Dívida 2010, Tabela 2). As
emissões sem contrapartida financeira somam {_br_bi(89_900)}, com destaque
para Petrobras ({_br_bi(42_900)}), BNDES ({_br_bi(24_800)}) e Petros
({_br_bi(16_300)}).

Remuneração legal do Tesouro nas transferências ao BNDES (Lei 11.948/2009,
§ 5º, com redação da Lei 12.096/2009):

1. até 30% do valor, pelo custo de captação externo em dólares, no prazo do
   ressarcimento;
2. o restante, pela TJLP.

O TCU observa o efeito colateral nas demonstrações do BNDES: o resultado com
títulos e valores mobiliários foi de {_br_bi(5_200)} em 2009 e {_br_bi(8_400)}
em 2010, contra lucro líquido de {_br_bi(6_700)} e {_br_bi(9_900)}. Parte do
lucro do banco é o carregamento dos títulos recebidos do Tesouro.

Entre 2003 e 2010, segundo a Nota de Inflação do Bacen (mar/2011) citada pelo
TCU, a Selic recuou 13,5 pontos percentuais e a taxa implícita da DLSP apenas
2,6 p.p. A taxa implícita dos ativos da União caiu 7,9 p.p.; a dos passivos,
4,9 p.p. O spread fiscal das operações oficiais se alargou.

## Quadro da dívida (R$ milhões)

Itens do quadro “Dívida Líquida e Bruta do Governo Geral” (Nota de Política
Fiscal / Bacen) mais próximos do recorte BNDES:

| Item | 2009 | % PIB | 2010 | % PIB | Δ R$ mi |
|---|---:|---:|---:|---:|---:|
"""
    foco = {
        "Dívida bruta do governo geral (DBGG)",
        "Dívida Líquida do Setor Público",
        "Dívida mobiliária do Tesouro Nacional",
        "Operações compromissadas do BCB",
        "Créditos concedidos a inst. financeiras oficiais",
        "Créditos junto ao BNDES",
        "Aplicações em fundos e programas",
        "Recursos do FAT na rede bancária",
        "Equalização cambial",
    }
    for _, r in cred.iterrows():
        if r["item"] not in foco:
            continue
        md += (
            f"| {r['item']} | {_br_num(r['valor_2009_r_mi'], 0)} | "
            f"{_br_num(r['pct_pib_2009'])} | {_br_num(r['valor_2010_r_mi'], 0)} | "
            f"{_br_num(r['pct_pib_2010'])} | {_br_num(r['var_r_mi'], 0)} |\n"
        )

    md += f"""
A DLSP sobe {_br_bi(113_109)} em termos nominais e cai 2,43 p.p. do PIB, puxada
pelo crescimento do produto (−5,52 p.p.). Os juros nominais ({_br_bi(195_369)},
5,34% do PIB) superam o superávit primário ({_br_bi(101_696)}, 2,78% do PIB).
Metade dos créditos às instituições oficiais de fomento, diz o TCU, **não tem
impacto sobre o estoque líquido**.

A DPF em mercado fecha 2010 em {_br_bi(1_694_000)}, dentro do PAF
(1,60–1,73 tri). Composição: prefixado 36,6%; preços 26,6%; Selic 30,8%;
câmbio 5,1%; TR e outros 0,8%. Prazo médio: 42,0 meses.

## Política creditícia

O saldo de crédito do SFN chega a {_br_bi(1_700_000)} (46,4% do PIB; +22,9%
sobre dez/2009). Em janeiro de 2007 essa relação era 30,7% do PIB. Recursos
direcionados crescem 31,1% (28,3% no texto de consolidação), puxados por
BNDES (+43,9%) e habitação (+50,4%). Livres: {_br_bi(1_100_000)} (+16,9%).
Direcionados: {_br_bi(589_800)} (+28,3%).

## Renúncia de receitas

Total projetado de {_br_bi(143_970)} em 2010, acima da despesa liquidada de
várias funções orçamentárias. Regionalização (R$ bilhões; o total inclui
{_br_bi(3_100)} sem classificação regional):

| Região | Tributários | Prev. | Fin./créd. | Total | Part. | Per capita |
|---|---:|---:|---:|---:|---:|---:|
"""
    for r in D.renuncia_regional():
        pc = "—" if r["per_capita_r"] is None else f"R$ {_br_num(r['per_capita_r'])}"
        md += (
            f"| {r['regiao']} | {_br_num(r['tributarios'])} | {_br_num(r['trib_prev'])} | "
            f"{_br_num(r['fin_cred'])} | {_br_num(r['total'])} | "
            f"{_br_num(r['participacao_pct'], 1)}% | {pc} |\n"
        )

    md += f"""
A média nacional é R$ 754,75 por habitante. O Norte (R$ 1.397,14) está 85%
acima da média, em grande parte pela Zona Franca de Manaus (> {_br_bi(15_000)},
mais de 75% do gasto tributário da região). O Nordeste (R$ 394,61) fica em 52%
da média. Nos benefícios sociais (Assistência, Saúde, Trabalho, Educação etc.)
a ordem se inverte: Sudeste R$ 362 e Norte/Nordeste R$ 61 / R$ 76, contra média
de R$ 229.

A renúncia tributária projetada tem sido subestimada em relação à estimada
(defasagem de −35% em 2006 e −3,7% em 2009). A partir de 2010 a RFB passa a
usar “projetado” / “estimado” no lugar de “estimado” / “efetivo”, e o IR passa
a ser alocado pelo ano-calendário.

Nos benefícios financeiros e creditícios (Portaria MF 379/2006), o total SPE
sobe 11,69%, de {_br_bi(16_901.39)} para **{_br_bi(18_877.65)}**. O FCVS
explica a maior parte da alta (de {_br_bi(693.59)} para {_br_bi(6_497.73)}).
Os fundos constitucionais (FNE, FNO, FCO) somam {_br_bi(6_192.91)}. O FIES
vai a {_br_bi(986.18)} (+67,4%).

O TCU registra, no TC 022.684/2010-7, levantamento sobre esses benefícios —
inclusive os da Lei 12.096/2009 e da Lei 11.948/2009, que autorizaram
subvenção econômica ao BNDES para financiamentos de até {_br_bi(124_000)} e
crédito ao banco **acima de {_br_bi(200_000)}**.

## PAC — desonerações e subsídios creditícios

Desonerações tributárias do PAC em 2010: **{_br_bi(23_318)}**, das quais
{_br_bi(20_745)} são medidas pré-PAC (reajuste da tabela do IR, {_br_bi(13_796)};
Lei Geral das PME, {_br_bi(4_500)}). Acumulado 2007–2010: {_br_bi(63_400)}.
Concentração no Sudeste: 63%.

Subsídios creditícios do PAC (SPE; NTN-F 2017 = 12,32% a.a.), por eixo:

| Eixo | Desembolsos | Part. | Contratações | Part. | Desemb. IPCA {data_ref.strftime('%b/%Y')} |
|---|---:|---:|---:|---:|---:|
"""
    for _, r in eixo.iterrows():
        md += (
            f"| {r['eixo']} | {_br_num(r['desembolsos'])} | "
            f"{_br_num(r['part_desemb_pct'], 1)}% | {_br_num(r['contratacoes'])} | "
            f"{_br_num(r['part_contr_pct'], 1)}% | {_br_num(r['desembolsos_ipca_jun2026'])} |\n"
        )

    md += f"""
Desembolsos físicos do PAC com recursos públicos federais: {_br_bi(21_700)}
(energia {_br_bi(15_940)}, inclusive Santo Antônio e Jirau). Contratações:
{_br_bi(17_400)}. Regionalização dos desembolsos: Norte 36%, Nordeste 29%,
Centro-Oeste 4%. Nas **contratações**, 78% dos subsídios vão ao Sudeste
(rodovias federais, embarcações e petroleiros em SP e RJ).

A SPE calcula o subsídio projeto a projeto (taxa, carência, prazo, sistema de
amortização). É o mesmo desenho — diferencial entre taxa cobrada e custo de
oportunidade do Tesouro — usado neste repositório com Selic e fator
30/06/2026. A diferença é o indexador: o TCU/SPE usa a NTN-F 2017 (12,32%);
aqui, a Selic efetiva na data de cada parcela.

## Como isso se conecta aos fluxos 2009–2010 do repositório

O gráfico da p. 33 (Selic, jan/2006–mar/2011) e o quadro da p. 35 (fatores
da base, 2003–2010) medem preço e quantidade da mesma liquidez. O
cotejamento está em `output/TCU_CG_2010_SELIC_BASE.md`: em 2007 a Selic
cai enquanto os títulos esterilizam o câmbio; em 2008 o compulsório injeta
com a Selic ainda alta; em 2010 o compulsório aperta à frente da Selic e
os títulos sobem como contrapartida, não como afrouxamento.

O gráfico da p. 43 (reservas, US$ 37,8 bi em 2002 → US$ 288,6 bi em 2010)
é o estoque em dólares do mesmo canal. Como as reservas eram ativo do
Bacen, a compra de dólares emitia reais (setor externo da p. 35). Essa
emissão não aparece como M1 descontrolado: a Selic e as compromissadas
a reciclam para dentro do M3. A série M1–M4 e o comentário estão em
`output/TCU_CG_2010_RESERVAS_M1M4.md`.

O relatório oficial confirma, no exercício de 2010:

1. o Tesouro virou o principal *funding* do BNDES, com estoque de
   {_br_bi(236_723)} e teto legal de {_br_bi(180_000)} (depois {_br_bi(210_000)}
   com a MP 505 e {_br_bi(293_000)} com a MP 526/2011);
2. o custo fiscal dessa alavancagem já era mensurado pelo TCU em
   {_br_bi(14_200)} ao ano no estoque e {_br_bi(3_100)} no primeiro lote da
   Lei 11.948/2009;
3. os subsídios creditícios do PAC de 2010 ({_br_bi(4_933.82)} nas
   contratações) estão quase inteiros no BNDES.

Os fluxos gerados por `scripts/gerar_fluxos.py` e
`scripts/contagil_fluxos.py` (indiretas automáticas 2009–2010) medem o mesmo
fenômeno no nível do contrato: `subsídio = saldo_fiscal × (Selic − taxa do
contrato)` e `impacto_fiscal` capitalizado até 30/06/2026. O número do TCU é
uma conta de estoque/ano; o deste repositório é a soma das parcelas.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010.xlsx` | Tabelas extraídas + coluna IPCA até {ref} |
| `output/TCU_CG_2010_RELATORIO.md` | Este relatório |
| `output/TCU_CG_2010_BASE_MONETARIA.md` | Análise da p. 35 — fatores da base monetária 2003–2010 |
| `output/TCU_CG_2010_SELIC_BASE.md` | Cotejamento Selic (p. 33) × fatores da base (p. 35) |
| `output/TCU_CG_2010_RESERVAS_M1M4.md` | Reservas (p. 43) × M1–M4 × Selic/repos |
| `scripts/tcu_cg_2010_dados.py` | Valores nominais extraídos do PDF |
| `scripts/build_tcu_cg_2010.py` | Regenera a planilha e este markdown |

```bash
python3 scripts/build_tcu_cg_2010.py
```

Abas da planilha: Fonte, Indicadores, Creditos_DLSP, Autorizacoes_Legais, DPF,
Fatores_DLSP, Superavit_Financeiro, Renuncia_Regional, Renuncia_Tributaria,
Renuncia_Projetada, Carga_vs_Renuncia_PIB, Renuncia_Previdenciaria,
Beneficios_Fin_Cred, PAC_Desoneracoes, PAC_Desoneracoes_Serie,
PAC_Subsidios_Eixo, Resumo_IPCA, Base_Monetaria, Base_Monetaria_Detalhe,
Base_Monetaria_Acum, Base_Monetaria_IPCA, Selic_Copom, Selic_Anual,
Cotejamento_Selic_Base, Selic_TCU_vs_oficial, Reservas_Internacionais,
Agregados_M1_M4.
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def build(
    ipca: pd.DataFrame,
    data_ref: datetime = DATA_REF_DEFAULT,
    xlsx: Path = OUTPUT_XLSX,
    md: Path = OUTPUT_MD,
) -> tuple[Path, Path, float]:
    fator = fator_dez2010_ref(ipca, data_ref=data_ref)
    base = analisar(ipca=ipca, data_ref=data_ref, pasta=xlsx.parent)
    cruz = cotejar(pasta=xlsx.parent)
    reservas = analisar_reservas(pasta=xlsx.parent)
    extra = {
        "Base_Monetaria": base["serie"],
        "Base_Monetaria_Detalhe": base["detalhe"],
        "Base_Monetaria_Acum": base["acumulado"],
        "Selic_Copom": cruz["decisoes"],
        "Selic_Anual": cruz["anual"],
        "Cotejamento_Selic_Base": cruz["cotejo"],
        "Selic_TCU_vs_oficial": cruz["tcu_vs_oficial"],
        "Reservas_Internacionais": reservas["reservas"],
        "Agregados_M1_M4": reservas["agregados"],
        "Reservas_vs_M1M4": reservas["quadro"],
    }
    if base["ipca"] is not None:
        extra["Base_Monetaria_IPCA"] = base["ipca"]
    p_xlsx = escrever_excel(xlsx, fator, data_ref, extra_sheets=extra)
    p_md = escrever_markdown(md, fator, data_ref)
    return p_xlsx, p_md, fator


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ipca", type=Path, default=None, help="Excel IPCA (Data, IPCA % a.m.)")
    parser.add_argument(
        "--data-ref",
        default=DATA_REF_DEFAULT.strftime("%Y-%m-%d"),
        help="Data de referência IPCA (padrão 2026-06-30)",
    )
    parser.add_argument("--xlsx", type=Path, default=OUTPUT_XLSX)
    parser.add_argument("--md", type=Path, default=OUTPUT_MD)
    args = parser.parse_args(argv)

    data_ref = datetime.strptime(args.data_ref, "%Y-%m-%d")
    ipca = carregar_ipca(args.ipca)
    xlsx, md, fator = build(ipca, data_ref=data_ref, xlsx=args.xlsx, md=args.md)
    print(f"[OK] fator IPCA dez/2010 → {data_ref.date()} = {fator:.6f}")
    print(f"[OK] {xlsx}")
    print(f"[OK] {md}")
    print(f"[OK] {xlsx.parent / 'TCU_CG_2010_BASE_MONETARIA.md'}")
    print(f"[OK] {xlsx.parent / 'TCU_CG_2010_SELIC_BASE.md'}")
    print(f"[OK] {xlsx.parent / 'TCU_CG_2010_RESERVAS_M1M4.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
