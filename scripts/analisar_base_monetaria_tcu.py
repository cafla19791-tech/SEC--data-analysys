#!/usr/bin/env python3
"""Análise dos fatores condicionantes da base monetária — TCU CG 2010, p. 35.

Quadro oficial: 2003–2010 (o título cita 2002–2010, mas 2002 não foi impresso).
Convenção: (+) expansão / (−) retração da base.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from scripts.calcular_diretas_ipca_selic import DATA_REF_DEFAULT, fator_ipca_entre
from scripts import tcu_cg_2010_dados as D

FATORES = (
    "tesouro_nacional",
    "titulos_publicos",
    "setor_externo",
    "demais_operacoes",
)
ROTULOS = {
    "tesouro_nacional": "Tesouro Nacional",
    "titulos_publicos": "Títulos públicos federais",
    "setor_externo": "Setor externo",
    "demais_operacoes": "Demais operações",
    "var_base": "Variação da base monetária",
}
CORES = {
    "tesouro_nacional": "#b45309",
    "titulos_publicos": "#1d4ed8",
    "setor_externo": "#047857",
    "demais_operacoes": "#6d28d9",
    "var_base": "#111827",
}


def df_fatores() -> pd.DataFrame:
    df = pd.DataFrame(D.fatores_base_monetaria())
    soma = df[list(FATORES)].sum(axis=1)
    df["identidade_ok"] = (soma - df["var_base"]).abs() <= 1.0
    df["residuo"] = soma - df["var_base"]
    return df


def df_detalhe_2009_2010() -> pd.DataFrame:
    df = pd.DataFrame(D.fatores_base_monetaria_detalhe_2009_2010())
    partes = [
        "tesouro_nacional",
        "titulos_publicos",
        "setor_externo",
        "depositos_inst_financ",
        "derivativos_ajustes",
        "outras_contas_ajustes",
    ]
    df["residuo"] = df[partes].sum(axis=1) - df["var_base"]
    return df


def aplicar_ipca_anual(
    df: pd.DataFrame,
    ipca: pd.DataFrame,
    data_ref: datetime = DATA_REF_DEFAULT,
) -> pd.DataFrame:
    out = df.copy()
    fatores = []
    for ano in out["ano"]:
        fatores.append(
            fator_ipca_entre(
                ipca,
                pd.Timestamp(year=int(ano), month=12, day=1),
                pd.Timestamp(data_ref),
            )
        )
    out["fator_ipca"] = fatores
    for col in list(FATORES) + ["var_base"]:
        out[f"{col}_ipca"] = out[col] * out["fator_ipca"]
    return out


def resumo_acumulado(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(FATORES) + ["var_base"]
    nominal = df[cols].sum()
    rows = []
    for col in cols:
        rows.append(
            {
                "fator": ROTULOS[col],
                "acumulado_r_mi": float(nominal[col]),
                "acumulado_r_bi": float(nominal[col]) / 1000.0,
                "anos_expansao": int((df[col] > 0).sum()) if col != "var_base" else int((df["var_base"] > 0).sum()),
                "anos_retracao": int((df[col] < 0).sum()) if col != "var_base" else int((df["var_base"] < 0).sum()),
                "min_r_mi": float(df[col].min()),
                "max_r_mi": float(df[col].max()),
                "ano_min": int(df.loc[df[col].idxmin(), "ano"]),
                "ano_max": int(df.loc[df[col].idxmax(), "ano"]),
            }
        )
    return pd.DataFrame(rows)


def fases() -> list[dict]:
    return [
        {
            "fase": "2003 — estabilização",
            "var_base_r_mi": -83,
            "leitura": (
                "Base praticamente estável. Tesouro quase neutro (−1,1 bi). "
                "Títulos expandem (+11,2 bi) e demais operações contraem (−10,8 bi)."
            ),
        },
        {
            "fase": "2004–2007 — acúmulo de reservas esterilizado",
            "var_base_r_mi": 14_102 + 12_514 + 19_854 + 25_516,
            "leitura": (
                "Setor externo vira o motor expansionista (12,6 → 155,4 bi). "
                "O Tesouro drena 42–60 bi/ano. Em 2007 os títulos públicos "
                "esterilizam −74,0 bi, o maior enxugamento da série até 2009."
            ),
        },
        {
            "fase": "2008 — crise",
            "var_base_r_mi": 933,
            "leitura": (
                "Único ano em que o setor externo contrai (−12,1 bi). "
                "Títulos (+34,1) e demais operações (+53,3, corte de compulsório) "
                "injetam liquidez. O Tesouro faz a maior drenagem da série (−74,3 bi). "
                "A base quase não cresce (+0,9 bi)."
            ),
        },
        {
            "fase": "2009 — estímulos mantidos",
            "var_base_r_mi": 18_523,
            "leitura": (
                "Volta o padrão pré-crise: externo +62,9 bi, Tesouro −52,3 bi, "
                "títulos +11,3 bi. Compulsórios ainda não foram restabelecidos."
            ),
        },
        {
            "fase": "2010 — troca de instrumento",
            "var_base_r_mi": 40_780,
            "leitura": (
                "Maior expansão da série (+40,8 bi, +24,6%). O Bacen sobe o "
                "compulsório (−236,9 bi em depósitos) e, com isso, resgata "
                "títulos/compromissadas (+249,5 bi). O setor externo segue "
                "expansionista (+75,6 bi)."
            ),
        },
    ]


def _br_bi(valor_milhoes: float, casas: int = 1) -> str:
    return (
        f"R$ {valor_milhoes / 1000.0:,.{casas}f} bi"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def _br_num(valor: float, casas: int = 1) -> str:
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def grafico_empilhado(df: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    anos = df["ano"].tolist()
    x = range(len(anos))
    fig, ax = plt.subplots(figsize=(11.2, 6.2))
    bottom_pos = [0.0] * len(anos)
    bottom_neg = [0.0] * len(anos)
    for col in FATORES:
        vals = df[col].to_numpy(dtype=float) / 1000.0
        pos = [v if v > 0 else 0.0 for v in vals]
        neg = [v if v < 0 else 0.0 for v in vals]
        ax.bar(x, pos, bottom=bottom_pos, color=CORES[col], width=0.72, label=ROTULOS[col])
        ax.bar(x, neg, bottom=bottom_neg, color=CORES[col], width=0.72)
        bottom_pos = [a + b for a, b in zip(bottom_pos, pos)]
        bottom_neg = [a + b for a, b in zip(bottom_neg, neg)]
    ax.plot(
        list(x),
        df["var_base"] / 1000.0,
        color=CORES["var_base"],
        marker="o",
        linewidth=2.0,
        label="Variação da base",
    )
    ax.axhline(0, color="#9ca3af", linewidth=0.8)
    ax.set_xticks(list(x), [str(a) for a in anos])
    ax.set_ylabel("R$ bilhões")
    ax.set_title(
        "Fatores condicionantes da base monetária — 2003 a 2010\n"
        "TCU, Contas do Governo 2010, p. 35 (fonte Bacen)"
    )
    ax.legend(loc="upper left", frameon=False, ncol=2)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(
    destino: Path,
    df: pd.DataFrame,
    detalhe: pd.DataFrame,
    acum: pd.DataFrame,
    df_ipca: pd.DataFrame | None,
    data_ref: datetime,
) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    linhas = [
        "| Ano | Tesouro | Títulos públ. | Setor externo | Demais | Var. base |",
        "|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in df.iterrows():
        linhas.append(
            f"| {int(r['ano'])} | {_br_num(r['tesouro_nacional'], 0)} | "
            f"{_br_num(r['titulos_publicos'], 0)} | {_br_num(r['setor_externo'], 0)} | "
            f"{_br_num(r['demais_operacoes'], 0)} | {_br_num(r['var_base'], 0)} |"
        )
    tot = df[list(FATORES) + ["var_base"]].sum()
    linhas.append(
        f"| **Soma 2003–2010** | **{_br_num(tot['tesouro_nacional'], 0)}** | "
        f"**{_br_num(tot['titulos_publicos'], 0)}** | **{_br_num(tot['setor_externo'], 0)}** | "
        f"**{_br_num(tot['demais_operacoes'], 0)}** | **{_br_num(tot['var_base'], 0)}** |"
    )

    linhas_acum = [
        "| Fator | Acumulado | Anos (+) | Anos (−) | Mínimo (ano) | Máximo (ano) |",
        "|---|---:|---:|---:|---:|---:|",
    ]
    for _, r in acum.iterrows():
        linhas_acum.append(
            f"| {r['fator']} | {_br_bi(r['acumulado_r_mi'])} | {int(r['anos_expansao'])} | "
            f"{int(r['anos_retracao'])} | {_br_bi(r['min_r_mi'])} ({int(r['ano_min'])}) | "
            f"{_br_bi(r['max_r_mi'])} ({int(r['ano_max'])}) |"
        )

    bloco_ipca = ""
    if df_ipca is not None:
        bloco_ipca = (
            f"\n## Mesmos fluxos em reais de {data_ref.strftime('%d/%m/%Y')} (IPCA)\n\n"
            "Cada fluxo anual é atualizado de dezembro daquele ano até a data de "
            "referência (Bacen SGS 433).\n\n"
            "| Ano | Tesouro | Títulos | Externo | Demais | Var. base | Fator |\n"
            "|---:|---:|---:|---:|---:|---:|---:|\n"
        )
        for _, r in df_ipca.iterrows():
            bloco_ipca += (
                f"| {int(r['ano'])} | {_br_num(r['tesouro_nacional_ipca'], 0)} | "
                f"{_br_num(r['titulos_publicos_ipca'], 0)} | "
                f"{_br_num(r['setor_externo_ipca'], 0)} | "
                f"{_br_num(r['demais_operacoes_ipca'], 0)} | "
                f"{_br_num(r['var_base_ipca'], 0)} | {_br_num(r['fator_ipca'], 4)} |\n"
            )
        s = df_ipca[[f"{c}_ipca" for c in list(FATORES) + ["var_base"]]].sum()
        bloco_ipca += (
            f"| **Soma** | **{_br_num(s['tesouro_nacional_ipca'], 0)}** | "
            f"**{_br_num(s['titulos_publicos_ipca'], 0)}** | "
            f"**{_br_num(s['setor_externo_ipca'], 0)}** | "
            f"**{_br_num(s['demais_operacoes_ipca'], 0)}** | "
            f"**{_br_num(s['var_base_ipca'], 0)}** | — |\n"
        )

    d2010 = detalhe.loc[detalhe["ano"] == 2010].iloc[0]
    md = f"""# Fatores condicionantes da base monetária (2002–2010)

**Fonte:** TCU, *Relatório e Parecer Prévio sobre as Contas do Governo da República — Exercício de 2010*, **p. 35** (e detalhe de 2009–2010 na **p. 34**).
**Origem dos números:** Banco Central, Nota para Imprensa — Política Monetária, março/2011.
**Convenção:** (+) expansão da base / (−) retração.
**Identidade:** Tesouro + Títulos públicos + Setor externo + Demais operações = variação da base.

O título do quadro é “2002 a 2010” e o texto fala em “últimos sete exercícios”.
A tabela impressa, porém, **começa em 2003**. Não há linha para 2002 na p. 35.
A análise abaixo cobre a série oficial impressa (2003–2010).

## O que o quadro mede

A base monetária é papel-moeda emitido + reservas bancárias. Os fatores
condicionantes decompõem a variação anual dessa base nas operações do Bacen
com o Tesouro (Conta Única, inclusive INSS), com títulos públicos federais
(mercado aberto / compromissadas), com o setor externo (compra e venda de
câmbio) e com o restante (compulsório, derivativos e ajustes).

Em 2010 a base cresceu **{_br_bi(40_780)}** (+24,6% sobre dez/2009):
papel-moeda +{_br_bi(19_300)} e reservas bancárias +{_br_bi(21_500)}.

## Série oficial da p. 35 (R$ milhões)

{chr(10).join(linhas)}

A identidade fecha em todos os anos (resíduo ≤ R$ 1 milhão; em 2008 o
arredondamento do TCU é de R$ 1 milhão).

{chr(10).join(linhas_acum)}

No acumulado 2003–2010 o setor externo **injetou** {_br_bi(float(tot['setor_externo']))}
e o Tesouro **retirou** {_br_bi(abs(float(tot['tesouro_nacional'])))}. Os títulos
públicos, no saldo do período, foram expansionistas ({_br_bi(float(tot['titulos_publicos']))})
porque o resgate de 2010 (+{_br_bi(249_513)}) mais do que compensou a
esterilização de 2007. A base cresceu {_br_bi(float(tot['var_base']))} no período.

## Evolução por fator

### Tesouro Nacional — drenagem permanente

O Tesouro é o único fator **contracionista em todos os oito anos**. Depois de
quase neutro em 2003 (−{_br_bi(1_064)}), passa a recolher {_br_bi(42_140)} a
{_br_bi(74_312)} por ano via Conta Única. O pico é 2008. Em 2009–2010 a
drenagem se estabiliza em torno de {_br_bi(52_000)}.

Isso é o superávit de caixa da União (e do INSS) depositado no Bacen: tira
reais de circulação sem ser política monetária no sentido estrito, mas
obriga o Bacen a decidir se esteriliza ou acomoda o resto da liquidez.

### Setor externo — o motor expansionista, exceto 2008

De 2004 a 2007 a compra de moeda estrangeira no interbancário cresce de
{_br_bi(12_599)} para **{_br_bi(155_390)}**. É o período clássico de acúmulo
de reservas. Em 2008 o sinal inverte (−{_br_bi(12_124)}): a crise reduz a
intervenção compradora. Em 2009–2010 o canal volta a expandir a base
({_br_bi(62_937)} e {_br_bi(75_553)}).

### Títulos públicos — o instrumento de esterilização (e o de 2010)

Os títulos fazem o ajuste fino e a esterilização do câmbio:

- 2004: +{_br_bi(52_111)} (injeção, ainda sem o grande ciclo de reservas);
- 2006–2007: −{_br_bi(687)} e **−{_br_bi(73_974)}** (esterilização do
  {_br_bi(155_390)} externo de 2007);
- 2008–2009: de novo expansionistas na crise;
- **2010: +{_br_bi(249_513)}** — o maior número da tabela, e o ponto da
  mudança de regime.

### Demais operações — compulsório e resíduos

Até 2007 o item é pequeno. Em 2008 explode para +{_br_bi(53_311)} (corte de
compulsório na crise). Em 2010 vira **−{_br_bi(233_082)}**: o TCU explica, na
p. 34, que o Bacen restabeleceu alíquotas de depósitos a prazo e exigibilidades
adicionais, com contração de {_br_bi(236_911)} só em depósitos de instituições
financeiras.

## 2010: a troca compulsório ↔ compromissadas

Detalhe da p. 34 (R$ milhões):

| Ano | Tesouro | Títulos | Externo | Depósitos IF | Derivativos | Outras | Var. base |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 2009 | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'tesouro_nacional'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'titulos_publicos'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'setor_externo'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'depositos_inst_financ'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'derivativos_ajustes'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'outras_contas_ajustes'].iloc[0]), 0)} | {_br_num(float(detalhe.loc[detalhe['ano']==2009,'var_base'].iloc[0]), 0)} |
| 2010 | {_br_num(float(d2010['tesouro_nacional']), 0)} | {_br_num(float(d2010['titulos_publicos']), 0)} | {_br_num(float(d2010['setor_externo']), 0)} | {_br_num(float(d2010['depositos_inst_financ']), 0)} | {_br_num(float(d2010['derivativos_ajustes']), 0)} | {_br_num(float(d2010['outras_contas_ajustes']), 0)} | {_br_num(float(d2010['var_base']), 0)} |

O TCU diz com clareza o mecanismo: a contração de {_br_bi(236_911)} no
compulsório “abriu espaço” para o resgate de títulos e compromissadas de
{_br_bi(249_513)}. Os dois movimentos quase se anulam; o que sobra para
expandir a base é, sobretudo, o setor externo (+{_br_bi(75_553)}) menos o
Tesouro (−{_br_bi(51_204)}).

Isso aparece no estoque das compromissadas (p. 36):

- dez/2009: {_br_bi(427_874)}
- dez/2010: {_br_bi(259_248)} (−{_br_bi(168_626)})
- prazo até 3 meses: {_br_bi(316_634)} → {_br_bi(116_509)}
- prazo acima de 3 meses: {_br_bi(79_394)} → {_br_bi(142_739)} (+{_br_bi(63_345)})

O TCU conclui que as compromissadas “têm servido cada vez menos ao propósito
de ajuste fino de liquidez”. O alongamento do prazo e a troca por compulsório
são a mesma história: o Bacen tira de circulação, de forma mais permanente, a
liquidez que ele mesmo criou ao comprar reservas e ao manter estímulos em 2009.

## Cinco fases

"""
    for f in fases():
        md += f"### {f['fase']}\n\n{f['leitura']}\n\n"

    md += f"""## Leitura conjunta

Três regularidades atravessam 2003–2010:

1. **O Tesouro sempre contrai a base.** A Conta Única é um dreno estrutural.
2. **O setor externo é o choque expansionista dominante**, interrompido só
   em 2008. A compra de reservas cria reais; alguém tem de enxugá-los.
3. **O instrumento de enxugamento muda.** Até 2007, títulos/compromissadas
   esterilizam o câmbio (2007 é o ano puro desse regime). Em 2010, o
   compulsório assume o dreno e os títulos são resgatados. A base, mesmo
   assim, cresce no maior ritmo da série (+24,6%), o que o próprio TCU liga
   à intempestividade da retirada dos estímulos de 2008–2009 e ao IPCA de
   5,91% (acima do centro da meta).

A p. 35, portanto, não é só um quadro contábil da base: é o mapa de como o
Tesouro, o câmbio e o Bacen se compensam. O mesmo relatório, no item 2.5,
mostra o outro lado dessa compensação — créditos da União no BNDES e o
custo Selic − TJLP. Aqui, o canal é monetário; lá, é fiscal.
{bloco_ipca}
## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_BASE_MONETARIA.md` | Esta análise |
| `output/TCU_CG_2010.xlsx` (abas `Base_Monetaria_*`) | Série, detalhe, acumulado, IPCA |
| `output/grafico_base_monetaria_2003_2010.png` | Barras empilhadas + variação da base |

```bash
python3 scripts/build_tcu_cg_2010.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def analisar(
    ipca: pd.DataFrame | None = None,
    data_ref: datetime = DATA_REF_DEFAULT,
    pasta: Path | None = None,
) -> dict:
    pasta = pasta or Path("output")
    df = df_fatores()
    detalhe = df_detalhe_2009_2010()
    acum = resumo_acumulado(df)
    df_ipca = aplicar_ipca_anual(df, ipca, data_ref) if ipca is not None else None
    png = grafico_empilhado(df, pasta / "grafico_base_monetaria_2003_2010.png")
    md = escrever_markdown(
        pasta / "TCU_CG_2010_BASE_MONETARIA.md",
        df,
        detalhe,
        acum,
        df_ipca,
        data_ref,
    )
    return {
        "serie": df,
        "detalhe": detalhe,
        "acumulado": acum,
        "ipca": df_ipca,
        "png": png,
        "md": md,
    }
