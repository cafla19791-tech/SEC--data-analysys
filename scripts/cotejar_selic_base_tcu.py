#!/usr/bin/env python3
"""Coteja a Selic (gráfico TCU p. 33) com os fatores da base (p. 35), 2003–2010.

A p. 33 trata a Selic overnight como instrumento primário e o compulsório
como secundário. A p. 35 decompõe quem cria ou destrói base. Este módulo
cruza o preço da liquidez com a quantidade.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Patch

from scripts import tcu_cg_2010_dados as D
from scripts.analisar_base_monetaria_tcu import CORES, FATORES, ROTULOS, df_fatores

JANELA_GRAFICO_TCU = (date(2006, 1, 1), date(2011, 3, 31))


def df_selic_decisoes() -> pd.DataFrame:
    return pd.DataFrame(D.selic_copom_decisoes())


def df_selic_anual() -> pd.DataFrame:
    return pd.DataFrame(D.selic_anual())


def ciclos_cotejamento() -> list[dict]:
    """Leitura ano a ano: sentido da Selic × sinais dos fatores da p. 35."""
    return [
        {
            "ano": 2003,
            "ciclo_selic": "Queda após pico (26,50% em fev.)",
            "instrumento_dominante": "Preço (Selic) + demais",
            "leitura": (
                "Copom herda 25% e sobe até 26,50% em fevereiro para ancorar a "
                "inflação de 2003; depois corta 10 pp, a 16,50% em dezembro. "
                "A base quase não se move (−83 milhões). Tesouro ainda é quase "
                "neutro; títulos injetam 11,2 bi e demais contraem 10,8 bi. "
                "O setor externo ainda não é o choque — o ajuste é de preço."
            ),
        },
        {
            "ano": 2004,
            "ciclo_selic": "Alta do ciclo 2004–05 (16,00 → 17,75%)",
            "instrumento_dominante": "Títulos + Tesouro",
            "leitura": (
                "Selic pausa em 16% em abril e retoma a alta no segundo "
                "semestre. O Tesouro vira dreno estrutural (−42,1 bi). O setor "
                "externo começa a expandir (+12,6 bi). Títulos injetam 52,1 bi "
                "— mais do que compensam Tesouro e demais. A quantidade já "
                "cresce (+14,1 bi) enquanto o preço volta a apertar."
            ),
        },
        {
            "ano": 2005,
            "ciclo_selic": "Aperto máximo (pico 19,75%), depois corte",
            "instrumento_dominante": "Preço (Selic) + Tesouro; externo lidera a quantidade",
            "leitura": (
                "Média anual mais alta do período (19,14%). O aperto é quase "
                "só via juros: títulos quase somem (+2,8 bi). Quem expande a "
                "base é o setor externo (+52,4 bi); o Tesouro drena 43,0 bi. "
                "A quantidade cresce apesar da Selic no teto do ciclo."
            ),
        },
        {
            "ano": 2006,
            "ciclo_selic": "Corte contínuo (18,00 → 13,25%) — janela do gráfico TCU",
            "instrumento_dominante": "Externo expansionista; títulos neutros",
            "leitura": (
                "Começa a janela do gráfico da p. 33. Copom corta 4,75 pp. "
                "Títulos deixam de injetar (−0,7 bi). Externo acelera "
                "(+74,4 bi) e o Tesouro drena 59,5 bi. Política de juros em "
                "afrouxamento com absorção cambial sem repos líquidos."
            ),
        },
        {
            "ano": 2007,
            "ciclo_selic": "Queda ao piso pré-crise (11,25%)",
            "instrumento_dominante": "Esterilização clássica via títulos",
            "leitura": (
                "Caso-escola do cotejamento. Selic ainda cai (13,25 → 11,25%) "
                "enquanto o setor externo injeta o recorde da série "
                "(+155,4 bi). Títulos contraem 74,0 bi para que a overnight "
                "efetiva possa sentar na meta. Sem essa esterilização a base "
                "teria crescido cerca de 100 bi; cresceu 25,5 bi."
            ),
        },
        {
            "ano": 2008,
            "ciclo_selic": "Alta a 13,75%; depois crise com Selic ainda alta",
            "instrumento_dominante": "Demais (compulsório) — instrumento secundário da p. 33",
            "leitura": (
                "No primeiro semestre a Selic sobe 2,50 pp e o Tesouro faz a "
                "maior drenagem da série (−74,3 bi). No segundo, o externo "
                "inverte (−12,1 bi, único ano) e o Bacen injeta via demais "
                "(+53,3 bi, corte de compulsório) e títulos (+34,1 bi). A "
                "base quase não cresce (+0,9 bi). Assincronia: o preço "
                "permanece alto (13,75% até janeiro de 2009) enquanto a "
                "quantidade já é de socorro — exatamente o uso secundário "
                "que o TCU descreve nas pp. 32–33."
            ),
        },
        {
            "ano": 2009,
            "ciclo_selic": "Corte agressivo ao piso de 8,75%",
            "instrumento_dominante": "Preço (Selic) + volta do externo",
            "leitura": (
                "Maior corte do período (−5,00 pp). O TCU lê o gráfico da "
                "p. 33 como o estímulo da crise. Externo volta (+62,9 bi). "
                "Títulos expandem pouco (+11,3 bi) com o estoque de "
                "compromissadas em 427,9 bi em dezembro. Demais deixa de "
                "injetar. Preço muito baixo e quantidade de novo liderada "
                "pelo câmbio."
            ),
        },
        {
            "ano": 2010,
            "ciclo_selic": "Alta tardia (8,75 → 10,75%); troca de instrumento",
            "instrumento_dominante": "Compulsório à frente da Selic; títulos como contrapartida",
            "leitura": (
                "Tese do TCU na p. 33: retirada intempestiva do estímulo. "
                "A Selic só sobe em 29/4 (9,50%), 10/6 (10,25%) e 22/7 "
                "(10,75%) e para aí, com IPCA em 5,91%. O compulsório aperta "
                "antes e mais forte (depósitos IF −236,9 bi; demais −233,1 bi). "
                "Títulos +249,5 bi não são afrouxamento: o TCU (p. 34) diz "
                "que a alta do compulsório 'abriu espaço' para resgatar "
                "repos (427,9 → 259,2 bi). A base ainda cresce 24,6% porque "
                "o externo (+75,6 bi) supera o Tesouro (−51,2 bi)."
            ),
        },
    ]


def df_cotejamento() -> pd.DataFrame:
    selic = df_selic_anual().set_index("ano")
    base = df_fatores().set_index("ano")
    ciclos = {c["ano"]: c for c in ciclos_cotejamento()}
    rows = []
    for ano in range(2003, 2011):
        s = selic.loc[ano]
        b = base.loc[ano]
        c = ciclos[ano]
        rows.append(
            {
                "ano": ano,
                "selic_ini": s["selic_ini"],
                "selic_fim": s["selic_fim"],
                "selic_media": s["selic_media"],
                "selic_max": s["selic_max"],
                "selic_min": s["selic_min"],
                "delta_selic_pp": s["delta_pp"],
                "sentido_selic": s["sentido"],
                "tesouro_nacional": b["tesouro_nacional"],
                "titulos_publicos": b["titulos_publicos"],
                "setor_externo": b["setor_externo"],
                "demais_operacoes": b["demais_operacoes"],
                "var_base": b["var_base"],
                "ciclo_selic": c["ciclo_selic"],
                "instrumento_dominante": c["instrumento_dominante"],
                "leitura": c["leitura"],
            }
        )
    return pd.DataFrame(rows)


def _br_bi(valor_milhoes: float, casas: int = 1) -> str:
    return (
        f"R$ {valor_milhoes / 1000.0:,.{casas}f} bi"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def _br_num(valor: float, casas: int = 1) -> str:
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _br_pp(valor: float) -> str:
    sinal = "+" if valor > 0 else ""
    return f"{sinal}{_br_num(valor, 2)} pp"


def grafico_selic_base(df: pd.DataFrame, destino: Path) -> Path:
    """Dois painéis: degrau da Selic (2003–mar/2011) e fatores anuais da p. 35."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    decisoes = D.selic_copom_decisoes()
    xs = [datetime.combine(d["vigencia"], datetime.min.time()) for d in decisoes]
    ys = [d["selic"] for d in decisoes]
    xs.append(datetime(2011, 3, 31))
    ys.append(ys[-1])

    fig, (ax_s, ax_b) = plt.subplots(
        2,
        1,
        figsize=(12.2, 8.6),
        gridspec_kw={"height_ratios": [1.05, 1.15], "hspace": 0.28},
    )

    ax_s.axvspan(
        datetime.combine(JANELA_GRAFICO_TCU[0], datetime.min.time()),
        datetime.combine(JANELA_GRAFICO_TCU[1], datetime.min.time()),
        color="#dbeafe",
        alpha=0.65,
        zorder=0,
    )
    ax_s.step(xs, ys, where="post", color="#b45309", linewidth=2.0, zorder=2)
    ax_s.set_xlim(datetime(2003, 1, 1), datetime(2011, 3, 31))
    ax_s.set_ylim(7.5, 28.5)
    ax_s.set_ylabel("% a.a.")
    ax_s.set_title(
        "Selic meta (Copom) e fatores condicionantes da base monetária — 2003 a 2010\n"
        "Faixa azul: janela do gráfico TCU p. 33 (jan/2006–mar/2011). "
        "Barras: TCU p. 35."
    )
    ax_s.xaxis.set_major_locator(mdates.YearLocator())
    ax_s.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax_s.axhline(8.75, color="#9ca3af", linewidth=0.6, linestyle=":")
    ax_s.axhline(11.25, color="#9ca3af", linewidth=0.6, linestyle=":")
    marcas = [
        (datetime(2003, 2, 20), 26.50, "26,50%"),
        (datetime(2005, 5, 19), 19.75, "19,75%"),
        (datetime(2007, 9, 6), 11.25, "11,25%"),
        (datetime(2008, 9, 11), 13.75, "13,75%"),
        (datetime(2009, 7, 23), 8.75, "8,75%"),
        (datetime(2010, 7, 22), 10.75, "10,75%"),
    ]
    for x, y, rotulo in marcas:
        ax_s.scatter([x], [y], color="#b45309", s=18, zorder=3)
        ax_s.annotate(
            rotulo,
            xy=(x, y),
            xytext=(0, 8),
            textcoords="offset points",
            ha="center",
            fontsize=8,
            color="#7c2d12",
        )
    ax_s.legend(
        handles=[
            plt.Line2D([0], [0], color="#b45309", lw=2, label="Selic meta (SGS 432)"),
            Patch(facecolor="#dbeafe", edgecolor="none", label="Gráfico TCU p. 33"),
        ],
        loc="upper right",
        frameon=False,
    )
    ax_s.spines["top"].set_visible(False)
    ax_s.spines["right"].set_visible(False)

    anos = df["ano"].tolist()
    x = list(range(len(anos)))
    bottom_pos = [0.0] * len(anos)
    bottom_neg = [0.0] * len(anos)
    for col in FATORES:
        vals = df[col].to_numpy(dtype=float) / 1000.0
        pos = [v if v > 0 else 0.0 for v in vals]
        neg = [v if v < 0 else 0.0 for v in vals]
        ax_b.bar(x, pos, bottom=bottom_pos, color=CORES[col], width=0.72, label=ROTULOS[col])
        ax_b.bar(x, neg, bottom=bottom_neg, color=CORES[col], width=0.72)
        bottom_pos = [a + b for a, b in zip(bottom_pos, pos)]
        bottom_neg = [a + b for a, b in zip(bottom_neg, neg)]
    ax_b.plot(
        x,
        df["var_base"] / 1000.0,
        color=CORES["var_base"],
        marker="o",
        linewidth=2.0,
        label="Variação da base",
    )
    ax_b.axhline(0, color="#9ca3af", linewidth=0.8)
    ax_b.set_xticks(x, [str(a) for a in anos])
    ax_b.set_ylabel("R$ bilhões")
    ax_b.set_xlabel("Exercício (fatores condicionantes da base — TCU p. 35)")
    ax_b.legend(loc="upper left", frameon=False, ncol=2, fontsize=8)
    ax_b.spines["top"].set_visible(False)
    ax_b.spines["right"].set_visible(False)

    fig.subplots_adjust(left=0.08, right=0.98, top=0.90, bottom=0.08, hspace=0.32)
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(destino: Path, cotejo: pd.DataFrame) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    linhas = [
        "| Ano | Selic ini | Selic fim | Média | Δ Selic | Tesouro | Títulos | Externo | Demais | Δ base |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in cotejo.iterrows():
        linhas.append(
            f"| {int(r['ano'])} | {_br_num(r['selic_ini'], 2)} | "
            f"{_br_num(r['selic_fim'], 2)} | {_br_num(r['selic_media'], 2)} | "
            f"{_br_pp(r['delta_selic_pp'])} | {_br_num(r['tesouro_nacional'], 0)} | "
            f"{_br_num(r['titulos_publicos'], 0)} | {_br_num(r['setor_externo'], 0)} | "
            f"{_br_num(r['demais_operacoes'], 0)} | {_br_num(r['var_base'], 0)} |"
        )

    linhas_tcu = [
        "| Evento | Leitura do TCU (p. 33) | Calendário oficial do Copom |",
        "|---|---|---|",
    ]
    for row in D.tcu_vs_oficial_selic_p33():
        linhas_tcu.append(
            f"| {row['evento']} | {row['leitura_tcu_p33']} | {row['oficial_copom']} |"
        )

    md = f"""# Selic (p. 33) × fatores da base monetária (p. 35) — 2003 a 2010

**Fonte:** TCU, *Relatório e Parecer Prévio sobre as Contas do Governo da República — Exercício de 2010*, gráfico da **p. 33** e quadro da **p. 35** (detalhe 2009–2010 na p. 34; compromissadas na p. 36).
**Selic:** meta do Copom, série Bacen SGS 432. O gráfico impresso pelo TCU cobre **janeiro de 2006 a março de 2011**; o cotejamento usa a mesma série oficial de 2003 a 2010, para casar com a tabela da p. 35.
**Convenção da base:** (+) expansão / (−) retração. Identidade: Tesouro + títulos + setor externo + demais = Δ base.

## Como ler os dois instrumentos juntos

A p. 33 e a p. 35 medem coisas diferentes:

- a **Selic** é o *preço* da liquidez overnight — o instrumento primário do Copom;
- os **fatores da base** são a *quantidade* e *quem* cria ou destrói essa liquidez.

O TCU escreve, na própria p. 33, que o compulsório e os limites de crédito
são instrumentos **secundários**. A Mesa do mercado aberto ajusta a liquidez
todos os dias para a taxa efetiva sentar na meta. Por isso o sinal dos
**títulos públicos** na p. 35 não se lê como “política frouxa” ou “apertada”:
é o braço que implementa a meta. Quando outro instrumento (câmbio, Tesouro
ou compulsório) choca a base, os títulos (e as compromissadas) compensam.

Três regras do cotejamento:

1. **Não se lê Δ base como stance.** Em 2003 a Selic cai 8,50 pp e a base
   fica estável. Em 2010 a Selic sobe 2,00 pp e a base cresce 24,6%.
2. **Títulos sobem quando o Bacen precisa devolver liquidez** — inclusive
   depois de um aperto de compulsório (2010) ou na crise (2008).
3. **Demais / compulsório** é o canal secundário da p. 33: corta em 2008
   com a Selic ainda alta; sobe em 2010 antes e mais forte que a Selic.

## Quadro cotejado (Selic em % a.a.; fatores em R$ milhões)

{chr(10).join(linhas)}

## O que o gráfico da p. 33 mostra — e o que o TCU lê nele

O TCU descreve o gráfico como o resultado das decisões do Copom, “com a
defasagem esperada”, de jan/2006 a mar/2011. A narrativa oficial do
relatório (pp. 32–33) é:

- até a reunião de 22/7/2009 a Selic vigente seria 8,75%;
- teria ido a 10,25% “somente em 1º/9/2010”;
- o +0,50 pp para 11,25% teria sido em 19/1/2011;
- a retirada tardia do estímulo de 2008 explicaria o IPCA de 5,91% em 2010
  e o acumulado de 6,3% em 12 meses até março de 2011.

O calendário oficial do Copom confirma a **tese** (o piso de 8,75% durou
até abril de 2010; a alta parou em 10,75% em julho) e corrige três datas
do texto:

{chr(10).join(linhas_tcu)}

No item 2.3.2 o próprio TCU já descreve 2010 com as datas certas: 8,75%
até o fim de abril, 9,50% em maio, 10,25% de junho a meados de julho e
10,75% até o encerramento do exercício. A leitura da p. 33 é, portanto,
mais um juízo sobre *intempestividade* do que um calendário.

## Cotejamento ano a ano

"""
    for _, r in cotejo.iterrows():
        md += (
            f"### {int(r['ano'])} — {r['ciclo_selic']}\n\n"
            f"**Selic:** {_br_num(r['selic_ini'], 2)}% → {_br_num(r['selic_fim'], 2)}% "
            f"(média {_br_num(r['selic_media'], 2)}%; {_br_pp(r['delta_selic_pp'])}; "
            f"máx. {_br_num(r['selic_max'], 2)}% / mín. {_br_num(r['selic_min'], 2)}%).\n\n"
            f"**Base:** Tesouro {_br_bi(r['tesouro_nacional'])}; "
            f"títulos {_br_bi(r['titulos_publicos'])}; "
            f"externo {_br_bi(r['setor_externo'])}; "
            f"demais {_br_bi(r['demais_operacoes'])}; "
            f"Δ base {_br_bi(r['var_base'])}.\n\n"
            f"**Instrumento que manda:** {r['instrumento_dominante']}\n\n"
            f"{r['leitura']}\n\n"
        )

    md += f"""## Síntese do cruzamento

O desce-e-sobe da Selic e os fatores da base não caminham no mesmo sentido.
Caminham em **camadas**:

| Camada | Papel no período | O que faz na crise e na saída |
|---|---|---|
| Tesouro | Dreno em todos os oito anos | Independente do ciclo Selic; pico em 2008 (−{_br_bi(74_312)}) |
| Setor externo | Choque expansionista (exceto 2008) | Recorde em 2007 (+{_br_bi(155_390)}); único ano negativo em 2008 |
| Títulos / repos | Implementam a meta Selic | Esterilizam o câmbio em 2007 (−{_br_bi(73_974)}); devolvem liquidez em 2010 (+{_br_bi(249_513)}) quando o compulsório aperta |
| Demais / compulsório | Instrumento secundário da p. 33 | Injeta em 2008 (+{_br_bi(53_311)}) com Selic ainda em 13,75%; drena em 2010 (−{_br_bi(233_082)}) à frente da alta da Selic |
| Selic | Instrumento primário (preço) | Piso 8,75% de 23/7/2009 a 28/4/2010; alta 2010 para em 10,75% |

Quatro cruzamentos fecham o argumento do TCU:

1. **2007** — Selic em queda e títulos em forte contração: esterilização
   clássica do choque externo. O preço pode cair porque a quantidade é
   enxugada no open market.
2. **2008** — Selic em alta e demais em expansão: o instrumento secundário
   socorre a liquidez enquanto o primário ainda está no ciclo de aperto.
3. **2009** — Selic no piso e externo de volta: estímulo de preço com a
   quantidade outra vez liderada pelo câmbio; compromissadas em
   {_br_bi(427_874)} (dez/2009).
4. **2010** — Selic sobe pouco e tarde; compulsório sobe cedo e forte;
   títulos sobem como *contrapartida*, não como afrouxamento. A base
   mesmo assim cresce {_br_bi(40_780)} (+24,6%). É a “intempestividade”
   da p. 33 vista pelo lado da quantidade da p. 35.

A alta de títulos em 2010 (+{_br_bi(249_513)}) e a queda das
compromissadas (dez/2009 {_br_bi(427_874)} → dez/2010 {_br_bi(259_248)})
são o mesmo movimento: o Bacen troca um instrumento de ajuste fino
(repos) por um dreno mais permanente (compulsório). Quem olha só o sinal
dos títulos na p. 35 lê afrouxamento; quem coteja com o gráfico da p. 33
e com o detalhe da p. 34 lê **troca de instrumento** — e, no juízo do
TCU, troca incompleta, porque a Selic parou em 10,75% com o IPCA em 5,91%.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_SELIC_BASE.md` | Este cotejamento |
| `output/grafico_selic_base_monetaria_2003_2010.png` | Selic em degrau + fatores empilhados |
| `output/TCU_CG_2010.xlsx` (abas `Selic_*`, `Cotejamento_Selic_Base`) | Decisões, totais anuais e quadro cruzado |

```bash
python3 scripts/build_tcu_cg_2010.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def cotejar(pasta: Path | None = None) -> dict:
    pasta = pasta or Path("output")
    cotejo = df_cotejamento()
    png = grafico_selic_base(df_fatores(), pasta / "grafico_selic_base_monetaria_2003_2010.png")
    md = escrever_markdown(pasta / "TCU_CG_2010_SELIC_BASE.md", cotejo)
    return {
        "decisoes": df_selic_decisoes(),
        "anual": df_selic_anual(),
        "tcu_vs_oficial": pd.DataFrame(D.tcu_vs_oficial_selic_p33()),
        "cotejo": cotejo,
        "png": png,
        "md": md,
    }
