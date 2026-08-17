#!/usr/bin/env python3
"""Por que a Selic não acelerou a queda após o superávit comercial de 2003.

Responde ao deslocamento de instrumento: até 2001 a Selic ajudava a fechar
o BP; depois de 1999/2003 ela é o instrumento da meta de inflação.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from scripts import selic_bp_2003_2016_dados as D


def df_externo() -> pd.DataFrame:
    bc = pd.DataFrame(D.balanca_comercial()).set_index("ano")
    tc = pd.DataFrame(D.transacoes_correntes()).set_index("ano")
    rs = pd.DataFrame(D.reservas_dezembro()).set_index("ano")
    out = pd.DataFrame({"ano": list(range(1995, 2017))}).set_index("ano")
    out["balanca"] = bc["usd_mi"]
    out["corrente"] = tc["usd_mi"]
    out["reservas"] = rs["usd_mi"]
    return out.reset_index()


def df_selic_ipca() -> pd.DataFrame:
    return pd.DataFrame(D.selic_ipca_anual())


def _br_num(valor: float, casas: int = 1) -> str:
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _usd_bi(valor_mi: float, casas: int = 1) -> str:
    return f"US$ {valor_mi / 1000.0:,.{casas}f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def grafico(ext: pd.DataFrame, si: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig, (ax0, ax1) = plt.subplots(2, 1, figsize=(12.2, 8.4))

    e = ext.loc[ext["ano"] >= 1999]
    x = e["ano"].tolist()
    cores = ["#047857" if v >= 0 else "#b45309" for v in e["balanca"]]
    ax0.bar(x, e["balanca"] / 1000.0, color=cores, width=0.7, label="Balança comercial")
    ax0.plot(x, e["corrente"] / 1000.0, color="#6d28d9", marker="o", linewidth=1.8, label="Transações correntes")
    ax0b = ax0.twinx()
    ax0b.plot(x, e["reservas"] / 1000.0, color="#111827", marker="s", linewidth=2.0, label="Reservas (eixo dir.)")
    ax0.axhline(0, color="#9ca3af", linewidth=0.7)
    ax0.set_ylabel("US$ bilhões")
    ax0b.set_ylabel("Reservas (US$ bilhões)")
    ax0.set_title(
        "Selic, contas externas e inflação — 1999 a 2016\n"
        "Bacen SGS 22710, 22701, 3546, 4390 e 433"
    )
    h1, l1 = ax0.get_legend_handles_labels()
    h2, l2 = ax0b.get_legend_handles_labels()
    ax0.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False)
    ax0.spines["top"].set_visible(False)

    ax1.plot(si["ano"], si["selic_pct"], color="#b45309", marker="o", linewidth=2.0, label="Selic efetiva (SGS 4390)")
    ax1.plot(si["ano"], si["ipca_pct"], color="#1d4ed8", marker="o", linewidth=2.0, label="IPCA")
    ax1.axhline(10, color="#9ca3af", linewidth=0.7, linestyle=":")
    ax1.axhspan(0, 10, color="#dbeafe", alpha=0.35, label="Selic de um dígito")
    ax1.set_ylabel("% no ano")
    ax1.set_xlabel("Exercício")
    ax1.legend(loc="upper right", frameon=False)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.set_ylim(0, 26)

    fig.subplots_adjust(left=0.08, right=0.92, top=0.90, bottom=0.08, hspace=0.28)
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(destino: Path, ext: pd.DataFrame, si: pd.DataFrame) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    acum_01_06 = float(ext.loc[ext["ano"].between(2001, 2006), "balanca"].sum())
    acum_95_00 = float(ext.loc[ext["ano"].between(1995, 2000), "balanca"].sum())
    y16 = si.loc[si["ano"] == 2016].iloc[0]
    fat_s = 1 + y16["selic_acum_pct"] / 100
    fat_i = 1 + y16["ipca_acum_pct"] / 100
    real = (fat_s / fat_i - 1) * 100
    geo_s = (fat_s ** (1 / 14) - 1) * 100
    geo_i = (fat_i ** (1 / 14) - 1) * 100

    linhas_ext = [
        "| Ano | Balança comercial | Transações correntes | Reservas dez | Selic efetiva | IPCA |",
        "|---:|---:|---:|---:|---:|---:|",
    ]
    selic_map = si.set_index("ano")
    for _, r in ext.iterrows():
        if int(r["ano"]) < 1999:
            continue
        ano = int(r["ano"])
        tc = "—" if pd.isna(r["corrente"]) else _br_num(r["corrente"], 0)
        rs = "—" if pd.isna(r["reservas"]) else _br_num(r["reservas"], 0)
        if ano in selic_map.index:
            s = _br_num(selic_map.loc[ano, "selic_pct"], 2)
            i = _br_num(selic_map.loc[ano, "ipca_pct"], 2)
        else:
            s, i = "—", "—"
        linhas_ext.append(
            f"| {ano} | {_br_num(r['balanca'], 0)} | {tc} | {rs} | {s} | {i} |"
        )

    md = f"""# Por que a Selic não acelerou a queda depois do superávit de 2003

A premissa externa está certa. A conclusão sobre o *instrumento* não está.
A partir de 2003 o Brasil **já não precisava** da Selic para captar dólares
e fechar o balanço de pagamentos. Precisava dela — e o Decreto 3.088/1999
já tinha mudado o mandato — para cumprir a **meta de inflação**. Por isso
a taxa caiu (25% no fim de 2002 → 11,25% em 2007 → 8,75% em 2009), mas
não foi ao piso e voltou a subir várias vezes. O “418%” (na série oficial
da Selic efetiva, **+{_br_num(y16['selic_acum_pct'], 1)}%** entre jan/2003 e
dez/2016) não é uma alta de 418 pontos da meta: é o **juro composto** de
quatorze anos com a taxa ainda alta. É o integral do nível, não um único
ciclo de aperto.

## O que a história até 2002 acerta

Até o fim dos anos 1990 a Selic alta tinha, de fato, uma função externa.
O real não é conversível; a pauta de importação (combustíveis, insumos,
medicamentos) exige dólares; o Plano Real usou o câmbio como âncora; a
balança comercial foi deficitária de 1995 a 2000 (acumulado
{_usd_bi(acum_95_00)}). Sem dólares, a economia para. Juro alto atraía
fluxo e segurava a âncora.

A flutuação de janeiro de 1999 e o regime de metas (Decreto 3.088/1999)
já separam os instrumentos: o **câmbio** passa a fechar o BP; a **Selic**
passa a perseguir o IPCA. A queda de 45% para 15% entre março de 1999 e
meados de 2001 é exatamente essa transição — interrompida pelo 11 de
setembro e, em 2002, pelo prêmio eleitoral que levou a meta a 25% e o
IPCA a 12,53%.

## O que muda em 2001–2007 — e o que não muda

A balança comercial vira superavitária em **2001** e cresce até 2006.
Acumulado 2001–2006: **{_usd_bi(acum_01_06)}** (mais de US$ 100 bi, como
no enunciado). Reservas (SGS 3546, gráfico TCU p. 43): US$ 33,0 bi em
2000 → US$ 85,8 bi em 2006 → US$ 180,3 bi em 2007 → US$ 288,6 bi em 2010.
A escassez de 2000 vira abundância. Nesse recorte, a Selic **deixa de ser**
o instrumento de obter dólares para importar diesel.

O que **não** deixa de ser é o instrumento do IPCA. Em janeiro de 2003 a
herança era inflação de dois dígitos e meta ajustada de 8,5%. O Copom
*sobe* primeiro (26,50% em fevereiro) e só então corta. Não havia como
“acelerar a redução a partir de 2003” sem reancorar preços. Em 2004–05,
com a demanda voltando, a Selic sobe de novo até 19,75%. O mandato é a
meta, não o superávit comercial.

{chr(10).join(linhas_ext)}

## A vulnerabilidade externa não some em 2003–2016

A balança comercial continua positiva até 2012, mas as **transações
correntes** — o conceito que fecha o BP — só são superavitárias em
2003–2006. Em 2007 já voltam a déficit (−{_usd_bi(2_754)}). Em 2010 o
rombo é {_usd_bi(86_718)}; em 2014, {_usd_bi(110_494)}. Serviços, rendas
e a absorção doméstica (PIB de 7,5% em 2010, crédito e BNDES) comem o
superávit de bens.

Reservas altas (US$ 350–370 bi em 2011–16) impedem a *escassez de 2000*.
Não impedem o outro risco: financiar déficit em conta corrente com fluxo
de portfólio e, quando o Fed aperta (taper de 2013) e o fiscal doméstico
piora (2014–16), o real deprecia e o IPCA sobe. Em 2015 o IPCA vai a
10,67% e a Selic a 14,25%. Isso não é “captar dólar para quimioterápico”.
É pass-through cambial + inflação de demanda/inércia com o mandato de
metas ainda em pé.

## Por que o acumulado é da ordem de 418–460%

Selic efetiva (SGS 4390), jan/2003–dez/2016: fator {_br_num(fat_s, 4)} →
**+{_br_num(y16['selic_acum_pct'], 1)}%** (média geométrica
{_br_num(geo_s, 2)}% a.a.). IPCA no mesmo período: **+{_br_num(y16['ipca_acum_pct'], 1)}%**
({_br_num(geo_i, 2)}% a.a.). Juro real composto: **+{_br_num(real, 1)}%**
(~6,4% a.a.). Os “418%” do enunciado são o mesmo objeto — o produto das
taxas — com pequena diferença de série ou de janela. Não são “a Selic
subiu 418%”.

Comparar esse acumulado com China, Índia e Romênia sem o IPCA de cada
país mistura regimes. A taxa de *empréstimo* do Banco Mundial (não é a
taxa básica) acumula +116% na China, +324% na Índia e +533% na Romênia;
o CPI acumula +45%, +162% e +113%. O IPCA brasileiro no período (+134%)
já é o triplo da inflação chinesa. Parte do “quase o quádruplo da China”
é simplesmente inflação e meta mais altas (centro 4,5% ±2 p.p.), não
uma Selic ainda usada para fechar o BP.

A outra parte — o juro *real* brasileiro, ainda alto — tem causas
domésticas, não cambiais:

1. **Crédito direcionado.** Enquanto BNDES, poupança, rural e habitacional
   pagam TJLP/TR, a Selic só opera sobre uma fatia do crédito. Para o
   mesmo IPCA, a taxa livre precisa ser mais alta. É o spread Selic−TJLP
   que o TCU 2010 já mede em R$ 14,2 bi ao ano sobre o estoque Tesouro–
   BNDES.
2. **Indexação residual.** Contratos, tarifas e salários ainda reajustam
   por inflação passada. A inércia obriga juro real positivo por mais
   tempo.
3. **Demanda 2006–10.** Reservas e superávit *afrouxam* a restrição
   externa e, com isso, *permitem* o boom de crédito e gasto. O TCU
   (p. 33) chama a retirada tardia do estímulo de 2008 de causa do
   IPCA de 5,91% em 2010. Superávit comercial não pede Selic baixa;
   às vezes pede o contrário, porque folga o BP e esquenta a demanda.
4. **A tentativa de cortar rápido já foi feita — e falhou.** Em 2011–12
   a “nova matriz” leva a meta a 7,25% sem ajuste fiscal. A Selic de um
   dígito dura de março/2012 a novembro/2013 e, somada à janela
   julho/2009–junho/2010, dá cerca de 950 dias em 14 anos (~19% do
   período). Em 2015–16 o Copom paga a conta: 14,25% e IPCA de 10,67%.
   Esse yo-yo não é mistério externo. É a função de reação do regime de
   metas quando a inflação foge do centro — e quando se tenta forçar o
   juro para baixo sem o fiscal acompanhar.

## Resposta direta

O Brasil não acelerou a redução da Selic a partir de 2003/04 **porque o
problema que a Selic passou a resolver já não era a falta de dólares**.
A falta de dólares acabou (balança + reservas). Sobrou a inflação, o
crédito direcionado, a indexação e, depois de 2008, um déficit em conta
corrente financiado por fluxo volátil. A Selic caiu quando o IPCA
permitiu e subiu quando o IPCA — não o BP — exigiu. Por isso ficou
pouquíssimo tempo na casa de um dígito entre 2003 e 2016, e por isso o
juro composto do período é da ordem de 400–460%, várias vezes o da China:
não como instrumento de captação de divisas, e sim como o preço de
quatorze anos de meta de inflação numa economia ainda indexada e com
metade do crédito fora da Selic.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_SELIC_BP_2003_2016.md` | Esta análise |
| `output/grafico_selic_bp_1999_2016.png` | Contas externas + Selic/IPCA |
| `scripts/selic_bp_2003_2016_dados.py` | Séries SGS 22710, 22701, 3546, 4390, 433 |

```bash
python3 scripts/build_tcu_cg_2010.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def analisar(pasta: Path | None = None) -> dict:
    pasta = pasta or Path("output")
    ext = df_externo()
    si = df_selic_ipca()
    png = grafico(ext, si, pasta / "grafico_selic_bp_1999_2016.png")
    md = escrever_markdown(pasta / "TCU_CG_2010_SELIC_BP_2003_2016.md", ext, si)
    return {"externo": ext, "selic_ipca": si, "comparacao": pd.DataFrame(D.comparacao_banco_mundial()), "png": png, "md": md}
