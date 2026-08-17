#!/usr/bin/env python3
"""Reservas internacionais (TCU p. 43) × agregados M1–M4 × Selic/repos.

O fator setor externo da p. 35 é o fluxo em reais da compra de dólares
pelo Bacen. As reservas da p. 43 são o estoque em dólares. Os agregados
mostram onde essa liquidez foi parar depois da esterilização.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from scripts import tcu_cg_2010_dados as D
from scripts.tcu_cg_2010_dados import selic_anual


def df_reservas() -> pd.DataFrame:
    df = pd.DataFrame(D.reservas_internacionais_liquidez())
    df["reservas_r_mi"] = df["reservas_usd_mi"] * df["ptax_fim"]
    df["delta_usd_mi"] = df["reservas_usd_mi"].diff()
    df["delta_r_mi"] = df["reservas_r_mi"].diff()
    return df


def df_agregados() -> pd.DataFrame:
    df = pd.DataFrame(D.agregados_monetarios_dezembro())
    for col in ("base", "m1_restrito", "m1", "m2", "m3", "m4"):
        df[f"var_{col}"] = df[col].diff()
        df[f"var_{col}_pct"] = df[col].pct_change() * 100.0
    df["m3_menos_m2"] = df["m3"] - df["m2"]
    df["m2_menos_m1"] = df["m2"] - df["m1"]
    df["mult_m1"] = df["m1"] / df["base"]
    df["mult_m3"] = df["m3"] / df["base"]
    df["mult_m4"] = df["m4"] / df["base"]
    return df


def df_compromissadas_dezembro() -> pd.DataFrame:
    rows = []
    for r in D.operacoes_compromissadas_prazos():
        if r["periodo"].startswith("dez/"):
            rows.append({"ano": 2000 + int(r["periodo"][-2:]), "compromissadas": r["total"]})
    return pd.DataFrame(rows)


def df_quadro() -> pd.DataFrame:
    res = df_reservas().set_index("ano")
    agg = df_agregados().set_index("ano")
    ext = pd.DataFrame(D.fatores_base_monetaria()).set_index("ano")
    sel = pd.DataFrame(selic_anual(2003, 2010)).set_index("ano")
    comp = df_compromissadas_dezembro().set_index("ano")
    tit = ext.reindex(range(2002, 2011))
    out = pd.DataFrame({"ano": list(range(2002, 2011))}).set_index("ano")
    out["reservas_usd_mi"] = res["reservas_usd_mi"]
    out["ptax_fim"] = res["ptax_fim"]
    out["reservas_r_mi"] = res["reservas_r_mi"]
    out["setor_externo"] = tit["setor_externo"]
    out["titulos_publicos"] = tit["titulos_publicos"]
    out["var_base"] = tit["var_base"]
    out["base"] = agg["base"]
    out["m1_restrito"] = agg["m1_restrito"]
    out["m1"] = agg["m1"]
    out["m2"] = agg["m2"]
    out["m3"] = agg["m3"]
    out["m4"] = agg["m4"]
    out["m3_menos_m2"] = agg["m3_menos_m2"]
    out["mult_m1"] = agg["mult_m1"]
    out["mult_m3"] = agg["mult_m3"]
    out["compromissadas"] = comp["compromissadas"]
    out["selic_fim"] = sel["selic_fim"]
    out["selic_media"] = sel["selic_media"]
    return out.reset_index()


def _br_bi(valor_milhoes: float, casas: int = 1) -> str:
    return (
        f"R$ {valor_milhoes / 1000.0:,.{casas}f} bi"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def _br_num(valor: float, casas: int = 1) -> str:
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _br_usd(valor: float, casas: int = 1) -> str:
    return f"US$ {valor / 1000.0:,.{casas}f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def grafico_reservas_agregados(df: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    anos = df["ano"].tolist()
    x = list(range(len(anos)))
    fig, axes = plt.subplots(
        3,
        1,
        figsize=(12.2, 10.4),
        gridspec_kw={"height_ratios": [1.05, 1.15, 1.05]},
    )

    ax0 = axes[0]
    ax0b = ax0.twinx()
    ax0.plot(
        x,
        df["reservas_usd_mi"] / 1000.0,
        color="#047857",
        marker="o",
        linewidth=2.2,
        label="Reservas (US$ bi) — p. 43",
    )
    ext = [v / 1000.0 if pd.notna(v) else 0.0 for v in df["setor_externo"]]
    cores_ext = ["#047857" if v >= 0 else "#b45309" for v in ext]
    ax0b.bar(x, ext, color=cores_ext, alpha=0.35, width=0.62, label="Setor externo (R$ bi) — p. 35")
    ax0.set_ylabel("Reservas (US$ bilhões)")
    ax0b.set_ylabel("Setor externo (R$ bilhões)")
    ax0.set_title(
        "Reservas internacionais, agregados M1–M4 e esterilização — 2002 a 2010\n"
        "TCU p. 43 (reservas) · p. 35 (setor externo) · Bacen SGS (M1–M4, base, PTAX)"
    )
    h1, l1 = ax0.get_legend_handles_labels()
    h2, l2 = ax0b.get_legend_handles_labels()
    ax0.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False)
    ax0.set_xticks(x, [str(a) for a in anos])
    ax0.spines["top"].set_visible(False)

    ax1 = axes[1]
    series = [
        ("m1", "M1", "#b45309"),
        ("m2", "M2", "#1d4ed8"),
        ("m3", "M3", "#6d28d9"),
        ("m4", "M4", "#111827"),
    ]
    for col, rotulo, cor in series:
        ax1.plot(x, df[col] / 1000.0, marker="o", linewidth=2.0, color=cor, label=rotulo)
    ax1.set_ylabel("R$ bilhões")
    ax1.set_xticks(x, [str(a) for a in anos])
    ax1.legend(loc="upper left", frameon=False, ncol=4)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.set_title("Agregados monetários — saldo de dezembro (metodologia Bacen 2018)")

    ax2 = axes[2]
    ax2.plot(x, df["base"] / 1000.0, color="#111827", marker="o", linewidth=2.0, label="Base monetária")
    ax2.plot(
        x,
        df["reservas_r_mi"] / 1000.0,
        color="#047857",
        marker="s",
        linewidth=2.0,
        label="Reservas × PTAX (R$)",
    )
    ax2.plot(
        x,
        df["compromissadas"] / 1000.0,
        color="#1d4ed8",
        marker="D",
        linewidth=2.0,
        label="Compromissadas (p. 36)",
    )
    ax2.set_ylabel("R$ bilhões")
    ax2.set_xlabel("Dezembro de cada exercício")
    ax2.set_xticks(x, [str(a) for a in anos])
    ax2.legend(loc="upper left", frameon=False)
    ax2.spines["top"].set_visible(False)
    ax2.spines["right"].set_visible(False)
    ax2.set_title("O espelho em reais: base, reservas convertidas e estoque de repos")

    fig.subplots_adjust(left=0.08, right=0.92, top=0.92, bottom=0.06, hspace=0.42)
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(destino: Path, df: pd.DataFrame) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    linhas_res = [
        "| Ano | Reservas US$ | Δ US$ | PTAX 31/12 | Reservas R$ | Setor externo | Títulos | Compromissadas | Selic fim |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    prev_usd = None
    for _, r in df.iterrows():
        d_usd = "" if prev_usd is None else _br_num(r["reservas_usd_mi"] - prev_usd, 0)
        prev_usd = r["reservas_usd_mi"]
        ext = "—" if pd.isna(r["setor_externo"]) else _br_num(r["setor_externo"], 0)
        tit = "—" if pd.isna(r["titulos_publicos"]) else _br_num(r["titulos_publicos"], 0)
        comp = "—" if pd.isna(r["compromissadas"]) else _br_num(r["compromissadas"], 0)
        sel = "—" if pd.isna(r["selic_fim"]) else _br_num(r["selic_fim"], 2)
        linhas_res.append(
            f"| {int(r['ano'])} | {_br_num(r['reservas_usd_mi'], 0)} | {d_usd} | "
            f"{_br_num(r['ptax_fim'], 4)} | {_br_num(r['reservas_r_mi'], 0)} | "
            f"{ext} | {tit} | {comp} | {sel} |"
        )

    linhas_m = [
        "| Ano | Base | M1 restrito | M1 | M2 | M3 | M4 | M3−M2 | M3/base |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in df.iterrows():
        linhas_m.append(
            f"| {int(r['ano'])} | {_br_num(r['base'], 0)} | {_br_num(r['m1_restrito'], 0)} | "
            f"{_br_num(r['m1'], 0)} | {_br_num(r['m2'], 0)} | {_br_num(r['m3'], 0)} | "
            f"{_br_num(r['m4'], 0)} | {_br_num(r['m3_menos_m2'], 0)} | "
            f"{_br_num(r['mult_m3'], 2)} |"
        )

    y02 = df.loc[df["ano"] == 2002].iloc[0]
    y03 = df.loc[df["ano"] == 2003].iloc[0]
    y07 = df.loc[df["ano"] == 2007].iloc[0]
    y10 = df.loc[df["ano"] == 2010].iloc[0]
    ext_0307 = float(df.loc[df["ano"].between(2003, 2007), "setor_externo"].sum())
    ext_0310 = float(df.loc[df["ano"].between(2003, 2010), "setor_externo"].sum())

    md = f"""# Reservas internacionais, agregados M1–M4 e a gestão da Selic (2002–2010)

**Fontes:** TCU, *Contas do Governo da República 2010*, gráfico da **p. 43** (reservas, conceito de liquidez), quadro da **p. 35** (fator setor externo) e quadro da **p. 36** (compromissadas). Estoques oficiais: Bacen SGS 3546 (reservas), SGS 1 (PTAX), SGS 1782 (base), SGS 1785 (M1 restrito) e SGS 27791/27810/27813/27815 (M1–M4, metodologia 2018 retroagida a dez/2001).
**Unidades:** reservas em US$ milhões; demais fluxos e estoques em R$ milhões, salvo indicação.

O fator setor externo saiu de **{_br_bi(643)}** em 2003 para **{_br_bi(155_390)}** em 2007 — não de “155.390 reais”. A unidade do quadro da p. 35 é R$ milhões. Esse salto é o mesmo movimento do gráfico da p. 43: as reservas (liquidez) passam de {_br_usd(y03['reservas_usd_mi'])} em 2003 para {_br_usd(y07['reservas_usd_mi'])} em 2007.

## O mecanismo legal e contábil

No recorte 2003–2010 as reservas em moeda estrangeira são **ativo do Banco Central**, não do Tesouro. A Lei 4.595/1964 (arts. 10 e 11) atribui ao Bacen o monopólio da emissão e a administração das reservas. Comprar dólares no interbancário, portanto, é pagar com reais que o próprio Bacen cria. Esse é exatamente o fator “setor externo” da p. 35: a contrapartida em reais da variação das reservas transacionada no mercado de câmbio.

A Lei 11.803/2008 **não transfere a propriedade** das reservas ao Tesouro. Ela equaliza o resultado do Bacen com a União (o TCU registra equalização cambial negativa de R$ 52,2 bi em 2009 e R$ 48,5 bi em 2010). O custo de carregar as reservas vira despesa fiscal explícita; o canal monetário permanece o mesmo: dólar entra, real é emitido.

No sistema de **depósitos fracionários**, o real recém-criado vira reserva bancária e, se ficar livre, lastreia um múltiplo de depósitos (M1) e, na sequência, poupança, títulos privados, fundos e repos (M2 a M4). A tese de que o acúmulo de reservas incrementa a liquidez ampla está correta. O que a p. 33 e a p. 36 acrescentam é *onde* essa liquidez foi parar: o Copom não deixa o real livre na overnight. Ele o estaciona em **operações compromissadas** para a Selic efetiva sentar na meta.

Três camadas, portanto:

1. **Estoque em dólares** (p. 43) — o ativo de reservas.
2. **Fluxo em reais** (p. 35, setor externo) — a emissão que paga esse ativo.
3. **Estoque de liquidez ampla** (M1–M4) e **estoque de repos** (p. 36) — o destino dessa emissão depois da esterilização.

O estoque em dólares e o fluxo em reais não são iguais. As reservas também rendem juros, sofrem avaliação e incluem operações que não passam pelo interbancário no mesmo exercício. Em 2007 o setor externo injeta {_br_bi(155_390)} enquanto as reservas sobem {_br_usd(180_334 - 85_839)}. Convertido pela PTAX de 31/12/2007 (1,7713), o estoque em reais vai a {_br_bi(float(y07['reservas_r_mi']))}. A apreciação do real (PTAX 2,89 em 2003 → 1,77 em 2007) *reduz* o valor em reais do estoque mesmo quando o volume em dólares explode. O fator da p. 35 mede a emissão, não a marcação a mercado.

## O gráfico da p. 43 em números

{chr(10).join(linhas_res)}

De 2002 a 2010 as reservas vão de {_br_usd(float(y02['reservas_usd_mi']))} para {_br_usd(float(y10['reservas_usd_mi']))} (×{_br_num(float(y10['reservas_usd_mi']/y02['reservas_usd_mi']), 2)}; o TCU arredonda 2010 para US$ 288,6 bi, +20,7% sobre 2009). O degrau decisivo é 2006–2007: {_br_usd(85_839)} → {_br_usd(180_334)}. É o ano do recorde do setor externo.

Soma do fator setor externo: 2003–2007 = {_br_bi(ext_0307)}; 2003–2010 = {_br_bi(ext_0310)} (único ano negativo: 2008, −{_br_bi(12_124)}).

## Agregados monetários M1 a M4

{chr(10).join(linhas_m)}

M1 restrito é a série contemporânea de 2010 (SGS 1785: papel-moeda em poder do público + depósitos à vista). M1–M4 são a série oficial vigente, com revisão metodológica de 2018 retroagida a 2001 — a mesma nestagem (M3 contém as compromissadas), com universo maior de emissores.

Variação 2002–2010:

| Agregado | dez/2002 | dez/2010 | Variação | Múltiplo |
|---|---:|---:|---:|---:|
| Base monetária | {_br_bi(float(y02['base']))} | {_br_bi(float(y10['base']))} | {_br_bi(float(y10['base']-y02['base']))} | ×{_br_num(float(y10['base']/y02['base']), 2)} |
| M1 restrito | {_br_bi(float(y02['m1_restrito']))} | {_br_bi(float(y10['m1_restrito']))} | {_br_bi(float(y10['m1_restrito']-y02['m1_restrito']))} | ×{_br_num(float(y10['m1_restrito']/y02['m1_restrito']), 2)} |
| M1 | {_br_bi(float(y02['m1']))} | {_br_bi(float(y10['m1']))} | {_br_bi(float(y10['m1']-y02['m1']))} | ×{_br_num(float(y10['m1']/y02['m1']), 2)} |
| M2 | {_br_bi(float(y02['m2']))} | {_br_bi(float(y10['m2']))} | {_br_bi(float(y10['m2']-y02['m2']))} | ×{_br_num(float(y10['m2']/y02['m2']), 2)} |
| M3 | {_br_bi(float(y02['m3']))} | {_br_bi(float(y10['m3']))} | {_br_bi(float(y10['m3']-y02['m3']))} | ×{_br_num(float(y10['m3']/y02['m3']), 2)} |
| M4 | {_br_bi(float(y02['m4']))} | {_br_bi(float(y10['m4']))} | {_br_bi(float(y10['m4']-y02['m4']))} | ×{_br_num(float(y10['m4']/y02['m4']), 2)} |
| Reservas (US$) | {_br_usd(float(y02['reservas_usd_mi']))} | {_br_usd(float(y10['reservas_usd_mi']))} | {_br_usd(float(y10['reservas_usd_mi']-y02['reservas_usd_mi']))} | ×{_br_num(float(y10['reservas_usd_mi']/y02['reservas_usd_mi']), 2)} |

A base fecha com o TCU: 2003 = −R$ 83 milhões; 2007 = +{_br_bi(25_516)}; 2010 = +{_br_bi(40_780)} (+24,6%).

## O fator externo influencia muito a base — o que não explode é o saldo líquido

Não se lê o salto de {_br_bi(643)} (2003) para {_br_bi(155_390)} (2007) como “influência pequena”. O quadro da p. 35 é uma **identidade contábil**, não um ranking de importância relativa:

`Tesouro + títulos + setor externo + demais = variação da base`

Cada fator entra pelo valor cheio. Em 2007 o setor externo **criou** {_br_bi(155_390)} de base — o maior choque expansionista da série, 6,1 vezes a variação líquida daquele ano ({_br_bi(25_516)}). Sem ele, e na mesma configuração dos outros fatores, a base teria *contraído* cerca de {_br_bi(155_390 - 25_516)}.

O que é pequeno é outra coisa: o **saldo líquido** depois da compensação. No mesmo exercício o Tesouro retirou {_br_bi(55_600)} e os títulos/repos retiraram {_br_bi(73_974)}. A conta fecha em +{_br_bi(25_516)}. Influência bruta e variação líquida não são o mesmo objeto.

| 2007 | R$ milhões | Papel |
|---|---:|---|
| Setor externo | +155.390 | cria base (compra de dólares) |
| Tesouro | −55.600 | drena (Conta Única) |
| Títulos públicos | −73.974 | esteriliza (repos / mercado aberto) |
| Demais | −300 | residual |
| **Variação da base** | **+25.516** | o que sobra no estoque de M0 |

Três afirmações distintas, portanto:

1. **Sobre a base (M0):** o fator externo influencia *muito*. É o motor. O net cresce pouco porque os outros fatores compensam no mesmo ano.
2. **Sobre o M1:** a influência é *atenuada*, não nula. M1 sobe 33% em 2007, mas não na mesma ordem do choque de R$ 155,4 bi, porque o compulsório e o open market não deixam a reserva livre lastrear o múltiplo inteiro.
3. **Sobre M3, repos e o custo da Selic:** a influência é *plena e persistente*. O real que saiu da base foi para as compromissadas (R$ 60,0 bi → R$ 165,8 bi) e permanece no M3. É aí que o salto de 643 para 155.390 continua a constranger a política monetária depois que a linha “variação da base” já fechou.

Dizer “influencia, mas não tanto” mistura (1) com o net de (1). O correto é: influencia tanto que o Bacen teve de montar um estoque de repos e, depois, de compulsório, para a Selic não perder o controle da overnight.

## A relação reservas → base → agregados → Selic/repos

A compra de reservas **cria** base. O depósito fracionário **pode** multiplicá-la em M1. A Mesa do mercado aberto **impede** que essa base livre derrube a overnight. O instrumento é a compromissada: o Bacen toma o real de volta, paga Selic, e devolve um título com recompra. A base se contrai (fator títulos); o M3 **não**. Na metodologia do Bacen, operações compromissadas entram no M3. Esterilizar a base é, em boa medida, **reclassificar** liquidez de reserva bancária para repo — ainda dinheiro, só que remunerado na meta.

Por isso o múltiplo M3/base sobe de {_br_num(float(y02['mult_m3']), 2)} em 2002 para {_br_num(float(y07['mult_m3']), 2)} em 2007 e {_br_num(float(y10['mult_m3']), 2)} em 2010, enquanto a base apenas pouco mais que dobra. A liquidez “brutal” está na camada ampla (M3−M2 vai de {_br_bi(float(y02['m3_menos_m2']))} para {_br_bi(float(y10['m3_menos_m2']))}), não no papel-moeda.

### 2003–2005 — reservas ainda pequenas, Selic no ciclo de aperto

Reservas saem de {_br_usd(37_823)} para {_br_usd(53_799)}. O setor externo ainda é modesto em 2003–2004 e só ganha corpo em 2005 (+{_br_bi(52_395)}). A Selic faz o pico de 19,75%. M1 cresce pouco; M3 já sobe mais rápido. O preço (Selic) segura a multiplicação; a quantidade ainda não é o problema.

### 2006–2007 — o degrau da p. 43 e o caso-escola da esterilização

Reservas quase dobram em 2007. Setor externo {_br_bi(155_390)}. Títulos contraem {_br_bi(73_974)}. Compromissadas: {_br_bi(60_030)} (dez/06) → {_br_bi(165_813)} (dez/07). Selic *cai* de 13,25% para 11,25%. M1 salta de {_br_bi(176_890)} para {_br_bi(235_075)} (+33%); M3 de {_br_bi(1_365_958)} para {_br_bi(1_600_006)}.

A leitura conjunta: o Bacen pode cortar a Selic *porque* enxuga a overnight com repos. Sem as compromissadas, os R$ 155,4 bi do setor externo teriam sobrado como reservas livres e a taxa efetiva teria ido ao piso. Com elas, a meta de 11,25% é cumprida e a liquidez migra para M3. O depósito fracionário opera, mas sobre uma base que o open market recicla todos os dias.

### 2008 — o teste da crise

Reservas sobem pouco em dólares ({_br_usd(180_334)} → {_br_usd(193_783)}) e muito em reais, porque o câmbio vai a 2,337. O setor externo **contrai** a base (−{_br_bi(12_124)}): o Bacen para de comprar e vende dólares. M1 cai ({_br_bi(235_075)} → {_br_bi(227_167)}). M2 explode ({_br_bi(779_566)} → {_br_bi(1_086_785)}) — fuga para poupança. Compromissadas continuam a subir ({_br_bi(165_813)} → {_br_bi(300_491)}): agora o repo injeta, não esteriliza. A Selic fica em 13,75% até janeiro de 2009. Preço alto, quantidade de socorro — o instrumento secundário da p. 33.

### 2009–2010 — estímulo, depois troca de instrumento

Reservas voltam a subir ({_br_usd(238_520)} → {_br_usd(288_575)}). Setor externo de novo expansionista. Compromissadas fazem o pico em dez/09 ({_br_bi(427_874)}) e caem em 2010 ({_br_bi(259_248)}) quando o compulsório assume o dreno (−{_br_bi(236_911)} em depósitos). M1, M2, M3 e M4 seguem em alta. A Selic só sai de 8,75% em abril de 2010 e para em 10,75%.

A troca compulsório ↔ repo *aperta mais* o multiplicador do que o repo sozinho: reserva compulsória não se empresta. Mesmo assim M3 cresce {_br_bi(float(y10['m3']-df.loc[df['ano']==2009,'m3'].iloc[0]))} em 2010, porque o setor externo continua a criar reais e porque M3 inclui os repos que restam e as quotas de fundos.

## Efeitos sobre a gestão da Selic e das compromissadas

O acúmulo de reservas **não determina** o nível da Selic. Determina o *volume* de operações que a Mesa precisa fazer para a Selic escolhida ser eficaz.

- **Selic** é o preço que torna o banco indiferente entre emprestar na overnight e fazer repo com o Bacen. Enquanto houver excesso de reais criado pela compra de dólares, a meta só se realiza se o Bacen oferecer um estoque suficiente de compromissadas a essa taxa.
- **Compromissadas** são o estoque-espelho da intervenção cambial (líquido do dreno do Tesouro e, em 2010, do compulsório). Dez/06 {_br_bi(60_030)} → dez/09 {_br_bi(427_874)} acompanha o gráfico da p. 43. Não é coincidência: é a contabilidade da esterilização.
- **Custo quasi-fiscal:** o Bacen paga Selic no repo e recebe o rendimento das reservas em dólar. A Lei 11.803/2008 transfere esse resultado à União. Quanto maior o estoque de repos e quanto maior a Selic, maior a conta — o TCU já a mede na equalização cambial e, no outro lado do balanço, no passivo de compromissadas da DBGG (R$ 454,7 bi em 2009 e R$ 288,7 bi em 2010 no quadro da dívida).
- **Limite do corte de juros (2006–07):** a Selic pode cair enquanto as reservas explodem *somente* porque os repos crescem. O risco não é “a base sair do controle” no sentido de M0; é a liquidez ampla (M3/M4) e o crédito que ela alimenta, com defasagem, pressionarem o IPCA — o juízo da p. 33 sobre 2010.
- **Limite da alta de juros (2008 e 2010):** subir a Selic encarece o estoque de repos que o próprio Bacen criou ao comprar reservas. Daí a tentação, em 2010, de trocar repo por compulsório: o compulsório não paga Selic (ou paga menos) e trava o multiplicador. O TCU descreve essa troca na p. 34.

A resposta à pergunta, em uma frase: **sim, a legislação fazia da compra de reservas uma emissão de reais, e o depósito fracionário levou essa emissão aos agregados; não, isso não aparece como um M1 descontrolado, porque a Selic e as compromissadas reciclaram a maior parte da emissão para dentro do M3 — e é aí, na liquidez ampla e no custo de carregá-la, que a política monetária do período foi de fato constrangida.**

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_RESERVAS_M1M4.md` | Esta análise |
| `output/grafico_reservas_agregados_2002_2010.png` | Reservas × setor externo, M1–M4, base/repos |
| `output/TCU_CG_2010.xlsx` (abas `Reservas_Internacionais`, `Agregados_M1_M4`) | Séries oficiais |

```bash
python3 scripts/build_tcu_cg_2010.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def analisar(pasta: Path | None = None) -> dict:
    pasta = pasta or Path("output")
    df = df_quadro()
    png = grafico_reservas_agregados(df, pasta / "grafico_reservas_agregados_2002_2010.png")
    md = escrever_markdown(pasta / "TCU_CG_2010_RESERVAS_M1M4.md", df)
    return {
        "quadro": df,
        "reservas": df_reservas(),
        "agregados": df_agregados(),
        "png": png,
        "md": md,
    }
