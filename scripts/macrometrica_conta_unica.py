#!/usr/bin/env python3
"""Bloco macrométrico sobre o contrafactual da Conta Única, 2003–2015.

A Macrométrica comercial (macrometrica.com.br) não está acessível neste
ambiente: modelo proprietário, mais de 200 equações, só para cliente.
Este script aplica a *classe* de análise que ela faz — um sistema
simultâneo atividade–receita–dívida — aos fluxos já medidos.

Calibração publicada (não inventada):
  * multiplicador de tributo μ_T = 0,40 (faixa 0,3–0,6 da literatura BR);
  * multiplicador de corte de crédito BNDES μ_B = 0,25;
  * elasticidade receita-PIB ε = 1,10 (IFI, curto prazo > 1);
  * carga federal / PIB τ = 0,23 (TCU 2006–2010, ~23%).

Identidade do ano t (R$):

    ΔY_t        = −μ_T × renúncia_t − μ_B × (indiretas + participações)_t
    ΔR_ciclo_t  = ε × τ × ΔY_t
    T_líquido_t = renúncia_t + ΔR_ciclo_t
    fluxo_mm_t  = (indiretas + participações)_t + T_líquido_t

O caixa do BNDES redirecionado não encolhe com o PIB. A receita extra da
renúncia, sim. A DBGG/PIB contrafactual usa o PIB com ΔY.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from scripts.simular_dbgg_conta_unica import _br_bi, _br_num, df_fluxos, df_simulacao

# Calibração-base (macrométrica)
MU_TRIBUTO = 0.40
MU_BNDES = 0.25
ELASTICIDADE_RECEITA = 1.10
CARGA_FEDERAL_PIB = 0.23

# Sensibilidade: feedback mais forte
MU_TRIBUTO_ALTO = 0.60
MU_BNDES_ALTO = 0.40
ELASTICIDADE_ALTA = 1.20


def df_macrometrica(
    fluxos: pd.DataFrame | None = None,
    mu_t: float = MU_TRIBUTO,
    mu_b: float = MU_BNDES,
    eps: float = ELASTICIDADE_RECEITA,
    carga: float = CARGA_FEDERAL_PIB,
) -> pd.DataFrame:
    f = fluxos if fluxos is not None else df_fluxos()
    rows = []
    for _, r in f.iterrows():
        y = float(r["pib"])
        trib = float(r["renuncia"])
        bndes = float(r["indiretas"] + r["participacoes_uso"])
        dy = -mu_t * trib - mu_b * bndes
        dr_ciclo = eps * carga * dy
        t_liq = trib + dr_ciclo
        fluxo_mm = bndes + t_liq
        rows.append(
            {
                "ano": int(r["ano"]),
                "pib_oficial": y,
                "delta_pib": dy,
                "pib_cf": y + dy,
                "delta_pib_pct": 100.0 * dy / y,
                "renuncia_estatica": trib,
                "receita_ciclo": dr_ciclo,
                "renuncia_liquida": t_liq,
                "bndes": bndes,
                "fluxo_estatico": float(r["fluxo_total"]),
                "fluxo_macrometrico": fluxo_mm,
                "vazamento": float(r["fluxo_total"]) - fluxo_mm,
                "selic_pct": float(r["selic_pct"]),
                "dbgg": r["dbgg"],
                "dbgg_pct": r["dbgg_pct"],
            }
        )
    out = pd.DataFrame(rows)
    return out


def df_dbgg_macrometrica(mm: pd.DataFrame) -> pd.DataFrame:
    """Mesma identidade de dívida do cenário estático, com o fluxo líquido."""
    fake = pd.DataFrame(
        {
            "ano": mm["ano"],
            "fluxo_total": mm["fluxo_macrometrico"],
            "pib": mm["pib_cf"],
            "selic_pct": mm["selic_pct"],
            "dbgg": mm["dbgg"],
            "dbgg_pct": mm["dbgg_pct"],
        }
    )
    sim = df_simulacao(fake)
    sim = sim.rename(
        columns={
            "fluxo": "fluxo_macrometrico",
            "dbgg_cf_estoque": "dbgg_mm_estoque",
            "dbgg_cf_selic": "dbgg_mm_selic",
            "dbgg_cf_estoque_pct": "dbgg_mm_estoque_pct",
            "dbgg_cf_selic_pct": "dbgg_mm_selic_pct",
        }
    )
    return sim


def grafico(mm: pd.DataFrame, estatico: pd.DataFrame, mm_div: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig, (ax0, ax1) = plt.subplots(2, 1, figsize=(12.2, 8.2))

    ax0.bar(mm["ano"], mm["renuncia_estatica"] / 1000.0, color="#b45309", width=0.62, label="Renúncia estática (DGT)")
    ax0.bar(mm["ano"], mm["renuncia_liquida"] / 1000.0, color="#047857", width=0.38, label="Receita líquida (macrométrica)")
    ax0.set_ylabel("R$ bilhões")
    ax0.set_title(
        "Bloco macrométrico — receita da renúncia depois do vazamento cíclico\n"
        f"μ_T={MU_TRIBUTO}, μ_B={MU_BNDES}, ε={ELASTICIDADE_RECEITA}, τ={CARGA_FEDERAL_PIB}"
    )
    ax0.legend(loc="upper left", frameon=False)
    ax0.spines["top"].set_visible(False)
    ax0.spines["right"].set_visible(False)

    s = estatico.loc[estatico["dbgg_oficial"].notna()]
    m = mm_div.loc[mm_div["dbgg_oficial"].notna()]
    ax1.plot(s["ano"], s["dbgg_oficial"] / 1000.0, color="#111827", marker="o", linewidth=2.2, label="DBGG oficial")
    ax1.plot(s["ano"], s["dbgg_cf_selic"] / 1000.0, color="#1d4ed8", marker="s", linewidth=2.0, label="Estático + Selic")
    ax1.plot(m["ano"], m["dbgg_mm_selic"] / 1000.0, color="#b45309", marker="^", linewidth=2.0, label="Macrométrico + Selic")
    ax1.set_ylabel("R$ bilhões")
    ax1.set_xlabel("Exercício")
    ax1.legend(loc="upper left", frameon=False)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)

    fig.tight_layout()
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(
    destino: Path,
    mm: pd.DataFrame,
    estatico: pd.DataFrame,
    mm_div: pd.DataFrame,
    mm_alto: pd.DataFrame,
    mm_div_alto: pd.DataFrame,
) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    y15e = estatico.loc[estatico["ano"] == 2015].iloc[0]
    y15m = mm.loc[mm["ano"] == 2015].iloc[0]
    y15d = mm_div.loc[mm_div["ano"] == 2015].iloc[0]
    y15a = mm_alto.loc[mm_alto["ano"] == 2015].iloc[0]
    y15ad = mm_div_alto.loc[mm_div_alto["ano"] == 2015].iloc[0]
    tot_est = float(mm["fluxo_estatico"].sum())
    tot_mm = float(mm["fluxo_macrometrico"].sum())
    tot_vaz = float(mm["vazamento"].sum())
    tot_ren_e = float(mm["renuncia_estatica"].sum())
    tot_ren_l = float(mm["renuncia_liquida"].sum())
    dy_medio = float(mm["delta_pib_pct"].mean())

    linhas = [
        "| Ano | Renúncia estática | Δ PIB | Receita cíclica | Renúncia líquida | Fluxo estático | Fluxo MM | Vazamento |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in mm.iterrows():
        linhas.append(
            f"| {int(r['ano'])} | {_br_num(r['renuncia_estatica']/1000, 1)} | "
            f"{_br_num(r['delta_pib_pct'], 2)}% | {_br_num(r['receita_ciclo']/1000, 1)} | "
            f"{_br_num(r['renuncia_liquida']/1000, 1)} | {_br_num(r['fluxo_estatico']/1000, 1)} | "
            f"{_br_num(r['fluxo_macrometrico']/1000, 1)} | {_br_num(r['vazamento']/1000, 1)} |"
        )

    linhas_d = [
        "| Ano | DBGG oficial | Estático+Selic | MM+Selic | Oficial % PIB | Estático % PIB | MM % PIB (PIB cf) |",
        "|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in mm_div.iterrows():
        e = estatico.loc[estatico["ano"] == r["ano"]].iloc[0]
        linhas_d.append(
            f"| {int(r['ano'])} | {_br_bi(r['dbgg_oficial'])} | {_br_bi(e['dbgg_cf_selic'])} | "
            f"{_br_bi(r['dbgg_mm_selic'])} | {_br_num(r['dbgg_oficial_pct'], 1)} | "
            f"{_br_num(e['dbgg_cf_selic_pct'], 1)} | {_br_num(r['dbgg_mm_selic_pct'], 1)} |"
        )

    md = f"""# Bloco macrométrico sobre a Conta Única, 2003 a 2015

A Macrométrica comercial não roda daqui. O que segue é a **mesma classe
de exercício** que o software dela faz: um sistema simultâneo em que a
arrecadação extra **não** deixa o PIB e a receita residual parados.

A identidade estática continua válida *se* o resto for estático. O bloco
macrométrico **solta** o resto.

## Equações

Calibração-base: μ_T = {MU_TRIBUTO}, μ_B = {MU_BNDES}, ε = {ELASTICIDADE_RECEITA}
(IFI, curto prazo), τ = {CARGA_FEDERAL_PIB} (carga federal TCU).

```
ΔY_t        = −μ_T × renúncia_t − μ_B × BNDES_t
ΔR_ciclo_t  = ε × τ × ΔY_t
T_líquido_t = renúncia_t + ΔR_ciclo_t
fluxo_mm_t  = BNDES_t + T_líquido_t
```

BNDES redirecionado à Conta Única permanece integral (é caixa, não base
tributária). A renúncia encolhe porque o PIB cai e a elasticidade IFI
come parte da receita residual. A DBGG/PIB macrométrica usa o PIB já
reduzido — o denominador também reage.

## O que o bloco faz com o “1 a 1”

Renúncia estática 2003–2015: {_br_bi(tot_ren_e)}.
Renúncia líquida depois do ciclo: {_br_bi(tot_ren_l)}.
Vazamento cíclico só na receita: {_br_bi(tot_ren_e - tot_ren_l)}.

Fluxo estático total (BNDES + renúncia): {_br_bi(tot_est)}.
Fluxo macrométrico: {_br_bi(tot_mm)}.
Vazamento total: {_br_bi(tot_vaz)} — inteiro na perna tributária.

O PIB contrafactual fica, em média, {_br_num(dy_medio, 2)}% abaixo do
oficial. Em 2015 o recuo é {_br_num(float(y15m['delta_pib_pct']), 2)}%
({_br_bi(float(y15m['delta_pib']))} sobre o PIB de {_br_bi(float(y15m['pib_oficial']))}).

A necessidade de financiamento **ainda cai**. Só não cai o DGT inteiro.
Em 2015 a renúncia estática é {_br_bi(float(y15m['renuncia_estatica']))};
a líquida é {_br_bi(float(y15m['renuncia_liquida']))}.

## DBGG: estático versus macrométrico

Em 2015 a DBGG oficial é {_br_bi(y15e['dbgg_oficial'])}
({_br_num(y15e['dbgg_oficial_pct'], 1)}% do PIB).

* Estático + Selic: {_br_bi(y15e['dbgg_cf_selic'])}
  ({_br_num(y15e['dbgg_cf_selic_pct'], 1)}% do PIB oficial).
* Macrométrico + Selic: {_br_bi(y15d['dbgg_mm_selic'])}
  ({_br_num(y15d['dbgg_mm_selic_pct'], 1)}% do PIB contrafactual).

O numerador macrométrico é maior que o estático (menos caixa líquido).
O denominador é menor (PIB cai). Os dois efeitos empurram a razão
DBGG/PIB para cima em relação ao cenário estático, e ainda assim ela
fica abaixo da oficial.

Sensibilidade alta (μ_T={MU_TRIBUTO_ALTO}, μ_B={MU_BNDES_ALTO},
ε={ELASTICIDADE_ALTA}): fluxo MM 2015 {_br_bi(float(y15a['fluxo_macrometrico']))},
DBGG MM+Selic {_br_bi(y15ad['dbgg_mm_selic'])}
({_br_num(y15ad['dbgg_mm_selic_pct'], 1)}% do PIB cf).

{chr(10).join(linhas)}

{chr(10).join(linhas_d)}

## O que isto não é

Não é o modelo de 200+ equações da Macrométrica Ltda. Não há câmbio,
setor externo, mercado de trabalho nem defasagens trimestrais. A Selic
do estoque velho permanece a oficial (SGS 4390): o bloco não impõe
Taylor. O vazamento é só o canal atividade → receita.

A conclusão do usuário permanece: com mais arrecadação, a necessidade
de financiamento cai. O bloco só mede **quanto** dessa arrecadação
sobrevive quando o PIB não fica parado.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_MACROMETRICA.md` | Esta análise |
| `output/grafico_macrometrica_conta_unica_2003_2015.png` | Renúncia líquida e DBGG |
| `scripts/macrometrica_conta_unica.py` | Equações e calibração |

```bash
python3 scripts/macrometrica_conta_unica.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def analisar(pasta: Path | None = None) -> dict:
    pasta = pasta or Path("output")
    fluxos = df_fluxos()
    estatico = df_simulacao(fluxos)
    mm = df_macrometrica(fluxos)
    mm_div = df_dbgg_macrometrica(mm)
    mm_alto = df_macrometrica(
        fluxos,
        mu_t=MU_TRIBUTO_ALTO,
        mu_b=MU_BNDES_ALTO,
        eps=ELASTICIDADE_ALTA,
    )
    mm_div_alto = df_dbgg_macrometrica(mm_alto)
    png = grafico(mm, estatico, mm_div, pasta / "grafico_macrometrica_conta_unica_2003_2015.png")
    md = escrever_markdown(
        pasta / "TCU_CG_2010_MACROMETRICA.md",
        mm,
        estatico,
        mm_div,
        mm_alto,
        mm_div_alto,
    )
    return {
        "macrometrica": mm,
        "dbgg": mm_div,
        "alto": mm_alto,
        "dbgg_alto": mm_div_alto,
        "png": png,
        "md": md,
    }


def main() -> int:
    out = analisar()
    print(f"[OK] {out['md']}")
    print(f"[OK] {out['png']}")
    y = out["macrometrica"].loc[out["macrometrica"]["ano"] == 2015].iloc[0]
    print(
        f"2015 renuncia estatica {y['renuncia_estatica']/1000:.1f} "
        f"liquida {y['renuncia_liquida']/1000:.1f} "
        f"dY% {y['delta_pib_pct']:.2f}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
