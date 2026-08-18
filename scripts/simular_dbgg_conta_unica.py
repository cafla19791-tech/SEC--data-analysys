#!/usr/bin/env python3
"""Simula a DBGG se indiretas, participações e três renúncias fossem à Conta Única.

Período: 2003–2015. Reservas internacionais ficam de fora.
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from scripts import sim_dbgg_conta_unica_dados as D


def _br_num(valor: float | None, casas: int = 1) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _br_bi(valor_mi: float | None, casas: int = 1) -> str:
    if valor_mi is None or (isinstance(valor_mi, float) and pd.isna(valor_mi)):
        return "—"
    return f"R$ {valor_mi / 1000.0:,.{casas}f} bi".replace(",", "X").replace(".", ",").replace("X", ".")


def df_fluxos() -> pd.DataFrame:
    rows = []
    for ano in D.ANOS:
        ren = D.renuncia_reconstruida(ano)
        part = D.PARTICIPACOES_R_MI[ano]
        part_uso = 0.0 if part is None else float(part)
        indireta = D.INDIRETAS_R_MI[ano]
        fluxo = indireta + part_uso + ren["renuncia_pedida"]
        rows.append(
            {
                "ano": ano,
                "indiretas": indireta,
                "participacoes": part,
                "participacoes_uso": part_uso,
                "regional_funcao": ren["desenvolvimento_regional"],
                "zfm_alc": ren["zfm_alc"],
                "regional_ampla": ren["regional_ampla"],
                "inovacao": ren["inovacao"],
                "imunes": ren["imunes_isentas"],
                "renuncia": ren["renuncia_pedida"],
                "fluxo_total": fluxo,
                "tesouro_captacao": D.TESOURO_CAPTACAO_R_MI[ano],
                "tesouro_devolucao": D.TESOURO_DEVOLUCAO_R_MI.get(ano, 0.0),
                "pib": D.PIB_DEZ_R_MI[ano],
                "selic_pct": D.SELIC_EFETIVA_PCT[ano],
                "dbgg": D.DBGG_DEZ_R_MI.get(ano),
                "dbgg_pct": D.DBGG_PCT_PIB.get(ano),
            }
        )
    return pd.DataFrame(rows)


def df_simulacao(fluxos: pd.DataFrame | None = None) -> pd.DataFrame:
    """Identidade dos fatores contrafactuais da DBGG.

    Convenção: o fluxo do ano entra na Conta Única no fim do exercício e é
    usado para não emitir / resgatar DPF (a DBGG é bruta: caixa parado no
    Bacen reduz a DLSP, não a DBGG).

    * ``emissao_evitada`` = −fluxo do ano (fator “emissões líquidas”).
    * ``juros_evitados`` = Selic do ano × saldo poupado no fim do ano anterior.
    * ``dbgg_cf_estoque`` = DBGG oficial − acumulado dos fluxos (sem juros).
    * ``dbgg_cf_selic`` = DBGG oficial − saldo capitalizado à Selic.
    """
    f = fluxos if fluxos is not None else df_fluxos()
    acum = 0.0
    saldo = 0.0
    rows = []
    for _, r in f.iterrows():
        ano = int(r["ano"])
        fluxo = float(r["fluxo_total"])
        selic = float(r["selic_pct"]) / 100.0
        juros = saldo * selic
        dbgg = r["dbgg"]
        pib = float(r["pib"])
        acum += fluxo
        saldo = saldo * (1.0 + selic) + fluxo
        dbgg_estoque = None if pd.isna(dbgg) else float(dbgg) - acum
        dbgg_selic = None if pd.isna(dbgg) else float(dbgg) - saldo
        rows.append(
            {
                "ano": ano,
                "fluxo": fluxo,
                "emissao_evitada": -fluxo,
                "juros_evitados": juros,
                "acum_fluxo": acum,
                "saldo_selic": saldo,
                "dbgg_oficial": None if pd.isna(dbgg) else float(dbgg),
                "dbgg_cf_estoque": dbgg_estoque,
                "dbgg_cf_selic": dbgg_selic,
                "dbgg_oficial_pct": r["dbgg_pct"] if pd.notna(r["dbgg_pct"]) else None,
                "dbgg_cf_estoque_pct": None if dbgg_estoque is None else 100.0 * dbgg_estoque / pib,
                "dbgg_cf_selic_pct": None if dbgg_selic is None else 100.0 * dbgg_selic / pib,
                "delta_oficial": None,
            }
        )
    out = pd.DataFrame(rows)
    out["delta_oficial"] = out["dbgg_oficial"].diff()
    return out


def grafico_fluxos(fluxos: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(12.2, 5.6))
    x = fluxos["ano"].tolist()
    a = fluxos["indiretas"] / 1000.0
    b = fluxos["participacoes_uso"] / 1000.0
    c = fluxos["renuncia"] / 1000.0
    ax.bar(x, a, color="#1d4ed8", width=0.72, label="Indiretas BNDES")
    ax.bar(x, b, bottom=a, color="#7c3aed", width=0.72, label="Participações acionárias")
    ax.bar(x, c, bottom=a + b, color="#b45309", width=0.72, label="Três famílias de renúncia")
    ax.set_ylabel("R$ bilhões correntes")
    ax.set_title(
        "Fluxos que iriam à Conta Única — 2003 a 2015\n"
        "BNDES desembolsos mensais + BNDESPAR + RFB/IFI (reconstrução)"
    )
    ax.legend(loc="upper left", frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def grafico_dbgg(sim: pd.DataFrame, destino: Path) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    s = sim.loc[sim["dbgg_oficial"].notna()]
    fig, ax = plt.subplots(figsize=(12.2, 5.6))
    ax.plot(s["ano"], s["dbgg_oficial"] / 1000.0, color="#111827", marker="o", linewidth=2.2, label="DBGG oficial (SGS 13761)")
    ax.plot(s["ano"], s["dbgg_cf_estoque"] / 1000.0, color="#1d4ed8", marker="s", linewidth=2.0, label="Contrafactual sem juros (não emissão)")
    ax.plot(s["ano"], s["dbgg_cf_selic"] / 1000.0, color="#b45309", marker="^", linewidth=2.0, label="Contrafactual capitalizado à Selic")
    ax.set_ylabel("R$ bilhões")
    ax.set_title(
        "DBGG oficial e contrafactual se os fluxos fossem à Conta Única\n"
        "Metodologia Bacen 2008 — dezembro"
    )
    ax.legend(loc="upper left", frameon=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(destino, dpi=140)
    plt.close(fig)
    return destino


def escrever_markdown(destino: Path, fluxos: pd.DataFrame, sim: pd.DataFrame) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    tot_ind = float(fluxos["indiretas"].sum())
    tot_part = float(fluxos["participacoes_uso"].sum())
    tot_ren = float(fluxos["renuncia"].sum())
    tot_fluxo = float(fluxos["fluxo_total"].sum())
    tot_tes = float(fluxos["tesouro_captacao"].sum())
    y15 = sim.loc[sim["ano"] == 2015].iloc[0]
    y06 = sim.loc[sim["ano"] == 2006].iloc[0]
    y10f = fluxos.loc[fluxos["ano"] == 2010].iloc[0]
    y13f = fluxos.loc[fluxos["ano"] == 2013].iloc[0]

    linhas_f = [
        "| Ano | Indiretas | Participações | Regional ampla | Inovação | Imunes/isentas | Fluxo total | % PIB |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in fluxos.iterrows():
        part = "—" if pd.isna(r["participacoes"]) else _br_num(r["participacoes"] / 1000.0, 1)
        linhas_f.append(
            f"| {int(r['ano'])} | {_br_num(r['indiretas']/1000, 1)} | {part} | "
            f"{_br_num(r['regional_ampla']/1000, 1)} | {_br_num(r['inovacao']/1000, 1)} | "
            f"{_br_num(r['imunes']/1000, 1)} | {_br_num(r['fluxo_total']/1000, 1)} | "
            f"{_br_num(100.0 * r['fluxo_total']/r['pib'], 2)} |"
        )

    linhas_s = [
        "| Ano | Fluxo | Emissões evitadas | Juros evitados | Acum. fluxos | Saldo à Selic | DBGG oficial | CF estoque | CF Selic | Oficial % PIB | CF Selic % PIB |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for _, r in sim.iterrows():
        linhas_s.append(
            f"| {int(r['ano'])} | {_br_num(r['fluxo']/1000, 1)} | {_br_num(r['emissao_evitada']/1000, 1)} | "
            f"{_br_num(r['juros_evitados']/1000, 1)} | {_br_num(r['acum_fluxo']/1000, 1)} | "
            f"{_br_num(r['saldo_selic']/1000, 1)} | {_br_bi(r['dbgg_oficial'])} | "
            f"{_br_bi(r['dbgg_cf_estoque'])} | {_br_bi(r['dbgg_cf_selic'])} | "
            f"{_br_num(r['dbgg_oficial_pct'], 1)} | {_br_num(r['dbgg_cf_selic_pct'], 1)} |"
        )

    delta_pp = None
    if pd.notna(y15["dbgg_oficial_pct"]) and pd.notna(y15["dbgg_cf_selic_pct"]):
        delta_pp = float(y15["dbgg_oficial_pct"]) - float(y15["dbgg_cf_selic_pct"])

    md = f"""# Simulação: Conta Única e fatores da DBGG, 2003 a 2015

Reservas internacionais ficam de lado. A pergunta é outra: **como
evoluiriam os fatores condicionantes da dívida bruta do governo geral**
se três fluxos tivessem ido para a Conta Única do Tesouro, em vez de
sair como desembolso do BNDES ou como renúncia tributária.

Os três fluxos, de 2003 a 2015, somam **{_br_bi(tot_fluxo)}** correntes:

1. orçamento do BNDES em **operações indiretas** — {_br_bi(tot_ind)};
2. orçamento do BNDES em **participações acionárias** (BNDESPAR, 2007–2015)
   — {_br_bi(tot_part)};
3. receita que a RFB teria lançado **sem** os benefícios de desenvolvimento
   regional (função + ZFM/ALC), **sem** o gasto tributário de inovação
   (Lei 10.973/2004 operacionalizada pela Lei do Bem 11.196/2005) e **sem**
   os benefícios a entidades imunes e isentas — {_br_bi(tot_ren)}.

Em 2015 a DBGG oficial (SGS 13761) era {_br_bi(y15["dbgg_oficial"])}
({_br_num(y15["dbgg_oficial_pct"], 1)}% do PIB). No contrafactual com o
caixa usado para não emitir / resgatar DPF e o saldo capitalizado à
Selic efetiva, a DBGG cairia para {_br_bi(y15["dbgg_cf_selic"])}
({_br_num(y15["dbgg_cf_selic_pct"], 1)}% do PIB) — cerca de
**{_br_num(delta_pp, 1)} p.p. do PIB** a menos. Sem capitalizar juros, o
corte é o estoque acumulado dos fluxos: {_br_bi(y15["acum_fluxo"])}, e a
DBGG ficaria em {_br_bi(y15["dbgg_cf_estoque"])}.

A Selic entre as quatro mais altas do BIS em cada ano do período não
entra como causa das reservas. Entra aqui só como **preço do estoque
poupado**: cada real que tivesse ficado na Conta Única e evitado DPF
deixava de carregar esse juro.

## Identidade que a simulação usa

A DBGG é **bruta**. Depósito na Conta Única (haver no Bacen) reduz a
**DLSP**, não a DBGG. Para o fator “emissões líquidas” da Nota de
Política Fiscal mudar, o caixa tem de **não emitir ou resgatar** título.

Identidade anual (R$):

```
Δ DBGG_cf = Δ DBGG_oficial + emissão_evitada + (− juros_evitados)
emissão_evitada_t = − fluxo_t
juros_evitados_t  = Selic_t × saldo_selic_{{t−1}}
saldo_selic_t     = saldo_selic_{{t−1}} × (1+Selic_t) + fluxo_t
```

O fluxo entra no **fim** do ano (convenção conservadora: menos juros no
próprio exercício). Não há efeito-atividade: a RFB também calcula gasto
tributário com base estática.

Dois caminhos, de propósito distintos:

* **CF estoque** — o Tesouro só deixa de emitir o fluxo do ano. A DBGG
  cai 1 a 1 com o acumulado. É o fator “emissões líquidas”.
* **CF Selic** — o estoque poupado deixa de pagar Selic. Soma o fator
  “juros nominais”. É o canal em que a Selic alta do período (e o
  ranking no BIS) multiplica o resultado.

## 1. Indiretas do BNDES

Fonte: *Desembolsos Mensais* do portal de dados abertos, campo
`forma_de_apoio = INDIRETA`. Finame, Automático, Finem indireto, Exim
indireto e Cartão BNDES entram. O CSV “indiretas e produto” **não**
serve: ele zera máquinas/serviços.

Soma 2003–2015: **{_br_bi(tot_ind)}**. Pico em 2013
({_br_bi(float(y13f["indiretas"]))}). Em 2010,
{_br_bi(float(y10f["indiretas"]))} — o ano do gráfico TCU da base
monetária e do salto Tesouro→BNDES.

Isso é **orçamento/desembolso do Sistema BNDES**, não “recurso do
Tesouro”. Até 2007 a fonte dominante é FAT/PIS/próprios. As captações
Tesouro→BNDES (página oficial) só aparecem em 2008–2014 e somam
**{_br_bi(tot_tes)}** (incluindo R$ 24,7 bi da capitalização da
Petrobras em 2010). São fungíveis entre direta, indireta e renda
variável. Redirecionar *toda* a indireta de 2003–2007 à Conta Única
exige mudar a destinação legal do FAT/PIS, não só “deixar de emitir
DPF”.

## 2. Participações acionárias

Fonte: BNDESPAR, desembolsos via renda variável, apenas
`PARTICIPAÇÃO ACIONÁRIA`. Debêntures (R$ 17,9 bi em 2007–2015) e cotas
de fundo (R$ 3,0 bi) ficam de fora — não são capital acionário.

A base começa em **2007**. 2003–2006 não têm microdado comparável; a
simulação trata esses anos como zero e o total de {_br_bi(tot_part)} é
portanto um **piso**. O salto de 2010 ({_br_bi(float(y10f["participacoes_uso"]))})
é a capitalização da Petrobras / ofertas daquela janela.

## 3. Três famílias de renúncia

Lei 10.973/2004 é o marco da inovação; o item que a RFB mensura no DGT
é sobretudo a **Lei do Bem** (Lei 11.196/2005). Informática (Lei 8.248)
**não** entra. Desenvolvimento regional na acepção da pergunta inclui a
função RFB *e* ZFM/ALC — no DGT elas vêm em linhas separadas; a tabela
mostra a soma (“regional ampla”) e as duas partes no código.

Âncora oficial de 2015 (IFI NT 17, bases efetivas, R$ milhões):
Desenvolvimento Regional 5.899, ZFM/ALC 23.232, Imunes/isentas 19.505,
Pesquisa científica e inovação 3.392.

A série 2003–2014 **não** é o DGT item a item de cada ano (os
demonstrativos anuais por modalidade não estão reproduzidos no
repositório). É uma reconstrução: a participação de cada família no PIB
de 2015 é aplicada ao PIB de cada ano (SGS 4382). Inovação = 0 em
2003–2005 (Lei do Bem ainda não opera). Em 2015 usam-se os valores
exatos da IFI. Isso preserva a ordem de grandeza e **não inventa** um
DGT anual que não foi transcrito.

Cheque de consistência: a isenção patronal das filantrópicas no TCU
2006–2010 (R$ 3,8 a 6,4 bi) é um *subconjunto* previdenciário das
imunes/isentas, e fica abaixo da reconstrução — como deve.

DGT PLOA 2015 (projeção, não efetiva): regional 7.274, ZFM 27.812,
imunes 22.323, inovação 3.403. A simulação usa a base efetiva.

## Fatores condicionantes — série anual

Valores da tabela em **R$ bilhões** correntes, salvo % PIB.

{chr(10).join(linhas_f)}

{chr(10).join(linhas_s)}

A DBGG oficial em metodologia 2008 só existe a partir de **dezembro de
2006** (SGS 13761). 2003–2005 entram como fluxo e saldo acumulado, sem
estoque oficial para subtrair. Em 2006 o contrafactual Selic já abre um
buraco de {_br_bi(float(y06["saldo_selic"]))} no estoque.

O fator que muda no ano *t* é quase todo **emissão líquida** (−fluxo).
O fator **juros nominais** só fica grande depois que o saldo acumulou —
e é aí que a Selic de dois dígitos (2003–2006, 2008, 2011, 2014–2015)
pesa. Em 2015 os juros evitados no ano são {_br_bi(float(y15["juros_evitados"]))};
o saldo capitalizado chega a {_br_bi(float(y15["saldo_selic"]))}.

## O que a simulação não é

Não é um modelo de equilíbrio geral. Sem indiretas e sem os três
benefícios, o PIB, a arrecadação residual e a própria Selic seriam
outros. A RFB também ignora essa reação no DGT.

Não substitui o crédito direcionado por nada: o exercício só devolve o
caixa à Conta Única e pergunta o que acontece com a **dívida bruta** se
esse caixa resgata DPF.

Não trata o crédito Tesouro→BNDES como se fosse igual às indiretas.
Aquele crédito ( {_br_bi(tot_tes)} em 2008–2014 ) é o fluxo que de fato
saiu da União e entrou no passivo do Tesouro. As indiretas de 2003–2015
({_br_bi(tot_ind)}) misturam Tesouro, FAT e recursos próprios.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_DBGG_CONTA_UNICA.md` | Esta análise |
| `output/grafico_dbgg_fluxos_conta_unica_2003_2015.png` | Empilhamento dos três fluxos |
| `output/grafico_dbgg_contrafactual_2006_2015.png` | DBGG oficial × contrafactuais |
| `scripts/sim_dbgg_conta_unica_dados.py` | Séries oficiais e reconstrução da renúncia |

```bash
python3 scripts/simular_dbgg_conta_unica.py
```
"""
    destino.write_text(md, encoding="utf-8")
    return destino


def analisar(pasta: Path | None = None) -> dict:
    pasta = pasta or Path("output")
    fluxos = df_fluxos()
    sim = df_simulacao(fluxos)
    png_f = grafico_fluxos(fluxos, pasta / "grafico_dbgg_fluxos_conta_unica_2003_2015.png")
    png_d = grafico_dbgg(sim, pasta / "grafico_dbgg_contrafactual_2006_2015.png")
    md = escrever_markdown(pasta / "TCU_CG_2010_DBGG_CONTA_UNICA.md", fluxos, sim)
    return {"fluxos": fluxos, "simulacao": sim, "png_fluxos": png_f, "png_dbgg": png_d, "md": md}


def main() -> int:
    out = analisar()
    print(f"[OK] {out['md']}")
    print(f"[OK] {out['png_fluxos']}")
    print(f"[OK] {out['png_dbgg']}")
    f = out["fluxos"]
    s = out["simulacao"]
    print(f"fluxo 2003-2015 R$ bi {f['fluxo_total'].sum()/1000:.1f}")
    y = s.loc[s["ano"] == 2015].iloc[0]
    print(f"DBGG 2015 oficial {y['dbgg_oficial']/1000:.1f}  CF Selic {y['dbgg_cf_selic']/1000:.1f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
