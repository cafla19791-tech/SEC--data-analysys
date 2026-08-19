"""Evolução do consignado, do CDC e do cartão de crédito (2002–2016).

Séries SGS (recursos livres, pessoas físicas, R$ milhões, saldo fim de período):

  20579  Crédito pessoal consignado total
  20583  Aquisição de bens total (CDC: veículos + outros bens)
  20581  Aquisição de veículos (principal componente do CDC)
  20590  Cartão de crédito total (rotativo + parcelado)
  20570  Crédito livre PF — total (base das participações)

O split por modalidade começa em março de 2007. De 2002 a 2006 o Bacen
não divulga essas rubricas.

Uso:
  python3 scripts/evolucao_consignado_cdc_cartao.py
  python3 scripts/evolucao_consignado_cdc_cartao.py --sem-download
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook

from scripts.evolucao_balanca_reservas import (
    _escrever_aba_excel,
    _fmt_numero,
    baixar_sgs,
    desenhar_tabela_png,
    tabela_html,
)

SERIES = {
    20579: "consignado",
    20583: "cdc",
    20581: "veiculos",
    20590: "cartao",
    20570: "pf_livres",
}

ANO_INICIO = 2002
ANO_FIM = 2016

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "output"
DATA_DIR = ROOT / "data"


def _fmt_bi(valor: float | None, casas: int = 1) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    n = valor / 1000.0
    return f"{n:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(valor: float | None, casas: int = 1) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    return _fmt_numero(float(valor), casas)


def carregar_series(cache_dir: Path | None = None, baixar: bool = True) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    for cod, nome in SERIES.items():
        cache = None if cache_dir is None else cache_dir / f"sgs_{cod}_{nome}.csv"
        if cache is not None and cache.exists():
            out[nome] = pd.read_csv(cache, parse_dates=["mes"])
            continue
        if not baixar:
            raise FileNotFoundError(f"Cache ausente para {nome}: {cache}")
        print(f"Baixando SGS {cod} ({nome}) 01/01/{ANO_INICIO}..31/12/{ANO_FIM}...")
        df = baixar_sgs(cod, f"01/01/{ANO_INICIO}", f"31/12/{ANO_FIM}")
        if cache is not None:
            cache.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(cache, index=False)
        out[nome] = df
    return out


def agregar_anual(series: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for nome, df in series.items():
        work = df.copy()
        work["ano"] = pd.to_datetime(work["mes"]).dt.year
        work = work[(work["ano"] >= ANO_INICIO) & (work["ano"] <= ANO_FIM)]
        ultimo = work.sort_values("mes").groupby("ano", as_index=False).last()
        frames.append(ultimo[["ano", "valor"]].rename(columns={"valor": nome}))
    out = frames[0]
    for extra in frames[1:]:
        out = out.merge(extra, on="ano", how="outer")
    out = out.sort_values("ano").reset_index(drop=True)
    out = (
        out.set_index("ano")
        .reindex(range(ANO_INICIO, ANO_FIM + 1))
        .rename_axis("ano")
        .reset_index()
    )
    for col, share in (
        ("consignado", "share_consignado"),
        ("cdc", "share_cdc"),
        ("cartao", "share_cartao"),
        ("veiculos", "share_veiculos"),
    ):
        if {col, "pf_livres"}.issubset(out.columns):
            out[share] = 100.0 * out[col] / out["pf_livres"]
    if {"veiculos", "cdc"}.issubset(out.columns):
        out["share_veic_cdc"] = 100.0 * out["veiculos"] / out["cdc"]
    for col in ("consignado", "cdc", "cartao"):
        out[f"var_{col}"] = out[col].pct_change() * 100.0
    out["soma_tres"] = out["consignado"] + out["cdc"] + out["cartao"]
    out["share_tres"] = 100.0 * out["soma_tres"] / out["pf_livres"]
    return out


def fases_historicas(anual: pd.DataFrame) -> list[dict]:
    recortes = [
        (2002, 2006, "Antes do split oficial por modalidade"),
        (2007, 2010, "Expansão do consignado e do CDC-veículos"),
        (2011, 2014, "Pico do CDC e avanço do cartão"),
        (2015, 2016, "Recessão: CDC recua, consignado resiste"),
    ]
    linhas = []
    for ini, fim, rotulo in recortes:
        bloco = anual[(anual["ano"] >= ini) & (anual["ano"] <= fim)]
        if bloco.empty:
            continue

        def _ext(col: str) -> tuple[float | None, float | None]:
            s = bloco[col].dropna()
            if s.empty:
                return None, None
            return float(s.iloc[0]), float(s.iloc[-1])

        c0, c1 = _ext("consignado")
        d0, d1 = _ext("cdc")
        k0, k1 = _ext("cartao")
        linhas.append(
            {
                "periodo": f"{ini}–{fim}",
                "rotulo": rotulo,
                "cons_ini": c0,
                "cons_fim": c1,
                "cdc_ini": d0,
                "cdc_fim": d1,
                "cart_ini": k0,
                "cart_fim": k1,
            }
        )
    return linhas


def cabecalhos_anual() -> list[str]:
    return [
        "Ano",
        "Consignado",
        "CDC (bens)",
        "  dos quais veículos",
        "Cartão",
        "Soma das 3",
        "% consignado*",
        "% CDC*",
        "% cartão*",
        "PF livres",
    ]


def cabecalhos_fases() -> list[str]:
    return [
        "Período",
        "Contexto",
        "Consignado início",
        "Consignado fim",
        "CDC início",
        "CDC fim",
        "Cartão início",
        "Cartão fim",
    ]


def linhas_tabela_anual(anual: pd.DataFrame) -> list[list[str]]:
    linhas = []
    for row in anual.itertuples(index=False):
        linhas.append(
            [
                str(int(row.ano)),
                _fmt_bi(row.consignado),
                _fmt_bi(row.cdc),
                _fmt_bi(row.veiculos),
                _fmt_bi(row.cartao),
                _fmt_bi(row.soma_tres),
                _fmt_pct(row.share_consignado),
                _fmt_pct(row.share_cdc),
                _fmt_pct(row.share_cartao),
                _fmt_bi(row.pf_livres),
            ]
        )
    return linhas


def linhas_tabela_fases(fases: list[dict]) -> list[list[str]]:
    return [
        [
            f["periodo"],
            f["rotulo"],
            _fmt_bi(f["cons_ini"]),
            _fmt_bi(f["cons_fim"]),
            _fmt_bi(f["cdc_ini"]),
            _fmt_bi(f["cdc_fim"]),
            _fmt_bi(f["cart_ini"]),
            _fmt_bi(f["cart_fim"]),
        ]
        for f in fases
    ]


def gerar_graficos(anual: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    plot = anual.dropna(subset=["consignado", "cdc", "cartao"])
    anos = plot["ano"].tolist()
    caminhos: list[Path] = []

    fig, ax = plt.subplots(figsize=(11.5, 5.3))
    ax.plot(anos, plot["consignado"] / 1000, color="#0b4f8a", linewidth=2.2, marker="o", markersize=4, label="Consignado")
    ax.plot(anos, plot["cdc"] / 1000, color="#b54708", linewidth=2.2, marker="o", markersize=4, label="CDC (aquisição de bens)")
    ax.plot(anos, plot["cartao"] / 1000, color="#1b7f4a", linewidth=2.2, marker="o", markersize=4, label="Cartão de crédito")
    ax.plot(anos, plot["veiculos"] / 1000, color="#b54708", linewidth=1.3, linestyle="--", label="Veículos (parte do CDC)")
    ax.set_title("Saldo PF — consignado, CDC e cartão (recursos livres)")
    ax.set_ylabel("R$ bilhões")
    ax.set_xlabel("Ano (estoque de dezembro)")
    ax.set_xticks(anos)
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p1 = output_dir / "grafico_consignado_cdc_cartao_2007_2016.png"
    fig.savefig(p1, dpi=140)
    plt.close(fig)
    caminhos.append(p1)

    fig, ax = plt.subplots(figsize=(11.5, 5.2))
    ax.stackplot(
        anos,
        plot["consignado"] / 1000,
        plot["cdc"] / 1000,
        plot["cartao"] / 1000,
        colors=["#0b4f8a", "#d97706", "#1b7f4a"],
        labels=["Consignado", "CDC", "Cartão"],
        alpha=0.9,
    )
    ax.set_title("Composição das três modalidades")
    ax.set_ylabel("R$ bilhões")
    ax.set_xticks(anos)
    ax.legend(loc="upper left", frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    fig.tight_layout()
    p2 = output_dir / "grafico_composicao_consignado_cdc_cartao_2007_2016.png"
    fig.savefig(p2, dpi=140)
    plt.close(fig)
    caminhos.append(p2)

    fig, ax = plt.subplots(figsize=(11.5, 5.2))
    ax.plot(anos, plot["share_consignado"], color="#0b4f8a", linewidth=2.2, label="% consignado no crédito livre PF")
    ax.plot(anos, plot["share_cdc"], color="#b54708", linewidth=2.2, label="% CDC no crédito livre PF")
    ax.plot(anos, plot["share_cartao"], color="#1b7f4a", linewidth=2.2, label="% cartão no crédito livre PF")
    ax.set_title("Participação no crédito livre às pessoas físicas")
    ax.set_ylabel("% do crédito livre PF")
    ax.set_xticks(anos)
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p3 = output_dir / "grafico_share_consignado_cdc_cartao_2007_2016.png"
    fig.savefig(p3, dpi=140)
    plt.close(fig)
    caminhos.append(p3)
    return caminhos


def _destaques(anual: pd.DataFrame) -> dict:
    bloco = anual.dropna(subset=["consignado", "cdc", "cartao"])
    a0, a1 = bloco.iloc[0], bloco.iloc[-1]
    max_cdc = bloco.loc[bloco["cdc"].idxmax()]
    return {
        "ano0": int(a0.ano),
        "ano1": int(a1.ano),
        "c0": float(a0.consignado),
        "c1": float(a1.consignado),
        "d0": float(a0.cdc),
        "d1": float(a1.cdc),
        "k0": float(a0.cartao),
        "k1": float(a1.cartao),
        "v0": float(a0.veiculos),
        "v1": float(a1.veiculos),
        "mc_c": float(a1.consignado / a0.consignado),
        "mc_d": float(a1.cdc / a0.cdc),
        "mc_k": float(a1.cartao / a0.cartao),
        "sc0": float(a0.share_consignado),
        "sc1": float(a1.share_consignado),
        "sd0": float(a0.share_cdc),
        "sd1": float(a1.share_cdc),
        "sk0": float(a0.share_cartao),
        "sk1": float(a1.share_cartao),
        "st0": float(a0.share_tres),
        "st1": float(a1.share_tres),
        "pf0": float(a0.pf_livres),
        "pf1": float(a1.pf_livres),
        "ano_max_cdc": int(max_cdc.ano),
        "max_cdc": float(max_cdc.cdc),
        "veic_cdc0": float(a0.share_veic_cdc),
        "veic_cdc1": float(a1.share_veic_cdc),
    }


def gerar_relatorio(anual: pd.DataFrame, output_dir: Path) -> Path:
    d = _destaques(anual)
    fases = fases_historicas(anual)
    gerado = datetime.now().strftime("%Y-%m-%d")
    html_anual = tabela_html(
        cabecalhos_anual(),
        linhas_tabela_anual(anual),
        ["center"] + ["right"] * 9,
    )
    html_fases = tabela_html(
        cabecalhos_fases(),
        linhas_tabela_fases(fases),
        ["center", "left"] + ["right"] * 6,
    )
    texto = f"""# Consignado, CDC e cartão de crédito (2002–2016)

**Fonte:** Banco Central do Brasil, SGS. Saldo em dezembro da carteira de
crédito com **recursos livres**, pessoas físicas, em R$ milhões, nas tabelas
em **R$ bilhões**. **Consulta:** {gerado}.

O Bacen só publica o corte por modalidade a partir de **março de 2007**.
De 2002 a 2006 essas rubricas não existem no SGS — o consignado privado
nasce com a Lei 10.820/2003, mas o estoque oficial mensal começa em 2007.

Definições:

- **Consignado** (SGS 20579): crédito pessoal com desconto em folha
  (INSS, servidores, CLT).
- **CDC** (SGS 20583): aquisição de bens total — o equivalente oficial do
  crédito direto ao consumidor (veículos + outros bens). Veículos = 20581.
- **Cartão** (SGS 20590): cartão de crédito total (rotativo + parcelado).
- Participações (*) sobre o crédito livre PF (SGS 20570).

Tabelas com **grade contínua**.

## Síntese (dez/{d['ano0']} → dez/{d['ano1']})

- Consignado: R$ {_fmt_bi(d['c0'])} bi → R$ {_fmt_bi(d['c1'])} bi
  ({_fmt_numero(d['mc_c'])} vezes). Fatia do crédito livre PF:
  {_fmt_pct(d['sc0'])}% → {_fmt_pct(d['sc1'])}%.
- CDC (bens): R$ {_fmt_bi(d['d0'])} bi → R$ {_fmt_bi(d['d1'])} bi
  ({_fmt_numero(d['mc_d'])} vezes). Pico em {d['ano_max_cdc']}
  (R$ {_fmt_bi(d['max_cdc'])} bi). Fatia: {_fmt_pct(d['sd0'])}% → {_fmt_pct(d['sd1'])}%.
- Veículos no CDC: {_fmt_pct(d['veic_cdc0'])}% → {_fmt_pct(d['veic_cdc1'])}%.
- Cartão: R$ {_fmt_bi(d['k0'])} bi → R$ {_fmt_bi(d['k1'])} bi
  ({_fmt_numero(d['mc_k'])} vezes). Fatia: {_fmt_pct(d['sk0'])}% → {_fmt_pct(d['sk1'])}%.
- As três modalidades juntas: {_fmt_pct(d['st0'])}% → {_fmt_pct(d['st1'])}%
  do crédito livre PF (R$ {_fmt_bi(d['pf0'])} bi → R$ {_fmt_bi(d['pf1'])} bi).

## Fases

{html_fases}

### 2002–2006 — sem série oficial

A Lei 10.820/2003 cria o consignado na folha do setor privado; o INSS
regulamenta o desconto de aposentadorias em seguida. CDC de veículos e
cartão já existiam, mas o SGS não decompõe o estoque até março de 2007.

### 2007–2010 — consignado e CDC em alta

O consignado se torna o produto de crédito pessoal dominante. O CDC —
quase todo em veículos — acompanha o ciclo de renda, emprego e IPI de
automóveis. O cartão cresce, ainda em estoque menor.

### 2011–2014 — pico do CDC, cartão acelera

O CDC atinge o máximo da janela. O consignado continua a subir com a
formalização e o teto de margem. O cartão ganha participação com o
parcelado lojista e o rotativo.

### 2015–2016 — recessão

A crise derruba o CDC (queda de vendas de veículos e de bens duráveis).
O consignado **resiste** — renda de aposentadoria e folha pública é mais
estável. O cartão segue em alta no estoque, apesar da inadimplência.

## Série anual (R$ bilhões; dezembro)

{html_anual}

\\* Participação no crédito livre às pessoas físicas (SGS 20570).

## Arquivos gerados

- `consignado_cdc_cartao_anual_2002_2016.csv`
- `consignado_cdc_cartao_fases_2002_2016.csv`
- `consignado_cdc_cartao_tabelas_2002_2016.xlsx`
- `tabela_consignado_cdc_cartao_anual.png` / `tabela_consignado_cdc_cartao_fases.png`
- `grafico_consignado_cdc_cartao_2007_2016.png`
- `grafico_composicao_consignado_cdc_cartao_2007_2016.png`
- `grafico_share_consignado_cdc_cartao_2007_2016.png`
"""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "evolucao_consignado_cdc_cartao_2002_2016.md"
    path.write_text(texto, encoding="utf-8")
    return path


def exportar_tabelas(anual: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_anual = output_dir / "consignado_cdc_cartao_anual_2002_2016.csv"
    out = anual.copy()
    for col in ["consignado", "cdc", "veiculos", "cartao", "pf_livres", "soma_tres"]:
        if col in out.columns:
            out[col] = out[col] / 1000.0
    out.to_csv(csv_anual, index=False, float_format="%.3f")
    fases = pd.DataFrame(fases_historicas(anual))
    csv_fases = output_dir / "consignado_cdc_cartao_fases_2002_2016.csv"
    for col in ["cons_ini", "cons_fim", "cdc_ini", "cdc_fim", "cart_ini", "cart_fim"]:
        if col in fases.columns:
            fases[col] = fases[col] / 1000.0
    fases.to_csv(csv_fases, index=False, float_format="%.3f")
    xlsx = output_dir / "consignado_cdc_cartao_tabelas_2002_2016.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Serie_anual"
    _escrever_aba_excel(ws1, cabecalhos_anual(), linhas_tabela_anual(anual))
    ws2 = wb.create_sheet("Fases")
    _escrever_aba_excel(ws2, cabecalhos_fases(), linhas_tabela_fases(fases_historicas(anual)))
    wb.save(xlsx)
    return [csv_anual, csv_fases, xlsx]


def gerar_tabelas_png(anual: pd.DataFrame, output_dir: Path) -> list[Path]:
    p1 = desenhar_tabela_png(
        cabecalhos_anual(),
        linhas_tabela_anual(anual),
        output_dir / "tabela_consignado_cdc_cartao_anual.png",
        "Consignado, CDC e cartão — estoque de dezembro (R$ bilhões)",
        larguras=[0.07, 0.11, 0.11, 0.14, 0.10, 0.10, 0.11, 0.09, 0.09, 0.10],
    )
    p2 = desenhar_tabela_png(
        cabecalhos_fases(),
        linhas_tabela_fases(fases_historicas(anual)),
        output_dir / "tabela_consignado_cdc_cartao_fases.png",
        "Fases — consignado, CDC e cartão (R$ bilhões)",
        larguras=[0.10, 0.28, 0.12, 0.12, 0.10, 0.10, 0.09, 0.09],
    )
    return [p1, p2]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--sem-download", action="store_true")
    parser.add_argument("--sem-graficos", action="store_true")
    args = parser.parse_args(argv)

    series = carregar_series(cache_dir=args.cache_dir, baixar=not args.sem_download)
    anual = agregar_anual(series)
    caminhos = exportar_tabelas(anual, args.output_dir)
    caminhos.append(gerar_relatorio(anual, args.output_dir))
    if not args.sem_graficos:
        caminhos.extend(gerar_graficos(anual, args.output_dir))
        caminhos.extend(gerar_tabelas_png(anual, args.output_dir))
    print(f"Anos na tabela: {int(anual['ano'].min())}–{int(anual['ano'].max())}")
    split = anual.dropna(subset=["consignado"])
    print(f"Split oficial: {int(split['ano'].min())}–{int(split['ano'].max())}")
    for p in caminhos:
        print(f"  {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
