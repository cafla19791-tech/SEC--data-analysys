"""Evolução dos recursos livres e direcionados no SFN (2002–2026).

O Banco Central não publica um estoque chamado “depósitos de recursos livres/
direcionados”. A estatística oficial é o **saldo da carteira de crédito** do
Sistema Financeiro Nacional segundo a origem dos recursos:

  20542  Recursos livres (R$ milhões) — desde mar/2007
  20593  Recursos direcionados (R$ milhões) — desde mar/2007
  20539  Crédito total (R$ milhões) — desde jan/2002
  20622  Crédito / PIB (%) — desde jan/2002
  20625  Livres / PIB (%) — desde mar/2007
  20628  Direcionados / PIB (%) — desde mar/2007

Recursos direcionados: operações regulamentadas pelo CMN ou lastreadas em
depósitos à vista, poupança, fundos e programas públicos (BNDES, rural,
habitação). Recursos livres: taxas pactuadas livremente, sem exigibilidade.

Uso:
  python3 scripts/evolucao_recursos_livres_direcionados.py
  python3 scripts/evolucao_recursos_livres_direcionados.py --sem-download
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
    20542: "livres",
    20593: "direcionados",
    20539: "total",
    20622: "credito_pib",
    20625: "livres_pib",
    20628: "direcionados_pib",
}

ANO_INICIO = 2002
ANO_FIM = 2026

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
    fim = datetime.now().strftime("%d/%m/%Y")
    for cod, nome in SERIES.items():
        cache = None if cache_dir is None else cache_dir / f"sgs_{cod}_{nome}.csv"
        if cache is not None and cache.exists():
            out[nome] = pd.read_csv(cache, parse_dates=["mes"])
            continue
        if not baixar:
            raise FileNotFoundError(f"Cache ausente para {nome}: {cache}")
        print(f"Baixando SGS {cod} ({nome}) 01/01/{ANO_INICIO}..{fim}...")
        df = baixar_sgs(cod, f"01/01/{ANO_INICIO}", fim)
        if cache is not None:
            cache.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(cache, index=False)
        out[nome] = df
    return out


def agregar_anual(series: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Estoque de dezembro; no ano corrente, último mês disponível."""
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
    if {"livres", "total"}.issubset(out.columns):
        out["share_livres"] = 100.0 * out["livres"] / out["total"]
    if {"direcionados", "total"}.issubset(out.columns):
        out["share_dir"] = 100.0 * out["direcionados"] / out["total"]
    out["var_livres"] = out["livres"].pct_change() * 100.0
    out["var_dir"] = out["direcionados"].pct_change() * 100.0
    out["var_total"] = out["total"].pct_change() * 100.0
    return out


def mes_referencia(series: dict[str, pd.DataFrame], ano: int) -> str | None:
    df = series.get("total")
    if df is None:
        return None
    bloco = df[pd.to_datetime(df["mes"]).dt.year == ano]
    if bloco.empty:
        return None
    return pd.to_datetime(bloco["mes"].max()).strftime("%b/%Y")


def fases_historicas(anual: pd.DataFrame) -> list[dict]:
    recortes = [
        (2002, 2006, "Antes do split oficial (só crédito total)"),
        (2007, 2010, "Início da série e resposta à crise"),
        (2011, 2016, "Expansão do direcionado"),
        (2017, 2019, "Rebalanceamento pós-recessão"),
        (2020, 2021, "Pandemia"),
        (2022, 2026, "Livres voltam a puxar o estoque"),
    ]
    linhas = []
    for ini, fim, rotulo in recortes:
        bloco = anual[(anual["ano"] >= ini) & (anual["ano"] <= fim)]
        if bloco.empty:
            continue
        linhas.append(
            {
                "periodo": f"{ini}–{fim}",
                "rotulo": rotulo,
                "livres_ini": float(bloco["livres"].dropna().iloc[0]) if bloco["livres"].notna().any() else None,
                "livres_fim": float(bloco["livres"].dropna().iloc[-1]) if bloco["livres"].notna().any() else None,
                "dir_ini": float(bloco["direcionados"].dropna().iloc[0]) if bloco["direcionados"].notna().any() else None,
                "dir_fim": float(bloco["direcionados"].dropna().iloc[-1]) if bloco["direcionados"].notna().any() else None,
                "total_ini": float(bloco["total"].iloc[0]),
                "total_fim": float(bloco["total"].iloc[-1]),
                "share_livres_fim": float(bloco["share_livres"].dropna().iloc[-1])
                if bloco["share_livres"].notna().any()
                else None,
                "share_dir_fim": float(bloco["share_dir"].dropna().iloc[-1])
                if bloco["share_dir"].notna().any()
                else None,
            }
        )
    return linhas


def cabecalhos_anual() -> list[str]:
    return [
        "Ano",
        "Livres",
        "Direcionados",
        "Total",
        "% livres",
        "% direcionados",
        "Livres/PIB",
        "Dir./PIB",
        "Crédito/PIB",
    ]


def cabecalhos_fases() -> list[str]:
    return [
        "Período",
        "Contexto",
        "Livres início",
        "Livres fim",
        "Direc. início",
        "Direc. fim",
        "Total início",
        "Total fim",
        "% livres fim",
        "% direc. fim",
    ]


def _rotulo_ano(ano: int, ultimo_ano: int, mes_ultimo: str | None) -> str:
    if ano == ultimo_ano and mes_ultimo and not mes_ultimo.lower().startswith("dec"):
        return f"{ano}*"
    return str(int(ano))


def linhas_tabela_anual(anual: pd.DataFrame, mes_ultimo: str | None = None) -> list[list[str]]:
    ultimo = int(anual["ano"].max())
    linhas = []
    for row in anual.itertuples(index=False):
        linhas.append(
            [
                _rotulo_ano(int(row.ano), ultimo, mes_ultimo),
                _fmt_bi(row.livres),
                _fmt_bi(row.direcionados),
                _fmt_bi(row.total),
                _fmt_pct(row.share_livres),
                _fmt_pct(row.share_dir),
                _fmt_pct(row.livres_pib),
                _fmt_pct(row.direcionados_pib),
                _fmt_pct(row.credito_pib),
            ]
        )
    return linhas


def linhas_tabela_fases(fases: list[dict]) -> list[list[str]]:
    return [
        [
            f["periodo"],
            f["rotulo"],
            _fmt_bi(f["livres_ini"]),
            _fmt_bi(f["livres_fim"]),
            _fmt_bi(f["dir_ini"]),
            _fmt_bi(f["dir_fim"]),
            _fmt_bi(f["total_ini"]),
            _fmt_bi(f["total_fim"]),
            _fmt_pct(f["share_livres_fim"]),
            _fmt_pct(f["share_dir_fim"]),
        ]
        for f in fases
    ]


def gerar_graficos(anual: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    anos = anual["ano"].tolist()
    livres = anual["livres"] / 1000.0
    direc = anual["direcionados"] / 1000.0
    total = anual["total"] / 1000.0
    caminhos: list[Path] = []

    fig, ax = plt.subplots(figsize=(12, 5.4))
    ax.plot(anos, livres, color="#0b4f8a", linewidth=2.2, marker="o", markersize=3.4, label="Recursos livres")
    ax.plot(anos, direc, color="#b54708", linewidth=2.2, marker="o", markersize=3.4, label="Recursos direcionados")
    ax.plot(anos, total, color="#333333", linewidth=1.6, linestyle="--", label="Crédito total")
    ax.set_title("Saldo da carteira de crédito do SFN — livres vs direcionados")
    ax.set_ylabel("R$ bilhões")
    ax.set_xlabel("Ano (estoque de dezembro; 2026 = último mês)")
    ax.set_xticks(anos[::2])
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p1 = output_dir / "grafico_recursos_livres_direcionados_2002_2026.png"
    fig.savefig(p1, dpi=140)
    plt.close(fig)
    caminhos.append(p1)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    split = anual.dropna(subset=["livres", "direcionados"])
    ax.stackplot(
        split["ano"],
        split["livres"] / 1000.0,
        split["direcionados"] / 1000.0,
        colors=["#0b4f8a", "#d97706"],
        labels=["Livres", "Direcionados"],
        alpha=0.9,
    )
    ax.set_title("Composição do crédito do SFN")
    ax.set_ylabel("R$ bilhões")
    ax.set_xticks(anos[::2])
    ax.legend(loc="upper left", frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    fig.tight_layout()
    p2 = output_dir / "grafico_composicao_credito_sfn_2002_2026.png"
    fig.savefig(p2, dpi=140)
    plt.close(fig)
    caminhos.append(p2)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    ax.plot(anos, anual["share_livres"], color="#0b4f8a", linewidth=2.2, label="% livres no total")
    ax.plot(anos, anual["share_dir"], color="#b54708", linewidth=2.2, label="% direcionados no total")
    ax.axhline(50, color="#666", linewidth=0.8, linestyle=":")
    ax.set_ylim(25, 75)
    ax.set_title("Participação no saldo total de crédito")
    ax.set_ylabel("% do crédito do SFN")
    ax.set_xticks(anos[::2])
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p3 = output_dir / "grafico_share_livres_direcionados_2002_2026.png"
    fig.savefig(p3, dpi=140)
    plt.close(fig)
    caminhos.append(p3)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    ax.plot(anos, anual["livres_pib"], color="#0b4f8a", linewidth=2.2, label="Livres / PIB")
    ax.plot(anos, anual["direcionados_pib"], color="#b54708", linewidth=2.2, label="Direcionados / PIB")
    ax.plot(anos, anual["credito_pib"], color="#333", linewidth=1.6, linestyle="--", label="Crédito total / PIB")
    ax.set_title("Crédito do SFN em relação ao PIB")
    ax.set_ylabel("% do PIB")
    ax.set_xticks(anos[::2])
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p4 = output_dir / "grafico_credito_pib_livres_direcionados_2002_2026.png"
    fig.savefig(p4, dpi=140)
    plt.close(fig)
    caminhos.append(p4)
    return caminhos


def _destaques(anual: pd.DataFrame) -> dict:
    com_split = anual.dropna(subset=["livres", "direcionados"])
    primeiro = com_split.iloc[0]
    ultimo = com_split.iloc[-1]
    total_ini = anual.iloc[0]
    cruzou = com_split[com_split["direcionados"] >= com_split["livres"]]
    return {
        "ano_total_ini": int(total_ini.ano),
        "total_ini": float(total_ini.total),
        "pib_ini": float(total_ini.credito_pib),
        "ano_split": int(primeiro.ano),
        "ano_fim": int(ultimo.ano),
        "livres_ini": float(primeiro.livres),
        "livres_fim": float(ultimo.livres),
        "dir_ini": float(primeiro.direcionados),
        "dir_fim": float(ultimo.direcionados),
        "total_fim": float(ultimo.total),
        "share_l_ini": float(primeiro.share_livres),
        "share_l_fim": float(ultimo.share_livres),
        "share_d_ini": float(primeiro.share_dir),
        "share_d_fim": float(ultimo.share_dir),
        "pib_fim": float(ultimo.credito_pib),
        "livres_pib_ini": float(primeiro.livres_pib),
        "livres_pib_fim": float(ultimo.livres_pib),
        "dir_pib_ini": float(primeiro.direcionados_pib),
        "dir_pib_fim": float(ultimo.direcionados_pib),
        "mult_livres": float(ultimo.livres / primeiro.livres),
        "mult_dir": float(ultimo.direcionados / primeiro.direcionados),
        "mult_total": float(ultimo.total / total_ini.total),
        "anos_dir_maior": ", ".join(str(int(a)) for a in cruzou["ano"]) if not cruzou.empty else "nenhum",
    }


def gerar_relatorio(anual: pd.DataFrame, output_dir: Path, mes_ultimo: str | None) -> Path:
    d = _destaques(anual)
    fases = fases_historicas(anual)
    gerado = datetime.now().strftime("%Y-%m-%d")
    nota_2026 = f" *{d['ano_fim']} = estoque de {mes_ultimo}." if mes_ultimo else ""
    html_anual = tabela_html(cabecalhos_anual(), linhas_tabela_anual(anual, mes_ultimo))
    html_fases = tabela_html(
        cabecalhos_fases(),
        linhas_tabela_fases(fases),
        ["center", "left"] + ["right"] * 8,
    )
    texto = f"""# Recursos livres e direcionados no Sistema Financeiro Nacional (2002–2026)

**Fonte:** Banco Central do Brasil, SGS. Saldo em final de período da carteira
de crédito do SFN, em R$ milhões, apresentado nas tabelas em **R$ bilhões**.
Participações e razões com o PIB em %. **Consulta:** {gerado}.{nota_2026}

O Bacen **não publica** uma série chamada “depósitos de recursos livres” ou
“depósitos de recursos direcionados”. A classificação oficial é a da
**carteira de crédito** segundo a origem dos recursos:

- **Recursos livres** (SGS 20542): taxas livremente pactuadas. Não inclui
  operações com taxas regulamentadas, BNDES nem lastro em compulsório ou
  recursos governamentais.
- **Recursos direcionados** (SGS 20593): operações regulamentadas pelo CMN
  ou vinculadas a orçamento. Fonte: parte dos **depósitos à vista** e da
  **caderneta de poupança**, mais fundos e programas públicos (BNDES, crédito
  rural, habitação).
- **Total** (SGS 20539) e **crédito/PIB** (SGS 20622) existem desde **2002**.
  O split livres/direcionados começa em **março de 2007**. Livres/PIB = 20625;
  direcionados/PIB = 20628.

Tabelas com **grade contínua** (borda sólida em todas as células).

## Síntese

O crédito total do SFN passou de **R$ {_fmt_bi(d['total_ini'])} bi** em
dezembro de {d['ano_total_ini']} ({_fmt_pct(d['pib_ini'])}% do PIB) para
**R$ {_fmt_bi(d['total_fim'])} bi** em {d['ano_fim']}
({_fmt_pct(d['pib_fim'])}% do PIB) — {_fmt_numero(d['mult_total'])} vezes.

Desde o início do split ({d['ano_split']}):

- Livres: R$ {_fmt_bi(d['livres_ini'])} bi → R$ {_fmt_bi(d['livres_fim'])} bi
  ({_fmt_numero(d['mult_livres'])} vezes). Participação:
  {_fmt_pct(d['share_l_ini'])}% → {_fmt_pct(d['share_l_fim'])}%.
- Direcionados: R$ {_fmt_bi(d['dir_ini'])} bi → R$ {_fmt_bi(d['dir_fim'])} bi
  ({_fmt_numero(d['mult_dir'])} vezes). Participação:
  {_fmt_pct(d['share_d_ini'])}% → {_fmt_pct(d['share_d_fim'])}%.
- Livres/PIB: {_fmt_pct(d['livres_pib_ini'])}% → {_fmt_pct(d['livres_pib_fim'])}%.
- Direcionados/PIB: {_fmt_pct(d['dir_pib_ini'])}% → {_fmt_pct(d['dir_pib_fim'])}%.
- Anos em que o estoque direcionado igualou ou superou o livre: {d['anos_dir_maior']}.

## Fases

{html_fases}

### 2002–2006 — só o total

Antes de março de 2007 o Bacen não divulgava o corte livres/direcionados.
O crédito total já crescia com a expansão bancária pós-Real e o ciclo de
crédito a pessoas físicas, mas o nível em relação ao PIB ainda era baixo
perante o que viria depois da crise de 2008.

### 2007–2010 — split oficial e anticíclico

A série oficial nasce com os **livres à frente**. Na crise de 2008–2009 o
direcionado (BNDES, habitação, rural) ganha papel anticíclico e começa a
fechar o hiato.

### 2011–2016 — auge do direcionado

BNDES, Minha Casa Minha Vida e crédito rural empurram o estoque
direcionado. Em dezembro de 2016 os dois blocos **quase empatam**
(50,1% livres e 49,9% direcionados) — o ponto mais próximo da série.
O direcionado cresceu 10,1 vezes entre 2007 e 2026, contra 6,6 vezes
nos livres.

### 2017–2019 — rebalanceamento

Com a recessão, o TLP no BNDES e o recuo do crédito oficial, os livres
voltam a crescer mais depressa e recuperam a maior fatia do estoque.

### 2020–2021 — pandemia

Programas públicos (Pronampe e afins) e moratórias sustentam o
direcionado; os livres acompanham a reabertura.

### 2022–2026 — livres na frente

O crédito livre (cartão, consignado, veículos, capital de giro a mercado)
volta a puxar o estoque. O direcionado segue grande em nível — habitação
e rural compensam a menor fatia do BNDES — mas deixa de ser o motor da
expansão.

## Série anual (R$ bilhões; estoque de dezembro)

{html_anual}

## Arquivos gerados

- `recursos_livres_direcionados_anual_2002_2026.csv`
- `recursos_livres_direcionados_fases_2002_2026.csv`
- `recursos_livres_direcionados_tabelas_2002_2026.xlsx`
- `tabela_recursos_anual_2002_2026.png` / `tabela_recursos_fases_2002_2026.png`
- `grafico_recursos_livres_direcionados_2002_2026.png`
- `grafico_composicao_credito_sfn_2002_2026.png`
- `grafico_share_livres_direcionados_2002_2026.png`
- `grafico_credito_pib_livres_direcionados_2002_2026.png`
"""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "evolucao_recursos_livres_direcionados_2002_2026.md"
    path.write_text(texto, encoding="utf-8")
    return path


def exportar_tabelas(anual: pd.DataFrame, output_dir: Path, mes_ultimo: str | None) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_anual = output_dir / "recursos_livres_direcionados_anual_2002_2026.csv"
    anual_out = anual.copy()
    for col in ["livres", "direcionados", "total"]:
        anual_out[col] = anual_out[col] / 1000.0
    anual_out.to_csv(csv_anual, index=False, float_format="%.3f")
    fases = pd.DataFrame(fases_historicas(anual))
    csv_fases = output_dir / "recursos_livres_direcionados_fases_2002_2026.csv"
    for col in ["livres_ini", "livres_fim", "dir_ini", "dir_fim", "total_ini", "total_fim"]:
        if col in fases.columns:
            fases[col] = fases[col] / 1000.0
    fases.to_csv(csv_fases, index=False, float_format="%.3f")
    xlsx = output_dir / "recursos_livres_direcionados_tabelas_2002_2026.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Serie_anual"
    _escrever_aba_excel(ws1, cabecalhos_anual(), linhas_tabela_anual(anual, mes_ultimo))
    ws2 = wb.create_sheet("Fases")
    _escrever_aba_excel(ws2, cabecalhos_fases(), linhas_tabela_fases(fases_historicas(anual)))
    wb.save(xlsx)
    return [csv_anual, csv_fases, xlsx]


def gerar_tabelas_png(anual: pd.DataFrame, output_dir: Path, mes_ultimo: str | None) -> list[Path]:
    p1 = desenhar_tabela_png(
        cabecalhos_anual(),
        linhas_tabela_anual(anual, mes_ultimo),
        output_dir / "tabela_recursos_anual_2002_2026.png",
        "Crédito do SFN — recursos livres e direcionados (R$ bilhões)",
        larguras=[0.08, 0.11, 0.13, 0.11, 0.10, 0.13, 0.11, 0.11, 0.12],
    )
    p2 = desenhar_tabela_png(
        cabecalhos_fases(),
        linhas_tabela_fases(fases_historicas(anual)),
        output_dir / "tabela_recursos_fases_2002_2026.png",
        "Fases — recursos livres e direcionados (R$ bilhões)",
        larguras=[0.09, 0.22, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.08, 0.08],
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
    ultimo = int(anual["ano"].max())
    mes_ultimo = mes_referencia(series, ultimo)
    caminhos = exportar_tabelas(anual, args.output_dir, mes_ultimo)
    caminhos.append(gerar_relatorio(anual, args.output_dir, mes_ultimo))
    if not args.sem_graficos:
        caminhos.extend(gerar_graficos(anual, args.output_dir))
        caminhos.extend(gerar_tabelas_png(anual, args.output_dir, mes_ultimo))
    print(f"Anos: {int(anual['ano'].min())}–{int(anual['ano'].max())} ({len(anual)} linhas)")
    if mes_ultimo:
        print(f"Último estoque: {mes_ultimo}")
    for p in caminhos:
        print(f"  {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
