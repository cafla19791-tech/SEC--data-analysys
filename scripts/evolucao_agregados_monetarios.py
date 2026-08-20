"""Evolução dos agregados monetários M1–M4 no fim de cada ano (1995–2026).

Fonte: Banco Central do Brasil, SGS, saldo em final de período
(milhares de R$). Nas tabelas os valores saem em **R$ bilhões**.

Metodologia atual (“Novo”, nota de ago/2018, histórico desde dez/2001):

  27791  M1
  27810  M2
  27813  M3
  27815  M4

A série nova não cobre 1995–2000. Nesse trecho usamos as séries
descontinuadas em jul/2018 (mesma unidade, saldo de fim de período):

  1827   M1
  1837   M2
  1840   M3
  1843   M4

A troca ocorre em **2001**. Os conceitos de M2–M4 mudaram na revisão
de 2001/2018; os níveis não são estritamente comparáveis através da
quebra.

Definição atual: M1 = papel-moeda em poder do público + depósitos à
vista; M2 = M1 + poupança + títulos privados de instituições
depositárias; M3 = M2 + quotas de fundos depositários + compromissadas;
M4 = M3 + títulos públicos federais em poder do público.

Uso:
  python3 scripts/evolucao_agregados_monetarios.py
  python3 scripts/evolucao_agregados_monetarios.py --sem-download
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

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

SERIES_NOVA = {
    27791: "m1",
    27810: "m2",
    27813: "m3",
    27815: "m4",
}
SERIES_ANTIGA = {
    1827: "m1",
    1837: "m2",
    1840: "m3",
    1843: "m4",
}
# Séries descontinuadas em jul/2018 — não pedir após essa data.
FIM_SERIE_ANTIGA = "31/07/2018"

ANO_INICIO = 1995
ANO_FIM = 2026
ANO_QUEBRA = 2001
NOMES = ("m1", "m2", "m3", "m4")

DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"


def _fmt_bi(valor: float | None, casas: int = 1) -> str:
    """SGS em milhares de R$ → R$ bilhões."""
    if valor is None or pd.isna(valor):
        return "—"
    n = valor / 1_000_000.0
    return f"{n:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(valor: float | None, casas: int = 1) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    return _fmt_numero(float(valor), casas)


def _carregar_bloco(
    series_map: dict[int, str],
    cache_dir: Path | None,
    baixar: bool,
    sufixo: str,
    fim: str | None = None,
) -> dict[str, pd.DataFrame]:
    out: dict[str, pd.DataFrame] = {}
    if fim is None:
        fim = datetime.now().strftime("%d/%m/%Y")
    for cod, nome in series_map.items():
        cache = None if cache_dir is None else cache_dir / f"sgs_{cod}_{nome}_{sufixo}.csv"
        if cache is not None and cache.exists():
            out[nome] = pd.read_csv(cache, parse_dates=["mes"])
            continue
        if not baixar:
            raise FileNotFoundError(f"Cache ausente para {nome} ({sufixo}): {cache}")
        print(f"Baixando SGS {cod} ({nome}, {sufixo}) 01/01/{ANO_INICIO}..{fim}...", flush=True)
        df = baixar_sgs(cod, f"01/01/{ANO_INICIO}", fim)
        if cache is not None:
            cache.parent.mkdir(parents=True, exist_ok=True)
            df.to_csv(cache, index=False)
        out[nome] = df
    return out


def carregar_series(cache_dir: Path | None = None, baixar: bool = True) -> dict[str, dict[str, pd.DataFrame]]:
    return {
        "antiga": _carregar_bloco(
            SERIES_ANTIGA, cache_dir, baixar, "antiga", fim=FIM_SERIE_ANTIGA
        ),
        "nova": _carregar_bloco(SERIES_NOVA, cache_dir, baixar, "nova"),
    }


def _estoque_ano(df: pd.DataFrame, ano: int) -> float | None:
    bloco = df[pd.to_datetime(df["mes"]).dt.year == ano]
    if bloco.empty:
        return None
    v = bloco.sort_values("mes").iloc[-1]["valor"]
    if pd.isna(v):
        return None
    return float(v)


def agregar_anual(blocos: dict[str, dict[str, pd.DataFrame]]) -> pd.DataFrame:
    """Dezembro de cada ano; no ano corrente, último mês disponível.

    Até 2000 usa a metodologia antiga. De 2001 em diante, a nova.
    Se a nova não tiver o ano (ex.: M1 em 2001), cai na antiga.
    """
    antiga = blocos["antiga"]
    nova = blocos["nova"]
    linhas = []
    for ano in range(ANO_INICIO, ANO_FIM + 1):
        fonte = "antiga" if ano < ANO_QUEBRA else "nova"
        src = antiga if fonte == "antiga" else nova
        row: dict = {"ano": ano, "fonte": fonte}
        for nome in NOMES:
            val = _estoque_ano(src[nome], ano)
            if val is None and fonte == "nova":
                val = _estoque_ano(antiga[nome], ano)
                if val is not None:
                    row[f"{nome}_fallback"] = "antiga"
            row[nome] = val
        if all(row[n] is None for n in NOMES):
            continue
        linhas.append(row)
    out = pd.DataFrame(linhas)
    if out.empty:
        return out
    for nome in ("m1", "m2", "m3"):
        out[f"share_{nome}"] = 100.0 * out[nome] / out["m4"]
    for nome in NOMES:
        out[f"var_{nome}"] = out[nome].pct_change() * 100.0
    return out


def mes_referencia(blocos: dict[str, dict[str, pd.DataFrame]], ano: int) -> str | None:
    src = blocos["antiga"] if ano < ANO_QUEBRA else blocos["nova"]
    df = src.get("m4")
    if df is None:
        return None
    bloco = df[pd.to_datetime(df["mes"]).dt.year == ano]
    if bloco.empty:
        return None
    return pd.to_datetime(bloco["mes"].max()).strftime("%b/%Y")


def fases_historicas(anual: pd.DataFrame) -> list[dict]:
    recortes = [
        (1995, 1998, "Pós-Real e âncora cambial"),
        (1999, 2002, "Metas de inflação e quebra metodológica"),
        (2003, 2008, "Expansão do crédito e da liquidez"),
        (2009, 2010, "Crise internacional"),
        (2011, 2016, "Desaceleração e recessão"),
        (2017, 2019, "Recomposição"),
        (2020, 2021, "Pandemia — salto do M1"),
        (2022, 2026, "Aperto e recomposição dos amplos"),
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
                "m1_ini": float(bloco["m1"].dropna().iloc[0]),
                "m1_fim": float(bloco["m1"].dropna().iloc[-1]),
                "m2_ini": float(bloco["m2"].dropna().iloc[0]),
                "m2_fim": float(bloco["m2"].dropna().iloc[-1]),
                "m3_ini": float(bloco["m3"].dropna().iloc[0]),
                "m3_fim": float(bloco["m3"].dropna().iloc[-1]),
                "m4_ini": float(bloco["m4"].dropna().iloc[0]),
                "m4_fim": float(bloco["m4"].dropna().iloc[-1]),
            }
        )
    return linhas


def cabecalhos_anual() -> list[str]:
    return ["Ano", "Fonte", "M1", "M2", "M3", "M4", "M1/M4 %", "M2/M4 %", "M3/M4 %"]


def cabecalhos_fases() -> list[str]:
    return [
        "Período",
        "Contexto",
        "M1 início",
        "M1 fim",
        "M2 início",
        "M2 fim",
        "M3 início",
        "M3 fim",
        "M4 início",
        "M4 fim",
    ]


def _rotulo_ano(ano: int, ultimo_ano: int, mes_ultimo: str | None) -> str:
    if ano != ultimo_ano or not mes_ultimo:
        return str(int(ano))
    prefixo = mes_ultimo.split("/")[0].strip().lower()[:3]
    if prefixo in {"dec", "dez"}:
        return str(int(ano))
    return f"{ano}*"


def linhas_tabela_anual(anual: pd.DataFrame, mes_ultimo: str | None = None) -> list[list[str]]:
    ultimo = int(anual["ano"].max())
    linhas = []
    for rec in anual.to_dict("records"):
        m1_txt = _fmt_bi(rec.get("m1"))
        if rec.get("m1_fallback") == "antiga":
            m1_txt += "†"
        linhas.append(
            [
                _rotulo_ano(int(rec["ano"]), ultimo, mes_ultimo),
                "antiga" if rec.get("fonte") == "antiga" else "nova",
                m1_txt,
                _fmt_bi(rec.get("m2")),
                _fmt_bi(rec.get("m3")),
                _fmt_bi(rec.get("m4")),
                _fmt_pct(rec.get("share_m1")),
                _fmt_pct(rec.get("share_m2")),
                _fmt_pct(rec.get("share_m3")),
            ]
        )
    return linhas


def linhas_tabela_fases(fases: list[dict]) -> list[list[str]]:
    return [
        [
            f["periodo"],
            f["rotulo"],
            _fmt_bi(f["m1_ini"]),
            _fmt_bi(f["m1_fim"]),
            _fmt_bi(f["m2_ini"]),
            _fmt_bi(f["m2_fim"]),
            _fmt_bi(f["m3_ini"]),
            _fmt_bi(f["m3_fim"]),
            _fmt_bi(f["m4_ini"]),
            _fmt_bi(f["m4_fim"]),
        ]
        for f in fases
    ]


def gerar_graficos(anual: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    anos = anual["ano"]
    m1 = anual["m1"] / 1_000_000.0
    m2 = anual["m2"] / 1_000_000.0
    m3 = anual["m3"] / 1_000_000.0
    m4 = anual["m4"] / 1_000_000.0
    caminhos: list[Path] = []

    fig, ax = plt.subplots(figsize=(12, 5.4))
    ax.plot(anos, m1, color="#0b4f8a", linewidth=2.2, marker="o", markersize=3.2, label="M1")
    ax.plot(anos, m2, color="#b54708", linewidth=2.2, marker="o", markersize=3.2, label="M2")
    ax.plot(anos, m3, color="#1b7f4a", linewidth=2.2, marker="o", markersize=3.2, label="M3")
    ax.plot(anos, m4, color="#111111", linewidth=2.0, linestyle="--", label="M4")
    ax.axvline(2000.5, color="#888", linewidth=0.9, linestyle=":")
    ax.set_title("Agregados monetários — saldo de fim de período")
    ax.set_ylabel("R$ bilhões")
    ax.set_xlabel("Ano (dezembro; 2026* = último mês). Linha pontilhada = quebra 2001")
    ax.set_xticks(list(anos[::2]))
    ax.legend(frameon=False, ncol=4)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p1 = output_dir / "grafico_agregados_monetarios_1995_2026.png"
    fig.savefig(p1, dpi=140)
    plt.close(fig)
    caminhos.append(p1)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    ax.stackplot(
        anos,
        m1,
        m2 - m1,
        m3 - m2,
        m4 - m3,
        colors=["#0b4f8a", "#d97706", "#1b7f4a", "#6b21a8"],
        labels=["M1", "M2 − M1", "M3 − M2", "M4 − M3"],
        alpha=0.9,
    )
    ax.axvline(2000.5, color="#888", linewidth=0.9, linestyle=":")
    ax.set_title("Composição dos agregados (camadas incrementais)")
    ax.set_ylabel("R$ bilhões")
    ax.set_xticks(list(anos[::2]))
    ax.legend(loc="upper left", frameon=False, ncol=2)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    fig.tight_layout()
    p2 = output_dir / "grafico_composicao_agregados_1995_2026.png"
    fig.savefig(p2, dpi=140)
    plt.close(fig)
    caminhos.append(p2)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    ax.plot(anos, anual["share_m1"], color="#0b4f8a", linewidth=2.2, label="M1 / M4")
    ax.plot(anos, anual["share_m2"], color="#b54708", linewidth=2.2, label="M2 / M4")
    ax.plot(anos, anual["share_m3"], color="#1b7f4a", linewidth=2.2, label="M3 / M4")
    ax.axvline(2000.5, color="#888", linewidth=0.9, linestyle=":")
    ax.set_title("Participação no M4")
    ax.set_ylabel("% do M4")
    ax.set_xticks(list(anos[::2]))
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p3 = output_dir / "grafico_share_agregados_m4_1995_2026.png"
    fig.savefig(p3, dpi=140)
    plt.close(fig)
    caminhos.append(p3)
    return caminhos


def _destaques(anual: pd.DataFrame) -> dict:
    primeiro = anual.iloc[0]
    ultimo = anual.iloc[-1]
    novo = anual[anual["fonte"] == "nova"].iloc[0]
    return {
        "ano_ini": int(primeiro.ano),
        "ano_fim": int(ultimo.ano),
        "m1_ini": float(primeiro.m1),
        "m2_ini": float(primeiro.m2),
        "m3_ini": float(primeiro.m3),
        "m4_ini": float(primeiro.m4),
        "m1_fim": float(ultimo.m1),
        "m2_fim": float(ultimo.m2),
        "m3_fim": float(ultimo.m3),
        "m4_fim": float(ultimo.m4),
        "share_m1_ini": float(primeiro.share_m1),
        "share_m1_fim": float(ultimo.share_m1),
        "share_m2_fim": float(ultimo.share_m2),
        "share_m3_fim": float(ultimo.share_m3),
        "mult_m1": float(ultimo.m1 / primeiro.m1),
        "mult_m2": float(ultimo.m2 / primeiro.m2),
        "mult_m3": float(ultimo.m3 / primeiro.m3),
        "mult_m4": float(ultimo.m4 / primeiro.m4),
        "ano_nova": int(novo.ano),
        "m1_nova": float(novo.m1),
        "m4_nova": float(novo.m4),
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
    texto = f"""# Agregados monetários M1–M4 (1995–2026)

**Fonte:** Banco Central do Brasil, SGS, saldo em final de período.
Unidade original: milhares de R$; tabelas em **R$ bilhões**.
**Consulta:** {gerado}.{nota_2026}

A metodologia **atual** (SGS 27791 / 27810 / 27813 / 27815) começa em
dezembro de 2001. De 1995 a 2000 usamos as séries descontinuadas em
jul/2018 (SGS 1827 / 1837 / 1840 / 1843). A coluna Fonte marca a
quebra. M2–M4 mudam de composição na revisão de 2001/2018: os níveis
**não são estritamente comparáveis** através de 2000–2001.
A API da SGS 27791 (M1 novo) começa em jan/2002; o M1 de dez/2001
usa a série antiga (1827) e aparece marcado com †.

Definição atual: **M1** = papel-moeda em poder do público + depósitos à
vista; **M2** = M1 + poupança + títulos privados de depositárias;
**M3** = M2 + quotas de fundos depositários + compromissadas; **M4** =
M3 + títulos públicos federais em poder do público.

Tabelas com **grade contínua**.

## Síntese

De dez/{d['ano_ini']} a {d['ano_fim']}:

- M1: R$ {_fmt_bi(d['m1_ini'])} bi → R$ {_fmt_bi(d['m1_fim'])} bi
  ({_fmt_numero(d['mult_m1'])} vezes). Fatia do M4:
  {_fmt_pct(d['share_m1_ini'])}% → {_fmt_pct(d['share_m1_fim'])}%.
- M2: R$ {_fmt_bi(d['m2_ini'])} bi → R$ {_fmt_bi(d['m2_fim'])} bi
  ({_fmt_numero(d['mult_m2'])} vezes).
- M3: R$ {_fmt_bi(d['m3_ini'])} bi → R$ {_fmt_bi(d['m3_fim'])} bi
  ({_fmt_numero(d['mult_m3'])} vezes).
- M4: R$ {_fmt_bi(d['m4_ini'])} bi → R$ {_fmt_bi(d['m4_fim'])} bi
  ({_fmt_numero(d['mult_m4'])} vezes).

Início da metodologia nova ({d['ano_nova']}): M1 R$ {_fmt_bi(d['m1_nova'])} bi;
M4 R$ {_fmt_bi(d['m4_nova'])} bi.

## Fases

{html_fases}

### 1995–1998 — âncora cambial

Com o Real já em vigor, os agregados crescem em nível mas ainda sob
câmbio administrado e juros altos. O M1 é pequeno diante do M4: a
liquidez está nos títulos e na poupança.

### 1999–2002 — metas e quebra de série

A flutuação do câmbio (1999) e as metas de inflação mudam o regime. Em
2001 o Bacen reformula M2–M4 (emissores em vez de grau de liquidez). A
série nova, refeita em 2018 até dez/2001, passa a ser a oficial.

### 2003–2008 — expansão

Ciclo de crédito, bancarização e crescimento do M3/M4 com fundos e
compromissadas.

### 2009–2010 — crise

A liquidez ampla se sustenta; o M1 reage à política anticíclica.

### 2011–2016 — desaceleração

O M4 segue subindo em nível nominal; o M1 perde fatia no agregado
amplo.

### 2017–2019 — recomposição

Juros em queda e recuperação lenta. Fundos e títulos puxam M3 e M4.

### 2020–2021 — pandemia

Auxílios, depósitos à vista e papel-moeda elevam o **M1**. A fatia do
restrito no M4 aumenta de forma excepcional.

### 2022–2026 — aperto e amplos

Com a Selic alta o M1 recua em participação; M2–M4 continuam a crescer
com fundos, compromissadas e títulos públicos.

## Série anual (R$ bilhões; estoque de dezembro)

{html_anual}

## Arquivos

- `agregados_monetarios_anual_1995_2026.csv`
- `agregados_monetarios_fases_1995_2026.csv`
- `agregados_monetarios_tabelas_1995_2026.xlsx`
- `tabela_agregados_anual_1995_2026.png` / `tabela_agregados_fases_1995_2026.png`
- `grafico_agregados_monetarios_1995_2026.png`
- `grafico_composicao_agregados_1995_2026.png`
- `grafico_share_agregados_m4_1995_2026.png`
"""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "evolucao_agregados_monetarios_1995_2026.md"
    path.write_text(texto, encoding="utf-8")
    return path


def exportar_tabelas(anual: pd.DataFrame, output_dir: Path, mes_ultimo: str | None) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_anual = output_dir / "agregados_monetarios_anual_1995_2026.csv"
    anual_out = anual.copy()
    for col in NOMES:
        anual_out[col] = anual_out[col] / 1_000_000.0
    anual_out.to_csv(csv_anual, index=False, float_format="%.3f")
    fases = pd.DataFrame(fases_historicas(anual))
    csv_fases = output_dir / "agregados_monetarios_fases_1995_2026.csv"
    for col in fases.columns:
        if col.startswith("m"):
            fases[col] = fases[col] / 1_000_000.0
    fases.to_csv(csv_fases, index=False, float_format="%.3f")
    xlsx = output_dir / "agregados_monetarios_tabelas_1995_2026.xlsx"
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
        output_dir / "tabela_agregados_anual_1995_2026.png",
        "Agregados monetários M1–M4 (R$ bilhões, fim de período)",
        larguras=[0.08, 0.09, 0.12, 0.12, 0.12, 0.13, 0.11, 0.11, 0.12],
    )
    p2 = desenhar_tabela_png(
        cabecalhos_fases(),
        linhas_tabela_fases(fases_historicas(anual)),
        output_dir / "tabela_agregados_fases_1995_2026.png",
        "Fases — agregados monetários (R$ bilhões)",
        larguras=[0.09, 0.22, 0.09, 0.09, 0.09, 0.09, 0.09, 0.09, 0.08, 0.07],
    )
    return [p1, p2]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--sem-download", action="store_true")
    parser.add_argument("--sem-graficos", action="store_true")
    args = parser.parse_args(argv)

    blocos = carregar_series(cache_dir=args.cache_dir, baixar=not args.sem_download)
    anual = agregar_anual(blocos)
    ultimo = int(anual["ano"].max())
    mes_ultimo = mes_referencia(blocos, ultimo)
    caminhos = exportar_tabelas(anual, args.output_dir, mes_ultimo)
    caminhos.append(gerar_relatorio(anual, args.output_dir, mes_ultimo))
    if not args.sem_graficos:
        caminhos.extend(gerar_graficos(anual, args.output_dir))
        caminhos.extend(gerar_tabelas_png(anual, args.output_dir, mes_ultimo))
    print(f"Anos: {int(anual['ano'].min())}–{int(anual['ano'].max())} ({len(anual)} linhas)")
    if mes_ultimo:
        print(f"Último estoque: {mes_ultimo}")
    print(anual[["ano", "fonte", "m1", "m4"]].to_string(index=False))
    for p in caminhos:
        print(f"  {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
