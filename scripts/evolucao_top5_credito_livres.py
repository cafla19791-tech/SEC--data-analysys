"""Carteira das 5 maiores IFs e proxy de crédito com recursos livres (2002–2026).

Fonte: Banco Central, IF.data (OData), conglomerados financeiros (tipo 2)
até dez/2024 e prudenciais (tipo 1) a partir de dez/2025.

O IF.data **não publica** a rubrica “recursos livres” por instituição.
O que existe:

  - Carteira de crédito classificada / carteira de crédito (Relatório Resumo)
    desde 2002, por conglomerado.
  - Modalidades PF (rel. 11) e PJ (rel. 13) a partir de dez/2014, usadas
    aqui como *proxy* de recursos livres: exclui habitação, rural e
    infraestrutura (tipicamente direcionados).

Os cinco nomes canônicos (fusões somadas): Banco do Brasil, Itaú Unibanco
(inclui Unibanco), Bradesco, Caixa e Santander (inclui ABN Amro/Banespa).
BNDES fica de fora do ranking comercial.

Uso:
  python3 scripts/evolucao_top5_credito_livres.py
  python3 scripts/evolucao_top5_credito_livres.py --sem-download
"""

from __future__ import annotations

import argparse
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import matplotlib.pyplot as plt
import pandas as pd
import requests
from openpyxl import Workbook

from scripts.evolucao_balanca_reservas import (
    _escrever_aba_excel,
    desenhar_tabela_png,
    tabela_html,
)

IFDATA = "https://olinda.bcb.gov.br/olinda/servico/IFDATA/versao/v1/odata"
DATA_DIR = ROOT / "data" / "ifdata"
OUTPUT_DIR = ROOT / "output"

PF_LIVRES = {
    "Cartão de Crédito",
    "Empréstimo com Consignação em Folha",
    "Empréstimo sem Consignação em Folha",
    "Veículos",
    "Outros Créditos",
}
PJ_LIVRES = {
    "Capital de Giro",
    "Capital de Giro Rotativo",
    "Cheque Especial e Conta Garantida",
    "Comércio Exterior",
    "Operações com Recebíveis",
    "Investimento",
    "Outros Créditos",
}

CANON_ORDEM = [
    "Banco do Brasil",
    "Itaú Unibanco",
    "Bradesco",
    "Caixa",
    "Santander",
]


def _norm(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return txt.upper()


def _e_ruido(n: str) -> bool:
    marcas = (
        "COOPERATIVA",
        "CECM",
        "CREDICOOP",
        "ITAUNA",
        "NOSSA CAIXA",
        "CAIXA GERAL",
        "CAIXA ESTADUAL",
        "CAIXA FORTE",
        "APCEF",
        "FOMENTO",
        "BNDES",
        "DESENVOLVIMENTO ECONOMICO",
    )
    return any(m in n for m in marcas)


def classificacao(nome: str) -> tuple[str, str] | None:
    """(banco canônico, entidade pré-fusão) ou None se não for uma das 5 IFs."""
    n = _norm(nome).strip()
    if n.endswith(" - PRUDENCIAL"):
        n = n[: -len(" - PRUDENCIAL")].strip()
    if _e_ruido(n):
        return None
    if n == "BB" or n.startswith("BB ") or n.startswith("BANCO DO BRASIL"):
        return ("Banco do Brasil", "BB")
    if n == "UNIBANCO" or n.startswith("UNIBANCO -") or n.startswith("UNIBANCO "):
        return ("Itaú Unibanco", "UNIBANCO")
    if n in {"ITAU", "ITAU UNIBANCO"} or n.startswith("ITAU UNIBANCO") or n.startswith("BANCO ITAU"):
        return ("Itaú Unibanco", "ITAU")
    if n == "BRADESCO" or n.startswith("BANCO BRADESCO") or n.startswith("BRADESCO"):
        return ("Bradesco", "BRADESCO")
    if n in {"CAIXA", "CAIXA ECONOMICA FEDERAL"} or n.startswith("CAIXA ECONOMICA FEDERAL"):
        return ("Caixa", "CAIXA")
    if n == "ABN AMRO" or n.startswith("ABN AMRO") or n.startswith("BANCO ABN"):
        return ("Santander", "ABN")
    if (
        n in {"SANTANDER", "SANTANDER BANESPA"}
        or n.startswith("BANCO SANTANDER")
        or n.startswith("SANTANDER")
    ):
        return ("Santander", "SANTANDER")
    if n.startswith("BANESPA") or "BANCO DO ESTADO DE SAO PAULO" in n:
        return ("Santander", "BANESPA")
    return None


def nome_canonico(nome: str) -> str | None:
    c = classificacao(nome)
    return None if c is None else c[0]


def _fmt_bi(valor: float | None, casas: int = 1) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    return f"{valor:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _get_json(url: str, tentativas: int = 4) -> list[dict]:
    ultimo = None
    for i in range(tentativas):
        try:
            resp = requests.get(url, timeout=180)
            resp.raise_for_status()
            return resp.json().get("value", [])
        except Exception as exc:  # noqa: BLE001
            ultimo = exc
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"Falha IF.data: {ultimo}") from ultimo


def datas_base() -> list[tuple[int, int, int, str]]:
    """(anomes, tipo, ano, coluna_credito).

    Até dez/2024 o crédito classificado está no conglomerado financeiro
    (tipo 2). A partir de 2025 o IF.data publica a carteira no
    conglomerado prudencial (tipo 1). Jun/2026 ainda não tinha dados
    na consulta; 2026 usa mar/2026.
    """
    out = []
    for ano in range(2002, 2025):
        out.append((ano * 100 + 12, 2, ano, "Carteira de Crédito Classificada"))
    out.append((202512, 1, 2025, "Carteira de Crédito"))
    out.append((202603, 1, 2026, "Carteira de Crédito"))
    return out


def baixar_cadastro(anomes: int, cache_dir: Path) -> pd.DataFrame:
    path = cache_dir / f"cadastro_{anomes}.csv"
    if path.exists():
        return pd.read_csv(path, dtype={"CodInst": str})
    url = f"{IFDATA}/IfDataCadastro(AnoMes=@AnoMes)?@AnoMes={anomes}&$format=json&$top=10000"
    df = pd.DataFrame(_get_json(url))
    if df.empty:
        return df
    df["CodInst"] = df["CodInst"].astype(str)
    keep = [c for c in ["CodInst", "NomeInstituicao"] if c in df.columns]
    df = df[keep]
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
    return df


def baixar_valores(
    anomes: int,
    tipo: int,
    relatorio: str,
    filtro: str,
    cache_name: str,
    cache_dir: Path,
) -> pd.DataFrame:
    path = cache_dir / f"{cache_name}_{anomes}_t{tipo}.csv"
    if path.exists():
        return pd.read_csv(path, dtype={"CodInst": str})
    rows: list[dict] = []
    skip = 0
    while True:
        url = (
            f"{IFDATA}/IfDataValores(AnoMes=@AnoMes,TipoInstituicao=@TipoInstituicao,Relatorio=@Relatorio)"
            f"?@AnoMes={anomes}&@TipoInstituicao={tipo}&@Relatorio='{relatorio}'"
            f"&$format=json&$top=5000&$skip={skip}&$filter={quote(filtro)}"
        )
        chunk = _get_json(url)
        rows.extend(chunk)
        if len(chunk) < 5000:
            break
        skip += 5000
        if skip > 40000:
            break
    df = pd.DataFrame(rows)
    if not df.empty:
        df["CodInst"] = df["CodInst"].astype(str)
        df["Saldo"] = pd.to_numeric(df["Saldo"], errors="coerce")
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
    return df


def _saldo_por_entidade(m: pd.DataFrame) -> pd.DataFrame:
    """Máximo por entidade pré-fusão; soma entidades materiais do mesmo banco.

    Evita o residual de 'Caixa Geral' no lugar da CEF e as cooperativas
    'Itaúna'. Mantém a soma Itaú+Unibanco e ABN+Santander antes das fusões.
    Entidade com menos de 5% da maior do banco é satélite e entra fora.
    """
    if m.empty:
        return pd.DataFrame(columns=["banco", "Saldo"])
    m = m.copy()
    parsed = m["NomeInstituicao"].map(classificacao)
    m["banco"] = parsed.map(lambda x: None if x is None else x[0])
    m["entidade"] = parsed.map(lambda x: None if x is None else x[1])
    m = m.dropna(subset=["banco", "entidade", "Saldo"])
    if m.empty:
        return pd.DataFrame(columns=["banco", "Saldo"])
    chaves = ["banco", "entidade"]
    if "Grupo" in m.columns and m["Grupo"].notna().any():
        chaves = ["banco", "entidade", "Grupo"]
    por = m.groupby(chaves, as_index=False)["Saldo"].max()
    teto = por.groupby([c for c in chaves if c != "entidade"])["Saldo"].transform("max")
    por = por[por["Saldo"] >= 0.05 * teto]
    g = por.groupby([c for c in chaves if c != "entidade"], as_index=False)["Saldo"].sum()
    return g


def agregar_big5(credito: pd.DataFrame, cadastro: pd.DataFrame) -> pd.DataFrame:
    if credito.empty or cadastro.empty:
        return pd.DataFrame(columns=["banco", "carteira"])
    m = credito.merge(cadastro, on="CodInst", how="left")
    g = _saldo_por_entidade(m)
    return g.rename(columns={"Saldo": "carteira"})


def proxy_livres(modal: pd.DataFrame, cadastro: pd.DataFrame, grupos: set[str]) -> pd.DataFrame:
    if modal.empty or cadastro.empty:
        return pd.DataFrame(columns=["banco", "livres"])
    m = modal.merge(cadastro, on="CodInst", how="left")
    m = m[m["Grupo"].isin(grupos)] if "Grupo" in m.columns else m
    if m.empty:
        return pd.DataFrame(columns=["banco", "livres"])
    por = _saldo_por_entidade(m)
    if por.empty:
        return pd.DataFrame(columns=["banco", "livres"])
    if "Grupo" in por.columns:
        por = por.groupby("banco", as_index=False)["Saldo"].sum()
    return por.rename(columns={"Saldo": "livres"})


def montar_painel(cache_dir: Path, baixar: bool) -> pd.DataFrame:
    if not baixar and not cache_dir.exists():
        raise FileNotFoundError(cache_dir)
    linhas = []
    for anomes, tipo, ano, coluna in datas_base():
        print(f"IF.data {anomes} tipo={tipo} ({coluna})...", flush=True)
        cad = baixar_cadastro(anomes, cache_dir)
        cred = baixar_valores(
            anomes, tipo, "1", f"NomeColuna eq '{coluna}'", "resumo_credito", cache_dir
        )
        big = agregar_big5(cred, cad)
        big = big.set_index("banco")["carteira"] / 1e9  # R$ bi
        pf_liv = pj_liv = None
        if ano >= 2014:
            pf = baixar_valores(anomes, tipo, "11", "NomeColuna eq 'Total'", "pf_total", cache_dir)
            pj = baixar_valores(anomes, tipo, "13", "NomeColuna eq 'Total'", "pj_total", cache_dir)
            pf_liv = proxy_livres(pf, cad, PF_LIVRES).set_index("banco")["livres"] / 1e9
            pj_liv = proxy_livres(pj, cad, PJ_LIVRES).set_index("banco")["livres"] / 1e9
        row: dict = {"ano": ano, "anomes": anomes}
        for banco in CANON_ORDEM:
            row[f"cart_{banco}"] = float(big.get(banco)) if banco in big.index else None
            if pf_liv is not None:
                pf_v = float(pf_liv.get(banco, 0) or 0)
                pj_v = float(pj_liv.get(banco, 0) or 0)
                livres = pf_v + pj_v
                row[f"liv_{banco}"] = livres if livres > 0 else None
        carts = [row.get(f"cart_{b}") or 0 for b in CANON_ORDEM]
        row["cart_top5"] = sum(carts) if any(carts) else None
        if ano >= 2014:
            livs = [row.get(f"liv_{b}") or 0 for b in CANON_ORDEM]
            row["liv_top5"] = sum(livs) if any(livs) else None
        linhas.append(row)
    return pd.DataFrame(linhas)


def cabecalhos_carteira() -> list[str]:
    return ["Ano", *CANON_ORDEM, "Soma das 5"]


def linhas_carteira(
    painel: pd.DataFrame, prefix: str, so_com_valor: bool = False
) -> list[list[str]]:
    col_soma = "cart_top5" if prefix == "cart" else "liv_top5"
    linhas: list[list[str]] = []
    for rec in painel.to_dict("records"):
        if so_com_valor and (rec.get(col_soma) is None or pd.isna(rec.get(col_soma))):
            continue
        rotulo = f"{int(rec['ano'])}*" if int(rec["ano"]) == 2026 else str(int(rec["ano"]))
        vals = [_fmt_bi(rec.get(f"{prefix}_{b}")) for b in CANON_ORDEM]
        linhas.append([rotulo, *vals, _fmt_bi(rec.get(col_soma))])
    return linhas


def gerar_graficos(painel: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    anos = painel["ano"]
    caminhos = []
    fig, ax = plt.subplots(figsize=(12, 5.4))
    cores = ["#0b4f8a", "#b54708", "#1b7f4a", "#6b21a8", "#b42318"]
    for banco, cor in zip(CANON_ORDEM, cores, strict=True):
        ax.plot(
            anos,
            painel[f"cart_{banco}"],
            marker="o",
            markersize=3.2,
            linewidth=2,
            color=cor,
            label=banco,
        )
    ax.set_title("Carteira de crédito das 5 maiores IFs (IF.data)")
    ax.set_ylabel("R$ bilhões")
    ax.set_xlabel("Ano (dezembro; 2026*=março/2026)")
    ax.legend(frameon=False, ncol=2)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p1 = output_dir / "grafico_top5_carteira_credito_2002_2026.png"
    fig.savefig(p1, dpi=140)
    plt.close(fig)
    caminhos.append(p1)

    fig, ax = plt.subplots(figsize=(12, 5.2))
    ax.plot(
        anos,
        painel["cart_top5"],
        color="#111",
        linewidth=2.3,
        marker="o",
        markersize=3.5,
        label="Soma das 5 — carteira total",
    )
    if "liv_top5" in painel.columns:
        ax.plot(
            anos,
            painel["liv_top5"],
            color="#0b4f8a",
            linewidth=2.2,
            marker="s",
            markersize=3.5,
            label="Soma das 5 — proxy recursos livres",
        )
    ax.set_title("Total das 5 maiores IFs")
    ax.set_ylabel("R$ bilhões")
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    p2 = output_dir / "grafico_top5_soma_carteira_livres_2002_2026.png"
    fig.savefig(p2, dpi=140)
    plt.close(fig)
    caminhos.append(p2)
    return caminhos


def gerar_relatorio(painel: pd.DataFrame, output_dir: Path) -> Path:
    gerado = datetime.now().strftime("%Y-%m-%d")
    primeiro = painel.iloc[0]
    ultimo = painel.iloc[-1]
    html_cart = tabela_html(cabecalhos_carteira(), linhas_carteira(painel, "cart"))
    html_liv = tabela_html(
        cabecalhos_carteira(), linhas_carteira(painel, "liv", so_com_valor=True)
    )
    tem_livres = "liv_top5" in painel.columns and painel["liv_top5"].notna().any()
    primeiro_liv = painel.loc[painel["liv_top5"].notna()].iloc[0] if tem_livres else None
    vezes = (
        ultimo["cart_top5"] / primeiro["cart_top5"]
        if primeiro.get("cart_top5") and ultimo.get("cart_top5")
        else None
    )
    liv_txt = "—"
    if primeiro_liv is not None and ultimo.get("liv_top5"):
        liv_txt = (
            f"R$ {_fmt_bi(primeiro_liv.get('liv_top5'))} bi em {int(primeiro_liv.ano)} "
            f"→ R$ {_fmt_bi(ultimo.get('liv_top5'))} bi em {int(ultimo.ano)}"
        )
    texto = f"""# Crédito das 5 maiores instituições financeiras (2002–2026)

**Fonte:** Banco Central do Brasil, IF.data (conglomerado financeiro até
dez/2024; prudencial em dez/2025 e mar/2026). Valores em **R$ bilhões**.
**Consulta:** {gerado}. 2026* = data-base **março/2026** (jun/2026 ainda
não publicado no IF.data).

O IF.data **não publica** “empréstimos com recursos livres” por
instituição. A série longa é a **carteira de crédito** do Relatório Resumo.
A partir de 2014 montamos um **proxy de recursos livres** somando
modalidades PF/PJ típicas de mercado e excluindo habitação, rural e
infraestrutura (em geral direcionadas). Não coincide com o SGS 20542.

Instituições (fusões somadas): Banco do Brasil, Itaú Unibanco (inclui
Unibanco), Bradesco, Caixa e Santander (inclui ABN Amro/Banespa).
BNDES excluído.

Tabelas com **grade contínua**.

## Síntese

Soma das 5 (carteira total): **R$ {_fmt_bi(primeiro.get("cart_top5"))} bi**
em {int(primeiro.ano)} → **R$ {_fmt_bi(ultimo.get("cart_top5"))} bi**
em {int(ultimo.ano)} ({_fmt_bi(vezes, 1)} vezes).

Proxy de recursos livres (desde 2014): {liv_txt}.

## Carteira de crédito (R$ bilhões)

{html_cart}

## Proxy de recursos livres (R$ bilhões, desde 2014)

{html_liv}

De 2012 a 2020 a CEF entra pela instituição 00360305 (o
conglomerado financeiro da Caixa só volta a publicar a carteira
cheia em 2021). Em 2025–2026 o IF.data passa ao conglomerado
prudencial (`Nome - PRUDENCIAL`). O proxy de 2024 (tipo 2) e o de
2025–2026 (tipo 1) **não são estritamente comparáveis**.

Proxy PF: cartão, consignado, pessoal sem consignação, veículos e outros.
Proxy PJ: capital de giro (inclui rotativo até a mudança de layout),
cheque especial/conta garantida, comércio exterior, recebíveis,
investimento e outros. Fora: habitação, rural e financiamento de
infraestrutura.

## Arquivos

- `top5_credito_livres_anual_2002_2026.csv`
- `top5_credito_tabelas_2002_2026.xlsx`
- `tabela_top5_carteira_2002_2026.png` / `tabela_top5_livres_2014_2026.png`
"""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "evolucao_top5_credito_livres_2002_2026.md"
    path.write_text(texto, encoding="utf-8")
    return path


def exportar(painel: pd.DataFrame, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv = output_dir / "top5_credito_livres_anual_2002_2026.csv"
    painel.to_csv(csv, index=False, float_format="%.3f")
    xlsx = output_dir / "top5_credito_tabelas_2002_2026.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Carteira"
    _escrever_aba_excel(ws, cabecalhos_carteira(), linhas_carteira(painel, "cart"))
    ws2 = wb.create_sheet("Proxy_livres")
    _escrever_aba_excel(
        ws2, cabecalhos_carteira(), linhas_carteira(painel, "liv", so_com_valor=True)
    )
    wb.save(xlsx)
    return [csv, xlsx]


def gerar_tabelas_png(painel: pd.DataFrame, output_dir: Path) -> list[Path]:
    p1 = desenhar_tabela_png(
        cabecalhos_carteira(),
        linhas_carteira(painel, "cart"),
        output_dir / "tabela_top5_carteira_2002_2026.png",
        "5 maiores IFs — carteira de crédito (R$ bilhões, IF.data)",
        larguras=[0.10, 0.16, 0.16, 0.14, 0.14, 0.14, 0.16],
    )
    p2 = desenhar_tabela_png(
        cabecalhos_carteira(),
        linhas_carteira(painel, "liv", so_com_valor=True),
        output_dir / "tabela_top5_livres_2014_2026.png",
        "5 maiores IFs — proxy de recursos livres (R$ bilhões)",
        larguras=[0.10, 0.16, 0.16, 0.14, 0.14, 0.14, 0.16],
    )
    return [p1, p2]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--sem-download", action="store_true")
    parser.add_argument("--sem-graficos", action="store_true")
    args = parser.parse_args(argv)
    painel = montar_painel(args.cache_dir, baixar=not args.sem_download)
    caminhos = exportar(painel, args.output_dir)
    caminhos.append(gerar_relatorio(painel, args.output_dir))
    if not args.sem_graficos:
        caminhos.extend(gerar_graficos(painel, args.output_dir))
        caminhos.extend(gerar_tabelas_png(painel, args.output_dir))
    print(painel[["ano", "cart_top5"]].to_string(index=False))
    for p in caminhos:
        print(" ", p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
