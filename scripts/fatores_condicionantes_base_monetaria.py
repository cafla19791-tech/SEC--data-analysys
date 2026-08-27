#!/usr/bin/env python3
"""Saldos e variações anuais dos fatores condicionantes da base monetária.

Fonte: Banco Central do Brasil, SGS (saldo em final de período).
Unidade original: milhares de unidades monetárias correntes (R$ mil).
Divulgação: R$ milhões (valor SGS ÷ 1.000).

As séries dos *fatores* medem a contribuição de cada item no mês
(variação da base atribuída ao fator). Por isso o script publica:

- valor de dezembro (última observação do ano na série SGS);
- variação acumulada no ano (soma dos meses) — fecha com Δ da base;
- saldo da base monetária restrita no fim do ano (estoque, SGS 1788).

O ano corrente usa o último mês disponível.
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

ANO_INICIO_DEFAULT = 1995
ANO_FIM_DEFAULT = 2026
SGS_API = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.{cod}/dados"

SERIES: list[dict] = [
    {
        "codigo": 1810,
        "coluna": "tesouro_conta_unica",
        "nome": "Tesouro Nacional — Conta única",
        "grupo": "fator",
        "inicio": "01/01/1995",
    },
    {
        "codigo": 1809,
        "coluna": "titulos_publicos_total",
        "nome": "Títulos públicos federais — Total",
        "grupo": "fator",
        "inicio": "01/01/1995",
    },
    {
        "codigo": 29004,
        "coluna": "titulos_mercado_primario",
        "nome": "Títulos públicos — mercado primário",
        "grupo": "desdobramento",
        "inicio": "01/01/2015",
    },
    {
        "codigo": 29006,
        "coluna": "titulos_mercado_secundario",
        "nome": "Títulos públicos — mercado secundário",
        "grupo": "desdobramento",
        "inicio": "01/01/2015",
    },
    {
        "codigo": 1811,
        "coluna": "setor_externo",
        "nome": "Operações com o setor externo",
        "grupo": "fator",
        "inicio": "01/01/1995",
    },
    {
        "codigo": 12487,
        "coluna": "derivativos_ajustes",
        "nome": "Derivativos — ajustes",
        "grupo": "fator",
        "inicio": "01/01/2000",
    },
    {
        "codigo": 12484,
        "coluna": "redesconto",
        "nome": "Redesconto do Banco Central",
        "grupo": "fator",
        "inicio": "01/03/2000",
    },
    {
        "codigo": 28724,
        "coluna": "linhas_temporarias_liquidez",
        "nome": "Linhas temporárias especiais de liquidez",
        "grupo": "fator",
        "inicio": "01/05/2020",
    },
    {
        "codigo": 1815,
        "coluna": "depositos_instituicoes_financeiras",
        "nome": "Depósitos de instituições financeiras",
        "grupo": "fator",
        "inicio": "01/01/1995",
    },
    {
        "codigo": 1818,
        "coluna": "outras_operacoes",
        "nome": "Autoridade Monetária — Outras operações",
        "grupo": "fator",
        "inicio": "01/01/1995",
    },
    {
        "codigo": 1788,
        "coluna": "base_monetaria_restrita",
        "nome": "Base monetária restrita",
        "grupo": "resultado",
        "inicio": "01/01/1995",
    },
]

FATORES_IDENTIDADE = [
    "tesouro_conta_unica",
    "titulos_publicos_total",
    "setor_externo",
    "derivativos_ajustes",
    "redesconto",
    "linhas_temporarias_liquidez",
    "depositos_instituicoes_financeiras",
    "outras_operacoes",
]


def anos_periodo(ano_inicio: int, ano_fim: int) -> list[int]:
    if ano_fim < ano_inicio:
        raise ValueError("ano_fim deve ser >= ano_inicio")
    return list(range(ano_inicio, ano_fim + 1))


def _blocos(inicio: pd.Timestamp, fim: pd.Timestamp, anos: int = 8):
    cursor = inicio
    while cursor <= fim:
        bloco_fim = min(cursor + pd.DateOffset(years=anos) - pd.DateOffset(days=1), fim)
        yield cursor, bloco_fim
        cursor = bloco_fim + pd.DateOffset(days=1)


def baixar_sgs(
    codigo: int,
    inicio: str = "01/01/1995",
    fim: str | None = None,
    session: requests.Session | None = None,
) -> pd.DataFrame:
    """Baixa série SGS em blocos e ignora janelas 404 (série ainda não existia)."""
    if fim is None:
        fim = datetime.now().strftime("%d/%m/%Y")
    inicio_ts = pd.Timestamp(datetime.strptime(inicio, "%d/%m/%Y"))
    fim_ts = pd.Timestamp(datetime.strptime(fim, "%d/%m/%Y"))
    http = session or requests
    partes: list[pd.DataFrame] = []
    url = SGS_API.format(cod=codigo)
    for bloco_ini, bloco_fim in _blocos(inicio_ts, fim_ts):
        params = {
            "formato": "json",
            "dataInicial": bloco_ini.strftime("%d/%m/%Y"),
            "dataFinal": bloco_fim.strftime("%d/%m/%Y"),
        }
        dados = _get_json(http, url, params)
        if not dados:
            continue
        df = pd.DataFrame(dados)
        df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
        partes.append(df.dropna(subset=["data", "valor"]))
    if not partes:
        return pd.DataFrame(columns=["codigo", "data", "valor"])
    out = (
        pd.concat(partes, ignore_index=True)
        .drop_duplicates(subset=["data"])
        .sort_values("data")
        .reset_index(drop=True)
    )
    out["codigo"] = codigo
    return out[["codigo", "data", "valor"]]


def _get_json(http, url: str, params: dict, tentativas: int = 4):
    """GET JSON do SGS; 404 = sem dados; corpo vazio/429 re-tenta."""
    ultimo_erro: Exception | None = None
    for tentativa in range(tentativas):
        resp = http.get(url, params=params, timeout=120)
        if resp.status_code == 404:
            return []
        corpo = getattr(resp, "content", b"x") or b""
        if resp.status_code in {429, 500, 502, 503} or not corpo.strip():
            ultimo_erro = RuntimeError(f"SGS HTTP {resp.status_code} vazio")
            time.sleep(1.5 * (tentativa + 1))
            continue
        resp.raise_for_status()
        try:
            dados = resp.json()
        except ValueError as exc:
            ultimo_erro = exc
            time.sleep(1.5 * (tentativa + 1))
            continue
        return dados or []
    if ultimo_erro:
        raise RuntimeError(f"Falha ao ler {url}: {ultimo_erro}") from ultimo_erro
    return []


def baixar_serie(
    codigo: int,
    inicio: str = "01/01/1995",
    fim: str | None = None,
    baixar=baixar_sgs,
) -> pd.DataFrame:
    return baixar(codigo, inicio=inicio, fim=fim)


def _rotulo_mes(ts: pd.Timestamp) -> str:
    return f"{ts.month:02d}/{ts.year}"


def agregados_anuais(df: pd.DataFrame, anos: list[int]) -> pd.DataFrame:
    """Para cada ano: valor de dezembro (ou último mês) e soma dos meses."""
    cols = [
        "codigo",
        "ano",
        "data",
        "valor_dezembro_rs_mil",
        "variacao_ano_rs_mil",
        "n_meses",
        "fechamento",
    ]
    if df.empty:
        return pd.DataFrame(columns=cols)
    work = df.copy()
    work["ano"] = work["data"].dt.year
    work = work[work["ano"].isin(anos)]
    if work.empty:
        return pd.DataFrame(columns=cols)
    last = work.sort_values("data").groupby("ano", as_index=False).tail(1)
    last["fechamento"] = last["data"].map(
        lambda ts: "dezembro" if ts.month == 12 else f"ultimo_{_rotulo_mes(ts)}"
    )
    last = last.rename(columns={"valor": "valor_dezembro_rs_mil"})
    soma = (
        work.groupby("ano", as_index=False)
        .agg(variacao_ano_rs_mil=("valor", "sum"), n_meses=("valor", "size"))
    )
    out = last.merge(soma, on="ano")
    out["codigo"] = df["codigo"].iloc[0]
    return out[cols]


def montar_tabelas(
    series_long: dict[int, pd.DataFrame],
    anos: list[int],
    catalogo: list[dict] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Devolve (longo, dezembro em R$ mi, variação no ano em R$ mi)."""
    catalogo = catalogo if catalogo is not None else SERIES
    partes: list[pd.DataFrame] = []
    for meta in catalogo:
        bruto = series_long.get(int(meta["codigo"]))
        if bruto is None or bruto.empty:
            continue
        anual = agregados_anuais(bruto, anos)
        if anual.empty:
            continue
        anual["serie"] = meta["nome"]
        anual["coluna"] = meta["coluna"]
        anual["grupo"] = meta["grupo"]
        anual["valor_dezembro_rs_milhoes"] = anual["valor_dezembro_rs_mil"] / 1_000.0
        anual["variacao_ano_rs_milhoes"] = anual["variacao_ano_rs_mil"] / 1_000.0
        partes.append(anual)
    if not partes:
        raise RuntimeError("Nenhuma série SGS retornou dados no período")
    longo = pd.concat(partes, ignore_index=True).sort_values(["ano", "codigo"])
    longo = longo.reset_index(drop=True)
    longo = _corrigir_variacao_estoque(longo)

    dez = _pivot(longo, "valor_dezembro_rs_milhoes", catalogo)
    var = _pivot(longo, "variacao_ano_rs_milhoes", catalogo)
    dez = _anexar_status(dez, longo)
    var = _anexar_status(var, longo)
    var = _anexar_identidade(var, dez)
    return longo, dez, var


def _corrigir_variacao_estoque(longo: pd.DataFrame) -> pd.DataFrame:
    """Base monetária é estoque: variação anual = Δ do saldo de dezembro."""
    out = longo.copy()
    mask = out["grupo"] == "resultado"
    if not mask.any():
        return out
    base = (
        out.loc[mask, ["codigo", "ano", "valor_dezembro_rs_mil"]]
        .sort_values(["codigo", "ano"])
        .copy()
    )
    base["variacao_ano_rs_mil"] = base.groupby("codigo")["valor_dezembro_rs_mil"].diff()
    out = out.merge(
        base[["codigo", "ano", "variacao_ano_rs_mil"]].rename(
            columns={"variacao_ano_rs_mil": "_var_estoque"}
        ),
        on=["codigo", "ano"],
        how="left",
    )
    mask = out["grupo"] == "resultado"
    out.loc[mask, "variacao_ano_rs_mil"] = out.loc[mask, "_var_estoque"]
    out.loc[mask, "variacao_ano_rs_milhoes"] = (
        out.loc[mask, "variacao_ano_rs_mil"] / 1_000.0
    )
    return out.drop(columns=["_var_estoque"])


def _pivot(longo: pd.DataFrame, valor: str, catalogo: list[dict]) -> pd.DataFrame:
    largos = []
    for meta in catalogo:
        sub = longo.loc[longo["codigo"] == meta["codigo"], ["ano", valor]]
        if sub.empty:
            continue
        largos.append(sub.rename(columns={valor: meta["coluna"]}))
    if not largos:
        return pd.DataFrame(columns=["ano"])
    out = largos[0]
    for extra in largos[1:]:
        out = out.merge(extra, on="ano", how="outer")
    return out.sort_values("ano").reset_index(drop=True)


def _anexar_status(largo: pd.DataFrame, longo: pd.DataFrame) -> pd.DataFrame:
    refs = (
        longo.groupby("ano", as_index=False)["data"]
        .max()
        .rename(columns={"data": "data_ref"})
    )
    status = (
        longo.groupby("ano")["fechamento"]
        .agg(lambda s: "dezembro" if (s == "dezembro").all() else "parcial")
        .reset_index()
    )
    out = largo.merge(refs, on="ano").merge(status, on="ano")
    frente = ["ano", "data_ref", "fechamento"]
    resto = [c for c in out.columns if c not in frente]
    return out[frente + resto]


def _anexar_identidade(var: pd.DataFrame, dez: pd.DataFrame) -> pd.DataFrame:
    out = var.copy()
    cols = [c for c in FATORES_IDENTIDADE if c in out.columns]
    if cols:
        out["soma_fatores"] = out[cols].fillna(0.0).sum(axis=1)
    if "base_monetaria_restrita" in dez.columns:
        base = dez[["ano", "base_monetaria_restrita"]].rename(
            columns={"base_monetaria_restrita": "saldo_base"}
        )
        out = out.merge(base, on="ano", how="left")
        out["variacao_base"] = out["saldo_base"].diff()
        if "soma_fatores" in out.columns:
            out["discrepancia"] = out["soma_fatores"] - out["variacao_base"]
    return out


def formatar_milhoes(valor: float | None) -> str:
    if valor is None or pd.isna(valor):
        return "—"
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _colunas_visiveis(largo: pd.DataFrame, catalogo: list[dict]) -> list[dict]:
    return [m for m in catalogo if m["coluna"] in largo.columns]


def markdown_tabela(
    largo: pd.DataFrame,
    titulo: str,
    nota: str,
    catalogo: list[dict] | None = None,
    extras: list[str] | None = None,
) -> str:
    catalogo = catalogo if catalogo is not None else SERIES
    visiveis = _colunas_visiveis(largo, catalogo)
    extra_cols = [c for c in (extras or []) if c in largo.columns]
    cab = ["Ano", "Mês ref."] + [m["nome"] for m in visiveis]
    nomes_extra = {
        "soma_fatores": "Soma dos fatores",
        "variacao_base": "Δ base monetária",
        "saldo_base": "Saldo da base",
        "discrepancia": "Discrepância",
    }
    cab += [nomes_extra[c] for c in extra_cols]
    linhas = [
        f"# {titulo}",
        "",
        nota,
        "",
        "| " + " | ".join(cab) + " |",
        "|" + "|".join(["---"] * 2 + ["---:"] * (len(visiveis) + len(extra_cols))) + "|",
    ]
    for _, row in largo.iterrows():
        mes = pd.Timestamp(row["data_ref"]).strftime("%m/%Y")
        vals = [formatar_milhoes(row[m["coluna"]]) for m in visiveis]
        vals += [formatar_milhoes(row[c]) for c in extra_cols]
        linhas.append(f"| {int(row['ano'])} | {mes} | " + " | ".join(vals) + " |")
    linhas.extend(
        [
            "",
            "## Códigos SGS",
            "",
            "| Código | Série | Grupo | Início na série |",
            "|-------:|-------|-------|-----------------|",
        ]
    )
    for m in catalogo:
        linhas.append(
            f"| {m['codigo']} | {m['nome']} | {m['grupo']} | {m.get('inicio', '')} |"
        )
    linhas.extend(
        [
            "",
            "API: `https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados`.",
            "Valores em **R$ milhões**. SGS original em R$ mil.",
            "",
        ]
    )
    return "\n".join(linhas)


def gravar_saidas(
    longo: pd.DataFrame,
    dezembro: pd.DataFrame,
    variacao: pd.DataFrame,
    pasta: Path,
    stem: str = "fatores_condicionantes_base_monetaria",
) -> dict[str, Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    csv_longo = pasta / f"{stem}_anual_longo.csv"
    csv_dez = pasta / f"{stem}_dezembro_r$_milhoes.csv"
    csv_var = pasta / f"{stem}_variacao_anual_r$_milhoes.csv"
    xlsx = pasta / f"{stem}_anual.xlsx"
    md_dez = pasta / f"{stem}_dezembro.md"
    md_var = pasta / f"{stem}_variacao_anual.md"

    longo_out = longo.copy()
    longo_out["data"] = pd.to_datetime(longo_out["data"]).dt.strftime("%Y-%m-%d")
    dez_out = dezembro.copy()
    var_out = variacao.copy()
    for frame in (dez_out, var_out):
        frame["data_ref"] = pd.to_datetime(frame["data_ref"]).dt.strftime("%Y-%m-%d")

    longo_out.to_csv(csv_longo, index=False)
    dez_out.to_csv(csv_dez, index=False)
    var_out.to_csv(csv_var, index=False)

    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        var_out.to_excel(writer, sheet_name="Variacao_no_ano", index=False)
        dez_out.to_excel(writer, sheet_name="Valor_dezembro_SGS", index=False)
        longo_out.to_excel(writer, sheet_name="Agregado_anual_longo", index=False)
        pd.DataFrame(SERIES).to_excel(writer, sheet_name="Codigos_SGS", index=False)

    nota_dez = (
        "Fonte: Banco Central do Brasil, SGS — *saldo em final de período*. "
        "Para os fatores, o valor de dezembro é a **contribuição daquele mês** "
        "(não é estoque). A base monetária restrita é estoque."
    )
    nota_var = (
        "Fonte: Banco Central do Brasil, SGS. "
        "Variação no ano = soma das contribuições mensais de cada fator. "
        "A soma dos fatores fecha com a variação da base (discrepância residual)."
    )
    md_dez.write_text(
        markdown_tabela(
            dezembro,
            "Fatores condicionantes — valor de dezembro (SGS)",
            nota_dez,
        ),
        encoding="utf-8",
    )
    md_var.write_text(
        markdown_tabela(
            variacao,
            "Fatores condicionantes — variação acumulada no ano",
            nota_var,
            extras=["soma_fatores", "variacao_base", "saldo_base", "discrepancia"],
        ),
        encoding="utf-8",
    )
    return {
        "csv_longo": csv_longo,
        "csv_dezembro": csv_dez,
        "csv_variacao": csv_var,
        "xlsx": xlsx,
        "md_dezembro": md_dez,
        "md_variacao": md_var,
    }


def coletar(
    ano_inicio: int = ANO_INICIO_DEFAULT,
    ano_fim: int = ANO_FIM_DEFAULT,
    baixar=baixar_sgs,
    catalogo: list[dict] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    catalogo = catalogo if catalogo is not None else SERIES
    anos = anos_periodo(ano_inicio, ano_fim)
    fim = datetime.now().strftime("%d/%m/%Y")
    series_long: dict[int, pd.DataFrame] = {}
    for meta in catalogo:
        codigo = int(meta["codigo"])
        inicio = meta.get("inicio") or f"01/01/{ano_inicio}"
        print(f"[INFO] SGS {codigo} — {meta['nome']} (desde {inicio})")
        try:
            series_long[codigo] = baixar_serie(
                codigo, inicio=inicio, fim=fim, baixar=baixar
            )
        except Exception as exc:
            print(f"[AVISO] SGS {codigo} falhou: {exc}")
            series_long[codigo] = pd.DataFrame(columns=["codigo", "data", "valor"])
        n = len(series_long[codigo])
        print(f"       {n} observações")
        time.sleep(0.4)
    return montar_tabelas(series_long, anos, catalogo=catalogo)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Saldos de dezembro e variação anual dos fatores condicionantes "
            "da base monetária (Bacen SGS)."
        )
    )
    parser.add_argument("--ano-inicio", type=int, default=ANO_INICIO_DEFAULT)
    parser.add_argument("--ano-fim", type=int, default=ANO_FIM_DEFAULT)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "output")
    parser.add_argument(
        "--stem", default="fatores_condicionantes_base_monetaria"
    )
    args = parser.parse_args(argv)

    longo, dezembro, variacao = coletar(args.ano_inicio, args.ano_fim)
    caminhos = gravar_saidas(longo, dezembro, variacao, args.output_dir, args.stem)
    print()
    print(
        markdown_tabela(
            dezembro,
            "Fatores condicionantes — valor de dezembro (SGS)",
            "Contribuição do mês de dezembro. Base monetária = estoque.",
        )
    )
    print(
        markdown_tabela(
            variacao,
            "Fatores condicionantes — variação acumulada no ano",
            "Soma dos meses. Fecha com a variação da base monetária.",
            extras=["soma_fatores", "variacao_base", "saldo_base", "discrepancia"],
        )
    )
    print("[OK] Arquivos:")
    for nome, path in caminhos.items():
        print(f"  {nome}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
