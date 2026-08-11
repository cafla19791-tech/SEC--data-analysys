#!/usr/bin/env python3
"""Discriminativo mensal — Operações de Financiamento BNDES (Dados Abertos).

Fonte:
  https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento

Arquivos:
  - operacoes-nao-automaticas.csv (direta + indireta não automática)
  - operacoes-indiretas-automaticas.csv (indireta automática)

Agrega por mês (data_da_contratacao) e forma_de_apoio (DIRETA / INDIRETA),
com valor contratado/operação e valor desembolsado, em corrente e IPCA jun/2026
(Ipeadata PRECOS12_IPCA12).

Cobertura do dataset: a partir de 2002; painel até jun/2026.
"""

from __future__ import annotations

import json
import urllib.request
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "bndes" / "dados_abertos"
OUT_XLSX = ROOT / "output" / "bndes_operacoes_financiamento_mensal_ipca.xlsx"
OUT_MD = ROOT / "output" / "bndes_operacoes_financiamento_mensal_ipca.md"
IPCA_CACHE = ROOT / "data" / "raw" / "ipeadata" / "PRECOS12_IPCA12.json"

NAO_AUT = DATA / "operacoes_nao_automaticas.csv"
IND_AUT = DATA / "operacoes_indiretas_automaticas.csv"

URL_NAO = (
    "https://dadosabertos.bndes.gov.br/dataset/"
    "10e21ad1-568e-45e5-a8af-43f2c05ef1a2/resource/"
    "6f56b78c-510f-44b6-8274-78a5b7e931f4/download/"
    "operacoes-financiamento-operacoes-nao-automaticas.csv"
)
URL_IND = (
    "https://dadosabertos.bndes.gov.br/dataset/"
    "10e21ad1-568e-45e5-a8af-43f2c05ef1a2/resource/"
    "612faa0b-b6be-4b2c-9317-da5dc2c0b901/download/"
    "operacoes-financiamento-operacoes-indiretas-automaticas.csv"
)


def ensure_files() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    for path, url in ((NAO_AUT, URL_NAO), (IND_AUT, URL_IND)):
        if path.exists() and path.stat().st_size > 1_000_000:
            continue
        print(f"Baixando {path.name} ...", flush=True)
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=600) as resp:
            path.write_bytes(resp.read())


def load_ipca() -> pd.Series:
    IPCA_CACHE.parent.mkdir(parents=True, exist_ok=True)
    if IPCA_CACHE.exists() and IPCA_CACHE.stat().st_size > 1000:
        payload = json.loads(IPCA_CACHE.read_text())
    else:
        url = (
            "https://www.ipeadata.gov.br/api/odata4/"
            "ValoresSerie(SERCODIGO='PRECOS12_IPCA12')"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            payload = json.loads(resp.read().decode())
        IPCA_CACHE.write_text(json.dumps(payload), encoding="utf-8")
    rows = payload["value"] if isinstance(payload, dict) else payload
    records = {}
    for r in rows:
        d = r["VALDATA"][:10]
        y, m = int(d[:4]), int(d[5:7])
        records[(y, m)] = float(r["VALVALOR"])
    s = pd.Series(records, dtype=float)
    s.index = pd.MultiIndex.from_tuples(s.index, names=["Ano", "Mes"])
    return s.sort_index()


def _parse_br_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _normalize_forma(s: pd.Series) -> pd.Series:
    u = s.astype(str).str.strip().str.upper()
    out = pd.Series(pd.NA, index=s.index, dtype="string")
    out = out.mask(u.str.contains("INDIRETA", na=False), "INDIRETA")
    out = out.mask(u.str.contains("DIRETA", na=False) & ~u.str.contains("INDIRETA", na=False), "DIRETA")
    return out


def aggregate_csv(
    path: Path,
    value_col: str,
    chunksize: int = 200_000,
) -> pd.DataFrame:
    """Retorna agregado Ano/Mes/Forma com contratado e desembolsado."""
    usecols = [
        "data_da_contratacao",
        value_col,
        "valor_desembolsado_reais",
        "forma_de_apoio",
    ]
    chunks = []
    reader = pd.read_csv(
        path,
        sep=";",
        encoding="latin-1",
        usecols=usecols,
        dtype=str,
        chunksize=chunksize,
    )
    for i, chunk in enumerate(reader, 1):
        chunk = chunk.rename(columns={value_col: "valor_contratado"})
        chunk["forma"] = _normalize_forma(chunk["forma_de_apoio"])
        chunk["data"] = pd.to_datetime(chunk["data_da_contratacao"], errors="coerce")
        chunk["valor_contratado"] = _parse_br_number(chunk["valor_contratado"])
        chunk["valor_desembolsado"] = _parse_br_number(chunk["valor_desembolsado_reais"])
        chunk = chunk.dropna(subset=["data", "forma"])
        chunk["Ano"] = chunk["data"].dt.year
        chunk["Mes"] = chunk["data"].dt.month
        g = chunk.groupby(["Ano", "Mes", "forma"], as_index=False).agg(
            contratado=("valor_contratado", "sum"),
            desembolsado=("valor_desembolsado", "sum"),
            n=("valor_contratado", "size"),
        )
        chunks.append(g)
        if i % 5 == 0:
            print(f"  {path.name}: chunk {i}", flush=True)
    if not chunks:
        return pd.DataFrame(
            columns=["Ano", "Mes", "forma", "contratado", "desembolsado", "n"]
        )
    out = pd.concat(chunks, ignore_index=True)
    out = out.groupby(["Ano", "Mes", "forma"], as_index=False).agg(
        contratado=("contratado", "sum"),
        desembolsado=("desembolsado", "sum"),
        n=("n", "sum"),
    )
    return out


def build() -> dict[str, pd.DataFrame]:
    ensure_files()
    ipca = load_ipca()
    ipca_jun = float(ipca.loc[(2026, 6)])

    print("Agregando não automáticas...", flush=True)
    a = aggregate_csv(NAO_AUT, "valor_contratado_reais")
    print("Agregando indiretas automáticas...", flush=True)
    b = aggregate_csv(IND_AUT, "valor_da_operacao_em_reais")
    agg = pd.concat([a, b], ignore_index=True)
    agg = agg.groupby(["Ano", "Mes", "forma"], as_index=False).agg(
        contratado=("contratado", "sum"),
        desembolsado=("desembolsado", "sum"),
        n=("n", "sum"),
    )

    # painel jan/2002 .. jun/2026
    months = pd.period_range("2002-01", "2026-06", freq="M")
    pivot_c = agg.pivot_table(
        index=["Ano", "Mes"], columns="forma", values="contratado", aggfunc="sum"
    ).fillna(0.0)
    pivot_d = agg.pivot_table(
        index=["Ano", "Mes"], columns="forma", values="desembolsado", aggfunc="sum"
    ).fillna(0.0)
    for p in (pivot_c, pivot_d):
        for col in ("DIRETA", "INDIRETA"):
            if col not in p.columns:
                p[col] = 0.0

    rows = []
    for per in months:
        y, m = per.year, per.month
        key = (y, m)
        dir_c = float(pivot_c.loc[key, "DIRETA"]) if key in pivot_c.index else 0.0
        ind_c = float(pivot_c.loc[key, "INDIRETA"]) if key in pivot_c.index else 0.0
        dir_d = float(pivot_d.loc[key, "DIRETA"]) if key in pivot_d.index else 0.0
        ind_d = float(pivot_d.loc[key, "INDIRETA"]) if key in pivot_d.index else 0.0
        if key not in ipca.index:
            raise RuntimeError(f"IPCA ausente {m:02d}/{y}")
        fator = ipca_jun / float(ipca.loc[key])
        rows.append(
            {
                "Mes_Ano": f"{m:02d}/{y}",
                "Ano": y,
                "Mes": m,
                "Direta_contratado_corrente": dir_c,
                "Direta_contratado_IPCA_jun2026": dir_c * fator,
                "Indireta_contratado_corrente": ind_c,
                "Indireta_contratado_IPCA_jun2026": ind_c * fator,
                "Direta_desembolsado_corrente": dir_d,
                "Direta_desembolsado_IPCA_jun2026": dir_d * fator,
                "Indireta_desembolsado_corrente": ind_d,
                "Indireta_desembolsado_IPCA_jun2026": ind_d * fator,
                "Fator_IPCA": fator,
            }
        )
    panel = pd.DataFrame(rows)

    demo = pd.DataFrame(
        {
            "Mês/ano": panel["Mes_Ano"],
            "Operações diretas — valor corrente (R$)": panel[
                "Direta_contratado_corrente"
            ],
            "Operações diretas — valor atual IPCA jun/2026 (R$)": panel[
                "Direta_contratado_IPCA_jun2026"
            ],
            "Operações indiretas — valor corrente (R$)": panel[
                "Indireta_contratado_corrente"
            ],
            "Operações indiretas — valor atual IPCA jun/2026 (R$)": panel[
                "Indireta_contratado_IPCA_jun2026"
            ],
        }
    )
    demo_des = pd.DataFrame(
        {
            "Mês/ano": panel["Mes_Ano"],
            "Operações diretas — desembolso corrente (R$)": panel[
                "Direta_desembolsado_corrente"
            ],
            "Operações diretas — desembolso IPCA jun/2026 (R$)": panel[
                "Direta_desembolsado_IPCA_jun2026"
            ],
            "Operações indiretas — desembolso corrente (R$)": panel[
                "Indireta_desembolsado_corrente"
            ],
            "Operações indiretas — desembolso IPCA jun/2026 (R$)": panel[
                "Indireta_desembolsado_IPCA_jun2026"
            ],
        }
    )

    metodologia = pd.DataFrame(
        [
            {"Item": "Portal", "Valor": "https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento"},
            {"Item": "Conceito principal", "Valor": "Valor contratado/operação na data_da_contratacao"},
            {"Item": "Aba Discriminativo_contratado", "Valor": "Pedido: Mês/ano + direta/indireta corrente e IPCA"},
            {"Item": "Aba Discriminativo_desembolso", "Valor": "Mesma agregação sobre valor_desembolsado_reais"},
            {"Item": "IPCA", "Valor": "Ipeadata PRECOS12_IPCA12; fator = IPCA_jun2026 / IPCA_mês"},
            {"Item": "Período", "Valor": "jan/2002 a jun/2026 (início do dataset de dados abertos)"},
            {
                "Item": "Soma direta contratada corrente",
                "Valor": float(panel["Direta_contratado_corrente"].sum()),
            },
            {
                "Item": "Soma indireta contratada corrente",
                "Valor": float(panel["Indireta_contratado_corrente"].sum()),
            },
            {
                "Item": "Soma direta contratada IPCA jun/2026",
                "Valor": float(panel["Direta_contratado_IPCA_jun2026"].sum()),
            },
            {
                "Item": "Soma indireta contratada IPCA jun/2026",
                "Valor": float(panel["Indireta_contratado_IPCA_jun2026"].sum()),
            },
        ]
    )

    return {
        "demo": demo,
        "demo_des": demo_des,
        "panel": panel,
        "metodologia": metodologia,
        "agg": agg,
    }


def write_outputs(tables: dict[str, pd.DataFrame]) -> None:
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    demo, panel = tables["demo"], tables["panel"]
    bi = pd.DataFrame(
        {
            "Mês/ano": demo["Mês/ano"],
            "Direta corrente (R$ bi)": demo.iloc[:, 1] / 1e9,
            "Direta IPCA jun/2026 (R$ bi)": demo.iloc[:, 2] / 1e9,
            "Indireta corrente (R$ bi)": demo.iloc[:, 3] / 1e9,
            "Indireta IPCA jun/2026 (R$ bi)": demo.iloc[:, 4] / 1e9,
        }
    )
    anual = panel.groupby("Ano", as_index=False).agg(
        Direta_corrente=("Direta_contratado_corrente", "sum"),
        Direta_IPCA=("Direta_contratado_IPCA_jun2026", "sum"),
        Indireta_corrente=("Indireta_contratado_corrente", "sum"),
        Indireta_IPCA=("Indireta_contratado_IPCA_jun2026", "sum"),
    )
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        demo.to_excel(writer, sheet_name="Discriminativo_contratado", index=False)
        tables["demo_des"].to_excel(
            writer, sheet_name="Discriminativo_desembolso", index=False
        )
        bi.to_excel(writer, sheet_name="Contratado_R$_bi", index=False)
        anual.to_excel(writer, sheet_name="Anual_contratado", index=False)
        panel.to_excel(writer, sheet_name="Serie_completa", index=False)
        tables["metodologia"].to_excel(writer, sheet_name="Metodologia", index=False)

    lines = [
        "# Operações de Financiamento BNDES — mensal direta × indireta",
        "",
        "Fonte: [Dados Abertos BNDES](https://dadosabertos.bndes.gov.br/dataset/operacoes-financiamento)",
        "",
        f"Arquivo: `{OUT_XLSX.relative_to(ROOT)}`",
        "",
        "Valor principal: **contratado** na `data_da_contratacao` (jan/2002–jun/2026).",
        "",
        f"- Direta corrente: R$ {panel['Direta_contratado_corrente'].sum()/1e9:.2f} bi",
        f"- Indireta corrente: R$ {panel['Indireta_contratado_corrente'].sum()/1e9:.2f} bi",
        f"- Direta IPCA jun/2026: R$ {panel['Direta_contratado_IPCA_jun2026'].sum()/1e9:.2f} bi",
        f"- Indireta IPCA jun/2026: R$ {panel['Indireta_contratado_IPCA_jun2026'].sum()/1e9:.2f} bi",
        "",
        "```bash",
        "python3 scripts/build_bndes_operacoes_financiamento_mensal.py",
        "```",
    ]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT_XLSX}")
    print(f"Wrote {OUT_MD}")


if __name__ == "__main__":
    write_outputs(build())
