#!/usr/bin/env python3
"""Discriminativo mensal de desembolsos BNDES (direta/indireta), jan/1995–jun/2026.

Fontes:
  Bases de Desembolso do Sistema BNDES (arquivos anuais/plurianuais 1995–2026)
  https://www.bndes.gov.br/

IPCA:
  Ipeadata PRECOS12_IPCA12 (índice mensal)
  http://www.ipeadata.gov.br/

Atualização: valor_atual = valor_corrente × (IPCA_jun/2026 / IPCA_mês)
"""

from __future__ import annotations

import json
import urllib.request
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "bndes" / "raw"
OUT_XLSX = ROOT / "output" / "bndes_desembolsos_mensal_direta_indireta_ipca.xlsx"
OUT_MD = ROOT / "output" / "bndes_desembolsos_mensal_direta_indireta_ipca.md"
IPCA_CACHE = ROOT / "data" / "raw" / "ipeadata" / "PRECOS12_IPCA12.json"

MONTH_MAP = {
    "JANEIRO": 1,
    "FEVEREIRO": 2,
    "MARCO": 3,
    "MARÇO": 3,
    "ABRIL": 4,
    "MAIO": 5,
    "JUNHO": 6,
    "JULHO": 7,
    "AGOSTO": 8,
    "SETEMBRO": 9,
    "OUTUBRO": 10,
    "NOVEMBRO": 11,
    "DEZEMBRO": 12,
}

USECOLS = ["ANO", "MÊS", "FORMA DE APOIO", "DESEMBOLSOS\n(R$)"]


def load_ipca() -> pd.Series:
    IPCA_CACHE.parent.mkdir(parents=True, exist_ok=True)
    if IPCA_CACHE.exists() and IPCA_CACHE.stat().st_size > 1000:
        payload = json.loads(IPCA_CACHE.read_text())
        rows = payload["value"] if isinstance(payload, dict) else payload
    else:
        url = (
            "https://www.ipeadata.gov.br/api/odata4/"
            "ValoresSerie(SERCODIGO='PRECOS12_IPCA12')"
        )
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            payload = json.loads(resp.read().decode())
        rows = payload["value"]
        IPCA_CACHE.write_text(json.dumps(payload), encoding="utf-8")

    records = []
    for r in rows:
        d = r["VALDATA"][:10]
        y, m = int(d[:4]), int(d[5:7])
        records.append(((y, m), float(r["VALVALOR"])))
    s = pd.Series(dict(records), dtype=float)
    s.index = pd.MultiIndex.from_tuples(s.index, names=["Ano", "Mes"])
    return s.sort_index()


def _normalize_month(val) -> int | None:
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        m = int(val)
        return m if 1 <= m <= 12 else None
    key = str(val).strip().upper()
    # remove accents variants already handled; strip trailing spaces
    key = (
        key.replace("Ç", "C")
        .replace("Ã", "A")
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
    )
    # map without accent
    plain = {
        "JANEIRO": 1,
        "FEVEREIRO": 2,
        "MARCO": 3,
        "ABRIL": 4,
        "MAIO": 5,
        "JUNHO": 6,
        "JULHO": 7,
        "AGOSTO": 8,
        "SETEMBRO": 9,
        "OUTUBRO": 10,
        "NOVEMBRO": 11,
        "DEZEMBRO": 12,
    }
    return plain.get(key) or MONTH_MAP.get(str(val).strip().upper())


def _normalize_forma(val) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip().upper()
    if "INDIRETA" in s:
        return "INDIRETA"
    if "DIRETA" in s:
        return "DIRETA"
    return None


def read_file(path: Path) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name="DESEMBOLSOS_BASE DE DADOS",
        skiprows=2,
        usecols=lambda c: str(c).strip() in {
            "ANO",
            "MÊS",
            "MES",
            "FORMA DE APOIO",
            "DESEMBOLSOS\n(R$)",
            "DESEMBOLSOS (R$)",
        }
        or "DESEMBOLSO" in str(c).upper(),
    )
    # standardize columns
    rename = {}
    for c in df.columns:
        cu = str(c).strip().upper().replace("\n", " ")
        if cu == "ANO":
            rename[c] = "ANO"
        elif cu in ("MÊS", "MES"):
            rename[c] = "MES"
        elif "FORMA" in cu:
            rename[c] = "FORMA"
        elif "DESEMBOLSO" in cu:
            rename[c] = "VALOR"
    df = df.rename(columns=rename)
    need = {"ANO", "MES", "FORMA", "VALOR"}
    if not need.issubset(df.columns):
        raise RuntimeError(f"{path.name}: colunas {df.columns.tolist()}")
    df = df[["ANO", "MES", "FORMA", "VALOR"]].copy()
    df["ANO"] = pd.to_numeric(df["ANO"], errors="coerce").astype("Int64")
    df["MES_N"] = df["MES"].map(_normalize_month)
    df["FORMA_N"] = df["FORMA"].map(_normalize_forma)
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["ANO", "MES_N", "FORMA_N"])
    df["ANO"] = df["ANO"].astype(int)
    df["MES_N"] = df["MES_N"].astype(int)
    return df


def aggregate_all() -> pd.DataFrame:
    files = sorted(RAW_DIR.glob("*.xlsx"), key=lambda p: int(p.name.split("_")[0]))
    if not files:
        raise FileNotFoundError(f"Nenhum xlsx em {RAW_DIR}")

    chunks = []
    for f in files:
        print(f"Lendo {f.name} ...", flush=True)
        df = read_file(f)
        g = (
            df.groupby(["ANO", "MES_N", "FORMA_N"], as_index=False)["VALOR"]
            .sum()
        )
        chunks.append(g)
        print(
            f"  linhas={len(df):,}  soma=R$ {df['VALOR'].sum()/1e9:.2f} bi  "
            f"anos={df['ANO'].min()}-{df['ANO'].max()}",
            flush=True,
        )

    allg = pd.concat(chunks, ignore_index=True)
    allg = allg.groupby(["ANO", "MES_N", "FORMA_N"], as_index=False)["VALOR"].sum()
    return allg


def build_panel(agg: pd.DataFrame, ipca: pd.Series) -> pd.DataFrame:
    # full calendar jan/1995 .. jun/2026
    months = pd.period_range("1995-01", "2026-06", freq="M")
    pivot = agg.pivot_table(
        index=["ANO", "MES_N"], columns="FORMA_N", values="VALOR", aggfunc="sum"
    ).fillna(0.0)
    for col in ("DIRETA", "INDIRETA"):
        if col not in pivot.columns:
            pivot[col] = 0.0

    rows = []
    ipca_jun = float(ipca.loc[(2026, 6)])
    for per in months:
        y, m = per.year, per.month
        direta = float(pivot.loc[(y, m), "DIRETA"]) if (y, m) in pivot.index else 0.0
        indireta = (
            float(pivot.loc[(y, m), "INDIRETA"]) if (y, m) in pivot.index else 0.0
        )
        if (y, m) not in ipca.index:
            raise RuntimeError(f"IPCA ausente para {m:02d}/{y}")
        fator = ipca_jun / float(ipca.loc[(y, m)])
        rows.append(
            {
                "Mes_Ano": f"{m:02d}/{y}",
                "Ano": y,
                "Mes": m,
                "Direta_corrente_R$": direta,
                "Direta_atual_IPCA_jun2026_R$": direta * fator,
                "Indireta_corrente_R$": indireta,
                "Indireta_atual_IPCA_jun2026_R$": indireta * fator,
                "Fator_IPCA_mes_para_jun2026": fator,
                "IPCA_mes_Ipeadata": float(ipca.loc[(y, m)]),
                "IPCA_jun2026": ipca_jun,
            }
        )
    return pd.DataFrame(rows)


def write_outputs(panel: pd.DataFrame) -> None:
    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    # Presentation sheet with exact column layout requested
    demo = pd.DataFrame(
        {
            "Mês/ano": panel["Mes_Ano"],
            "Operações diretas — valor corrente (R$)": panel["Direta_corrente_R$"],
            "Operações diretas — valor atual IPCA jun/2026 (R$)": panel[
                "Direta_atual_IPCA_jun2026_R$"
            ],
            "Operações indiretas — valor corrente (R$)": panel["Indireta_corrente_R$"],
            "Operações indiretas — valor atual IPCA jun/2026 (R$)": panel[
                "Indireta_atual_IPCA_jun2026_R$"
            ],
        }
    )
    bi = pd.DataFrame(
        {
            "Mês/ano": panel["Mes_Ano"],
            "Direta corrente (R$ bi)": panel["Direta_corrente_R$"] / 1e9,
            "Direta atual IPCA jun/2026 (R$ bi)": panel["Direta_atual_IPCA_jun2026_R$"]
            / 1e9,
            "Indireta corrente (R$ bi)": panel["Indireta_corrente_R$"] / 1e9,
            "Indireta atual IPCA jun/2026 (R$ bi)": panel[
                "Indireta_atual_IPCA_jun2026_R$"
            ]
            / 1e9,
        }
    )
    anual = (
        panel.groupby("Ano", as_index=False)
        .agg(
            Direta_corrente=("Direta_corrente_R$", "sum"),
            Direta_atual=("Direta_atual_IPCA_jun2026_R$", "sum"),
            Indireta_corrente=("Indireta_corrente_R$", "sum"),
            Indireta_atual=("Indireta_atual_IPCA_jun2026_R$", "sum"),
        )
    )
    anual["Direta_corrente_bi"] = anual["Direta_corrente"] / 1e9
    anual["Direta_atual_bi"] = anual["Direta_atual"] / 1e9
    anual["Indireta_corrente_bi"] = anual["Indireta_corrente"] / 1e9
    anual["Indireta_atual_bi"] = anual["Indireta_atual"] / 1e9

    metodologia = pd.DataFrame(
        [
            {
                "Item": "Fonte desembolsos",
                "Valor": "Bases de Desembolso do Sistema BNDES (FORMA DE APOIO)",
            },
            {
                "Item": "Período painel",
                "Valor": "jan/1995 a jun/2026 (meses sem desembolso = 0)",
            },
            {
                "Item": "Cobertura BNDES 2026",
                "Valor": "Arquivo 2026 cobre até mar/2026 na CAPA; abr–jun/2026 = 0 se não houver lançamentos",
            },
            {
                "Item": "IPCA",
                "Valor": "Ipeadata PRECOS12_IPCA12 — http://www.ipeadata.gov.br/",
            },
            {
                "Item": "Fórmula",
                "Valor": "valor_atual = valor_corrente × (IPCA_jun/2026 / IPCA_mês)",
            },
            {
                "Item": "IPCA jun/2026",
                "Valor": float(panel["IPCA_jun2026"].iloc[0]),
            },
            {
                "Item": "Soma direta corrente",
                "Valor": float(panel["Direta_corrente_R$"].sum()),
            },
            {
                "Item": "Soma indireta corrente",
                "Valor": float(panel["Indireta_corrente_R$"].sum()),
            },
            {
                "Item": "Soma direta atual IPCA jun/2026",
                "Valor": float(panel["Direta_atual_IPCA_jun2026_R$"].sum()),
            },
            {
                "Item": "Soma indireta atual IPCA jun/2026",
                "Valor": float(panel["Indireta_atual_IPCA_jun2026_R$"].sum()),
            },
        ]
    )

    fatores = panel[
        [
            "Mes_Ano",
            "IPCA_mes_Ipeadata",
            "IPCA_jun2026",
            "Fator_IPCA_mes_para_jun2026",
        ]
    ].copy()

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        demo.to_excel(writer, sheet_name="Discriminativo_mensal", index=False)
        bi.to_excel(writer, sheet_name="Mensal_R$_bi", index=False)
        anual.to_excel(writer, sheet_name="Anual", index=False)
        fatores.to_excel(writer, sheet_name="Fatores_IPCA", index=False)
        panel.to_excel(writer, sheet_name="Serie_completa", index=False)
        metodologia.to_excel(writer, sheet_name="Metodologia", index=False)

    # short markdown
    lines = [
        "# Desembolsos BNDES mensais — Direta × Indireta (jan/1995–jun/2026)",
        "",
        f"Arquivo: `{OUT_XLSX.relative_to(ROOT)}`",
        "",
        "IPCA: Ipeadata `PRECOS12_IPCA12`, atualizado até jun/2026.",
        "",
        f"- Soma operações diretas (corrente): R$ {panel['Direta_corrente_R$'].sum()/1e9:.2f} bi",
        f"- Soma operações indiretas (corrente): R$ {panel['Indireta_corrente_R$'].sum()/1e9:.2f} bi",
        f"- Soma diretas (IPCA jun/2026): R$ {panel['Direta_atual_IPCA_jun2026_R$'].sum()/1e9:.2f} bi",
        f"- Soma indiretas (IPCA jun/2026): R$ {panel['Indireta_atual_IPCA_jun2026_R$'].sum()/1e9:.2f} bi",
        "",
        "```bash",
        "python3 scripts/build_bndes_desembolsos_mensal.py",
        "```",
    ]
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {OUT_XLSX}")
    print(f"Wrote {OUT_MD}")


def main() -> None:
    ipca = load_ipca()
    agg = aggregate_all()
    panel = build_panel(agg, ipca)
    write_outputs(panel)


if __name__ == "__main__":
    main()
