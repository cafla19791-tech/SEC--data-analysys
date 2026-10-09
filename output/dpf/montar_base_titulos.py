#!/usr/bin/env python3
"""Monta planilha DPF por título: estoque dez/2022, resgates 2019-2022, base individualizada."""

from __future__ import annotations

import re
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent


def parse_br_num(s) -> float:
    if pd.isna(s) or s is None or str(s).strip() == "":
        return 0.0
    s = str(s).strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def familia(titulo: str) -> str:
    t = titulo.upper().strip()
    for fam in (
        "NTN-B1",
        "NTN-B",
        "NTN-C",
        "NTN-D",
        "NTN-F",
        "NTN-I",
        "LFT",
        "LTN",
        "TDA",
        "NBC",
        "CFT",
    ):
        if t.startswith(fam) or t.startswith(fam.replace("-", "")):
            return fam
    if t.startswith("GLOBAL"):
        return "Global"
    if "EURO" in t:
        return "Euro"
    if t.startswith("TDAD") or t.startswith("TDA"):
        return "TDA"
    m = re.match(r"([A-Za-zÀ-ú0-9\-\+]+)", t)
    return m.group(1) if m else t[:20]


def load_estoque() -> pd.DataFrame:
    chunks = []
    for chunk in pd.read_csv(
        ROOT / "EstoqueDPF.csv", sep=";", decimal=",", chunksize=80000, dtype=str
    ):
        chunks.append(chunk)
    raw = pd.concat(chunks, ignore_index=True)
    raw["valor"] = raw["Valor do Estoque"].map(parse_br_num)
    raw["quantidade"] = raw["Quantidade do Estoque"].map(parse_br_num)
    raw["vencimento"] = pd.to_datetime(
        raw["Vencimento do Titulo/Contrato"], dayfirst=True, errors="coerce"
    )
    raw["mes"] = pd.to_datetime(raw["Mes do Estoque"], format="%m/%Y", errors="coerce")
    raw["titulo"] = raw["Titulo/Contrato"].astype(str).str.strip()
    raw["classe"] = raw["Classe da Carteira"].astype(str).str.strip()
    raw["tipo_divida"] = raw["Tipo de Divida"].astype(str).str.strip()
    raw["familia"] = raw["titulo"].map(familia)
    return raw


def build_estoque_dez2022(raw: pd.DataFrame):
    dez22 = raw[raw["Mes do Estoque"] == "12/2022"].copy()

    est_det = (
        dez22.groupby(
            ["titulo", "vencimento", "familia", "classe", "tipo_divida"], dropna=False
        )
        .agg(quantidade=("quantidade", "sum"), valor_estoque=("valor", "sum"))
        .reset_index()
    )
    est_det["pu_implicito"] = np.where(
        est_det["quantidade"] > 0,
        est_det["valor_estoque"] / est_det["quantidade"],
        np.nan,
    )
    est_det = est_det.sort_values(["classe", "familia", "vencimento", "titulo"])
    est_det["vencimento_fmt"] = est_det["vencimento"].dt.strftime("%d/%m/%Y")
    est_det["valor_bi"] = est_det["valor_estoque"] / 1e9
    est_det["posicao"] = "Dez/2022"

    merc = dez22[dez22["classe"] == "Mercado"]
    est_merc = (
        merc.groupby(["titulo", "vencimento", "familia", "tipo_divida"], dropna=False)
        .agg(quantidade=("quantidade", "sum"), valor_estoque=("valor", "sum"))
        .reset_index()
    )
    est_merc["pu_implicito"] = np.where(
        est_merc["quantidade"] > 0,
        est_merc["valor_estoque"] / est_merc["quantidade"],
        np.nan,
    )
    est_merc = est_merc.sort_values(["familia", "vencimento", "titulo"])
    est_merc["vencimento_fmt"] = est_merc["vencimento"].dt.strftime("%d/%m/%Y")
    est_merc["valor_bi"] = est_merc["valor_estoque"] / 1e9

    est_fam = (
        merc.groupby(["familia", "tipo_divida"])
        .agg(
            n_papeis=("titulo", "nunique"),
            quantidade=("quantidade", "sum"),
            valor_estoque=("valor", "sum"),
        )
        .reset_index()
    )
    est_fam["valor_bi"] = est_fam["valor_estoque"] / 1e9
    est_fam = est_fam.sort_values("valor_estoque", ascending=False)

    return dez22, est_det, est_merc, est_fam


def load_leiloes_compra() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Parse auction workbooks. Data starts at column B (index 1)."""
    auctions = []
    for y in [2019, 2020, 2021, 2022]:
        wb = load_workbook(ROOT / f"leiloes/historico-leiloes-{y}.xlsx", data_only=True)
        ws = wb.active
        for r in ws.iter_rows(min_row=8, values_only=True):
            # Leading empty col A → fields are B..N
            if r[1] is None:
                continue
            auctions.append(
                {
                    "ano": y,
                    "data_leilao": r[1],
                    "titulo_familia": r[2],
                    "tipo_leilao": r[3],
                    "volta": r[4],
                    "data_liquidacao": r[5],
                    "vencimento": r[6],
                    "oferta_qtd": r[7],
                    "taxa_media": r[8],
                    "taxa_corte": r[9],
                    "qtd_aceita": r[10],
                    "financeiro": r[11],
                    "qtd_bacen": r[12],
                    "financeiro_bacen": r[13] if len(r) > 13 else None,
                }
            )
    auc = pd.DataFrame(auctions)
    if auc.empty:
        empty = pd.DataFrame(
            columns=[
                "Ano",
                "Data_Leilao",
                "Data_Liquidacao",
                "Titulo",
                "Tipo_Leilao",
                "Volta",
                "Vencimento",
                "Quantidade",
                "Valor_R$",
                "Fonte",
            ]
        )
        return empty, empty.copy()

    auc["tipo_leilao"] = auc["tipo_leilao"].astype(str).str.strip()
    # COMPRA / EXTRA COMPRA = early redemptions; TROCA also retires bonds
    compra = auc[
        auc["tipo_leilao"].str.upper().str.contains("COMPRA|TROCA", na=False, regex=True)
    ].copy()
    compra["fonte"] = compra["tipo_leilao"].map(
        lambda t: (
            "Leilão de troca (oferta pública)"
            if "TROCA" in str(t).upper()
            else "Leilão de compra (oferta pública)"
        )
    )
    compra["vencimento_fmt"] = pd.to_datetime(
        compra["vencimento"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    compra["data_leilao_fmt"] = pd.to_datetime(
        compra["data_leilao"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")
    compra["data_liquidacao_fmt"] = pd.to_datetime(
        compra["data_liquidacao"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    det = compra[
        [
            "ano",
            "data_leilao_fmt",
            "data_liquidacao_fmt",
            "titulo_familia",
            "tipo_leilao",
            "volta",
            "vencimento_fmt",
            "qtd_aceita",
            "financeiro",
            "fonte",
        ]
    ].rename(
        columns={
            "ano": "Ano",
            "data_leilao_fmt": "Data_Leilao",
            "data_liquidacao_fmt": "Data_Liquidacao",
            "titulo_familia": "Titulo",
            "tipo_leilao": "Tipo_Leilao",
            "volta": "Volta",
            "vencimento_fmt": "Vencimento",
            "qtd_aceita": "Quantidade",
            "financeiro": "Valor_R$",
            "fonte": "Fonte",
        }
    )
    # Keep only positive accepted quantities
    det = det[pd.to_numeric(det["Quantidade"], errors="coerce").fillna(0) > 0].copy()

    if det.empty:
        agg = det.copy()
        agg["n_operacoes"] = []
        agg["Valor_R$_bi"] = []
        return det, agg

    agg = (
        det.groupby(["Titulo", "Vencimento"], dropna=False)
        .agg({"Quantidade": "sum", "Valor_R$": "sum"})
        .reset_index()
    )
    counts = det.groupby(["Titulo", "Vencimento"], dropna=False).size().reset_index(name="n_operacoes")
    agg = agg.merge(counts, on=["Titulo", "Vencimento"], how="left")
    agg["Fonte"] = "Leilão compra/troca 2019-2022"
    agg["Valor_R$_bi"] = agg["Valor_R$"] / 1e9
    return det, agg


def resgates_vencimento(raw: pd.DataFrame) -> pd.DataFrame:
    merc = raw[raw["classe"] == "Mercado"].copy()
    merc = merc[(merc["mes"] >= "2018-12-01") & (merc["mes"] <= "2023-01-31")]
    merc["ano_venc"] = merc["vencimento"].dt.year
    titles = merc[merc["ano_venc"].between(2019, 2022)][
        ["titulo", "vencimento", "familia", "tipo_divida"]
    ].drop_duplicates()

    rows = []
    for _, t in titles.iterrows():
        sub = merc[
            (merc["titulo"] == t["titulo"]) & (merc["vencimento"] == t["vencimento"])
        ].sort_values("mes")
        if sub.empty or pd.isna(t["vencimento"]):
            continue
        venc = t["vencimento"]
        before = sub[sub["mes"] <= (venc + pd.offsets.MonthEnd(0))]
        before = before[before["quantidade"] > 0]
        if before.empty:
            continue
        last = before.iloc[-1]
        after = sub[sub["mes"] > last["mes"]]
        vanished = (
            after.empty
            or (after["quantidade"].max() == 0)
            or (after["quantidade"].iloc[0] < last["quantidade"] * 0.05)
        )
        months_to_mat = (venc.year - last["mes"].year) * 12 + (
            venc.month - last["mes"].month
        )
        if months_to_mat > 2 and not vanished:
            continue
        if not (2019 <= venc.year <= 2022):
            continue
        rows.append(
            {
                "Titulo/Contrato": t["titulo"],
                "Familia": t["familia"],
                "Vencimento": venc.strftime("%d/%m/%Y"),
                "Tipo_Divida": t["tipo_divida"],
                "Mes_Ultima_Posicao": last["mes"].strftime("%m/%Y"),
                "Ano_Resgate": int(venc.year),
                "Quantidade": last["quantidade"],
                "Valor_R$": last["valor"],
                "PU_implicito": (
                    last["valor"] / last["quantidade"] if last["quantidade"] else np.nan
                ),
                "Fonte": "Vencimento (estoque Mercado — última posição)",
            }
        )
    return pd.DataFrame(rows)


def load_tesouro_direto() -> tuple[pd.DataFrame, pd.DataFrame]:
    def prep(path: Path, tipo: str) -> pd.DataFrame:
        df = pd.read_csv(path, sep=";", dtype=str)
        df["Data Resgate"] = pd.to_datetime(
            df["Data Resgate"], dayfirst=True, errors="coerce"
        )
        df["Vencimento do Titulo"] = pd.to_datetime(
            df["Vencimento do Titulo"], dayfirst=True, errors="coerce"
        )
        df["Quantidade"] = df["Quantidade"].map(parse_br_num)
        df["Valor"] = df["Valor"].map(parse_br_num)
        df = df[df["Data Resgate"].dt.year.between(2019, 2022)].copy()
        df["Tipo_Resgate"] = tipo
        df["Ano"] = df["Data Resgate"].dt.year
        return df

    rec = prep(ROOT / "td/recompras.csv", "Recompra TD")
    venc = prep(ROOT / "td/vencimentos.csv", "Vencimento TD")
    cols = [
        "Tipo Titulo",
        "Vencimento do Titulo",
        "Data Resgate",
        "Quantidade",
        "Valor",
        "Tipo_Resgate",
        "Ano",
    ]
    all_td = pd.concat([rec[cols], venc[cols]], ignore_index=True)
    all_td["Vencimento"] = all_td["Vencimento do Titulo"].dt.strftime("%d/%m/%Y")
    all_td["Data_Resgate"] = all_td["Data Resgate"].dt.strftime("%d/%m/%Y")
    det = all_td.rename(
        columns={"Tipo Titulo": "Titulo_TD", "Quantidade": "Quantidade", "Valor": "Valor_R$"}
    )[["Ano", "Data_Resgate", "Titulo_TD", "Vencimento", "Tipo_Resgate", "Quantidade", "Valor_R$"]]

    agg = (
        det.groupby(["Titulo_TD", "Vencimento", "Tipo_Resgate", "Ano"], dropna=False)
        .agg({"Quantidade": "sum", "Valor_R$": "sum"})
        .reset_index()
    )
    agg["Valor_R$_bi"] = agg["Valor_R$"] / 1e9
    return det, agg


def main() -> None:
    print("Loading estoque...")
    raw = load_estoque()
    dez22, est_det, est_merc, est_fam = build_estoque_dez2022(raw)
    print(
        f"Dez/2022 Mercado: {len(est_merc)} papéis, "
        f"R$ {est_merc['valor_estoque'].sum()/1e9:.2f} bi"
    )

    print("Loading leilões compra...")
    leilao_det, leilao_agg = load_leiloes_compra()
    print(
        f"Leilões compra: {len(leilao_det)} ops, "
        f"R$ {leilao_det['Valor_R$'].sum()/1e9:.2f} bi"
    )

    print("Inferring maturity redemptions...")
    resg_venc = resgates_vencimento(raw)
    print(
        f"Vencimentos: {len(resg_venc)} papéis, "
        f"R$ {resg_venc['Valor_R$'].sum()/1e9:.2f} bi"
        if len(resg_venc)
        else "Vencimentos: 0"
    )

    print("Loading Tesouro Direto...")
    td_det, td_agg = load_tesouro_direto()
    print(f"TD: {len(td_det)} linhas, R$ {td_det['Valor_R$'].sum()/1e9:.2f} bi")

    # Consolidated
    parts = []
    if len(leilao_agg):
        a = leilao_agg.rename(columns={"Titulo": "titulo", "Vencimento": "vencimento"})
        a["tipo_resgate"] = "Compra/troca em leilão"
        a["fonte"] = "Leilão compra/troca 2019-2022"
        parts.append(
            a[
                ["titulo", "vencimento", "tipo_resgate", "Quantidade", "Valor_R$", "fonte"]
            ].rename(
                columns={"Quantidade": "quantidade", "Valor_R$": "financeiro_R$"}
            )
        )

    if len(resg_venc):
        b = resg_venc.rename(
            columns={
                "Titulo/Contrato": "titulo",
                "Vencimento": "vencimento",
                "Quantidade": "quantidade",
                "Valor_R$": "financeiro_R$",
                "Fonte": "fonte",
            }
        )
        b["tipo_resgate"] = "Vencimento"
        parts.append(
            b[["titulo", "vencimento", "tipo_resgate", "quantidade", "financeiro_R$", "fonte"]]
        )

    c = td_agg.rename(
        columns={
            "Titulo_TD": "titulo",
            "Vencimento": "vencimento",
            "Tipo_Resgate": "tipo_resgate",
            "Quantidade": "quantidade",
            "Valor_R$": "financeiro_R$",
        }
    )
    c["fonte"] = "Tesouro Direto"
    parts.append(
        c[["titulo", "vencimento", "tipo_resgate", "quantidade", "financeiro_R$", "fonte"]]
    )

    consol = pd.concat(parts, ignore_index=True)
    consol["financeiro_R$_bi"] = consol["financeiro_R$"] / 1e9
    consol = consol.sort_values(["tipo_resgate", "titulo", "vencimento"])

    # Export frames
    est_merc_out = pd.DataFrame(
        {
            "Titulo/Contrato": est_merc["titulo"],
            "Familia": est_merc["familia"],
            "Vencimento": est_merc["vencimento_fmt"],
            "Tipo_Divida": est_merc["tipo_divida"],
            "Quantidade": est_merc["quantidade"],
            "Valor_Estoque_R$": est_merc["valor_estoque"],
            "Valor_R$_bi": est_merc["valor_bi"],
            "PU_implicito": est_merc["pu_implicito"],
        }
    )

    base_out = pd.DataFrame(
        {
            "Titulo/Contrato": est_det["titulo"],
            "Familia": est_det["familia"],
            "Vencimento": est_det["vencimento_fmt"],
            "Classe_Carteira": est_det["classe"],
            "Tipo_Divida": est_det["tipo_divida"],
            "Quantidade": est_det["quantidade"],
            "Valor_Estoque_R$": est_det["valor_estoque"],
            "Valor_R$_bi": est_det["valor_bi"],
            "PU_implicito": est_det["pu_implicito"],
            "Posicao": est_det["posicao"],
        }
    )

    resumo = pd.DataFrame(
        [
            ["Estoque Mercado Dez/2022 (R$ bi)", round(est_merc["valor_estoque"].sum() / 1e9, 2)],
            [
                "Estoque BC Dez/2022 (R$ bi)",
                round(dez22.loc[dez22["classe"] == "Banco Central", "valor"].sum() / 1e9, 2),
            ],
            ["Nº papéis Mercado Dez/2022", len(est_merc_out)],
            ["Nº papéis Base (Mercado+BC)", len(base_out)],
            [
                "Resgates leilão compra 2019-22 (R$ bi)",
                round(leilao_det["Valor_R$"].sum() / 1e9, 2),
            ],
            [
                "Resgates vencimento (estoque) 2019-22 (R$ bi)",
                round(resg_venc["Valor_R$"].sum() / 1e9, 2) if len(resg_venc) else 0,
            ],
            [
                "Resgates Tesouro Direto 2019-22 (R$ bi)",
                round(td_det["Valor_R$"].sum() / 1e9, 2),
            ],
        ],
        columns=["Indicador", "Valor"],
    )

    readme = pd.DataFrame(
        {
            "Item": [
                "1_Estoque_Dez2022_Mercado",
                "1b_Estoque_por_Familia",
                "2_Resgates_Leiloes_Compra",
                "2b_Resgates_Leiloes_Agg",
                "2c_Resgates_Vencimento",
                "2d_Resgates_Tesouro_Direto",
                "2e_Resgates_Consolidados",
                "3_Base_Individualizada",
                "Fontes",
                "Nota — vencimento",
                "Nota — leilão compra",
                "Nota — TD",
            ],
            "Descricao": [
                "Estoque DPF em poder do público (Mercado) em dez/2022, por título e vencimento.",
                "Agregado do estoque Mercado por família de título.",
                "Resgates via leilões COMPRA/EXTRA COMPRA 2019–2022 (financeiro efetivo).",
                "Leilões de compra agregados por família + vencimento.",
                "Resgates por vencimento: última quantidade/valor no estoque Mercado antes do vencimento (2019–2022). Valor = marcação a mercado da última posição.",
                "Recompras e vencimentos do Tesouro Direto 2019–2022.",
                "Consolidação das fontes de resgate por papel.",
                "Posição individualizada por papel em dez/2022 (Mercado e Banco Central).",
                "EstoqueDPF.csv (Tesouro Transparente); ds013 leilões DPMFi; Resgates do Tesouro Direto.",
                "Aproximação do principal resgatado no vencimento; não inclui cupons isolados.",
                "Resgates antecipados oficiais com valor financeiro liquidado.",
                "Parcela específica dos resgates da DPMFi no programa Tesouro Direto.",
            ],
        }
    )

    out_path = ROOT / "DPF_base_titulos_2019_2022.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="0_README", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        est_merc_out.to_excel(writer, sheet_name="1_Estoque_Dez2022_Mercado", index=False)
        est_fam.to_excel(writer, sheet_name="1b_Estoque_por_Familia", index=False)
        leilao_det.to_excel(writer, sheet_name="2_Resgates_Leiloes_Compra", index=False)
        leilao_agg.to_excel(writer, sheet_name="2b_Resgates_Leiloes_Agg", index=False)
        resg_venc.to_excel(writer, sheet_name="2c_Resgates_Vencimento", index=False)
        td_det.to_excel(writer, sheet_name="2d_Resgates_Tesouro_Direto", index=False)
        consol.to_excel(writer, sheet_name="2e_Resgates_Consolidados", index=False)
        base_out.to_excel(writer, sheet_name="3_Base_Individualizada", index=False)

    est_merc_out.to_csv(ROOT / "estoque_dez2022_por_titulo.csv", index=False)
    base_out.to_csv(ROOT / "base_individualizada_dez2022.csv", index=False)
    consol.to_csv(ROOT / "resgates_por_titulo_2019_2022.csv", index=False)
    resg_venc.to_csv(ROOT / "resgates_vencimento_2019_2022.csv", index=False)
    leilao_det.to_csv(ROOT / "resgates_leiloes_compra_2019_2022.csv", index=False)

    print("Saved", out_path)
    print(resumo.to_string(index=False))
    print("Top 10 Mercado:")
    print(
        est_merc_out.nlargest(10, "Valor_Estoque_R$")[
            ["Titulo/Contrato", "Vencimento", "Valor_R$_bi"]
        ].to_string(index=False)
    )


if __name__ == "__main__":
    main()
