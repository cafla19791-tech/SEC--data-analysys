#!/usr/bin/env python3
"""Discriminativo 2º turno 2014/2018/2022 no interior do Nordeste (até 40 mil eleitores).

Interior = municípios das 9 UFs do Nordeste, excluídas as capitais.
Universo = até 40.000 eleitores aptos no 2º turno de 2022.
Em cada pleito: diferença percentual sobre votos válidos e sobre votos
totais (comparecimento = válidos + brancos + nulos).

Saída:
  output/tse_planilhas/discriminativo_interior_nordeste_ate_40mil.xlsx

Uso:
  python3 scripts/discriminativo_interior_nordeste_ate_40mil.py
  python discriminativo_interior_nordeste_ate_40mil.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.discriminativo_interior_nordeste import (
        marcar_capital_interior,
        pct,
        recorte_nordeste,
    )
    from scripts.planilha_resultados_presidente import (
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )
except ImportError:  # ContAgil
    from discriminativo_interior_nordeste import (  # type: ignore
        marcar_capital_interior,
        pct,
        recorte_nordeste,
    )
    from planilha_resultados_presidente import (  # type: ignore
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )

ANOS = (2014, 2018, 2022)
LIMITE_APTOS = 40_000

LADOS = {
    2014: ("DILMA", "AECIO", "Dilma", "Aécio"),
    2018: ("HADDAD", "BOLSONARO", "Haddad", "Bolsonaro"),
    2022: ("LULA", "BOLSONARO", "Lula", "Bolsonaro"),
}

CAMPOS_ANO = (
    "CAND_PT",
    "CAND_OPP",
    "QT_SECOES",
    "QT_VOTOS_PT",
    "QT_VOTOS_OPP",
    "QT_DIF_PT_OPP",
    "QT_VOTOS_VALIDOS",
    "PCT_PT_VALIDOS",
    "PCT_OPP_VALIDOS",
    "DIF_PCT_VALIDOS",
    "QT_VOTOS_BRANCO",
    "QT_VOTOS_NULO",
    "QT_VOTOS_TOTAIS",
    "PCT_PT_TOTAIS",
    "PCT_OPP_TOTAIS",
    "DIF_PCT_TOTAIS",
    "QT_APTOS",
    "VENCEDOR",
)

ROTOLOS_CAMPO = {
    "CAND_PT": "Candidato (PT)",
    "CAND_OPP": "Candidato (opp.)",
    "QT_SECOES": "Seções",
    "QT_VOTOS_PT": "Votos PT",
    "QT_VOTOS_OPP": "Votos oposição",
    "QT_DIF_PT_OPP": "Diferença (PT − opp.)",
    "QT_VOTOS_VALIDOS": "Votos válidos",
    "PCT_PT_VALIDOS": "% PT (válidos)",
    "PCT_OPP_VALIDOS": "% opp. (válidos)",
    "DIF_PCT_VALIDOS": "Diferença p.p. (válidos)",
    "QT_VOTOS_BRANCO": "Brancos",
    "QT_VOTOS_NULO": "Nulos",
    "QT_VOTOS_TOTAIS": "Votos totais",
    "PCT_PT_TOTAIS": "% PT (totais)",
    "PCT_OPP_TOTAIS": "% opp. (totais)",
    "DIF_PCT_TOTAIS": "Diferença p.p. (totais)",
    "QT_APTOS": "Eleitores aptos",
    "VENCEDOR": "Vencedor",
}

TIPOS_CAMPO = {
    "CAND_PT": "text",
    "CAND_OPP": "text",
    "QT_SECOES": "int",
    "QT_VOTOS_PT": "int",
    "QT_VOTOS_OPP": "int",
    "QT_DIF_PT_OPP": "int",
    "QT_VOTOS_VALIDOS": "int",
    "PCT_PT_VALIDOS": "pct",
    "PCT_OPP_VALIDOS": "pct",
    "DIF_PCT_VALIDOS": "pp",
    "QT_VOTOS_BRANCO": "int",
    "QT_VOTOS_NULO": "int",
    "QT_VOTOS_TOTAIS": "int",
    "PCT_PT_TOTAIS": "pct",
    "PCT_OPP_TOTAIS": "pct",
    "DIF_PCT_TOTAIS": "pp",
    "QT_APTOS": "int",
    "VENCEDOR": "text",
}

COR_ANO = {2014: "#548235", 2018: "#C00000", 2022: "#1F4E79"}


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconf = getattr(stream, "reconfigure", None)
        if reconf is None:
            continue
        try:
            reconf(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dados", type=Path, default=None)
    p.add_argument("--saida", type=Path, default=None)
    p.add_argument("--limite-aptos", type=int, default=LIMITE_APTOS)
    return p.parse_args(argv)


def com_diferencas_pares(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    tinha_aptos = "QT_APTOS" in out.columns
    for col in (
        "QT_VOTOS_PT",
        "QT_VOTOS_OPP",
        "QT_VOTOS_VALIDOS",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_COMPARECIMENTO",
        "QT_APTOS",
        "QT_SECOES",
    ):
        if col not in out.columns:
            out[col] = pd.NA if col == "QT_APTOS" else 0
        out[col] = pd.to_numeric(out[col], errors="coerce")
    if not tinha_aptos:
        out["QT_APTOS"] = pd.NA
    fallback = (
        out["QT_VOTOS_VALIDOS"].fillna(0)
        + out["QT_VOTOS_BRANCO"].fillna(0)
        + out["QT_VOTOS_NULO"].fillna(0)
    )
    if "QT_COMPARECIMENTO" in out.columns:
        totais = pd.to_numeric(out["QT_COMPARECIMENTO"], errors="coerce")
        totais = totais.where(totais.fillna(0) > 0, fallback)
    else:
        totais = fallback
    out["QT_VOTOS_TOTAIS"] = totais
    out["QT_DIF_PT_OPP"] = out["QT_VOTOS_PT"].fillna(0) - out["QT_VOTOS_OPP"].fillna(0)
    out["PCT_PT_VALIDOS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_PT"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_OPP_VALIDOS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_OPP"], out["QT_VOTOS_VALIDOS"])
    ]
    out["DIF_PCT_VALIDOS"] = [
        pct(d, v) for d, v in zip(out["QT_DIF_PT_OPP"], out["QT_VOTOS_VALIDOS"])
    ]
    out["PCT_PT_TOTAIS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_PT"], out["QT_VOTOS_TOTAIS"])
    ]
    out["PCT_OPP_TOTAIS"] = [
        pct(a, b) for a, b in zip(out["QT_VOTOS_OPP"], out["QT_VOTOS_TOTAIS"])
    ]
    out["DIF_PCT_TOTAIS"] = [
        pct(d, t) for d, t in zip(out["QT_DIF_PT_OPP"], out["QT_VOTOS_TOTAIS"])
    ]
    nomes_pt = out["CAND_PT"] if "CAND_PT" in out.columns else "PT"
    nomes_opp = out["CAND_OPP"] if "CAND_OPP" in out.columns else "OPP"
    vencedor = []
    for pt, opp, n_pt, n_opp in zip(
        out["QT_VOTOS_PT"], out["QT_VOTOS_OPP"], nomes_pt, nomes_opp
    ):
        if pd.isna(pt) or pd.isna(opp):
            vencedor.append(None)
        elif pt > opp:
            vencedor.append(str(n_pt))
        elif opp > pt:
            vencedor.append(str(n_opp))
        else:
            vencedor.append("Empate")
    out["VENCEDOR"] = vencedor
    return out


def padronizar_municipios(urnas: pd.DataFrame, ano: int) -> pd.DataFrame:
    _pt, _opp, pt_nome, opp_nome = LADOS[ano]
    ne = recorte_nordeste(urnas)
    mun = agregar(ne, ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"])
    mun = marcar_capital_interior(mun)
    col_pt = f"QT_VOTOS_{_pt}"
    col_opp = f"QT_VOTOS_{_opp}"
    if col_pt not in mun.columns or col_opp not in mun.columns:
        raise KeyError(f"{ano} 2T: faltam {col_pt} / {col_opp}")
    mun["QT_VOTOS_PT"] = pd.to_numeric(mun[col_pt], errors="coerce")
    mun["QT_VOTOS_OPP"] = pd.to_numeric(mun[col_opp], errors="coerce")
    mun["CAND_PT"] = pt_nome
    mun["CAND_OPP"] = opp_nome
    mun = com_diferencas_pares(mun)
    mun["CD_MUNICIPIO"] = pd.to_numeric(mun["CD_MUNICIPIO"], errors="coerce")
    mun["SG_UF"] = mun["SG_UF"].astype(str).str.strip().str.upper()
    keep = ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO", "EH_CAPITAL", *CAMPOS_ANO]
    out = mun[keep].copy()
    renome = {c: f"{c}_{ano}" for c in CAMPOS_ANO}
    renome["NM_MUNICIPIO"] = f"NM_MUNICIPIO_{ano}"
    return out.rename(columns=renome)


def cruzar_anos(blocos: dict[int, pd.DataFrame]) -> pd.DataFrame:
    chaves = ["SG_UF", "CD_MUNICIPIO"]
    out: pd.DataFrame | None = None
    for ano in ANOS:
        parte = blocos[ano]
        out = parte if out is None else out.merge(
            parte.drop(columns=["EH_CAPITAL"], errors="ignore"),
            on=chaves,
            how="outer",
        )
    assert out is not None
    nomes = [f"NM_MUNICIPIO_{ano}" for ano in (2022, 2018, 2014)]
    out["NM_MUNICIPIO"] = out[nomes[0]]
    for col in nomes[1:]:
        if col in out.columns:
            out["NM_MUNICIPIO"] = out["NM_MUNICIPIO"].where(
                out["NM_MUNICIPIO"].notna() & (out["NM_MUNICIPIO"].astype(str) != ""),
                out[col],
            )
    out = out.drop(columns=[c for c in nomes if c in out.columns])
    if "EH_CAPITAL" not in out.columns:
        out = marcar_capital_interior(out)
    return out


def filtrar_interior_ate_limite(df: pd.DataFrame, limite: int) -> pd.DataFrame:
    aptos = pd.to_numeric(df["QT_APTOS_2022"], errors="coerce")
    return df[(~df["EH_CAPITAL"]) & aptos.notna() & (aptos <= limite)].copy()


def agregar_ufs(municipios: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for uf, grupo in municipios.groupby("SG_UF", dropna=False):
        rec: dict[str, object] = {
            "SG_UF": uf,
            "CD_MUNICIPIO": pd.NA,
            "NM_MUNICIPIO": (
                f"{len(grupo)} município do interior"
                if len(grupo) == 1
                else f"{len(grupo)} municípios do interior"
            ),
            "QT_APTOS_2022": int(grupo["QT_APTOS_2022"].fillna(0).sum()),
        }
        for ano in ANOS:
            rec[f"CAND_PT_{ano}"] = LADOS[ano][2]
            rec[f"CAND_OPP_{ano}"] = LADOS[ano][3]
            soma_pt = float(grupo[f"QT_VOTOS_PT_{ano}"].fillna(0).sum())
            soma_opp = float(grupo[f"QT_VOTOS_OPP_{ano}"].fillna(0).sum())
            validos = float(grupo[f"QT_VOTOS_VALIDOS_{ano}"].fillna(0).sum())
            totais = float(grupo[f"QT_VOTOS_TOTAIS_{ano}"].fillna(0).sum())
            rec[f"QT_SECOES_{ano}"] = int(grupo[f"QT_SECOES_{ano}"].fillna(0).sum())
            rec[f"QT_VOTOS_PT_{ano}"] = soma_pt
            rec[f"QT_VOTOS_OPP_{ano}"] = soma_opp
            rec[f"QT_DIF_PT_OPP_{ano}"] = soma_pt - soma_opp
            rec[f"QT_VOTOS_VALIDOS_{ano}"] = validos
            rec[f"PCT_PT_VALIDOS_{ano}"] = pct(soma_pt, validos)
            rec[f"PCT_OPP_VALIDOS_{ano}"] = pct(soma_opp, validos)
            rec[f"DIF_PCT_VALIDOS_{ano}"] = pct(soma_pt - soma_opp, validos)
            rec[f"QT_VOTOS_BRANCO_{ano}"] = float(
                grupo[f"QT_VOTOS_BRANCO_{ano}"].fillna(0).sum()
            )
            rec[f"QT_VOTOS_NULO_{ano}"] = float(
                grupo[f"QT_VOTOS_NULO_{ano}"].fillna(0).sum()
            )
            rec[f"QT_VOTOS_TOTAIS_{ano}"] = totais
            rec[f"PCT_PT_TOTAIS_{ano}"] = pct(soma_pt, totais)
            rec[f"PCT_OPP_TOTAIS_{ano}"] = pct(soma_opp, totais)
            rec[f"DIF_PCT_TOTAIS_{ano}"] = pct(soma_pt - soma_opp, totais)
            aptos_col = f"QT_APTOS_{ano}"
            rec[aptos_col] = (
                float(grupo[aptos_col].fillna(0).sum())
                if aptos_col in grupo.columns and grupo[aptos_col].fillna(0).sum() > 0
                else None
            )
            rec[f"VENCEDOR_{ano}"] = (
                LADOS[ano][2] if soma_pt > soma_opp else LADOS[ano][3] if soma_opp > soma_pt else "Empate"
            )
        linhas.append(rec)
    return pd.DataFrame(linhas).sort_values("SG_UF", kind="mergesort").reset_index(drop=True)


def montar_resumo(municipios: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for ano in ANOS:
        pt = float(municipios[f"QT_VOTOS_PT_{ano}"].fillna(0).sum())
        opp = float(municipios[f"QT_VOTOS_OPP_{ano}"].fillna(0).sum())
        validos = float(municipios[f"QT_VOTOS_VALIDOS_{ano}"].fillna(0).sum())
        totais = float(municipios[f"QT_VOTOS_TOTAIS_{ano}"].fillna(0).sum())
        aptos_col = f"QT_APTOS_{ano}"
        aptos = (
            float(municipios[aptos_col].fillna(0).sum())
            if aptos_col in municipios.columns
            else None
        )
        linhas.append(
            {
                "Pleito": f"2º turno {ano}",
                "Candidato PT": LADOS[ano][2],
                "Candidato oposição": LADOS[ano][3],
                "Municípios": len(municipios),
                "Votos PT": pt,
                "Votos oposição": opp,
                "Diferença (PT − opp.)": pt - opp,
                "Votos válidos": validos,
                "% PT (válidos)": pct(pt, validos),
                "% opp. (válidos)": pct(opp, validos),
                "Diferença p.p. (válidos)": pct(pt - opp, validos),
                "Votos totais": totais,
                "% PT (totais)": pct(pt, totais),
                "% opp. (totais)": pct(opp, totais),
                "Diferença p.p. (totais)": pct(pt - opp, totais),
                "Eleitores aptos": aptos if aptos and aptos > 0 else None,
                "Vitórias PT": int((municipios[f"VENCEDOR_{ano}"] == LADOS[ano][2]).sum()),
                "Vitórias oposição": int(
                    (municipios[f"VENCEDOR_{ano}"] == LADOS[ano][3]).sum()
                ),
            }
        )
    return pd.DataFrame(linhas)


def _fmts(wb):
    base_header = {
        "bold": True,
        "font_color": "white",
        "border": 1,
        "valign": "vcenter",
        "align": "center",
        "text_wrap": True,
    }
    out = {
        "title": wb.add_format(
            {
                "bold": True,
                "font_size": 13,
                "font_color": "white",
                "bg_color": "#1F4E79",
                "valign": "vcenter",
            }
        ),
        "sub": wb.add_format(
            {"italic": True, "font_size": 10, "font_color": "#1F4E79", "valign": "vcenter"}
        ),
        "id_header": wb.add_format({**base_header, "bg_color": "#595959"}),
        "text": wb.add_format({"border": 1, "valign": "vcenter"}),
        "int": wb.add_format({"border": 1, "num_format": "#,##0", "valign": "vcenter"}),
        "pct": wb.add_format({"border": 1, "num_format": "0.00", "valign": "vcenter"}),
        "pp": wb.add_format(
            {"border": 1, "num_format": "+0.00;-0.00;0.00", "valign": "vcenter"}
        ),
        "label": wb.add_format(
            {"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1}
        ),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
    }
    for ano, cor in COR_ANO.items():
        out[f"ano_{ano}"] = wb.add_format({**base_header, "bg_color": cor, "font_size": 11})
        out[f"campo_{ano}"] = wb.add_format({**base_header, "bg_color": cor})
    return out


def _escrever_celula(ws, r, c, val, tipo, fmts) -> None:
    if val is None or (not isinstance(val, str) and pd.isna(val)) or val == "":
        ws.write_blank(r, c, None, fmts["text"])
        return
    if tipo == "int":
        ws.write_number(r, c, int(val), fmts["int"])
    elif tipo in ("pct", "pp"):
        ws.write_number(r, c, float(val), fmts[tipo])
    else:
        ws.write_string(r, c, str(val), fmts["text"])


def escrever_municipios(ws, df: pd.DataFrame, fmts, titulo: str, sub: str) -> None:
    id_cols = [
        ("SG_UF", "UF", 6, "text"),
        ("CD_MUNICIPIO", "Código TSE", 12, "int"),
        ("NM_MUNICIPIO", "Município", 32, "text"),
        ("QT_APTOS_2022", "Eleitores 2022", 14, "int"),
    ]
    ncols = len(id_cols) + len(ANOS) * len(CAMPOS_ANO)
    ws.merge_range(0, 0, 0, ncols - 1, titulo, fmts["title"])
    ws.merge_range(1, 0, 1, ncols - 1, sub, fmts["sub"])
    ws.set_row(0, 22)
    ws.set_row(1, 18)
    ws.set_row(2, 18)
    ws.set_row(3, 30)
    col = 0
    ws.merge_range(2, 0, 2, len(id_cols) - 1, "Identificação", fmts["id_header"])
    for campo, rotulo, largura, _ in id_cols:
        ws.write(3, col, rotulo, fmts["id_header"])
        ws.set_column(col, col, largura)
        col += 1
    for ano in ANOS:
        inicio = col
        fim = col + len(CAMPOS_ANO) - 1
        ws.merge_range(2, inicio, 2, fim, f"2º turno {ano}", fmts[f"ano_{ano}"])
        for campo in CAMPOS_ANO:
            ws.write(3, col, ROTOLOS_CAMPO[campo], fmts[f"campo_{ano}"])
            ws.set_column(col, col, 14 if TIPOS_CAMPO[campo] != "text" else 14)
            col += 1
    dados = df.sort_values(["SG_UF", "NM_MUNICIPIO"], kind="mergesort").reset_index(drop=True)
    for r, rec in enumerate(dados.itertuples(index=False), 4):
        row = dados.iloc[r - 4]
        c = 0
        for campo, _, _, tipo in id_cols:
            _escrever_celula(ws, r, c, row[campo], tipo, fmts)
            c += 1
        for ano in ANOS:
            for campo in CAMPOS_ANO:
                _escrever_celula(
                    ws, r, c, row.get(f"{campo}_{ano}"), TIPOS_CAMPO[campo], fmts
                )
                c += 1
    last = 3 + len(dados)
    if len(dados):
        ws.autofilter(3, 0, last, ncols - 1)
    ws.freeze_panes(4, 4)
    ws.set_column(2, 2, 32)


def gravar_xlsx(
    destino: Path,
    municipios: pd.DataFrame,
    por_uf: pd.DataFrame,
    resumo: pd.DataFrame,
    fontes: dict[int, str],
    limite: int,
    excluidos: int,
) -> Path:
    import xlsxwriter

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)

    leia = wb.add_worksheet("Leia-me")
    leia.set_column(0, 0, 38)
    leia.set_column(1, 1, 100)
    linhas = [
        ("Pleitos", "2º turno de Presidente em 2014, 2018 e 2022"),
        ("Recorte", "Interior do Nordeste (AL BA CE MA PB PE PI RN SE, sem as capitais)"),
        (
            "Capitais excluídas",
            "Maceió, Salvador, Fortaleza, São Luís, João Pessoa, Recife, Teresina, Natal, Aracaju",
        ),
        (
            "Filtro de eleitores",
            f"Até {limite:,} eleitores aptos no 2º turno de 2022 (inclusive)".replace(",", "."),
        ),
        ("Municípios no filtro", f"{len(municipios):,}".replace(",", ".")),
        (
            "Interior acima do limite",
            f"{excluidos:,} municípios do interior com mais de {limite:,} eleitores em 2022".replace(",", "."),
        ),
        ("Pares", "2014 Dilma × Aécio · 2018 Haddad × Bolsonaro · 2022 Lula × Bolsonaro"),
        (
            "Votos válidos",
            "Soma dos dois candidatos. Diferença p.p. = (PT − oposição) / válidos × 100",
        ),
        (
            "Votos totais",
            "Comparecimento (válidos + brancos + nulos). Em 2018 o comparecimento nacional "
            "é reconstruído assim, porque o arquivo de seção não traz QT_APTOS.",
        ),
        (
            "Eleitores aptos 2018",
            "Não disponíveis no arquivo nacional de seção; a coluna fica em branco. "
            "O filtro de 40 mil usa os aptos de 2022.",
        ),
        ("Sinal da diferença", "Positivo = vantagem do lado PT; negativo = vantagem da oposição"),
        ("Fonte 2014", fontes[2014]),
        ("Fonte 2018", fontes[2018]),
        ("Fonte 2022", fontes[2022]),
        ("Abas", "Municipios, Por_UF, Resumo"),
    ]
    leia.write(0, 0, "Campo", fmts["id_header"])
    leia.write(0, 1, "Valor", fmts["id_header"])
    for i, (k, v) in enumerate(linhas, 1):
        leia.write(i, 0, k, fmts["label"])
        leia.write(i, 1, v, fmts["wrap"])
        leia.set_row(i, 20)
    leia.freeze_panes(1, 0)

    escrever_municipios(
        wb.add_worksheet("Municipios"),
        municipios,
        fmts,
        f"Interior do Nordeste com até {limite:,} eleitores — 2º turno 2014, 2018 e 2022".replace(",", "."),
        "Uma linha por município. Diferença percentual sobre votos válidos e sobre votos totais.",
    )
    escrever_municipios(
        wb.add_worksheet("Por_UF"),
        por_uf,
        fmts,
        "Interior do Nordeste (até 40 mil eleitores) por UF",
        "Soma só dos municípios do recorte.",
    )

    res = wb.add_worksheet("Resumo")
    res.set_column(0, 0, 18)
    res.set_column(1, 20, 16)
    res.merge_range(
        0,
        0,
        0,
        len(resumo.columns) - 1,
        "Totais do recorte — interior do Nordeste com até 40 mil eleitores (2022)",
        fmts["title"],
    )
    res.set_row(0, 22)
    res.set_row(2, 28)
    for c, col in enumerate(resumo.columns):
        res.write(2, c, col, fmts["id_header"])
    tipos_int = {
        "Municípios",
        "Votos PT",
        "Votos oposição",
        "Diferença (PT − opp.)",
        "Votos válidos",
        "Votos totais",
        "Eleitores aptos",
        "Vitórias PT",
        "Vitórias oposição",
    }
    tipos_pp = {"Diferença p.p. (válidos)", "Diferença p.p. (totais)"}
    tipos_pct = {
        "% PT (válidos)",
        "% opp. (válidos)",
        "% PT (totais)",
        "% opp. (totais)",
    }
    for r, rec in enumerate(resumo.itertuples(index=False), 3):
        row = resumo.iloc[r - 3]
        for c, col in enumerate(resumo.columns):
            val = row[col]
            if col in tipos_int:
                _escrever_celula(res, r, c, val, "int", fmts)
            elif col in tipos_pp:
                _escrever_celula(res, r, c, val, "pp", fmts)
            elif col in tipos_pct:
                _escrever_celula(res, r, c, val, "pct", fmts)
            else:
                _escrever_celula(res, r, c, val, "text", fmts)
    res.freeze_panes(3, 1)
    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    limite = int(args.limite_aptos)
    blocos: dict[int, pd.DataFrame] = {}
    fontes: dict[int, str] = {}
    for ano in ANOS:
        urnas, fonte = carregar_pleito(dados, ano, 2)
        blocos[ano] = padronizar_municipios(urnas, ano)
        fontes[ano] = fonte
        print(f"  {ano} 2T: {len(blocos[ano])} municípios NE ← {fonte}", flush=True)
    cruzado = cruzar_anos(blocos)
    interior = cruzado[~cruzado["EH_CAPITAL"]].copy()
    recorte = filtrar_interior_ate_limite(cruzado, limite)
    excluidos = int(len(interior) - len(recorte))
    if recorte.empty:
        raise RuntimeError("Nenhum município no recorte.")
    if recorte["QT_APTOS_2022"].max() > limite:
        raise RuntimeError("Filtro de eleitores falhou.")
    por_uf = agregar_ufs(recorte)
    resumo = montar_resumo(recorte)
    destino = saida / "discriminativo_interior_nordeste_ate_40mil.xlsx"
    gravar_xlsx(destino, recorte, por_uf, resumo, fontes, limite, excluidos)
    print(f"Municípios no filtro: {len(recorte)}")
    print(f"Interior acima de {limite}: {excluidos}")
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
