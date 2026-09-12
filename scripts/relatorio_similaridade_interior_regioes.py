#!/usr/bin/env python3
"""Similaridade 2014 × 2018 × 2022 no interior das demais regiões.

Mesma metodologia do relatório do Nordeste: 2º turno de Presidente,
lado PT (Dilma/Haddad/Lula) × oposição (Aécio/Bolsonaro), municípios
do interior (capitais estaduais excluídas). Cobre Norte, Centro-Oeste,
Sudeste e Sul; o Nordeste entra só no comparativo.

Saída:
  output/tse_planilhas/relatorio_similaridade_interior_regioes.xlsx

Uso:
  python3 scripts/relatorio_similaridade_interior_regioes.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from scripts.discriminativo_interior_nordeste import eh_capital
    from scripts.discriminativo_interior_nordeste_ate_40mil import (
        ANOS,
        CAMPOS_ANO,
        LADOS,
        com_diferencas_pares,
        cruzar_anos,
    )
    from scripts.planilha_resultados_presidente import (
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )
    from scripts.relatorio_similaridade_interior_nordeste import (
        PARES,
        icc_consistency,
        lado_pt,
        resumo_anos,
        stats_par,
        tabela_porte,
        tabela_sequencias,
        tabela_uf,
    )
except ImportError:  # ContAgil
    from discriminativo_interior_nordeste import eh_capital  # type: ignore
    from discriminativo_interior_nordeste_ate_40mil import (  # type: ignore
        ANOS,
        CAMPOS_ANO,
        LADOS,
        com_diferencas_pares,
        cruzar_anos,
    )
    from planilha_resultados_presidente import (  # type: ignore
        agregar,
        carregar_pleito,
        pasta_dados,
        pastas_saida,
        rotulo_regiao,
    )
    from relatorio_similaridade_interior_nordeste import (  # type: ignore
        PARES,
        icc_consistency,
        lado_pt,
        resumo_anos,
        stats_par,
        tabela_porte,
        tabela_sequencias,
        tabela_uf,
    )

REGIOES_FOCO = ("Norte", "Centro-Oeste", "Sudeste", "Sul")
REGIOES_TODAS = ("Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul")


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
    return p.parse_args(argv)


def padronizar_municipios_pais(urnas: pd.DataFrame, ano: int) -> pd.DataFrame:
    _pt, _opp, pt_nome, opp_nome = LADOS[ano]
    base = urnas.copy()
    base["SG_UF"] = base["SG_UF"].astype(str).str.strip().str.upper()
    mun = agregar(base, ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO"])
    mun["EH_CAPITAL"] = [
        eh_capital(uf, nome) for uf, nome in zip(mun["SG_UF"], mun["NM_MUNICIPIO"])
    ]
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
    keep = ["SG_UF", "CD_MUNICIPIO", "NM_MUNICIPIO", "EH_CAPITAL", *CAMPOS_ANO]
    out = mun[keep].copy()
    renome = {c: f"{c}_{ano}" for c in CAMPOS_ANO}
    renome["NM_MUNICIPIO"] = f"NM_MUNICIPIO_{ano}"
    return out.rename(columns=renome)


def enriquecer(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["REGIAO"] = out["SG_UF"].map(rotulo_regiao)
    for ano in ANOS:
        out[f"LADO_{ano}"] = lado_pt(out[f"VENCEDOR_{ano}"], ano)
    cols = [f"PCT_PT_VALIDOS_{ano}" for ano in ANOS]
    out["AMPLITUDE_PT"] = out[cols].max(axis=1) - out[cols].min(axis=1)
    out["MEDIA_PT"] = out[cols].mean(axis=1)
    out["DESVIO_MEDIO_PT"] = out[cols].sub(out["MEDIA_PT"], axis=0).abs().mean(axis=1)
    out["SEQUENCIA"] = out["LADO_2014"] + "–" + out["LADO_2018"] + "–" + out["LADO_2022"]
    out["MESMO_VENCEDOR_3"] = (out["LADO_2014"] == out["LADO_2018"]) & (
        out["LADO_2018"] == out["LADO_2022"]
    )
    return out


def carregar_interior_brasil(dados: Path) -> tuple[pd.DataFrame, dict[int, str]]:
    blocos: dict[int, pd.DataFrame] = {}
    fontes: dict[int, str] = {}
    for ano in ANOS:
        urnas, fonte = carregar_pleito(dados, ano, 2)
        blocos[ano] = padronizar_municipios_pais(urnas, ano)
        fontes[ano] = fonte
    cruz = cruzar_anos(blocos)
    cruz["EH_CAPITAL"] = [
        eh_capital(uf, nome) for uf, nome in zip(cruz["SG_UF"], cruz["NM_MUNICIPIO"])
    ]
    interior = cruz[~cruz["EH_CAPITAL"]].copy()
    interior = enriquecer(interior)
    interior = interior[interior["REGIAO"].isin(REGIOES_TODAS)].copy()
    return interior, fontes


def linha_comparativo(nome: str, g: pd.DataFrame) -> dict:
    icc = icc_consistency(g[[f"PCT_PT_VALIDOS_{a}" for a in ANOS]].to_numpy())
    rec: dict[str, object] = {
        "Região": nome,
        "Municípios": len(g),
        "ICC(3,1)": round(icc, 4),
        "Mesmo vencedor nas 3 (%)": round(100.0 * float(g["MESMO_VENCEDOR_3"].mean()), 2),
        "Amplitude média (p.p.)": round(float(g["AMPLITUDE_PT"].mean()), 2),
        "Amplitude mediana (p.p.)": round(float(g["AMPLITUDE_PT"].median()), 2),
        "Oscilação ≤5 p.p. (%)": round(100.0 * float((g["AMPLITUDE_PT"] <= 5).mean()), 2),
        "Oscilação ≤10 p.p. (%)": round(100.0 * float((g["AMPLITUDE_PT"] <= 10).mean()), 2),
    }
    for ano in ANOS:
        rec[f"% PT médio {ano}"] = round(float(g[f"PCT_PT_VALIDOS_{ano}"].mean()), 2)
        rec[f"% PT ponderado {ano}"] = round(
            100.0
            * float(g[f"QT_VOTOS_PT_{ano}"].sum())
            / float(g[f"QT_VOTOS_VALIDOS_{ano}"].sum()),
            2,
        )
        rec[f"Vitórias PT {ano}"] = int((g[f"LADO_{ano}"] == "PT").sum())
        rec[f"Vitórias opp. {ano}"] = int((g[f"LADO_{ano}"] == "OPP").sum())
    for a, b in PARES:
        s = stats_par(g, a, b)
        rec[f"Pearson {a}×{b}"] = s["Pearson (r)"]
        rec[f"Pearson pond. {a}×{b}"] = s["Pearson ponderado"]
        rec[f"MAE {a}–{b}"] = s["MAE (p.p.)"]
        rec[f"Persistência {a}–{b} (%)"] = s["Persistência do vencedor (%)"]
        rec[f"Inversões {a}–{b}"] = s["Inversões"]
    return rec


def tabela_comparativo(interior: pd.DataFrame) -> pd.DataFrame:
    linhas = [linha_comparativo(reg, interior[interior["REGIAO"] == reg]) for reg in REGIOES_TODAS]
    linhas.append(linha_comparativo("Brasil (interior, 5 regiões)", interior))
    return pd.DataFrame(linhas)


def pares_por_regiao(interior: pd.DataFrame) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_TODAS:
        g = interior[interior["REGIAO"] == reg]
        parte = pd.DataFrame([stats_par(g, a, b) for a, b in PARES])
        parte.insert(0, "Região", reg)
        blocos.append(parte)
    return pd.concat(blocos, ignore_index=True)


def resumos_por_regiao(interior: pd.DataFrame) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_TODAS:
        parte = resumo_anos(interior[interior["REGIAO"] == reg])
        parte.insert(0, "Região", reg)
        blocos.append(parte)
    return pd.concat(blocos, ignore_index=True)


def sequencias_por_regiao(interior: pd.DataFrame) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_TODAS:
        parte = tabela_sequencias(interior[interior["REGIAO"] == reg])
        parte.insert(0, "Região", reg)
        blocos.append(parte)
    return pd.concat(blocos, ignore_index=True)


def ufs_com_regiao(interior: pd.DataFrame) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_TODAS:
        parte = tabela_uf(interior[interior["REGIAO"] == reg])
        parte.insert(0, "Região", reg)
        blocos.append(parte)
    return pd.concat(blocos, ignore_index=True)


def porte_por_regiao(interior: pd.DataFrame) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_FOCO:
        parte = tabela_porte(interior[interior["REGIAO"] == reg])
        parte.insert(0, "Região", reg)
        blocos.append(parte)
    return pd.concat(blocos, ignore_index=True)


def municipios_detalhe(df: pd.DataFrame) -> pd.DataFrame:
    out = df[
        [
            "REGIAO",
            "SG_UF",
            "CD_MUNICIPIO",
            "NM_MUNICIPIO",
            "QT_APTOS_2022",
            "PCT_PT_VALIDOS_2014",
            "PCT_PT_VALIDOS_2018",
            "PCT_PT_VALIDOS_2022",
            "VENCEDOR_2014",
            "VENCEDOR_2018",
            "VENCEDOR_2022",
            "SEQUENCIA",
            "AMPLITUDE_PT",
            "DESVIO_MEDIO_PT",
            "MESMO_VENCEDOR_3",
        ]
    ].copy()
    out["DIF_2014_2018"] = (out["PCT_PT_VALIDOS_2018"] - out["PCT_PT_VALIDOS_2014"]).round(2)
    out["DIF_2018_2022"] = (out["PCT_PT_VALIDOS_2022"] - out["PCT_PT_VALIDOS_2018"]).round(2)
    out["DIF_2014_2022"] = (out["PCT_PT_VALIDOS_2022"] - out["PCT_PT_VALIDOS_2014"]).round(2)
    return out.sort_values(["REGIAO", "AMPLITUDE_PT", "SG_UF", "NM_MUNICIPIO"]).reset_index(
        drop=True
    )


def extremos_por_regiao(df: pd.DataFrame, n: int = 10) -> pd.DataFrame:
    blocos = []
    for reg in REGIOES_FOCO:
        g = df[df["REGIAO"] == reg]
        est = g.nsmallest(n, "AMPLITUDE_PT").copy()
        est["Tipo"] = "Mais estável"
        osc = g.nlargest(n, "AMPLITUDE_PT").copy()
        osc["Tipo"] = "Maior oscilação"
        blocos.append(pd.concat([est, osc], ignore_index=True))
    return pd.concat(blocos, ignore_index=True)


def texto_relatorio(comp: pd.DataFrame, fontes: dict[int, str]) -> list[tuple[str, str]]:
    def cel(reg: str, col: str):
        return comp.loc[comp["Região"] == reg, col].iloc[0]

    foco = ", ".join(REGIOES_FOCO)
    partes = [
        (
            "Objeto",
            "Mesma análise do interior do Nordeste, agora para o interior das demais regiões: "
            f"{foco}. 2º turno de 2014, 2018 e 2022; lado PT (Dilma/Haddad/Lula) × oposição "
            "(Aécio/Bolsonaro). Interior = todos os municípios da região, excluída a capital "
            "estadual (e Brasília, no Centro-Oeste). O Nordeste entra no comparativo só como "
            "referência do relatório anterior.",
        ),
        (
            "Leitura geral",
            "O interior do Nordeste continua o recorte mais estável e mais petista. "
            "Nas outras regiões a % PT municipal também é persistente — sobretudo de 2018 "
            "para 2022 — mas o nível e o vencedor mudam de região para região: Sul e "
            "Centro-Oeste são estruturalmente de oposição; o Sudeste é misto (Minas puxa "
            "para o PT, o interior de SP e RJ para Bolsonaro); o Norte fica no meio, "
            "mais petista que o Centro-Sul e menos estável que o Nordeste.",
        ),
    ]
    for reg in REGIOES_FOCO:
        partes.append(
            (
                reg,
                (
                    f"{int(cel(reg, 'Municípios'))} municípios do interior. "
                    f"ICC = {cel(reg, 'ICC(3,1)')}; "
                    f"mesmo vencedor nas três eleições em {cel(reg, 'Mesmo vencedor nas 3 (%)')}%. "
                    f"Pearson 2018×2022 = {cel(reg, 'Pearson 2018×2022')} "
                    f"(MAE {cel(reg, 'MAE 2018–2022')} p.p.); "
                    f"2014×2022 = {cel(reg, 'Pearson 2014×2022')}. "
                    f"% PT médio 2014/2018/2022 = "
                    f"{cel(reg, '% PT médio 2014')} / {cel(reg, '% PT médio 2018')} / "
                    f"{cel(reg, '% PT médio 2022')}. "
                    f"Vitórias PT em 2022: {int(cel(reg, 'Vitórias PT 2022'))} "
                    f"× oposição {int(cel(reg, 'Vitórias opp. 2022'))}."
                ),
            )
        )
    partes.extend(
        [
            (
                "Como ler",
                "Pearson: a cidade mais petista num ano continua mais petista no outro. "
                "MAE: erro típico da % PT em pontos percentuais. Persistência: mesmo lado "
                "vencedor. ICC: consistência das três séries no mesmo município. "
                "Kappa sofre com prevalência (quase todos PT ou quase todos oposição) — "
                "prefira persistência bruta e correlação da %.",
            ),
            ("Fonte 2014", fontes[2014]),
            ("Fonte 2018", fontes[2018]),
            ("Fonte 2022", fontes[2022]),
        ]
    )
    return partes


def _fmts(wb):
    header = {
        "bold": True,
        "bg_color": "#1F4E79",
        "font_color": "white",
        "border": 1,
        "valign": "vcenter",
        "align": "center",
        "text_wrap": True,
    }
    return {
        "title": wb.add_format(
            {"bold": True, "font_size": 13, "font_color": "white", "bg_color": "#1F4E79"}
        ),
        "header": wb.add_format(header),
        "label": wb.add_format(
            {"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1, "valign": "top"}
        ),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "text": wb.add_format({"border": 1}),
        "int": wb.add_format({"border": 1, "num_format": "#,##0"}),
        "num": wb.add_format({"border": 1, "num_format": "0.00"}),
    }


def _escrever_df(ws, df: pd.DataFrame, fmts, titulo: str, linha0: int = 0) -> int:
    ncol = max(len(df.columns) - 1, 0)
    ws.merge_range(linha0, 0, linha0, ncol, titulo, fmts["title"])
    for c, col in enumerate(df.columns):
        ws.write(linha0 + 2, c, col, fmts["header"])
        ws.set_column(c, c, min(26, max(11, len(str(col)) + 2)))
    for r, _ in enumerate(df.itertuples(index=False), linha0 + 3):
        row = df.iloc[r - (linha0 + 3)]
        for c, col in enumerate(df.columns):
            val = row[col]
            if val is None or (not isinstance(val, str) and pd.isna(val)):
                ws.write_blank(r, c, None, fmts["text"])
            elif isinstance(val, bool):
                ws.write(r, c, "S" if val else "N", fmts["text"])
            elif isinstance(val, (int,)) and not isinstance(val, bool):
                ws.write_number(r, c, int(val), fmts["int"])
            elif isinstance(val, float):
                ws.write_number(r, c, float(val), fmts["num"])
            else:
                ws.write(r, c, str(val), fmts["text"])
    last = linha0 + 2 + len(df)
    if len(df):
        ws.autofilter(linha0 + 2, 0, last, ncol)
    ws.freeze_panes(linha0 + 3, 1)
    return last


def gravar_xlsx(destino: Path, interior: pd.DataFrame, fontes: dict[int, str]) -> Path:
    import xlsxwriter

    dest_foco = interior[interior["REGIAO"].isin(REGIOES_FOCO)].copy()
    comp = tabela_comparativo(interior)
    pares = pares_por_regiao(interior)
    resumos = resumos_por_regiao(interior)
    seq = sequencias_por_regiao(interior)
    ufs = ufs_com_regiao(interior)
    porte = porte_por_regiao(interior)
    detalhe = municipios_detalhe(dest_foco)
    extremos = extremos_por_regiao(dest_foco)
    narrativa = texto_relatorio(comp, fontes)

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)

    rel = wb.add_worksheet("Relatorio")
    rel.set_column(0, 0, 22)
    rel.set_column(1, 1, 118)
    rel.merge_range(
        0,
        0,
        0,
        1,
        "Similaridade no interior das regiões — 2º turno 2014, 2018 e 2022",
        fmts["title"],
    )
    rel.write(2, 0, "Campo", fmts["header"])
    rel.write(2, 1, "Texto", fmts["header"])
    for i, (k, v) in enumerate(narrativa, 3):
        rel.write(i, 0, k, fmts["label"])
        rel.write(i, 1, v, fmts["wrap"])
        rel.set_row(i, 52 if k in {"Objeto", "Leitura geral", "Como ler"} else 36)

    _escrever_df(wb.add_worksheet("Comparativo"), comp, fmts, "Interior por região — mesmo recorte do Nordeste")
    _escrever_df(wb.add_worksheet("Resumo_anos"), resumos, fmts, "Nível da votação petista no interior, por região")
    _escrever_df(wb.add_worksheet("Pares"), pares, fmts, "Similaridade par a par da % PT, por região")
    _escrever_df(wb.add_worksheet("Sequencias"), seq, fmts, "Trajetória do vencedor municipal, por região")
    _escrever_df(wb.add_worksheet("Por_UF"), ufs, fmts, "Similaridade por UF (interior, capitais fora)")
    _escrever_df(wb.add_worksheet("Por_porte"), porte, fmts, "Similaridade por porte eleitoral (demais regiões)")
    _escrever_df(
        wb.add_worksheet("Municipios"),
        detalhe,
        fmts,
        "Norte, Centro-Oeste, Sudeste e Sul — interior, ordenado pela amplitude da % PT",
    )
    _escrever_df(
        wb.add_worksheet("Extremos"),
        extremos[
            [
                "Tipo",
                "REGIAO",
                "SG_UF",
                "NM_MUNICIPIO",
                "PCT_PT_VALIDOS_2014",
                "PCT_PT_VALIDOS_2018",
                "PCT_PT_VALIDOS_2022",
                "AMPLITUDE_PT",
                "SEQUENCIA",
                "QT_APTOS_2022",
            ]
        ],
        fmts,
        "10 mais estáveis e 10 com maior oscilação em cada região (exceto Nordeste)",
    )

    dados = wb.add_worksheet("Dados_grafico")
    dados.write_row(0, 0, ["Região", "Pearson 2018×2022", "Persistência 3 anos (%)", "% PT médio 2022"])
    for i, reg in enumerate(REGIOES_TODAS, 1):
        row = comp[comp["Região"] == reg].iloc[0]
        dados.write(i, 0, reg)
        dados.write_number(i, 1, float(row["Pearson 2018×2022"]))
        dados.write_number(i, 2, float(row["Mesmo vencedor nas 3 (%)"]))
        dados.write_number(i, 3, float(row["% PT médio 2022"]))
    graf = wb.add_worksheet("Graficos")
    graf.merge_range(0, 0, 0, 10, "Comparação entre regiões (interior)", fmts["title"])
    for col, titulo, ymax in (
        (1, "Pearson da % PT — 2018 × 2022", 1),
        (2, "Mesmo vencedor nas 3 eleições (%)", 100),
        (3, "% PT médio municipal em 2022", 100),
    ):
        chart = wb.add_chart({"type": "column"})
        chart.add_series(
            {
                "name": titulo,
                "categories": ["Dados_grafico", 1, 0, 5, 0],
                "values": ["Dados_grafico", 1, col, 5, col],
            }
        )
        chart.set_title({"name": titulo})
        chart.set_y_axis({"min": 0, "max": ymax})
        chart.set_legend({"none": True})
        chart.set_size({"width": 480, "height": 320})
        graf.insert_chart(2, (col - 1) * 8, chart)

    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    interior, fontes = carregar_interior_brasil(dados)
    esperado = {"Norte": 443, "Nordeste": 1785, "Centro-Oeste": 463, "Sudeste": 1664, "Sul": 1188}
    cont = interior.groupby("REGIAO").size().to_dict()
    print("Municípios por região:", cont, flush=True)
    for reg, n in esperado.items():
        if cont.get(reg) != n:
            print(f"aviso: {reg} esperado {n}, obtido {cont.get(reg)}", flush=True)
    destino = saida / "relatorio_similaridade_interior_regioes.xlsx"
    gravar_xlsx(destino, interior, fontes)
    print(tabela_comparativo(interior)[["Região", "Municípios", "ICC(3,1)", "Pearson 2018×2022", "Mesmo vencedor nas 3 (%)"]].to_string(index=False))
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
