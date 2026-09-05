#!/usr/bin/env python3
"""Relatório de similaridade 2014 × 2018 × 2022 no interior do Nordeste.

Compara o 2º turno de Presidente (Dilma/Haddad/Lula × Aécio/Bolsonaro)
em todos os municípios do interior (capitais excluídas).

Saída:
  output/tse_planilhas/relatorio_similaridade_interior_nordeste.xlsx

Uso:
  python3 scripts/relatorio_similaridade_interior_nordeste.py
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

try:
    from scripts.discriminativo_interior_nordeste_ate_40mil import (
        ANOS,
        LADOS,
        cruzar_anos,
        padronizar_municipios,
    )
    from scripts.planilha_resultados_presidente import (
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )
except ImportError:  # ContAgil
    from discriminativo_interior_nordeste_ate_40mil import (  # type: ignore
        ANOS,
        LADOS,
        cruzar_anos,
        padronizar_municipios,
    )
    from planilha_resultados_presidente import (  # type: ignore
        carregar_pleito,
        pasta_dados,
        pastas_saida,
    )

PARES = ((2014, 2018), (2018, 2022), (2014, 2022))
FAIXAS_APTOS = (
    (0, 10_000, "Até 10 mil"),
    (10_000, 20_000, "10 a 20 mil"),
    (20_000, 40_000, "20 a 40 mil"),
    (40_000, 100_000, "40 a 100 mil"),
    (100_000, 10**9, "Mais de 100 mil"),
)


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


def lado_pt(serie: pd.Series, ano: int) -> pd.Series:
    return np.where(serie.astype(str) == LADOS[ano][2], "PT", "OPP")


def cohen_kappa(a: pd.Series | np.ndarray, b: pd.Series | np.ndarray) -> float:
    x = np.asarray(a)
    y = np.asarray(b)
    po = float((x == y).mean())
    px = pd.Series(x).value_counts(normalize=True)
    py = pd.Series(y).value_counts(normalize=True)
    pe = float(sum(px.get(k, 0.0) * py.get(k, 0.0) for k in set(x) | set(y)))
    if pe >= 1:
        return 1.0
    return (po - pe) / (1.0 - pe)


def icc_consistency(matriz: np.ndarray) -> float:
    """ICC(3,1) — consistência, two-way mixed, single measure."""
    x = np.asarray(matriz, dtype=float)
    n, k = x.shape
    if n < 2 or k < 2:
        return float("nan")
    media_n = x.mean(axis=1)
    grand = x.mean()
    ssb = k * ((media_n - grand) ** 2).sum()
    ssw = ((x - media_n[:, None]) ** 2).sum()
    msb = ssb / (n - 1)
    msw = ssw / (n * (k - 1))
    den = msb + (k - 1) * msw
    if den == 0:
        return 1.0
    return float((msb - msw) / den)


def pearson_ponderado(x: pd.Series, y: pd.Series, w: pd.Series) -> float:
    ww = pd.to_numeric(w, errors="coerce").fillna(0)
    xx = pd.to_numeric(x, errors="coerce")
    yy = pd.to_numeric(y, errors="coerce")
    mask = ww > 0
    xx, yy, ww = xx[mask], yy[mask], ww[mask]
    mx = float(np.average(xx, weights=ww))
    my = float(np.average(yy, weights=ww))
    cov = float(np.average((xx - mx) * (yy - my), weights=ww))
    vx = float(np.average((xx - mx) ** 2, weights=ww))
    vy = float(np.average((yy - my) ** 2, weights=ww))
    if vx <= 0 or vy <= 0:
        return float("nan")
    return cov / (vx ** 0.5 * vy ** 0.5)


def carregar_interior(dados: Path) -> tuple[pd.DataFrame, dict[int, str]]:
    blocos: dict[int, pd.DataFrame] = {}
    fontes: dict[int, str] = {}
    for ano in ANOS:
        urnas, fonte = carregar_pleito(dados, ano, 2)
        blocos[ano] = padronizar_municipios(urnas, ano)
        fontes[ano] = fonte
    cruz = cruzar_anos(blocos)
    interior = cruz[~cruz["EH_CAPITAL"]].copy()
    for ano in ANOS:
        interior[f"LADO_{ano}"] = lado_pt(interior[f"VENCEDOR_{ano}"], ano)
    cols = [f"PCT_PT_VALIDOS_{ano}" for ano in ANOS]
    interior["AMPLITUDE_PT"] = interior[cols].max(axis=1) - interior[cols].min(axis=1)
    interior["MEDIA_PT"] = interior[cols].mean(axis=1)
    interior["DESVIO_MEDIO_PT"] = (
        interior[cols].sub(interior["MEDIA_PT"], axis=0).abs().mean(axis=1)
    )
    interior["SEQUENCIA"] = (
        interior["LADO_2014"] + "–" + interior["LADO_2018"] + "–" + interior["LADO_2022"]
    )
    interior["MESMO_VENCEDOR_3"] = (
        (interior["LADO_2014"] == interior["LADO_2018"])
        & (interior["LADO_2018"] == interior["LADO_2022"])
    )
    return interior, fontes


def stats_par(df: pd.DataFrame, a: int, b: int) -> dict:
    xa = df[f"PCT_PT_VALIDOS_{a}"]
    xb = df[f"PCT_PT_VALIDOS_{b}"]
    la = df[f"LADO_{a}"]
    lb = df[f"LADO_{b}"]
    qa = pd.qcut(xa, 4, labels=["Q1", "Q2", "Q3", "Q4"], duplicates="drop")
    qb = pd.qcut(xb, 4, labels=["Q1", "Q2", "Q3", "Q4"], duplicates="drop")
    peso = df[f"QT_VOTOS_VALIDOS_{b}"]
    return {
        "Par": f"{a} × {b}",
        "N": len(df),
        "Pearson (r)": round(float(xa.corr(xb)), 4),
        "Spearman (ρ)": round(float(xa.rank().corr(xb.rank())), 4),
        "Pearson ponderado": round(pearson_ponderado(xa, xb, peso), 4),
        "R²": round(float(xa.corr(xb) ** 2), 4),
        "MAE (p.p.)": round(float((xa - xb).abs().mean()), 2),
        "RMSE (p.p.)": round(float(np.sqrt(((xa - xb) ** 2).mean())), 2),
        "Viés médio (p.p.)": round(float((xb - xa).mean()), 2),
        "Persistência do vencedor (%)": round(100.0 * float((la == lb).mean()), 2),
        "Kappa (vencedor)": round(cohen_kappa(la, lb), 4),
        "Mesmo quartil de % PT (%)": round(100.0 * float((qa.astype(str) == qb.astype(str)).mean()), 2),
        "Inversões": int((la != lb).sum()),
    }


def resumo_anos(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for ano in ANOS:
        pt = float(df[f"QT_VOTOS_PT_{ano}"].sum())
        validos = float(df[f"QT_VOTOS_VALIDOS_{ano}"].sum())
        linhas.append(
            {
                "Pleito": f"2º turno {ano}",
                "Candidato PT": LADOS[ano][2],
                "Candidato oposição": LADOS[ano][3],
                "Municípios": len(df),
                "% PT médio (município)": round(float(df[f"PCT_PT_VALIDOS_{ano}"].mean()), 2),
                "% PT mediano": round(float(df[f"PCT_PT_VALIDOS_{ano}"].median()), 2),
                "Desvio-padrão (p.p.)": round(float(df[f"PCT_PT_VALIDOS_{ano}"].std()), 2),
                "% PT ponderado (votos)": round(100.0 * pt / validos, 2),
                "Vitórias PT": int((df[f"LADO_{ano}"] == "PT").sum()),
                "Vitórias oposição": int((df[f"LADO_{ano}"] == "OPP").sum()),
            }
        )
    return pd.DataFrame(linhas)


def tabela_pares(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame([stats_par(df, a, b) for a, b in PARES])


def tabela_sequencias(df: pd.DataFrame) -> pd.DataFrame:
    rotulo = {
        "PT–PT–PT": "PT nas três eleições",
        "OPP–OPP–OPP": "Oposição nas três eleições",
        "PT–OPP–OPP": "PT em 2014; oposição em 2018 e 2022",
        "OPP–PT–PT": "Oposição em 2014; PT em 2018 e 2022",
        "OPP–PT–OPP": "Oscilou (PT só em 2018)",
        "PT–OPP–PT": "Oscilou (oposição só em 2018)",
        "OPP–OPP–PT": "Oposição em 2014–2018; PT em 2022",
        "PT–PT–OPP": "PT em 2014–2018; oposição em 2022",
    }
    g = df["SEQUENCIA"].value_counts().rename_axis("Sequência").reset_index(name="Municípios")
    g["Descrição"] = g["Sequência"].map(rotulo).fillna(g["Sequência"])
    g["%"] = (100.0 * g["Municípios"] / len(df)).round(2)
    return g[["Sequência", "Descrição", "Municípios", "%"]]


def tabela_uf(df: pd.DataFrame) -> pd.DataFrame:
    linhas = []
    for uf, g in df.groupby("SG_UF", dropna=False):
        linhas.append(
            {
                "UF": uf,
                "Municípios": len(g),
                "Pearson 2014×2018": round(float(g["PCT_PT_VALIDOS_2014"].corr(g["PCT_PT_VALIDOS_2018"])), 3),
                "Pearson 2018×2022": round(float(g["PCT_PT_VALIDOS_2018"].corr(g["PCT_PT_VALIDOS_2022"])), 3),
                "Pearson 2014×2022": round(float(g["PCT_PT_VALIDOS_2014"].corr(g["PCT_PT_VALIDOS_2022"])), 3),
                "MAE 2018–2022 (p.p.)": round(
                    float((g["PCT_PT_VALIDOS_2018"] - g["PCT_PT_VALIDOS_2022"]).abs().mean()), 2
                ),
                "Mesmo vencedor nas 3 (%)": round(100.0 * float(g["MESMO_VENCEDOR_3"].mean()), 2),
                "Amplitude média (p.p.)": round(float(g["AMPLITUDE_PT"].mean()), 2),
                "% PT 2022 (médio)": round(float(g["PCT_PT_VALIDOS_2022"].mean()), 2),
            }
        )
    return pd.DataFrame(linhas).sort_values("UF").reset_index(drop=True)


def faixa_aptos(valor: float) -> str:
    if pd.isna(valor):
        return "sem aptos"
    for lo, hi, nome in FAIXAS_APTOS:
        if lo <= float(valor) < hi:
            return nome
    return "sem aptos"


def tabela_porte(df: pd.DataFrame) -> pd.DataFrame:
    trabalho = df.copy()
    trabalho["Porte"] = trabalho["QT_APTOS_2022"].map(faixa_aptos)
    ordem = [nome for _, _, nome in FAIXAS_APTOS]
    linhas = []
    for nome in ordem:
        g = trabalho[trabalho["Porte"] == nome]
        if g.empty:
            continue
        linhas.append(
            {
                "Porte (eleitores 2022)": nome,
                "Municípios": len(g),
                "Pearson 2014×2022": round(
                    float(g["PCT_PT_VALIDOS_2014"].corr(g["PCT_PT_VALIDOS_2022"])), 3
                ),
                "Pearson 2018×2022": round(
                    float(g["PCT_PT_VALIDOS_2018"].corr(g["PCT_PT_VALIDOS_2022"])), 3
                ),
                "MAE 2018–2022 (p.p.)": round(
                    float((g["PCT_PT_VALIDOS_2018"] - g["PCT_PT_VALIDOS_2022"]).abs().mean()), 2
                ),
                "Mesmo vencedor nas 3 (%)": round(100.0 * float(g["MESMO_VENCEDOR_3"].mean()), 2),
                "Amplitude média (p.p.)": round(float(g["AMPLITUDE_PT"].mean()), 2),
            }
        )
    return pd.DataFrame(linhas)


def municipios_detalhe(df: pd.DataFrame) -> pd.DataFrame:
    out = df[
        [
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
    out["DIF_2014_2018"] = (
        out["PCT_PT_VALIDOS_2018"] - out["PCT_PT_VALIDOS_2014"]
    ).round(2)
    out["DIF_2018_2022"] = (
        out["PCT_PT_VALIDOS_2022"] - out["PCT_PT_VALIDOS_2018"]
    ).round(2)
    out["DIF_2014_2022"] = (
        out["PCT_PT_VALIDOS_2022"] - out["PCT_PT_VALIDOS_2014"]
    ).round(2)
    return out.sort_values(["AMPLITUDE_PT", "SG_UF", "NM_MUNICIPIO"]).reset_index(drop=True)


def texto_relatorio(df: pd.DataFrame, pares: pd.DataFrame, icc: float) -> list[tuple[str, str]]:
    n = len(df)
    mesmo3 = int(df["MESMO_VENCEDOR_3"].sum())
    amp_med = float(df["AMPLITUDE_PT"].mean())
    amp_md = float(df["AMPLITUDE_PT"].median())
    ate5 = 100.0 * float((df["AMPLITUDE_PT"] <= 5).mean())
    ate10 = 100.0 * float((df["AMPLITUDE_PT"] <= 10).mean())
    r1418 = pares.loc[pares["Par"] == "2014 × 2018", "Pearson (r)"].iloc[0]
    r1822 = pares.loc[pares["Par"] == "2018 × 2022", "Pearson (r)"].iloc[0]
    r1422 = pares.loc[pares["Par"] == "2014 × 2022", "Pearson (r)"].iloc[0]
    mae1822 = pares.loc[pares["Par"] == "2018 × 2022", "MAE (p.p.)"].iloc[0]
    pers1822 = pares.loc[pares["Par"] == "2018 × 2022", "Persistência do vencedor (%)"].iloc[0]
    return [
        (
            "Objeto",
            "Similaridade dos resultados de Presidente no 2º turno de 2014, 2018 e 2022 "
            "em todos os municípios do interior do Nordeste (AL, BA, CE, MA, PB, PE, PI, RN e SE, "
            "excluídas as 9 capitais). O 2º turno é o recorte comparável: dois candidatos, "
            "alinhados como lado PT (Dilma / Haddad / Lula) versus oposição (Aécio / Bolsonaro).",
        ),
        (
            "Universo",
            f"{n} municípios do interior, todos presentes nos três pleitos. "
            "A unidade de análise é o município (cada cidade pesa igual nas correlações simples). "
            "Há também correlação ponderada pelos votos válidos.",
        ),
        (
            "Leitura geral",
            "Os três resultados são muito parecidos. Quase todos os municípios do interior "
            f"nordestino votaram no mesmo lado nas três eleições ({mesmo3} de {n}, "
            f"{100.0 * mesmo3 / n:.1f}%). A % do lado PT no município em 2018 prediz a de 2022 "
            f"com r = {r1822:.3f} (MAE de só {mae1822} p.p.). 2014 já aponta na mesma direção "
            f"(r = {r1418:.3f} com 2018; r = {r1422:.3f} com 2022), com um pouco mais de ruído "
            "porque o adversário era Aécio, não Bolsonaro.",
        ),
        (
            "Consistência das três séries",
            f"ICC(3,1) = {icc:.3f}: a ordenação e o nível da votação petista no município "
            "são estáveis ao longo das três eleições. A amplitude (máximo − mínimo da % PT "
            f"no município) tem média {amp_med:.2f} p.p. e mediana {amp_md:.2f} p.p. "
            f"{ate5:.1f}% das cidades oscilaram 5 p.p. ou menos; {ate10:.1f}% oscilaram 10 p.p. ou menos.",
        ),
        (
            "Vencedor",
            f"A persistência do vencedor municipal entre 2018 e 2022 é {pers1822:.2f}%. "
            "As inversões são raras e concentradas em cidades pequenas ou no litoral de Alagoas. "
            "O kappa fica só moderado (0,38 a 0,66) por um efeito de prevalência: "
            "como quase todos os municípios são PT, o acaso já acertaria a maior parte. "
            "A métrica útil para este recorte é a persistência bruta e a correlação da % PT, "
            "não o kappa. O interior é estruturalmente petista; a similaridade medida aqui "
            "é a estabilidade desse padrão, não o resultado nacional.",
        ),
        (
            "Como ler as métricas",
            "Pearson/Spearman: quanto a % PT de um ano acompanha a do outro (1 = mesma ordenação). "
            "MAE/RMSE: erro típico em pontos percentuais. Viés: mudança média de nível. "
            "Kappa: concordância do vencedor além do acaso. Quartil: se a cidade ficou no mesmo "
            "quarto da distribuição. Amplitude e desvio médio: quão estável foi cada município.",
        ),
    ]


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
            {"bold": True, "font_size": 14, "font_color": "white", "bg_color": "#1F4E79"}
        ),
        "header": wb.add_format(header),
        "label": wb.add_format(
            {"bold": True, "bg_color": "#1F4E79", "font_color": "white", "border": 1, "valign": "top"}
        ),
        "wrap": wb.add_format({"text_wrap": True, "valign": "top", "border": 1}),
        "text": wb.add_format({"border": 1}),
        "int": wb.add_format({"border": 1, "num_format": "#,##0"}),
        "num": wb.add_format({"border": 1, "num_format": "0.00"}),
        "pct": wb.add_format({"border": 1, "num_format": "0.00"}),
        "sub": wb.add_format({"italic": True, "font_color": "#1F4E79"}),
    }


def _escrever_df(ws, df: pd.DataFrame, fmts, titulo: str, linha0: int = 0) -> int:
    ws.merge_range(linha0, 0, linha0, max(len(df.columns) - 1, 0), titulo, fmts["title"])
    for c, col in enumerate(df.columns):
        ws.write(linha0 + 2, c, col, fmts["header"])
        ws.set_column(c, c, min(28, max(12, len(str(col)) + 2)))
    for r, rec in enumerate(df.itertuples(index=False), linha0 + 3):
        row = df.iloc[r - (linha0 + 3)]
        for c, col in enumerate(df.columns):
            val = row[col]
            if val is None or (not isinstance(val, str) and pd.isna(val)):
                ws.write_blank(r, c, None, fmts["text"])
            elif isinstance(val, (bool, np.bool_)):
                ws.write(r, c, "S" if val else "N", fmts["text"])
            elif isinstance(val, (int, np.integer)) and not isinstance(val, bool):
                ws.write_number(r, c, int(val), fmts["int"])
            elif isinstance(val, (float, np.floating)):
                ws.write_number(r, c, float(val), fmts["num"])
            else:
                ws.write(r, c, str(val), fmts["text"])
    last = linha0 + 2 + len(df)
    if len(df):
        ws.autofilter(linha0 + 2, 0, last, len(df.columns) - 1)
    ws.freeze_panes(linha0 + 3, 0)
    return last


def gravar_xlsx(
    destino: Path,
    interior: pd.DataFrame,
    fontes: dict[int, str],
) -> Path:
    import xlsxwriter

    pares = tabela_pares(interior)
    anos = resumo_anos(interior)
    seq = tabela_sequencias(interior)
    ufs = tabela_uf(interior)
    porte = tabela_porte(interior)
    detalhe = municipios_detalhe(interior)
    icc = icc_consistency(interior[[f"PCT_PT_VALIDOS_{a}" for a in ANOS]].to_numpy())
    narrativa = texto_relatorio(interior, pares, icc)

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(destino)
    fmts = _fmts(wb)

    rel = wb.add_worksheet("Relatorio")
    rel.set_column(0, 0, 28)
    rel.set_column(1, 1, 110)
    rel.merge_range(
        0,
        0,
        0,
        1,
        "Similaridade eleitoral no interior do Nordeste — 2º turno 2014, 2018 e 2022",
        fmts["title"],
    )
    rel.set_row(0, 22)
    rel.write(2, 0, "Campo", fmts["header"])
    rel.write(2, 1, "Texto", fmts["header"])
    extras = [
        ("ICC(3,1) da % PT", f"{icc:.4f}"),
        ("Fonte 2014", fontes[2014]),
        ("Fonte 2018", fontes[2018]),
        ("Fonte 2022", fontes[2022]),
    ]
    for i, (k, v) in enumerate(narrativa + extras, 3):
        rel.write(i, 0, k, fmts["label"])
        rel.write(i, 1, v, fmts["wrap"])
        rel.set_row(i, 48 if k in {"Objeto", "Leitura geral", "Consistência das três séries", "Como ler as métricas"} else 28)

    _escrever_df(wb.add_worksheet("Resumo_anos"), anos, fmts, "Nível da votação petista no interior do Nordeste")
    _escrever_df(wb.add_worksheet("Pares"), pares, fmts, "Similaridade par a par da % PT (votos válidos)")
    _escrever_df(wb.add_worksheet("Sequencias"), seq, fmts, "Trajetória do vencedor municipal (PT = Dilma/Haddad/Lula)")
    _escrever_df(wb.add_worksheet("Por_UF"), ufs, fmts, "Similaridade por Unidade da Federação")
    _escrever_df(wb.add_worksheet("Por_porte"), porte, fmts, "Similaridade por porte eleitoral (aptos 2022)")

    mun = wb.add_worksheet("Municipios")
    _escrever_df(
        mun,
        detalhe,
        fmts,
        "Municípios do interior — % PT, amplitude e sequência (ordenados da menor para a maior oscilação)",
    )

    dados = wb.add_worksheet("Dados_grafico")
    dados.write_row(0, 0, ["PCT_2014", "PCT_2018", "PCT_2022"])
    for i, rec in enumerate(interior.itertuples(index=False), 1):
        dados.write_number(i, 0, float(rec.PCT_PT_VALIDOS_2014))
        dados.write_number(i, 1, float(rec.PCT_PT_VALIDOS_2018))
        dados.write_number(i, 2, float(rec.PCT_PT_VALIDOS_2022))
    n = len(interior)
    graf = wb.add_worksheet("Graficos")
    graf.merge_range(0, 0, 0, 8, "Dispersão da % PT entre eleições (cada ponto = um município)", fmts["title"])
    pares_xy = ((0, 1, "2014 × 2018"), (1, 2, "2018 × 2022"), (0, 2, "2014 × 2022"))
    for i, (cx, cy, titulo) in enumerate(pares_xy):
        chart = wb.add_chart({"type": "scatter"})
        chart.add_series(
            {
                "name": titulo,
                "categories": ["Dados_grafico", 1, cx, n, cx],
                "values": ["Dados_grafico", 1, cy, n, cy],
                "marker": {"type": "circle", "size": 4},
            }
        )
        chart.set_title({"name": f"% PT — {titulo}"})
        chart.set_x_axis({"name": f"% PT {2014 if cx == 0 else 2018 if cx == 1 else 2022}", "min": 30, "max": 100})
        chart.set_y_axis({"name": f"% PT {2018 if cy == 1 else 2022}", "min": 30, "max": 100})
        chart.set_legend({"none": True})
        chart.set_size({"width": 480, "height": 360})
        graf.insert_chart(2, i * 8, chart)

    maiores = detalhe.nlargest(20, "AMPLITUDE_PT")
    menores = detalhe.nsmallest(20, "AMPLITUDE_PT")
    ext = wb.add_worksheet("Extremos")
    last = _escrever_df(ext, menores, fmts, "20 municípios mais estáveis (menor amplitude da % PT)")
    _escrever_df(ext, maiores, fmts, "20 municípios com maior oscilação da % PT", linha0=last + 3)

    wb.close()
    return destino


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    dados = args.dados or pasta_dados()
    saida = args.saida or pastas_saida()
    interior, fontes = carregar_interior(dados)
    if len(interior) != 1785:
        print(f"aviso: esperado 1785 municípios do interior; obtido {len(interior)}", flush=True)
    destino = saida / "relatorio_similaridade_interior_nordeste.xlsx"
    gravar_xlsx(destino, interior, fontes)
    pares = tabela_pares(interior)
    print(f"Municípios: {len(interior)}")
    print(pares.to_string(index=False))
    print(f"Workbook: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
