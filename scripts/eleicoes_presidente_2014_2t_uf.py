#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Votos válidos dos dois candidatos à Presidência no 2º turno de 2014, por UF.

Dilma Rousseff e Aécio Neves, em cada uma das 27 Unidades da Federação,
a partir das tabelas de totalização do TSE (Election Resources).

Uso::

  PYTHONPATH=. python3 scripts/eleicoes_presidente_2014_2t_uf.py
"""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlencode

import pandas as pd
import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

CACHE = ROOT / "data" / "eleicoes_presidente_2014_2t_uf.json"
URL = "http://electionresources.org/br/president.php"
UA = "SEC-data-analysys/eleicoes-presidente-2014-2t"
ANO = 2014
DATA_2T = "26/10/2014"

# Totalização oficial do TSE (proclamação definitiva, 09/12/2014).
TSE_BRASIL_2T = {
    "dilma": 54_501_118,
    "aecio": 51_041_155,
    "vv": 105_542_273,
}

UFS = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
]

NOMES = {
    "AC": "Acre",
    "AL": "Alagoas",
    "AP": "Amapá",
    "AM": "Amazonas",
    "BA": "Bahia",
    "CE": "Ceará",
    "DF": "Distrito Federal",
    "ES": "Espírito Santo",
    "GO": "Goiás",
    "MA": "Maranhão",
    "MT": "Mato Grosso",
    "MS": "Mato Grosso do Sul",
    "MG": "Minas Gerais",
    "PA": "Pará",
    "PB": "Paraíba",
    "PR": "Paraná",
    "PE": "Pernambuco",
    "PI": "Piauí",
    "RJ": "Rio de Janeiro",
    "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul",
    "RO": "Rondônia",
    "RR": "Roraima",
    "SC": "Santa Catarina",
    "SP": "São Paulo",
    "SE": "Sergipe",
    "TO": "Tocantins",
    "ZZ": "Exterior",
    "BR": "Brasil",
}

REGIAO = {
    "AC": "Norte",
    "AL": "Nordeste",
    "AP": "Norte",
    "AM": "Norte",
    "BA": "Nordeste",
    "CE": "Nordeste",
    "DF": "Centro-Oeste",
    "ES": "Sudeste",
    "GO": "Centro-Oeste",
    "MA": "Nordeste",
    "MT": "Centro-Oeste",
    "MS": "Centro-Oeste",
    "MG": "Sudeste",
    "PA": "Norte",
    "PB": "Nordeste",
    "PR": "Sul",
    "PE": "Nordeste",
    "PI": "Nordeste",
    "RJ": "Sudeste",
    "RN": "Nordeste",
    "RS": "Sul",
    "RO": "Norte",
    "RR": "Norte",
    "SC": "Sul",
    "SP": "Sudeste",
    "SE": "Nordeste",
    "TO": "Norte",
    "ZZ": "Exterior",
    "BR": "Brasil",
}

META_LABELS = {
    "Registered Electors",
    "Voters",
    "Blank Votes",
    "Invalid Votes",
    "Valid Votes",
    "Candidate",
}

COLUNAS = [
    "uf",
    "unidade",
    "regiao",
    "c1_votos",
    "c1_pct",
    "c2_votos",
    "c2_pct",
    "votos_validos",
    "vencedor",
]


@dataclass(frozen=True)
class Candidato:
    chave: str
    nome: str
    votos_brasil: int
    pct_brasil: float


def parse_int(texto: str) -> int:
    return int(re.sub(r"[^0-9]", "", str(texto)))


def nome_curto(nome_completo: str) -> str:
    base = re.sub(r"\s*\(.*\)\s*$", "", nome_completo).strip()
    if base.startswith("Dilma"):
        return "Dilma"
    if "cio Neves" in base or base.startswith("Aécio") or base.startswith("Aecio"):
        return "Aécio Neves"
    return base


def chave_candidato(nome_completo: str) -> str:
    if "Dilma" in nome_completo:
        return "dilma"
    if "cio Neves" in nome_completo or "Aécio" in nome_completo or "Aecio" in nome_completo:
        return "aecio"
    return nome_curto(nome_completo).lower()


def parse_bloco(texto: str) -> dict[str, Any]:
    title = re.search(r"(?:Presidential|Runoff) Election Results - ([^<]+)", texto)
    unidade = html.unescape(title.group(1)).strip() if title else None
    rows: list[tuple[str, int]] = []
    for m in re.finditer(
        r"<TD ALIGN=LEFT NOWRAP>&nbsp;([^<]+)&nbsp;</TD>\s*<TD>&nbsp;&nbsp;([^<]+)</TD>",
        texto,
    ):
        nome = html.unescape(m.group(1)).strip()
        votos = parse_int(m.group(2))
        rows.append((nome, votos))
    meta = {n: v for n, v in rows if n in META_LABELS}
    cands = [{"nome": n, "votos": v} for n, v in rows if n not in META_LABELS]
    return {
        "unidade": unidade,
        "eleitores": meta.get("Registered Electors"),
        "comparecimento": meta.get("Voters"),
        "brancos": meta.get("Blank Votes"),
        "nulos": meta.get("Invalid Votes"),
        "validos": meta.get("Valid Votes"),
        "candidatos": cands,
    }


def parse_pagina(texto: str) -> dict[str, dict[str, Any]]:
    partes = re.split(r"Runoff Election Results", texto, maxsplit=1)
    primeiro = parse_bloco(partes[0])
    segundo = parse_bloco("Runoff Election Results" + partes[1]) if len(partes) > 1 else None
    if segundo is None:
        raise ValueError("Página sem bloco de 2º turno (Runoff).")
    return {"1t": primeiro, "2t": segundo}


def votos_de(cands: list[dict], chave: str) -> int:
    for item in cands:
        if chave_candidato(item["nome"]) == chave:
            return int(item["votos"])
    raise KeyError(chave)


def dois_mais_votados(payload: dict) -> tuple[Candidato, Candidato]:
    ordenados = sorted(payload["candidatos"], key=lambda c: int(c["votos"]), reverse=True)
    if len(ordenados) < 2:
        raise ValueError("Sem dois candidatos à Presidência.")
    validos = int(payload["validos"])
    top: list[Candidato] = []
    for item in ordenados[:2]:
        votos = int(item["votos"])
        top.append(
            Candidato(
                chave=chave_candidato(item["nome"]),
                nome=nome_curto(item["nome"]),
                votos_brasil=votos,
                pct_brasil=round(100.0 * votos / validos, 2) if validos else 0.0,
            )
        )
    return top[0], top[1]


def vencedor_por_votos(v1: int, v2: int, nome1: str, nome2: str) -> str:
    if v1 > v2:
        return nome1
    if v2 > v1:
        return nome2
    return "Empate"


def pct(votos: int, validos: int) -> float | None:
    if not validos:
        return None
    return round(100.0 * votos / validos, 2)


def baixar_html(session: requests.Session, url: str, tentativas: int = 5) -> str:
    ultimo: Exception | None = None
    for i in range(tentativas):
        try:
            resp = session.get(url, timeout=60)
            resp.raise_for_status()
            return resp.text
        except (requests.RequestException, ValueError) as exc:
            ultimo = exc
            time.sleep(min(2 ** i, 8))
    raise RuntimeError(f"Falha ao baixar {url}: {ultimo}") from ultimo


def baixar_tse(session: requests.Session | None = None) -> dict[str, dict]:
    http = session or requests.Session()
    http.headers.update({"User-Agent": UA})
    bruto: dict[str, dict] = {}
    for uf in [*UFS, "ZZ", "BR"]:
        url = f"{URL}?{urlencode({'election': ANO, 'state': uf})}"
        bruto[uf] = parse_pagina(baixar_html(http, url))
    return bruto


def conferir_totais(bruto: dict[str, dict]) -> None:
    br = bruto["BR"]["2t"]
    got = {
        "dilma": votos_de(br["candidatos"], "dilma"),
        "aecio": votos_de(br["candidatos"], "aecio"),
        "vv": int(br["validos"]),
    }
    if got != TSE_BRASIL_2T:
        raise ValueError(f"Totais nacionais 2014-2T divergem do TSE: {got} vs {TSE_BRASIL_2T}.")
    for uf in [*UFS, "ZZ"]:
        bloco = bruto[uf]["2t"]
        if len(bloco["candidatos"]) < 2:
            raise ValueError(f"{uf} sem os dois candidatos do 2º turno.")
        votos_de(bloco["candidatos"], "dilma")
        votos_de(bloco["candidatos"], "aecio")


def carregar_dados(session: requests.Session | None = None, *, forcar: bool = False) -> dict[str, dict]:
    if CACHE.exists() and not forcar:
        bruto = json.loads(CACHE.read_text(encoding="utf-8"))
        conferir_totais(bruto)
        return bruto
    bruto = baixar_tse(session)
    conferir_totais(bruto)
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(bruto, ensure_ascii=False, indent=2), encoding="utf-8")
    return bruto


def linha_uf(uf: str, bruto: dict[str, dict], cand1: Candidato, cand2: Candidato) -> dict:
    bloco = bruto[uf]["2t"]
    validos = int(bloco["validos"])
    v1 = votos_de(bloco["candidatos"], cand1.chave)
    v2 = votos_de(bloco["candidatos"], cand2.chave)
    return {
        "uf": uf,
        "unidade": NOMES.get(uf, bloco.get("unidade") or uf),
        "regiao": REGIAO.get(uf, ""),
        "c1_votos": v1,
        "c1_pct": pct(v1, validos),
        "c2_votos": v2,
        "c2_pct": pct(v2, validos),
        "votos_validos": validos,
        "vencedor": vencedor_por_votos(v1, v2, cand1.nome, cand2.nome),
    }


def totais_bloco(
    origem: pd.DataFrame,
    uf: str,
    unidade: str,
    regiao: str,
    cand1: Candidato,
    cand2: Candidato,
) -> dict:
    validos = int(origem["votos_validos"].sum())
    v1 = int(origem["c1_votos"].sum())
    v2 = int(origem["c2_votos"].sum())
    return {
        "uf": uf,
        "unidade": unidade,
        "regiao": regiao,
        "c1_votos": v1,
        "c1_pct": pct(v1, validos),
        "c2_votos": v2,
        "c2_pct": pct(v2, validos),
        "votos_validos": validos,
        "vencedor": vencedor_por_votos(v1, v2, cand1.nome, cand2.nome),
    }


def montar_tabelas(
    bruto: dict[str, dict],
    cand1: Candidato,
    cand2: Candidato,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    ufs = pd.DataFrame([linha_uf(uf, bruto, cand1, cand2) for uf in UFS])
    exterior = pd.DataFrame([linha_uf("ZZ", bruto, cand1, cand2)])
    brasil = pd.DataFrame([linha_uf("BR", bruto, cand1, cand2)])
    total = pd.DataFrame(
        [
            totais_bloco(
                ufs,
                "UFs",
                "Total das 27 Unidades da Federação",
                "Brasil (sem exterior)",
                cand1,
                cand2,
            )
        ]
    )
    extra = pd.concat([brasil, total], ignore_index=True)
    return ufs.reset_index(drop=True), exterior.reset_index(drop=True), extra.reset_index(drop=True)


def _borda_fina(ws) -> None:
    thin = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin


def _titulos(cand1: Candidato, cand2: Candidato) -> list[str]:
    return [
        "UF",
        "Unidade da Federação",
        "Região",
        f"{cand1.nome} (votos válidos)",
        f"{cand1.nome} (% válidos)",
        f"{cand2.nome} (votos válidos)",
        f"{cand2.nome} (% válidos)",
        "Total de votos válidos",
        "Mais votado",
    ]


def _escrever_cabecalho(ws, cand1: Candidato, cand2: Candidato) -> None:
    fill_id = PatternFill("solid", fgColor="1B4F72")
    fill_t2 = PatternFill("solid", fgColor="6C3483")
    font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells("A1:C1")
    ws.merge_cells("D1:I1")
    ws["A1"] = "Unidade da Federação"
    ws["D1"] = f"2º turno — {DATA_2T}"
    for col in range(1, 10):
        fill = fill_id if col <= 3 else fill_t2
        cell = ws.cell(1, col)
        cell.fill = fill
        cell.font = font
        cell.alignment = center
    for col, titulo in enumerate(_titulos(cand1, cand2), start=1):
        cell = ws.cell(2, col, titulo)
        fill = fill_id if col <= 3 else fill_t2
        cell.fill = fill
        cell.font = font
        cell.alignment = center
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 36
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:I{ws.max_row}"


def _pintar_vencedor(ws, cand1: Candidato, cand2: Candidato) -> None:
    fill_c1 = PatternFill("solid", fgColor="F5B7B1")
    fill_c2 = PatternFill("solid", fgColor="AED6F1")
    for row in range(3, ws.max_row + 1):
        vencedor = ws.cell(row, 9).value
        if vencedor == cand1.nome:
            ws.cell(row, 4).fill = fill_c1
            ws.cell(row, 9).fill = fill_c1
        elif vencedor == cand2.nome:
            ws.cell(row, 6).fill = fill_c2
            ws.cell(row, 9).fill = fill_c2


def _formatar_corpo(ws) -> None:
    inteiros = {4, 6, 8}
    percentuais = {5, 7}
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    for row in range(3, ws.max_row + 1):
        for col in range(1, 10):
            ws.cell(row, col).alignment = center
        ws.cell(row, 2).alignment = left
        for col in inteiros:
            ws.cell(row, col).number_format = "#,##0"
        for col in percentuais:
            ws.cell(row, col).number_format = "0.00"
    larguras = [8, 38, 22, 24, 16, 24, 16, 22, 16]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _preencher_aba(
    ws,
    df: pd.DataFrame,
    cand1: Candidato,
    cand2: Candidato,
    destacar_ultima: bool = False,
) -> None:
    corpo = df.loc[:, COLUNAS]
    for r_idx, valores in enumerate(corpo.itertuples(index=False, name=None), start=3):
        for c_idx, value in enumerate(valores, start=1):
            ws.cell(r_idx, c_idx, value)
    _escrever_cabecalho(ws, cand1, cand2)
    _formatar_corpo(ws)
    _pintar_vencedor(ws, cand1, cand2)
    if destacar_ultima and ws.max_row >= 3:
        fill = PatternFill("solid", fgColor="D5F5E3")
        font = Font(bold=True)
        for col in range(1, 10):
            cell = ws.cell(ws.max_row, col)
            cell.fill = fill
            cell.font = font
    _borda_fina(ws)


def fmt_int_br(valor: int) -> str:
    return f"{int(valor):,}".replace(",", ".")


def _notas(cand1: Candidato, cand2: Candidato) -> list[str]:
    t = TSE_BRASIL_2T
    return [
        "Tribunal Superior Eleitoral (TSE) — 2º turno das Eleições Gerais 2014, cargo de Presidente da República.",
        (
            "Os dois candidatos do 2º turno (os mais votados no 1º turno nacional): "
            f"{cand1.nome} (Dilma Rousseff, PT, nº 13) e {cand2.nome} (Aécio Neves, PSDB, nº 45)."
        ),
        f"Data do 2º turno: {DATA_2T}. Proclamação definitiva do resultado pelo plenário do TSE em 09/12/2014.",
        "Votos válidos: votos nominais nos candidatos (excluídos brancos e nulos). Percentuais sobre os válidos da UF.",
        "Aba Por_UF: 26 estados + Distrito Federal. A última linha soma só as 27 UFs (sem voto no exterior).",
        "Aba Brasil_e_exterior: total nacional oficial do TSE (inclui ZZ/exterior), voto no exterior e a soma das 27 UFs.",
        "Microdados por UF: http://electionresources.org/br/president.php?election=2014 (tabelas compiladas da totalização do TSE).",
        (
            f"Brasil oficial 2º turno: {cand1.nome} {fmt_int_br(t['dilma'])} (51,64%); "
            f"{cand2.nome} {fmt_int_br(t['aecio'])} (48,36%); válidos {fmt_int_br(t['vv'])}."
        ),
        (
            "A soma das 27 UFs + voto no exterior fica 65.695 votos válidos (0,06%) abaixo do total Brasil "
            "proclamado pelo TSE — o mesmo recorte das tabelas por estado (Election Resources / consolidações "
            "públicas). A linha Brasil usa a totalização oficial da Corte."
        ),
        "Notícia TSE da proclamação: https://www.tse.jus.br/comunicacao/noticias/2014/Dezembro/plenario-do-tse-proclama-resultado-definitivo-do-segundo-turno-da-eleicao-presidencial",
    ]


def gravar_xlsx(
    ufs: pd.DataFrame,
    exterior: pd.DataFrame,
    extra: pd.DataFrame,
    cand1: Candidato,
    cand2: Candidato,
    caminho: Path,
) -> Path:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Por_UF"
    ufs_com_total = pd.concat([ufs, extra.loc[extra["uf"] == "UFs"]], ignore_index=True)
    _preencher_aba(ws, ufs_com_total, cand1, cand2, destacar_ultima=True)
    ws.sheet_view.showGridLines = False

    ws2 = wb.create_sheet("Brasil_e_exterior")
    brasil_ext = pd.concat(
        [extra.loc[extra["uf"] == "BR"], exterior, extra.loc[extra["uf"] == "UFs"]],
        ignore_index=True,
    )
    _preencher_aba(ws2, brasil_ext, cand1, cand2)
    ws2.sheet_view.showGridLines = False

    ws3 = wb.create_sheet("Fonte")
    ws3["A1"] = "Fonte, recorte e conceito"
    ws3["A1"].font = Font(bold=True, size=14)
    for i, texto in enumerate(_notas(cand1, cand2), start=3):
        ws3[f"A{i}"] = texto
        ws3[f"A{i}"].alignment = Alignment(wrap_text=True)
        ws3.row_dimensions[i].height = 18
    ws3.column_dimensions["A"].width = 140
    wb.save(caminho)
    return caminho


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Planilha TSE: votos válidos de Dilma e Aécio no 2º turno de 2014, em cada UF."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=ROOT / "output" / "eleicoes_2014_presidente_2t_por_uf.xlsx",
    )
    parser.add_argument("--baixar", action="store_true", help="Força novo download.")
    args = parser.parse_args(argv)
    bruto = carregar_dados(forcar=args.baixar)
    cand1, cand2 = dois_mais_votados(bruto["BR"]["2t"])
    ufs, exterior, extra = montar_tabelas(bruto, cand1, cand2)
    caminho = gravar_xlsx(ufs, exterior, extra, cand1, cand2, args.output)
    print(f"[OK] {caminho}")
    print(f"Candidatos: {cand1.nome} e {cand2.nome} (2º turno 2014)")
    print(ufs.to_string(index=False))
    br = extra.loc[extra["uf"] == "BR"].iloc[0]
    print(
        "Brasil TSE 2º turno:",
        int(br["c1_votos"]),
        int(br["c2_votos"]),
        "válidos",
        int(br["votos_validos"]),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
