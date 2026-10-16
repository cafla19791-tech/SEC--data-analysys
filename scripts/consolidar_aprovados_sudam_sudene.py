#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolida o que Sudam (e cruzamento RFB) publicam sobre incentivos
de redução/isenção de IRPJ.

Sudam — listas anuais de aprovados (2010–2023, PDFs do repositório):
  empresa, CNPJ, município, UF, modalidade/pleito, laudo
  (sem valor de renúncia)

RFB — quando disponível, anexa valores de renúncia 75% (2015–2023)
já calculados no discriminativo / fonte local.

Sudene — não há lista pública equivalente em PDF/planilha no portal;
fica registrado na capa.

Uso::

  python scripts/consolidar_aprovados_sudam_sudene.py
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

MARKER = "aprovados-sudam-sudene-20260817a"

SUDAM_BASE = (
    "http://repositorio.sudam.gov.br/sudam/incentivos-fiscais/relatorios/"
    "relacao-incentivos-fiscais-reducao-e-isencao-por-ano"
)
SUDAM_URLS = {
    2010: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2010.pdf",
    2011: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2011.pdf",
    2012: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2012.pdf",
    2013: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2013.pdf",
    2014: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2014.pdf",
    2015: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2015.pdf",
    2016: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2016.pdf",
    2017: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2017.pdf",
    2018: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2018.pdf",
    2019: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2019.pdf",
    2020: f"{SUDAM_BASE}/relacao-incentivos-fiscais-reducao-e-isencao-2020.pdf",
    2021: f"{SUDAM_BASE}/reducao_e_isencao___2021.pdf",
    2022: f"{SUDAM_BASE}/planilha-aprovados-2022-reducao-e-reinvestimento.pdf",
    2023: f"{SUDAM_BASE}/planilha-aprovados-2023-reducao-e-reinvestimento.pdf",
}

CNPJ_RE = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})")
LAUDO_END_RE = re.compile(r"\d{2}/\d{2}/\d{4}\s+\d{1,4}/\d{4}")
PROC_RE = re.compile(r"\d+\s+\d{5}\.\d+/\d{4}-\d+\s*")
UFS = ("AC", "AP", "AM", "PA", "RO", "RR", "TO", "MT", "MA")


def baixar_pdfs(pasta: Path, anos: list[int] | None = None) -> list[Path]:
    pasta.mkdir(parents=True, exist_ok=True)
    paths = []
    alvo = anos or sorted(SUDAM_URLS)
    for ano in alvo:
        url = SUDAM_URLS[ano]
        dest = pasta / f"sudam_{ano}.pdf"
        if dest.exists() and dest.stat().st_size > 10_000:
            print(f"[CACHE] {dest.name}")
        else:
            print(f"[DOWNLOAD] {ano} ...")
            r = requests.get(url, timeout=120)
            r.raise_for_status()
            dest.write_bytes(r.content)
            print(f"  OK {dest.stat().st_size / 1e3:.0f} KB")
        paths.append(dest)
    return paths


def pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join((p.extract_text() or "") for p in reader.pages).replace("\xa0", " ")


def parse_sudam_pdf(path: Path, ano: int) -> list[dict]:
    text = pdf_text(path)
    text = re.sub(
        r"EMPRESA CNPJ/MF.*?N\.º/ANO", " ", text, flags=re.S | re.I
    )
    text = re.sub(
        r"N PROCESSO EMPRESA CNPJ MUNICIPIO UF MODALIDADE PRODUTO/\s*SERVIÇO LAUDO N\.º",
        " ",
        text,
        flags=re.I,
    )
    matches = list(CNPJ_RE.finditer(text))
    rows: list[dict] = []
    uf_alt = "|".join(UFS)
    muni_pat = re.compile(
        rf"([A-Za-zÁÉÍÓÚÃÕÂÊÔáéíóúãõâêôçÇ][^0-9]{{1,45}}?)\s+({uf_alt})\b"
    )

    for i, m in enumerate(matches):
        prev_end = matches[i - 1].end() if i else 0
        chunk = PROC_RE.sub(" ", text[prev_end : m.start()])
        ends = list(LAUDO_END_RE.finditer(chunk))
        empresa = chunk[ends[-1].end() :] if ends else chunk
        empresa = re.sub(r"\s+", " ", empresa).strip(" -\n\t|")
        empresa = re.sub(r"^\d+\s+", "", empresa)
        # remove remnant headers
        for junk in (
            "ENQUADRAMENTO",
            "PLEITO",
            "MODALIDADE",
            "LAUDO DATA",
            "LAUDO N",
            "/ANO",
        ):
            if junk in empresa.upper():
                empresa = re.split(junk, empresa, flags=re.I)[-1].strip(" -")
        if len(empresa) > 100:
            empresa = " ".join(empresa.split()[-12:])
        if len(empresa) < 3:
            continue

        nxt = matches[i + 1].start() if i + 1 < len(matches) else min(len(text), m.end() + 320)
        after = re.sub(r"\s+", " ", text[m.end() : nxt]).strip()
        muni_uf = muni_pat.match(after)
        municipio = muni_uf.group(1).strip() if muni_uf else ""
        uf = muni_uf.group(2) if muni_uf else ""

        modalidade = ""
        if re.search(r"Reinvestimento", after, re.I):
            modalidade = "Reinvestimento"
        elif re.search(r"Redu[cç][aã]o", after, re.I):
            modalidade = "Redução"
        elif re.search(r"Isen[cç][aã]o", after, re.I):
            modalidade = "Isenção"

        pleito_m = re.search(
            r"\b(Implanta[cç][aã]o|Moderniza[cç][aã]o(?:\s+Total)?|"
            r"Amplia[cç][aã]o|Diversifica[cç][aã]o|"
            r"Incorpora[cç][aã]o(?:\s*-\s*[\wÇçãáéíóú]+)?)\b",
            after,
            re.I,
        )
        pleito = pleito_m.group(1) if pleito_m else ""
        if not modalidade and pleito:
            modalidade = pleito

        laudo_m = re.search(r"(\d{1,4}/\d{4})", after)
        data_m = re.search(r"(\d{2}/\d{2}/\d{4})", after)

        rows.append(
            {
                "Órgão": "SUDAM",
                "Ano da lista": ano,
                "Empresa": empresa,
                "CNPJ": m.group(1),
                "Município": municipio,
                "UF": uf,
                "Modalidade": modalidade,
                "Pleito": pleito,
                "Data do laudo": data_m.group(1) if data_m else "",
                "Laudo nº/ano": laudo_m.group(1) if laudo_m else "",
            }
        )
    return rows


def cnpj_digits(s: str) -> str:
    d = re.sub(r"\D", "", str(s))
    return d.zfill(14) if d.isdigit() and len(d) <= 14 else d


def carregar_renuncia_rfb(path: Path | None) -> pd.DataFrame:
    """Carrega renúncia 75% da fonte RFB local, se existir."""
    if path is None or not path.exists():
        return pd.DataFrame()
    try:
        from scripts.discriminativo_sudam_sudene_75_ipca import (
            COL_ANO,
            COL_CNPJ,
            COL_IPCA,
            COL_NOME,
            COL_VALOR,
            aplicar_ipca,
            carregar_ipca_desde_2002,
            carregar_renuncias_75,
        )
    except Exception as exc:
        print(f"[AVISO] Não carregou módulo RFB: {exc}")
        return pd.DataFrame()

    df = carregar_renuncias_75(path)
    ipca = carregar_ipca_desde_2002(None)
    df = aplicar_ipca(df, ipca)
    df["_cnpj"] = df[COL_CNPJ].map(cnpj_digits)
    return df.rename(
        columns={
            COL_ANO: "Ano-Calendário RFB",
            COL_CNPJ: "CNPJ RFB",
            COL_NOME: "Beneficiário RFB",
            COL_VALOR: "Valor Renunciado (R$)",
            COL_IPCA: "Valor Renunciado IPCA 31/07/2026 (R$)",
        }
    )


def consolidar(
    pasta_pdf: Path,
    *,
    fonte_rfb: Path | None,
    saida: Path,
    baixar: bool = True,
) -> dict:
    print(f"[{MARKER}]")
    if baixar:
        baixar_pdfs(pasta_pdf)

    todos: list[dict] = []
    for ano, url in sorted(SUDAM_URLS.items()):
        path = pasta_pdf / f"sudam_{ano}.pdf"
        if not path.exists():
            print(f"[AVISO] Falta {path.name}")
            continue
        rows = parse_sudam_pdf(path, ano)
        print(f"[SUDAM {ano}] {len(rows)} registros")
        todos.extend(rows)

    aprovados = pd.DataFrame(todos)
    if aprovados.empty:
        raise RuntimeError("Nenhum registro extraído dos PDFs Sudam.")

    aprovados["_cnpj"] = aprovados["CNPJ"].map(cnpj_digits)

    # empresas únicas Sudam
    emp = (
        aprovados.groupby("_cnpj", sort=False)
        .agg(
            CNPJ=("CNPJ", "first"),
            Empresa=("Empresa", lambda s: s.value_counts().index[0]),
            UFs=("UF", lambda s: ", ".join(sorted({x for x in s if x}))),
            Anos=("Ano da lista", lambda s: ", ".join(str(x) for x in sorted(set(s)))),
            Qtd_laudos=("Laudo nº/ano", "size"),
            Modalidades=("Modalidade", lambda s: ", ".join(sorted({x for x in s if x}))),
        )
        .reset_index()
        .sort_values("Empresa")
    )

    rfb = carregar_renuncia_rfb(fonte_rfb)
    cruzado = emp.copy()
    if not rfb.empty:
        tot_rfb = (
            rfb.groupby("_cnpj")
            .agg(
                **{
                    "Anos com renúncia RFB": ("Ano-Calendário RFB", lambda s: ", ".join(str(int(x)) for x in sorted(set(s)))),
                    "Valor Renunciado (R$)": ("Valor Renunciado (R$)", "sum"),
                    "Valor Renunciado IPCA 31/07/2026 (R$)": (
                        "Valor Renunciado IPCA 31/07/2026 (R$)",
                        "sum",
                    ),
                }
            )
            .reset_index()
        )
        cruzado = emp.merge(tot_rfb, on="_cnpj", how="left")
        cruzado["Valor Renunciado (R$)"] = cruzado["Valor Renunciado (R$)"].round(2)
        cruzado["Valor Renunciado IPCA 31/07/2026 (R$)"] = cruzado[
            "Valor Renunciado IPCA 31/07/2026 (R$)"
        ].round(2)

    # por ano lista
    por_ano = (
        aprovados.groupby("Ano da lista")
        .agg(Qtd_registros=("CNPJ", "size"), Qtd_CNPJ=("CNPJ", "nunique"))
        .reset_index()
    )

    saida.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(saida))
    fmt_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white"})
    fmt_num = wb.add_format({"num_format": "#,##0.00"})
    fmt_bold = wb.add_format({"bold": True})

    def dump(name: str, df: pd.DataFrame, money: set[str] | None = None):
        money = money or set()
        ws = wb.add_worksheet(name[:31])
        cols = [c for c in df.columns if not str(c).startswith("_")]
        for j, c in enumerate(cols):
            ws.write(0, j, str(c), fmt_hdr)
        for i, row in enumerate(df[cols].itertuples(index=False, name=None), start=1):
            for j, val in enumerate(row):
                col = cols[j]
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    continue
                if col in money and isinstance(val, (int, float)):
                    ws.write_number(i, j, float(val), fmt_num)
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    ws.write_number(i, j, float(val))
                else:
                    ws.write(i, j, str(val))
        ws.set_column(0, 0, 14)
        ws.set_column(1, 2, 40)

    # Capa
    ws = wb.add_worksheet("Capa")
    capa = [
        ("Título", "Aprovados Sudam (+ cruzamento RFB) — redução/isenção IRPJ"),
        ("Sudam", "Listas anuais de laudos aprovados 2010–2023 (repositório Sudam)"),
        (
            "Sudene",
            "Não há planilha/PDF público equivalente de aprovados no portal; "
            "consulta via SIBF. Valores de renúncia por CNPJ vêm da RFB (quando houver).",
        ),
        ("O que a Sudam traz", "Empresa, CNPJ, município, UF, modalidade/pleito, laudo — SEM valor R$"),
        ("O que a RFB traz", "Valor renunciado por CNPJ/ano (75%) 2015–2023, com IPCA até 31/07/2026"),
        ("Registros Sudam", f"{len(aprovados):,}"),
        ("CNPJ distintos Sudam", f"{aprovados['CNPJ'].nunique():,}"),
        ("Marker", MARKER),
        ("Gerado em", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    ws.write(0, 0, "Campo", fmt_bold)
    ws.write(0, 1, "Valor", fmt_bold)
    for i, (k, v) in enumerate(capa, start=1):
        ws.write(i, 0, k)
        ws.write(i, 1, v)
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 100)

    dump("Sudam_Aprovados", aprovados)
    dump("Sudam_Empresas", cruzado, {"Valor Renunciado (R$)", "Valor Renunciado IPCA 31/07/2026 (R$)"})
    dump("Sudam_Por_Ano_Lista", por_ano)
    if not rfb.empty:
        dump(
            "RFB_Renuncia_75",
            rfb,
            {"Valor Renunciado (R$)", "Valor Renunciado IPCA 31/07/2026 (R$)"},
        )

    # Sudene placeholder sheet
    ws = wb.add_worksheet("Sudene_Nota")
    ws.write(0, 0, "Situação", fmt_bold)
    ws.write(
        1,
        0,
        "A Sudene não disponibiliza, no portal público, uma relação anual em PDF/XLSX "
        "equivalente à da Sudam (aprovados por ano). O acompanhamento é via SIBF. "
        "Beneficiários Sudene com valor de renúncia aparecem na aba RFB_Renuncia_75 "
        "(quando o CNPJ está na base da Receita), misturados a Sudam.",
    )
    ws.set_column(0, 0, 120)

    wb.close()
    print(f"[OK] {saida} ({saida.stat().st_size / 1e6:.2f} MB)")
    return {"aprovados": aprovados, "empresas": cruzado, "por_ano": por_ano}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--pasta-pdf",
        type=Path,
        default=ROOT / "data" / "sudam_sudene" / "aprovados",
    )
    p.add_argument(
        "--fonte-rfb",
        type=Path,
        default=ROOT / "data" / "sudam_sudene" / "renuncia_sudam_sudene.xlsx",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=ROOT
        / "output"
        / "sudam_sudene"
        / "APROVADOS_SUDAM_CRUZAMENTO_RFB.xlsx",
    )
    p.add_argument("--sem-baixar", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        # copy cached pdfs if present in /tmp
        tmp = Path("/tmp/sudam_pdfs")
        if tmp.exists() and not args.sem_baixar:
            args.pasta_pdf.mkdir(parents=True, exist_ok=True)
            for p in tmp.glob("sudam_*.pdf"):
                dest = args.pasta_pdf / p.name
                if not dest.exists() or dest.stat().st_size < p.stat().st_size:
                    dest.write_bytes(p.read_bytes())
        info = consolidar(
            args.pasta_pdf,
            fonte_rfb=args.fonte_rfb if args.fonte_rfb.exists() else None,
            saida=args.saida,
            baixar=not args.sem_baixar,
        )
        print(info["por_ano"].to_string(index=False))
        print(f"Empresas/CNPJ: {len(info['empresas']):,}")
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
