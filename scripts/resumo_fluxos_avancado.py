#!/usr/bin/env python3
"""
Resumo avançado de fluxos ContAgil — contratos + ano + agente + SELIC.

Lê os ``fluxos_*.csv/.xlsx`` da pasta de saída ContAgil, cruza com o Excel
original de operações indiretas e recalcula o impacto fiscal com a série
SELIC (coluna D / FATOR_30_06_2026).

Saídas (na pasta ``--pasta`` ou ``--output-dir``):
  - resumo_fluxos_avancado.xlsx  (abas Contratos, Por_Ano, Por_Agente,
    Impacto_Por_Ano, Totais)
  - resumo_contratos.xlsx / .csv
  - resumo_por_ano.xlsx / .csv
  - resumo_por_agente.xlsx / .csv
  - impacto_fiscal_por_ano.xlsx / .csv

Uso (WinPython ContAgil):
  python scripts/resumo_fluxos_avancado.py \\
      --pasta "C:\\Arquivos de Programas RFB\\ContAgilAppBeta64\\python_jep\\winpython\\saida" \\
      --original "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx" \\
      --selic "STP-20260716182715078.xlsx"

Uso (repo / cloud):
  python3 scripts/resumo_fluxos_avancado.py \\
      --pasta output \\
      --original data/sample_operacoes_com_agente.csv \\
      --baixar-selic
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from scripts.gerar_fluxos import (
    CONTAGIL_PASTA_DADOS,
    CONTAGIL_PASTA_SAIDA,
    CONTAGIL_SELIC_ALT,
    CONTAGIL_SELIC_DEFAULT,
    CONTAGIL_WINPYTHON,
    DATA_DIR,
    OUTPUT_DIR,
    SELIC_BACEN_CACHE,
    SelicSerie,
    agregar_por_agente,
    load_from_csv,
    load_from_excel,
    resolver_arquivo_selic,
    resolver_excel_operacoes,
)
from scripts.impacto_fiscal_por_ano import agregar_impacto_por_ano
from scripts.resumo_fluxos import (
    carregar_fluxos,
    normalizar_colunas,
    resumo_por_ano,
    resumo_por_contrato,
    salvar_resumos,
)

NOME_ORIGINAL_DEFAULT = "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx"
NOME_SELIC_DEFAULT = "STP-20260716182715078.xlsx"
WORKBOOK_NAME = "resumo_fluxos_avancado.xlsx"


def _parece_contagil(path: Path | None) -> bool:
    if path is None:
        return False
    texto = str(path).replace("/", "\\").upper()
    return "CONTAGIL" in texto or "WINPYTHON" in texto or texto.startswith("C:\\ARQUIVOS")


def resolver_pasta(pasta: Path | None) -> Path:
    """Resolve pasta de saída ContAgil (saida) ou fallback local."""
    if pasta is not None:
        if pasta.exists() and pasta.is_dir():
            return pasta
        if _parece_contagil(pasta):
            espelho = DATA_DIR / "contagil_winpython" / "saida"
            if espelho.exists():
                print(f"⚠️ Pasta ContAgil ausente: {pasta}")
                print(f"   Usando espelho local: {espelho}")
                return espelho
            if OUTPUT_DIR.exists():
                print(f"⚠️ Pasta ContAgil ausente: {pasta}")
                print(f"   Usando output/ do repo: {OUTPUT_DIR}")
                return OUTPUT_DIR
        raise FileNotFoundError(f"Pasta de fluxos não encontrada: {pasta}")
    if CONTAGIL_PASTA_SAIDA.exists():
        return CONTAGIL_PASTA_SAIDA
    return OUTPUT_DIR


def _candidatos_nome(nome: str | Path, bases: list[Path]) -> list[Path]:
    path = Path(nome)
    out: list[Path] = []
    seen: set[str] = set()

    def _add(p: Path) -> None:
        key = str(p)
        if key in seen:
            return
        seen.add(key)
        out.append(p)

    if path.is_absolute() or path.parent != Path("."):
        _add(path)
    for base in bases:
        _add(base / path.name)
        _add(base / path)
    return out


def resolver_original(nome: str | Path | None, pasta: Path) -> Path:
    """Resolve Excel/CSV original de operações (nome curto ContAgil ok)."""
    if nome is None:
        found = resolver_excel_operacoes(None)
        if found is not None:
            return found
        sample = DATA_DIR / "sample_operacoes_com_agente.csv"
        if sample.exists():
            return sample
        raise FileNotFoundError(
            "Arquivo original de operações não encontrado. "
            f"Informe --original (ex.: {NOME_ORIGINAL_DEFAULT})."
        )

    bases = [
        Path.cwd(),
        pasta,
        pasta.parent,
        CONTAGIL_WINPYTHON,
        CONTAGIL_PASTA_DADOS,
        DATA_DIR / "contagil_winpython" / "dados",
        DATA_DIR,
        ROOT / "attachments",
        Path("/home/workdir/attachments"),
    ]
    for cand in _candidatos_nome(nome, bases):
        if cand.exists() and cand.is_file():
            return cand

    # Auto ContAgil / attachments
    found = resolver_excel_operacoes(Path(nome))
    if found is not None:
        return found

    sample = DATA_DIR / "sample_operacoes_com_agente.csv"
    if sample.exists() and _parece_contagil(Path(nome)):
        print(f"⚠️ Original ContAgil ausente: {nome}")
        print(f"   Usando amostra local: {sample}")
        return sample

    raise FileNotFoundError(f"Arquivo original não encontrado: {nome}")


def resolver_selic(
    nome: str | Path | None,
    pasta: Path,
    *,
    baixar_selic: bool = False,
) -> tuple[Path | None, SelicSerie | None]:
    """Resolve STP ContAgil e carrega SelicSerie (ou Bacen se --baixar-selic)."""
    if nome is not None:
        bases = [
            Path.cwd(),
            pasta,
            pasta.parent,
            CONTAGIL_WINPYTHON,
            DATA_DIR / "contagil_winpython",
            DATA_DIR,
            ROOT / "attachments",
            Path("/home/workdir/attachments"),
        ]
        for cand in _candidatos_nome(nome, bases):
            if cand.exists() and cand.is_file():
                serie = SelicSerie.from_excel(cand)
                return cand, serie
        # Tenta também variante com " (1)"
        alt = Path(str(nome).replace(".xlsx", " (1).xlsx"))
        for cand in _candidatos_nome(alt, bases):
            if cand.exists() and cand.is_file():
                serie = SelicSerie.from_excel(cand)
                return cand, serie

    resolvido = resolver_arquivo_selic(
        Path(nome) if nome is not None else None
    )
    if resolvido is not None:
        return resolvido, SelicSerie.from_excel(resolvido)

    # Defaults ContAgil
    for cand in (CONTAGIL_SELIC_ALT, CONTAGIL_SELIC_DEFAULT):
        if cand.exists():
            return cand, SelicSerie.from_excel(cand)

    if baixar_selic or nome is not None:
        # Nome ContAgil ausente neste ambiente → Bacen
        if nome is not None:
            print(f"⚠️ SELIC ContAgil ausente: {nome}")
            print("   Baixando SELIC Bacen SGS 11 (fatores ContAgil)...")
        serie = SelicSerie.from_bacen(cache_path=SELIC_BACEN_CACHE)
        return SELIC_BACEN_CACHE if SELIC_BACEN_CACHE.exists() else None, serie

    return None, None


def listar_arquivos_fluxos(pasta: Path) -> list[Path]:
    """Lista fluxos_*.csv/.xlsx na pasta (exclui diários e resumos)."""
    matches: list[Path] = []
    for pattern in ("fluxos_*.csv", "fluxos_*.xlsx", "fluxos_*.xls"):
        matches.extend(sorted(pasta.glob(pattern)))
    # Também aceita nomes sem underscore numérico (fluxos_amostra, etc.)
    for pattern in ("fluxos*.csv", "fluxos*.xlsx"):
        for p in sorted(pasta.glob(pattern)):
            if p not in matches:
                matches.append(p)

    filtrados: list[Path] = []
    for p in matches:
        stem = p.stem.lower()
        if "diario" in stem or "diarios" in stem:
            continue
        if stem.startswith("resumo") or "resumo_" in stem:
            continue
        filtrados.append(p)
    return filtrados


def _somente_amostra_resumo(path: Path) -> bool:
    """True se o Excel é workbook agregado (só aba Amostra_Parcelas)."""
    suffix = path.suffix.lower()
    if suffix not in {".xlsx", ".xls"}:
        return False
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return False
    nomes = set(xl.sheet_names)
    return "Amostra_Parcelas" in nomes and "Sheet1" not in nomes and "Parcelas" not in nomes


def carregar_fluxos_pasta(pasta: Path) -> tuple[pd.DataFrame, list[Path]]:
    """Concatena todos os arquivos de fluxos da pasta (só detalhe de parcelas)."""
    arquivos = listar_arquivos_fluxos(pasta)
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum fluxos_*.csv/.xlsx em {pasta}. "
            "Gere com scripts/contagil_fluxos.py ou scripts/gerar_fluxos.py."
        )

    # Se há detalhe completo, ignora workbooks só com Amostra_Parcelas
    tem_completo = any(not _somente_amostra_resumo(p) for p in arquivos)
    if tem_completo:
        filtrados = [p for p in arquivos if not _somente_amostra_resumo(p)]
        if filtrados:
            arquivos = filtrados

    partes: list[pd.DataFrame] = []
    usados: list[Path] = []
    for path in arquivos:
        try:
            df = carregar_fluxos(path)
        except ValueError as exc:
            print(f"  Ignorando {path.name}: {exc}")
            continue
        print(f"  Lendo fluxos: {path.name} ({len(df):,} linhas)")
        df["_arquivo_origem"] = path.name
        partes.append(df)
        usados.append(path)
    if not partes:
        raise FileNotFoundError(
            f"Nenhum arquivo de parcelas válido em {pasta} "
            f"(candidatos: {[p.name for p in arquivos]})."
        )
    return pd.concat(partes, ignore_index=True), usados


def carregar_original(path: Path) -> pd.DataFrame:
    """Carrega Excel portal (header=5) ou CSV filtrado/amostra."""
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return load_from_excel(path)
    return load_from_csv(path)


def enriquecer_resumo_contratos(
    resumo: pd.DataFrame,
    contratos: pd.DataFrame,
) -> pd.DataFrame:
    """Anexa metadados do Excel original ao resumo por contrato."""
    out = resumo.reset_index()
    if "contrato" not in out.columns:
        out = out.rename(columns={out.columns[0]: "contrato"})

    cols_meta = [
        c
        for c in (
            "contrato",
            "agente",
            "valor_desembolsado",
            "data_contratacao",
            "juros",
            "prazo_carencia",
            "prazo_amortizacao",
            "custo_financeiro",
        )
        if c in contratos.columns
    ]
    meta = contratos[cols_meta].drop_duplicates("contrato")
    out = out.merge(meta, on="contrato", how="left")

    # Ordem amigável
    front = [
        c
        for c in (
            "contrato",
            "agente",
            "valor_desembolsado",
            "data_contratacao",
            "custo_financeiro",
            "juros",
            "prazo_carencia",
            "prazo_amortizacao",
        )
        if c in out.columns
    ]
    rest = [c for c in out.columns if c not in front]
    return out[front + rest]


def aplicar_impacto_contagil(
    df: pd.DataFrame,
    selic_serie: SelicSerie,
) -> pd.DataFrame:
    """Recalcula coluna impacto via fatores ContAgil (col D)."""
    from scripts.impacto_fiscal_por_ano import _impacto_contagil

    work = df.copy()
    work["impacto_contagil"] = _impacto_contagil(
        work["subsidio"], work["data_fluxo"], selic_serie
    )
    # Para agregações ContAgil, usa o impacto recalculado
    work["impacto"] = work["impacto_contagil"]
    work["impacto_fiscal"] = work["impacto_contagil"]
    return work


def montar_totais(
    resumo_contrato: pd.DataFrame,
    resumo_agente: pd.DataFrame,
    impacto_ano: pd.DataFrame,
    n_arquivos: int,
    n_parcelas: int,
) -> pd.DataFrame:
    """Tabela-resumo executiva (aba Totais)."""
    subsidio_col = "Total Subsídio (R$)"
    impacto_col = "Impacto Fiscal 2026 (R$)"
    rows = [
        {"Indicador": "Arquivos de fluxos", "Valor": int(n_arquivos)},
        {"Indicador": "Parcelas", "Valor": int(n_parcelas)},
        {"Indicador": "Contratos", "Valor": int(len(resumo_contrato))},
        {
            "Indicador": "Agentes financeiros",
            "Valor": int(resumo_agente["Agente"].nunique())
            if not resumo_agente.empty
            else 0,
        },
        {
            "Indicador": "Total Subsídio (R$)",
            "Valor": round(float(resumo_contrato[subsidio_col].sum()), 2),
        },
        {
            "Indicador": "Total Impacto Fiscal 2026 (R$)",
            "Valor": round(float(resumo_contrato[impacto_col].sum()), 2),
        },
    ]
    if not impacto_ano.empty and "Impacto Fiscal 2026 (R$)" in impacto_ano.columns:
        rows.append(
            {
                "Indicador": "Anos com pagamento",
                "Valor": int(impacto_ano["Ano"].nunique()),
            }
        )
    # object dtype evita coerção int→float (senão tudo vira R$ no print)
    return pd.DataFrame(rows).astype({"Indicador": "string", "Valor": "object"})


def formatar_valor_total(indicador: str, valor) -> str:
    """Formata linha da aba Totais para o console."""
    if "R$" in str(indicador):
        return f"R$ {float(valor):,.2f}"
    try:
        return f"{int(valor):,}"
    except (TypeError, ValueError):
        return str(valor)


def salvar_workbook(
    output_dir: Path,
    *,
    resumo_contrato: pd.DataFrame,
    resumo_ano: pd.DataFrame,
    resumo_agente: pd.DataFrame,
    impacto_ano: pd.DataFrame,
    totais: pd.DataFrame,
) -> Path:
    """Grava Excel multi-aba + espelhos individuais."""
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook = output_dir / WORKBOOK_NAME

    # Contratos / ano no formato índice (compatível com resumo_fluxos)
    rc_idx = resumo_contrato
    if "contrato" in rc_idx.columns:
        rc_for_simple = resumo_contrato.set_index("contrato")[
            [
                c
                for c in (
                    "Total Subsídio (R$)",
                    "Impacto Fiscal 2026 (R$)",
                    "Saldo Final (R$)",
                    "Quantidade de Parcelas",
                )
                if c in resumo_contrato.columns
            ]
        ]
    else:
        rc_for_simple = resumo_contrato

    ra = resumo_ano
    if not isinstance(ra.index, pd.MultiIndex):
        # já é multiindex de resumo_por_ano
        pass

    salvar_resumos(rc_for_simple, ra, output_dir)

    if not resumo_agente.empty:
        resumo_agente.to_excel(output_dir / "resumo_por_agente.xlsx", index=False)
        resumo_agente.to_csv(output_dir / "resumo_por_agente.csv", index=False)

    if not impacto_ano.empty:
        impacto_ano.to_excel(output_dir / "impacto_fiscal_por_ano.xlsx", index=False)
        impacto_ano.to_csv(output_dir / "impacto_fiscal_por_ano.csv", index=False)

    with pd.ExcelWriter(workbook, engine="openpyxl") as writer:
        # Aba Contratos com metadados (sem índice)
        if "contrato" in resumo_contrato.columns:
            resumo_contrato.to_excel(writer, sheet_name="Contratos", index=False)
        else:
            resumo_contrato.to_excel(writer, sheet_name="Contratos")

        ra_out = ra.reset_index() if isinstance(ra.index, pd.MultiIndex) else ra
        ra_out.to_excel(writer, sheet_name="Por_Ano", index=False)

        if not resumo_agente.empty:
            resumo_agente.to_excel(writer, sheet_name="Por_Agente", index=False)
        if not impacto_ano.empty:
            impacto_ano.to_excel(writer, sheet_name="Impacto_Por_Ano", index=False)
        totais.to_excel(writer, sheet_name="Totais", index=False)

    return workbook


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--pasta",
        type=Path,
        default=None,
        help=(
            "Pasta ContAgil saida com fluxos_*.csv/.xlsx "
            f"(default: {CONTAGIL_PASTA_SAIDA} ou output/)."
        ),
    )
    p.add_argument(
        "--original",
        type=str,
        default=None,
        help=(
            "Excel/CSV de operações originais (nome curto ContAgil ok), "
            f"ex.: {NOME_ORIGINAL_DEFAULT}"
        ),
    )
    p.add_argument(
        "--selic",
        type=str,
        default=None,
        help=f"Excel STP ContAgil (col D), ex.: {NOME_SELIC_DEFAULT}",
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help="Baixa SELIC Bacen se o STP ContAgil não estiver disponível.",
    )
    p.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Pasta de saída (default: mesma de --pasta).",
    )
    p.add_argument(
        "--sem-recalcular",
        action="store_true",
        help="Usa impacto já gravado nos fluxos (não recalcula com SELIC).",
    )
    p.add_argument(
        "--top",
        type=int,
        default=10,
        help="Linhas no preview do console (default 10).",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        pasta = resolver_pasta(args.pasta)
        original_path = resolver_original(args.original, pasta)
        selic_path, selic_serie = resolver_selic(
            args.selic, pasta, baixar_selic=args.baixar_selic
        )
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    output_dir = args.output_dir if args.output_dir is not None else pasta

    print("Resumo avançado ContAgil — contratos / ano / agente / SELIC")
    print(f"Pasta fluxos : {pasta}")
    print(f"Original     : {original_path}")
    if selic_path is not None:
        print(f"SELIC        : {selic_path}")
    elif selic_serie is not None:
        print(f"SELIC        : série em memória ({selic_serie.origem})")
    else:
        print("SELIC        : (ausente — usando impacto da coluna)")

    try:
        df_raw, arquivos = carregar_fluxos_pasta(pasta)
        df = normalizar_colunas(df_raw)
        contratos = carregar_original(original_path)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"Arquivos     : {len(arquivos)}")
    print(f"Parcelas     : {len(df):,}")
    print(f"Contratos orig.: {len(contratos):,}")

    modo_selic = selic_serie is not None and not args.sem_recalcular
    if modo_selic:
        assert selic_serie is not None
        print(f"Recalculando impacto ContAgil ({len(selic_serie.datas):,} pontos)...")
        df = aplicar_impacto_contagil(df, selic_serie)
        impacto_ano = agregar_impacto_por_ano(
            df, modo="contagil", selic_serie=selic_serie
        )
    else:
        impacto_ano = agregar_impacto_por_ano(df, modo="coluna")

    resumo_c = resumo_por_contrato(df)
    resumo_c = enriquecer_resumo_contratos(resumo_c, contratos)
    resumo_a = resumo_por_ano(df)

    # Ranking por agente (Instituição nos fluxos ou mapa do original)
    try:
        resumo_agente = agregar_por_agente(df, contratos)
    except ValueError:
        # Sem instituição nos fluxos e sem agente no original
        resumo_agente = pd.DataFrame(
            columns=[
                "Agente",
                "Qtd Contratos",
                "Total Subsídio (R$)",
                "Impacto Fiscal 2026 (R$)",
            ]
        )

    rc_for_totais = (
        resumo_c.set_index("contrato")
        if "contrato" in resumo_c.columns
        else resumo_c
    )
    totais = montar_totais(
        rc_for_totais,
        resumo_agente,
        impacto_ano,
        n_arquivos=len(arquivos),
        n_parcelas=len(df),
    )

    print("\nResumo por Contrato (top):")
    print(resumo_c.head(args.top).to_string(index=False))

    print("\nPor Agente (top):")
    if not resumo_agente.empty:
        print(resumo_agente.head(args.top).to_string(index=False))
    else:
        print("  (sem agentes)")

    print("\nImpacto Fiscal por Ano:")
    print(impacto_ano.head(args.top).to_string(index=False))

    workbook = salvar_workbook(
        output_dir,
        resumo_contrato=resumo_c,
        resumo_ano=resumo_a,
        resumo_agente=resumo_agente,
        impacto_ano=impacto_ano,
        totais=totais,
    )

    print(f"\n✅ Workbook: {workbook}")
    print(f"   Pasta: {output_dir}")
    print("\n" + "=" * 60)
    print("TOTAIS")
    print("=" * 60)
    for _, row in totais.iterrows():
        print(f"{row['Indicador']}: {formatar_valor_total(row['Indicador'], row['Valor'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
