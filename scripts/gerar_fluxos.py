#!/usr/bin/env python3
"""
Gera fluxos mensais detalhados (carência + amortização SAC) e impacto fiscal
a valor de 30/06/2026, a partir de operações indiretas automáticas do BNDES.

Metodologia ContAgil (lógica corrigida) + carência corrigida:
  - taxa_contrato_efetiva (mensal):
      * TAXA FIXA / demais: (1 + juros)^(1/12) − 1
      * TJLP / TLP: (1 + 0,06)^(1/12) × (1 + juros)^(1/12) − 1
  - gerar_fluxos(df, df) → df_original (Instituição); gerar_fluxos(df, selic) → fatores
  - Aceita planilha bruta ContAgil (header=5, colunas em português)
  - Fluxos em TODOS os meses (carência + amortização) — corrige bug p=1..n
  - Amortização constante só após a carência
  - Dual balance: saldo_fiscal (principal) e saldo_contrato (com juros)
  - spread = (1 + (SELIC_m − taxa_contrato_m))^n
  - subsídio = saldo_fiscal × (SELIC_m − taxa_contrato_m)  [antes da amortização]
  - impacto_fiscal (calcular_impacto_fiscal_real):
      * STP ContAgil (col D): subsídio × FATOR_30_06_2026 / fator(nearest data_parcela)
      * Bacen/outros: subsídio × fator(nearest 30/06/2026) / fator(nearest data_parcela)
      * sem fatores: subsídio × (1 + SELIC_m)^(meses até 30/06/2026)

Entrada:
  - Excel do portal (header=5), ou
  - CSV aberto do BNDES (download automático 2009–2010)
  - (opcional) Excel SELIC ContAgil STP-*.xlsx ou --baixar-selic (Bacen SGS 11)

Saídas:
  - output/fluxos_completos_final.csv/.xlsx   (detalhe por parcela + colunas extras)
  - output/resumo_por_agente.csv|.xlsx       (agregado por instituição financeira)
  - output/fluxos_diarios_detalhados.xlsx    (opcional: --fluxo-diario, dia a dia)
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# ===================== CONFIGURAÇÕES =====================
TAXA_SELIC_ANUAL = 0.145  # 14,5% a.a.
TJLP_TLP_BASE = 0.06  # ContAgil: TJLP/TLP = 6% + juros do contrato
DATA_IMPACTO = datetime(2026, 6, 30)
# Fator ContAgil de referência em 30/06/2026 (coluna D do STP)
FATOR_30_06_2026 = 82.84819
AGENTE_NAO_INFORMADO = "Não informado"

# Caminhos ContAgil Windows (WinPython / RFB)
CONTAGIL_WINPYTHON = Path(
    r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
)
CONTAGIL_SELIC_DEFAULT = CONTAGIL_WINPYTHON / "STP-20260716182715078 (1).xlsx"
CONTAGIL_SELIC_ALT = CONTAGIL_WINPYTHON / "STP-20260716182715078.xlsx"
CONTAGIL_PASTA_DADOS = CONTAGIL_WINPYTHON / "dados"
CONTAGIL_PASTA_SAIDA = CONTAGIL_WINPYTHON / "saida"

BCB_SELIC_DIARIA_URL = (
    "https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados"
    "?formato=json&dataInicial={inicio}&dataFinal={fim}"
)


def taxa_mensal_composta(taxa_aa: float) -> float:
    """Converte taxa anual em taxa mensal composta: (1+r)^(1/12)-1."""
    return (1.0 + float(taxa_aa)) ** (1.0 / 12.0) - 1.0


def taxa_diaria_composta(taxa_aa: float) -> float:
    """Converte taxa anual em taxa diária composta (calendário): (1+r)^(1/365)-1."""
    return (1.0 + float(taxa_aa)) ** (1.0 / 365.0) - 1.0


def taxa_contrato_anual(custo_financeiro: str | None, juros_pct: float) -> float:
    """Taxa anual aproximada do contrato (legado / resumos).

    - TJLP / TLP: 6% + juros (%) do contrato
    - demais (ex.: TAXA FIXA): só o juros do contrato

    Para o cronograma mensal use ``taxa_contrato_efetiva``.
    """
    juros = float(juros_pct) / 100.0
    custo = str(custo_financeiro or "").upper()
    if "TJLP" in custo or "TLP" in custo:
        return TJLP_TLP_BASE + juros
    return juros


def taxa_contrato_efetiva(
    custo_financeiro: str | None | pd.Series | dict = None,
    juros_pct: float | None = None,
) -> float:
    """Taxa mensal efetiva do contrato (lógica corrigida ContAgil).

    Aceita ``taxa_contrato_efetiva(custo, juros_pct)`` ou o rascunho
    ContAgil ``taxa_contrato_efetiva(row)`` (Series/dict com
    ``Custo financeiro`` / ``Juros``).

    ``juros_pct`` vem da coluna Juros em % a.a. (ex.: 6.0 → 6%).

    - TAXA FIXA / demais: ``(1 + juros)^(1/12) − 1``
    - TJLP / TLP: ``(1 + 0,06)^(1/12) × (1 + juros)^(1/12) − 1``

    Nota: o rascunho ContAgil às vezes escreve
    ``(1,06)^(1/12)×(1+juros)−1`` (juros anual sem mensalizar). Aqui ambos
    os fatores anuais são compostos mensalmente — caso contrário a taxa
    mensal fica ~2–3%/mês e o subsídio vira fortemente negativo.
    """
    # ContAgil paste: taxa_contrato_efetiva(row)
    if juros_pct is None and custo_financeiro is not None and not isinstance(
        custo_financeiro, str
    ):
        row = custo_financeiro
        get = row.get if hasattr(row, "get") else lambda k, d=None: (
            row[k] if k in getattr(row, "index", ()) else d
        )
        custo_financeiro = get("Custo financeiro") or get("custo_financeiro") or ""
        raw_juros = get("Juros", get("juros", 0))
        try:
            juros_pct = float(str(raw_juros).replace("%", "").replace(",", "."))
        except (TypeError, ValueError):
            juros_pct = 0.0

    try:
        juros = float(juros_pct or 0.0) / 100.0
    except (TypeError, ValueError):
        juros = 0.0

    custo = str(custo_financeiro or "").upper()
    # TJLP/TLP antes de TAXA FIXA (ex.: "TLP + TAXA FIXA")
    if "TJLP" in custo:
        return (1.0 + TJLP_TLP_BASE) ** (1.0 / 12.0) * (1.0 + juros) ** (1.0 / 12.0) - 1.0
    if "TLP" in custo:
        return (1.0 + TJLP_TLP_BASE) ** (1.0 / 12.0) * (1.0 + juros) ** (1.0 / 12.0) - 1.0
    return (1.0 + juros) ** (1.0 / 12.0) - 1.0


TAXA_SELIC_MENSAL = taxa_mensal_composta(TAXA_SELIC_ANUAL)


class SelicSerie:
    """Lookup de fatores Selic acumulados (Excel STP ContAgil / Bacen)."""

    def __init__(
        self,
        datas: np.ndarray,
        fatores: np.ndarray,
        origem: str = "stp",
        fator_referencia: float | None = None,
    ):
        order = np.argsort(datas)
        self.datas = np.asarray(datas)[order]
        self.fatores = np.asarray(fatores, dtype=float)[order]
        self.origem = origem
        # ContAgil STP col D: usa FATOR_30_06_2026; Bacen: None (lookup na série)
        self.fator_referencia = fator_referencia

    @classmethod
    def from_dataframe(cls, selic: pd.DataFrame, origem: str = "dataframe") -> "SelicSerie":
        """Monta série a partir de DataFrame ContAgil (col A=data, col D=fator).

        Preferência:
          1) coluna nomeada 'fator_acumulado' (cache Bacen)
          2) coluna nomeada 'fator'
          3) coluna D (índice 3) — layout ContAgil corrigido / script RFB
          4) última coluna numérica com valores > 0
        """
        datas = pd.to_datetime(selic.iloc[:, 0], dayfirst=True, errors="coerce").values.astype(
            "datetime64[ns]"
        )

        fatores: np.ndarray | None = None
        fator_ref: float | None = None
        cols_lower = {str(c).strip().lower(): c for c in selic.columns}
        if "fator_acumulado" in cols_lower:
            fatores = pd.to_numeric(
                selic[cols_lower["fator_acumulado"]], errors="coerce"
            ).values
        elif "fator" in cols_lower:
            fatores = pd.to_numeric(selic[cols_lower["fator"]], errors="coerce").values
        elif selic.shape[1] >= 4:
            # ContAgil corrigido: fatores na coluna D (índice 3)
            fatores = pd.to_numeric(selic.iloc[:, 3], errors="coerce").values
            if np.any((~pd.isna(fatores)) & (fatores > 0)):
                fator_ref = FATOR_30_06_2026
            elif selic.shape[1] >= 5:
                fatores = pd.to_numeric(selic.iloc[:, 4], errors="coerce").values
            else:
                fatores = pd.to_numeric(selic.iloc[:, -1], errors="coerce").values
        else:
            fatores = pd.to_numeric(selic.iloc[:, -1], errors="coerce").values

        mask = ~pd.isna(datas) & ~pd.isna(fatores) & (fatores > 0)
        return cls(
            datas[mask],
            fatores[mask].astype(float),
            origem=origem,
            fator_referencia=fator_ref,
        )

    @classmethod
    def from_excel(cls, path: Path) -> "SelicSerie":
        """Lê STP ContAgil: col A = data, col D = fator acumulado."""
        return cls.from_dataframe(pd.read_excel(path), origem=f"stp:{path}")

    @classmethod
    def from_taxas_diarias(
        cls,
        datas: np.ndarray,
        taxas_pct_a_d: np.ndarray,
        origem: str = "bacen:sgs.11",
    ) -> "SelicSerie":
        """Monta fator acumulado ContAgil a partir da SELIC diária (% a.d.)."""
        datas = np.asarray(datas, dtype="datetime64[ns]")
        taxas = np.asarray(taxas_pct_a_d, dtype=float)
        mask = ~pd.isna(datas) & ~pd.isna(taxas)
        datas = datas[mask]
        taxas = taxas[mask]
        order = np.argsort(datas)
        datas = datas[order]
        taxas = taxas[order]
        # fator_t = Π (1 + i_k/100)
        fatores = np.cumprod(1.0 + taxas / 100.0)
        return cls(datas, fatores, origem=origem, fator_referencia=None)

    @classmethod
    def from_bacen(
        cls,
        inicio: str = "01/01/2008",
        fim: str | None = None,
        cache_path: Path | None = None,
    ) -> "SelicSerie":
        """Baixa SELIC diária (SGS 11) e gera fatores acumulados no formato ContAgil.

        A API do Bacen rejeita intervalos muito longos (HTTP 406); baixamos ano a ano.
        """
        import requests

        if fim is None:
            fim = DATA_IMPACTO.strftime("%d/%m/%Y")
        if cache_path is not None and cache_path.exists():
            print(f"Usando cache SELIC: {cache_path}")
            return cls.from_excel(cache_path)

        start_ts = pd.to_datetime(inicio, dayfirst=True)
        end_ts = pd.to_datetime(fim, dayfirst=True)
        print(f"Baixando SELIC diária Bacen (SGS 11): {inicio} .. {fim} (por ano)")

        parts: list[pd.DataFrame] = []
        for year in range(start_ts.year, end_ts.year + 1):
            y0 = max(start_ts, pd.Timestamp(year=year, month=1, day=1))
            y1 = min(end_ts, pd.Timestamp(year=year, month=12, day=31))
            url = BCB_SELIC_DIARIA_URL.format(
                inicio=y0.strftime("%d/%m/%Y"),
                fim=y1.strftime("%d/%m/%Y"),
            )
            last_err: Exception | None = None
            rows = None
            for attempt in range(1, 6):
                try:
                    resp = requests.get(
                        url, timeout=120, headers={"Accept": "application/json"}
                    )
                    resp.raise_for_status()
                    rows = resp.json()
                    break
                except Exception as exc:  # noqa: BLE001 — retries em rede/Bacen
                    last_err = exc
                    wait = min(2**attempt, 32)
                    print(f"  {year}: falha tentativa {attempt}/5 ({exc}); retry {wait}s")
                    time.sleep(wait)
            if rows is None:
                raise RuntimeError(f"Falha ao baixar SELIC {year}: {last_err}")
            if not rows:
                continue
            chunk = pd.DataFrame(rows)
            parts.append(chunk)
            print(f"  {year}: {len(chunk):,} dias")

        if not parts:
            raise RuntimeError("Bacen retornou série SELIC vazia.")

        df = pd.concat(parts, ignore_index=True)
        df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
        df["taxa_pct_a_d"] = pd.to_numeric(df["valor"], errors="coerce")
        df = (
            df.dropna(subset=["data", "taxa_pct_a_d"])
            .drop_duplicates(subset=["data"], keep="last")
            .sort_values("data")
            .reset_index(drop=True)
        )

        serie = cls.from_taxas_diarias(
            df["data"].values.astype("datetime64[ns]"),
            df["taxa_pct_a_d"].values,
            origem="bacen:sgs.11",
        )

        if cache_path is not None:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            # Cache Bacen: fator nomeado (não usa col D do STP ContAgil)
            out = pd.DataFrame(
                {
                    "data": pd.to_datetime(serie.datas),
                    "col_b": np.nan,
                    "col_c": np.nan,
                    "taxa_pct_a_d": df["taxa_pct_a_d"].values,
                    "fator_acumulado": serie.fatores,
                }
            )
            out.to_excel(cache_path, index=False)
            print(f"SELIC fatores salvos: {cache_path} ({len(out):,} linhas)")

        return serie

    def idx_proximo(self, data) -> int:
        """Equivalente a pd.Index(...).get_indexer([data], method='nearest')."""
        ts = np.datetime64(pd.Timestamp(data), "ns")
        if len(self.datas) == 0:
            raise ValueError("Série SELIC vazia")
        idx = int(np.searchsorted(self.datas, ts, side="left"))
        if idx <= 0:
            return 0
        if idx >= len(self.datas):
            return len(self.datas) - 1
        before = idx - 1
        after = idx
        if abs(self.datas[after].astype("int64") - ts.astype("int64")) < abs(
            ts.astype("int64") - self.datas[before].astype("int64")
        ):
            return after
        return before

    def fator_rapido(self, datas) -> np.ndarray:
        """Fatores na data mais próxima (method='nearest' ContAgil)."""
        datas_arr = np.array(datas, dtype="datetime64[ns]")
        out = np.empty(len(datas_arr), dtype=float)
        for i, ts in enumerate(datas_arr):
            out[i] = self.fatores[self.idx_proximo(ts)]
        return out

    def capitalizar(
        self,
        valor: float,
        data_fluxo: datetime,
        data_impacto: datetime = DATA_IMPACTO,
    ) -> float:
        """impacto = valor × fator(impacto) / fator(fluxo), como no ContAgil."""
        return calcular_impacto_fiscal_real(valor, data_fluxo, self, data_impacto)


def calcular_impacto_fiscal_real(
    subsidio: float,
    data_parcela,
    selic_serie: SelicSerie,
    data_impacto: datetime = DATA_IMPACTO,
) -> float:
    """
    ContAgil (fator Selic correto — coluna D):

      idx = nearest(data_parcela)           # própria data da parcela
      fator_parcela = fatores_coluna_d[idx]
      fator_acumulado = FATOR_30_06_2026 / fator_parcela   # STP ContAgil
      # ou, sem fator de referência (Bacen): fator(data_impacto) / fator_parcela
      impacto = subsidio * fator_acumulado
    """
    if subsidio <= 0:
        return 0.0
    idx = selic_serie.idx_proximo(data_parcela)
    fator_parcela = float(selic_serie.fatores[idx])
    if fator_parcela <= 0:
        return 0.0

    if selic_serie.fator_referencia is not None:
        fator_fim = float(selic_serie.fator_referencia)
    else:
        fator_fim = float(selic_serie.fatores[selic_serie.idx_proximo(data_impacto)])

    if fator_fim <= 0:
        return round(float(subsidio), 2)
    return round(float(subsidio) * (fator_fim / fator_parcela), 2)


def candidatos_arquivo_selic(explicit: Path | None = None) -> list[Path]:
    """Ordem de busca do Excel STP ContAgil / cache Bacen."""
    found: list[Path] = []
    seen: set[str] = set()

    def _add(p: Path | None) -> None:
        if p is None:
            return
        key = str(p)
        if key in seen:
            return
        seen.add(key)
        found.append(p)

    _add(explicit)
    env = os.environ.get("CONTAGIL_SELIC") or os.environ.get("SELIC_STP")
    if env:
        _add(Path(env))
    _add(CONTAGIL_SELIC_DEFAULT)
    _add(CONTAGIL_SELIC_ALT)
    _add(DATA_DIR / "STP-20260716182715078 (1).xlsx")
    _add(DATA_DIR / "STP-20260716182715078.xlsx")
    _add(DATA_DIR / "selic_fatores_bacen.xlsx")
    # Anexos de cloud agents / ContAgil exportados localmente
    for base in (
        Path("/home/workdir/attachments"),
        Path.cwd() / "attachments",
        ROOT / "attachments",
    ):
        _add(base / "STP-20260716182715078 (1).xlsx")
        _add(base / "STP-20260716182715078.xlsx")
        for pattern in ("STP-*.xlsx", "STP*.xlsx", "*selic*.xlsx"):
            if base.is_dir():
                for p in sorted(base.glob(pattern)):
                    _add(p)
    for pattern in ("STP-*.xlsx", "STP*.xlsx", "*selic*.xlsx"):
        for p in sorted(DATA_DIR.glob(pattern)):
            _add(p)
        for p in sorted(Path.cwd().glob(pattern)):
            _add(p)
    return found


def resolver_arquivo_selic(explicit: Path | None = None) -> Path | None:
    """Retorna o primeiro STP/cache existente na lista de candidatos."""
    for path in candidatos_arquivo_selic(explicit):
        if path.exists() and path.is_file():
            return path
    return None


def candidatos_excel_operacoes(explicit: Path | None = None) -> list[Path]:
    """Ordem de busca do Excel de operações indiretas automáticas."""
    found: list[Path] = []
    seen: set[str] = set()

    def _add(p: Path | None) -> None:
        if p is None:
            return
        key = str(p)
        if key in seen:
            return
        seen.add(key)
        found.append(p)

    _add(explicit)
    env = os.environ.get("BNDES_OPERACOES_XLSX") or os.environ.get("OPERACOES_XLSX")
    if env:
        _add(Path(env))
    nome = "operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx"
    for base in (
        Path("/home/workdir/attachments"),
        Path.cwd() / "attachments",
        ROOT / "attachments",
        DATA_DIR,
        Path.cwd(),
    ):
        _add(base / nome)
        if base.is_dir():
            for p in sorted(base.glob("operacoes_indiretas_automaticas*.xlsx")):
                _add(p)
    return found


def resolver_excel_operacoes(explicit: Path | None = None) -> Path | None:
    """Retorna o primeiro Excel de operações existente."""
    for path in candidatos_excel_operacoes(explicit):
        if path.exists() and path.is_file():
            return path
    return None


BNDES_CSV_URL = (
    "https://dadosabertos.bndes.gov.br/dataset/"
    "10e21ad1-568e-45e5-a8af-43f2c05ef1a2/resource/"
    "612faa0b-b6be-4b2c-9317-da5dc2c0b901/download/"
    "operacoes-financiamento-operacoes-indiretas-automaticas.csv"
)

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
FILTERED_CSV = DATA_DIR / "operacoes_indiretas_automaticas_2009-2010.csv"
SELIC_BACEN_CACHE = DATA_DIR / "selic_fatores_bacen.xlsx"
OUTPUT_DIR = ROOT / "output"

EXCEL_COLUMNS = {
    "Data da contratação": "data_contratacao",
    "Data da Contratação": "data_contratacao",
    "Data da contratacao": "data_contratacao",
    "Valor Desembolsado R$ (*)": "valor_desembolsado",
    "Valor desembolsado Reais": "valor_desembolsado",
    "Valor Desembolsado Reais": "valor_desembolsado",
    "Valor da operação em Reais": "valor_desembolsado",
    "Valor da Operação em Reais": "valor_desembolsado",
    # ContAgil BNDES INDIRETAS (massa winpython/dados)
    "Valor histórico": "valor_desembolsado",
    "Valor Histórico": "valor_desembolsado",
    "Valor  Histórico": "valor_desembolsado",
    "Valor Histórico R$ (*)": "valor_desembolsado",
    "Valor Histórico R$ ": "valor_desembolsado",
    "Valor Histórico R$": "valor_desembolsado",
    "Valor Histórico em R$": "valor_desembolsado",
    "Juros": "juros",
    "Prazo - Carência (meses)": "prazo_carencia",
    "Prazo de Carência (meses)": "prazo_carencia",
    "Prazo - Carencia (meses)": "prazo_carencia",
    "Prazo - Amortização (meses)": "prazo_amortizacao",
    "Prazo de Amortização (meses)": "prazo_amortizacao",
    "Prazo - Amortizacao (meses)": "prazo_amortizacao",
    "Prazo - Amortizacao(meses)": "prazo_amortizacao",
    "Instituição Financeira Credenciada": "agente",
    "Instituicao Financeira Credenciada": "agente",
    "Custo financeiro": "custo_financeiro",
    "Custo Financeiro": "custo_financeiro",
    "Encargo financeiro": "custo_financeiro",
    "Encargo Financeiro": "custo_financeiro",
    "encargo financeiro": "custo_financeiro",
}

CSV_COLUMNS = {
    "data_da_contratacao": "data_contratacao",
    "valor_desembolsado_reais": "valor_desembolsado",
    "valor_da_operacao_em_reais": "valor_desembolsado",
    "juros": "juros",
    "prazo_carencia_meses": "prazo_carencia",
    "prazo_amortizacao_meses": "prazo_amortizacao",
    "instituicao_financeira_credenciada": "agente",
    "custo_financeiro": "custo_financeiro",
}

# Aliases após normalização (minúsculas, sem acento/símbolos).
# Cobre ContAgil, portal BNDES e variações "BNDES INDIRETAS *.xlsx".
NORM_COLUMN_ALIASES: dict[str, str] = {
    "data_da_contratacao": "data_contratacao",
    "data_contratacao": "data_contratacao",
    "data_de_contratacao": "data_contratacao",
    "valor_desembolsado_reais": "valor_desembolsado",
    "valor_desembolsado_r": "valor_desembolsado",
    "valor_desembolsado": "valor_desembolsado",
    "valor_da_operacao_em_reais": "valor_desembolsado",
    "valor_da_operacao_reais": "valor_desembolsado",
    "valor_da_operacao": "valor_desembolsado",
    # ContAgil BNDES INDIRETAS: "Valor histórico" / "Valor Histórico R$ (*)"
    "valor_historico": "valor_desembolsado",
    "valor_historico_r": "valor_desembolsado",
    "valor_historico_em_r": "valor_desembolsado",
    "valor_historico_reais": "valor_desembolsado",
    "juros": "juros",
    "taxa_juros": "juros",
    "prazo_carencia_meses": "prazo_carencia",
    "prazo_carencia": "prazo_carencia",
    "prazo_de_carencia_meses": "prazo_carencia",
    "carencia_meses": "prazo_carencia",
    "prazo_amortizacao_meses": "prazo_amortizacao",
    "prazo_amortizacao": "prazo_amortizacao",
    "prazo_de_amortizacao_meses": "prazo_amortizacao",
    "amortizacao_meses": "prazo_amortizacao",
    "instituicao_financeira_credenciada": "agente",
    "instituicao_financeira": "agente",
    "agente_financeiro": "agente",
    "agente": "agente",
    "custo_financeiro": "custo_financeiro",
    "custo_financeiro_da_operacao": "custo_financeiro",
    "encargo_financeiro": "custo_financeiro",
}


def _normalize_nome_coluna(name: object) -> str:
    """Normaliza nome de coluna: minúsculas, sem acentos/símbolos."""
    import unicodedata

    text = str(name).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    cleaned: list[str] = []
    for ch in text:
        if ch.isalnum():
            cleaned.append(ch)
        else:
            cleaned.append("_")
    text = "".join(cleaned)
    while "__" in text:
        text = text.replace("__", "_")
    return text.strip("_")


def _mapear_colunas_contratos(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str]]:
    """Renomeia colunas ContAgil/BNDES → canônicas (exato + normalizado)."""
    rename: dict[str, str] = {}
    used_targets: set[str] = set()

    for k, v in {**EXCEL_COLUMNS, **CSV_COLUMNS}.items():
        if k in df.columns and v not in used_targets:
            rename[k] = v
            used_targets.add(v)

    for col in df.columns:
        if col in rename:
            continue
        key = _normalize_nome_coluna(col)
        target = NORM_COLUMN_ALIASES.get(key)
        if target is None:
            # ContAgil: "Valor Desembolsado R$ (*)" → valor_desembolsado_r_*
            if key.startswith("valor_desembolsado"):
                target = "valor_desembolsado"
            elif key.startswith("valor_da_operacao"):
                target = "valor_desembolsado"
            elif key.startswith("valor_historico"):
                # BNDES INDIRETAS ContAgil: Valor histórico / Valor Histórico R$ (*)
                target = "valor_desembolsado"
            elif "prazo" in key and "carencia" in key:
                target = "prazo_carencia"
            elif "prazo" in key and "amortizacao" in key:
                target = "prazo_amortizacao"
            elif "instituicao" in key and "financeira" in key:
                target = "agente"
            elif key.startswith("custo_financeiro") or key.startswith("encargo_financeiro"):
                target = "custo_financeiro"
        if target is not None and target not in used_targets:
            rename[col] = target
            used_targets.add(target)

    return df.rename(columns=rename), rename


def limpar_valor(series: pd.Series) -> pd.Series:
    """Converte BR (1.234,56), US (1234.56) ou já numérico."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = (
        series.astype(str)
        .str.replace("R$", "", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
    )

    def _one(v: str):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        text = str(v).strip()
        if not text or text.lower() in {"nan", "none"}:
            return np.nan
        if "," in text and "." in text:
            text = text.replace(".", "").replace(",", ".")
        elif "," in text:
            text = text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return np.nan

    return pd.to_numeric(s.map(_one), errors="coerce")


def parse_datas(series: pd.Series) -> pd.Series:
    """Parseia ISO (YYYY-MM-DD) ou BR (DD/MM/YYYY) sem misturar."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    s = series.astype(str).str.strip()
    iso_mask = s.str.match(r"^\d{4}-\d{2}-\d{2}", na=False).fillna(False)

    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if iso_mask.any():
        out.loc[iso_mask] = pd.to_datetime(
            s[iso_mask], errors="coerce", format="ISO8601"
        ).values
    if (~iso_mask).any():
        out.loc[~iso_mask] = pd.to_datetime(
            s[~iso_mask], dayfirst=True, errors="coerce"
        ).values
    return out


def meses_ate_impacto(data_fluxo: datetime, data_impacto: datetime = DATA_IMPACTO) -> int:
    """Meses de data_fluxo até data_impacto (pode ser negativo se fluxo for futuro)."""
    return (data_impacto.year - data_fluxo.year) * 12 + (data_impacto.month - data_fluxo.month)


def _stream_download(url: str, dest: Path, retries: int = 4) -> Path:
    """Baixa arquivo grande via HTTP streaming com retries (pandas URL buffer falha em ~1GB)."""
    import requests

    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".part")
    last_err: Exception | None = None

    for attempt in range(1, retries + 1):
        try:
            resume_from = tmp.stat().st_size if tmp.exists() else 0
            headers = {"Range": f"bytes={resume_from}-"} if resume_from else {}
            mode = "ab" if resume_from else "wb"

            print(
                f"Download tentativa {attempt}/{retries}"
                + (f" (retomando de {resume_from:,} bytes)" if resume_from else "")
                + "..."
            )
            with requests.get(url, stream=True, timeout=120, headers=headers) as resp:
                if resp.status_code not in (200, 206):
                    resp.raise_for_status()
                if resp.status_code == 200 and resume_from:
                    mode = "wb"
                    resume_from = 0

                total = resp.headers.get("Content-Length")
                expected = (
                    int(total) + resume_from
                    if total and resp.status_code == 206
                    else (int(total) if total else None)
                )

                downloaded = resume_from
                with tmp.open(mode) as f:
                    for chunk in resp.iter_content(chunk_size=8 * 1024 * 1024):
                        if not chunk:
                            continue
                        f.write(chunk)
                        downloaded += len(chunk)
                        if expected and downloaded % (64 * 1024 * 1024) < 8 * 1024 * 1024:
                            pct = 100.0 * downloaded / expected
                            print(f"  baixados {downloaded:,}/{expected:,} ({pct:.1f}%)")

            if expected is not None and tmp.stat().st_size < expected:
                raise IOError(
                    f"Download incompleto: {tmp.stat().st_size:,} < {expected:,} bytes"
                )

            tmp.replace(dest)
            print(f"CSV bruto salvo: {dest} ({dest.stat().st_size:,} bytes)")
            return dest
        except Exception as exc:  # noqa: BLE001 — retries em rede
            last_err = exc
            print(f"  falha: {exc}")
            time.sleep(min(2**attempt, 32))

    raise RuntimeError(f"Falha ao baixar {url}: {last_err}")


def download_and_filter_csv(
    url: str = BNDES_CSV_URL,
    start: str = "2009-01-01",
    end: str = "2010-12-31",
    dest: Path = FILTERED_CSV,
) -> Path:
    """Baixa o CSV aberto do BNDES em streaming e grava só o período pedido."""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Baixando e filtrando {start} .. {end} ...")
    print(f"URL: {url}")

    raw_path = RAW_DIR / "operacoes-indiretas-automaticas.csv"
    if not raw_path.exists() or raw_path.stat().st_size < 100_000_000:
        _stream_download(url, raw_path)
    else:
        print(f"Usando cache local: {raw_path} ({raw_path.stat().st_size:,} bytes)")

    encoding = "utf-8"
    try:
        with raw_path.open("r", encoding="utf-8") as f:
            f.read(2048)
    except UnicodeDecodeError:
        encoding = "cp1252"

    reader = pd.read_csv(
        raw_path,
        sep=";",
        encoding=encoding,
        dtype=str,
        chunksize=100_000,
        low_memory=False,
    )

    start_ts = pd.Timestamp(start)
    end_ts = pd.Timestamp(end)
    parts: list[pd.DataFrame] = []
    total_rows = 0
    kept_rows = 0

    for i, chunk in enumerate(reader, start=1):
        total_rows += len(chunk)
        dates = pd.to_datetime(chunk["data_da_contratacao"], errors="coerce")
        mask = (dates >= start_ts) & (dates <= end_ts)
        filtered = chunk.loc[mask].copy()
        kept_rows += len(filtered)
        if not filtered.empty:
            parts.append(filtered)
        if i % 5 == 0:
            print(f"  chunks={i:,} lidas={total_rows:,} mantidas={kept_rows:,}")

    if not parts:
        raise RuntimeError("Nenhum contrato encontrado no período solicitado.")

    df = pd.concat(parts, ignore_index=True)
    dest.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(dest, index=False, sep=";", encoding="utf-8")
    print(f"Arquivo filtrado salvo: {dest} ({len(df):,} contratos)")
    return dest


def _excel_tem_colunas_contratos(df: pd.DataFrame) -> bool:
    """True se o DataFrame já tem as colunas ContAgil / portal / BNDES."""
    mapped, _ = _mapear_colunas_contratos(df)
    hits = sum(
        1
        for k in (
            "data_contratacao",
            "valor_desembolsado",
            "juros",
            "prazo_carencia",
            "prazo_amortizacao",
        )
        if k in mapped.columns
    )
    return hits >= 3


def normalizar_colunas(df: pd.DataFrame, *, preparar: bool = True) -> pd.DataFrame:
    """Normaliza colunas de contratos ContAgil / BNDES INDIRETAS.

    Mapeia headers em português/CSV (com acentos, aliases e variações) para
    nomes canônicos (``data_contratacao``, ``valor_desembolsado``, etc.).

    Usada pelo script ContAgil WinPython após ``read_excel`` — a ausência
    desta função gerava ``NameError: name 'normalizar_colunas' is not defined``.
    """
    mapped, rename = _mapear_colunas_contratos(df)
    if rename:
        print(f"    Colunas mapeadas: {rename}")
    if not preparar:
        return mapped
    return _prepare_contracts(mapped)


def load_from_excel(
    path: Path,
    sheet_name: str | int = "operacoes_indiretas_automaticas",
    header: int | None = None,
) -> pd.DataFrame:
    """Carrega Excel ContAgil / portal / BNDES INDIRETAS.

    - header=None (default): tenta header=0..8 (ContAgil, títulos extras, portal).
    - header explícito: usa só esse valor.
    - Mapeia colunas por nome exato e por forma normalizada (acentos/espaços).
    """

    def _read(h: int) -> pd.DataFrame:
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=h)
        except ValueError:
            return pd.read_excel(path, sheet_name=0, header=h)

    if header is not None:
        candidatos = [header]
    else:
        # 0 = ContAgil/dados; 5 = portal; demais = planilhas com título/metadados
        candidatos = [0, 5, 1, 2, 3, 4, 6, 7, 8]

    df = None
    header_usado: int | None = None
    for h in candidatos:
        try:
            candidato = _read(h)
        except Exception:  # noqa: BLE001 — tenta próximo header
            continue
        if _excel_tem_colunas_contratos(candidato):
            df = candidato
            header_usado = h
            break
        if df is None:
            df = candidato
            header_usado = h

    if df is None:
        raise ValueError(f"Não foi possível ler Excel: {path}")

    df, rename = _mapear_colunas_contratos(df)
    if header_usado is not None and header_usado != 0:
        print(f"  Header Excel detectado na linha {header_usado}")
    if rename:
        print(f"  Colunas mapeadas: {rename}")
    return _prepare_contracts(df)


def load_from_csv(path: Path) -> pd.DataFrame:
    """Carrega CSV do portal de dados abertos."""
    df = pd.read_csv(path, sep=";", encoding="utf-8", dtype=str, low_memory=False)
    df, _ = _mapear_colunas_contratos(df)
    return _prepare_contracts(df)


def _normalizar_agente(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace(
        {
            "": AGENTE_NAO_INFORMADO,
            "nan": AGENTE_NAO_INFORMADO,
            "None": AGENTE_NAO_INFORMADO,
            "NaT": AGENTE_NAO_INFORMADO,
        }
    )
    return s.fillna(AGENTE_NAO_INFORMADO)


def _prepare_contracts(df: pd.DataFrame) -> pd.DataFrame:
    required = [
        "data_contratacao",
        "valor_desembolsado",
        "juros",
        "prazo_carencia",
        "prazo_amortizacao",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas ausentes: {missing}. Disponíveis: {list(df.columns)}")

    if "agente" in df.columns:
        agente = _normalizar_agente(df["agente"])
    else:
        agente = pd.Series([AGENTE_NAO_INFORMADO] * len(df), index=df.index)

    if "custo_financeiro" in df.columns:
        custo = df["custo_financeiro"].astype(str).fillna("")
    else:
        custo = pd.Series([""] * len(df), index=df.index)

    out = pd.DataFrame(
        {
            "data_contratacao": parse_datas(df["data_contratacao"]),
            "valor_desembolsado": limpar_valor(df["valor_desembolsado"]),
            "juros": limpar_valor(df["juros"]),
            "prazo_carencia": limpar_valor(df["prazo_carencia"]).fillna(0),
            "prazo_amortizacao": limpar_valor(df["prazo_amortizacao"]),
            "agente": agente.values,
            "custo_financeiro": custo.values,
        }
    )

    before = len(out)
    out = out.dropna(
        subset=["data_contratacao", "valor_desembolsado", "juros", "prazo_amortizacao"]
    )
    out = out[(out["valor_desembolsado"] > 0) & (out["prazo_amortizacao"] > 0)]
    out = out.reset_index(drop=True)
    out["contrato"] = out.index

    print(f"Contratos na entrada: {before:,}")
    print(f"Contratos válidos: {len(out):,}")
    print(f"Agentes distintos: {out['agente'].nunique():,}")
    return out


def _coluna_impacto(df: pd.DataFrame) -> str:
    if "impacto_fiscal" in df.columns:
        return "impacto_fiscal"
    if "impacto" in df.columns:
        return "impacto"
    raise ValueError("fluxos precisam de 'impacto_fiscal' ou 'impacto'")


def agregar_por_agente(df_fluxos: pd.DataFrame, contratos: pd.DataFrame) -> pd.DataFrame:
    """
    Resume fluxos por Agente Financeiro (Instituição Financeira Credenciada).

    Correção vs script com merge por índice: o CSV de fluxos é por parcela;
    o vínculo correto é contrato → agente (não left_index/right_index).
    """
    df = df_fluxos.copy()
    impacto_col = _coluna_impacto(df)

    if "Instituição Financeira" in df.columns:
        df["Agente"] = df["Instituição Financeira"].fillna(AGENTE_NAO_INFORMADO)
    elif "agente" in contratos.columns:
        mapa = contratos.set_index("contrato")["agente"]
        df["Agente"] = df["contrato"].map(mapa).fillna(AGENTE_NAO_INFORMADO)
    else:
        raise ValueError("contratos precisa da coluna 'agente' ou fluxos de 'Instituição Financeira'")

    resumo = (
        df.groupby("Agente", dropna=False)
        .agg(
            qtd_contratos=("contrato", "nunique"),
            subsidio=("subsidio", "sum"),
            impacto=(impacto_col, "sum"),
        )
        .round(2)
        .reset_index()
    )
    resumo.columns = [
        "Agente",
        "Qtd Contratos",
        "Total Subsídio (R$)",
        "Impacto Fiscal 2026 (R$)",
    ]
    return resumo.sort_values("Total Subsídio (R$)", ascending=False).reset_index(drop=True)


def resumo_from_agent_agg(agent_agg: dict) -> pd.DataFrame:
    """Converte acumulador interno {agente: {contratos, subsidio, impacto}} em DataFrame."""
    if not agent_agg:
        return pd.DataFrame(
            columns=[
                "Agente",
                "Qtd Contratos",
                "Total Subsídio (R$)",
                "Impacto Fiscal 2026 (R$)",
            ]
        )

    rows = [
        {
            "Agente": agente,
            "Qtd Contratos": vals["contratos"],
            "Total Subsídio (R$)": round(vals["subsidio"], 2),
            "Impacto Fiscal 2026 (R$)": round(vals["impacto"], 2),
        }
        for agente, vals in agent_agg.items()
    ]
    resumo = pd.DataFrame(rows)
    return resumo.sort_values("Total Subsídio (R$)", ascending=False).reset_index(drop=True)


def gerar_fluxos_contrato(
    data_contr: pd.Timestamp,
    valor: float,
    taxa_juros_aa: float,
    carencia: int,
    n: int,
    contrato_id: int,
    instituicao: str = AGENTE_NAO_INFORMADO,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
    custo_financeiro: str | None = None,
    juros_pct: float | None = None,
) -> list[dict]:
    """
    Gera fluxos detalhados de UM contrato (carência + amortização).

    Correção vs script ContAgil com bug:
      O original fazia `data = contr+(carencia+p)` E `em_carencia = p <= carencia`
      no loop `p=1..n`, o que zera amortização nas primeiras parcelas pós-carência
      e deixa saldo residual. Aqui o cronograma cobre carência+n meses.

    Lógica corrigida (dual balance):
      - saldo_fiscal: só principal (base do subsídio)
      - saldo_contrato: principal + juros do contrato
      - taxa via ``taxa_contrato_efetiva`` (TJLP/TLP / TAXA FIXA)
    """
    if n <= 0 or valor <= 0:
        return []

    amort_mensal = valor / n
    saldo_fiscal = valor
    saldo_contrato = valor

    if juros_pct is not None or custo_financeiro:
        pct = float(juros_pct) if juros_pct is not None else float(taxa_juros_aa) * 100.0
        taxa_contrato_mensal = taxa_contrato_efetiva(custo_financeiro, pct)
    else:
        # Compat testes: taxa_juros_aa já em decimal a.a. (TAXA FIXA)
        taxa_contrato_mensal = taxa_mensal_composta(taxa_juros_aa)

    taxa_selic_mensal = taxa_mensal_composta(selic_aa)
    spread = (1.0 + (taxa_selic_mensal - taxa_contrato_mensal)) ** n
    fluxos: list[dict] = []

    # Dia 15 como no ContAgil (evita deslocamento de fim de mês)
    try:
        data_base = data_contr.replace(day=15)
    except ValueError:
        data_base = data_contr

    total_meses = carencia + n
    for p in range(1, total_meses + 1):
        data_fluxo = data_base + relativedelta(months=p - 1)
        em_carencia = p <= carencia
        amort = 0.0 if em_carencia else amort_mensal

        # Subsídio sobre saldo fiscal ANTES da amortização do mês
        subsidio = saldo_fiscal * (taxa_selic_mensal - taxa_contrato_mensal)

        if selic_serie is not None:
            impacto = selic_serie.capitalizar(
                subsidio, data_fluxo.to_pydatetime(), data_impacto
            )
        else:
            meses = meses_ate_impacto(data_fluxo.to_pydatetime(), data_impacto)
            impacto = round(subsidio * ((1.0 + taxa_selic_mensal) ** meses), 2)

        fluxos.append(
            {
                "contrato": contrato_id,
                "Instituição Financeira": instituicao,
                "mes": p,
                "data_fluxo": data_fluxo.date(),
                "saldo_fiscal": round(saldo_fiscal, 2),
                "saldo_contrato": round(saldo_contrato, 2),
                "saldo": round(saldo_fiscal, 2),  # alias compat
                "amortizacao": round(amort, 2),
                "taxa_selic_mensal": round(taxa_selic_mensal, 8),
                # ContAgil: taxa do contrato só na 1ª parcela do cronograma
                "taxa_contrato_mensal": (
                    round(taxa_contrato_mensal, 8) if p == 1 else None
                ),
                "spread": round(spread, 6),
                "subsidio": round(subsidio, 2),
                "impacto_fiscal": impacto,
                "em_carencia": em_carencia,
            }
        )

        # Atualização dos saldos
        if not em_carencia:
            saldo_fiscal -= amort
            saldo_contrato = (saldo_contrato - amort) * (1.0 + taxa_contrato_mensal)
        else:
            saldo_contrato = saldo_contrato * (1.0 + taxa_contrato_mensal)

        if saldo_fiscal <= 1e-9:
            break

    return fluxos


def gerar_fluxos_diarios_contrato(
    data_contr: pd.Timestamp,
    valor: float,
    taxa_juros_aa: float,
    carencia: int,
    n: int,
    contrato_id: int,
    instituicao: str = AGENTE_NAO_INFORMADO,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
    custo_financeiro: str | None = None,
    juros_pct: float | None = None,
) -> list[dict]:
    """Expande o cronograma mensal em linhas dia a dia (entre parcelas ContAgil).

    Em cada período mensal (data da parcela → véspera da próxima):
      - saldo fiscal constante (SAC ContAgil)
      - amortização só no dia da parcela
      - subsídio diário = saldo_fiscal × (SELIC_d − taxa_contrato_d)
      - impacto_fiscal capitalizado na data da parcela (col D / FATOR_30_06_2026)
    """
    mensais = gerar_fluxos_contrato(
        data_contr=data_contr,
        valor=valor,
        taxa_juros_aa=taxa_juros_aa,
        carencia=carencia,
        n=n,
        contrato_id=contrato_id,
        instituicao=instituicao,
        selic_aa=selic_aa,
        data_impacto=data_impacto,
        selic_serie=selic_serie,
        custo_financeiro=custo_financeiro,
        juros_pct=juros_pct,
    )
    if not mensais:
        return []

    # Taxa mensal efetiva (1ª parcela); diária equivalente composta em 30 dias
    taxa_contrato_mensal = float(mensais[0]["taxa_contrato_mensal"] or 0.0)
    taxa_contrato_diaria = (1.0 + taxa_contrato_mensal) ** (1.0 / 30.0) - 1.0
    taxa_selic_mensal = taxa_mensal_composta(selic_aa)
    taxa_selic_diaria = taxa_diaria_composta(selic_aa)
    diarios: list[dict] = []

    for i, parcela in enumerate(mensais):
        data_ini = pd.Timestamp(parcela["data_fluxo"])
        if i + 1 < len(mensais):
            data_fim = pd.Timestamp(mensais[i + 1]["data_fluxo"]) - timedelta(days=1)
        else:
            data_fim = data_ini + relativedelta(months=1) - timedelta(days=1)

        saldo_fiscal = float(parcela["saldo_fiscal"])
        em_carencia = bool(parcela["em_carencia"])
        amort_parcela = float(parcela["amortizacao"])
        spread = float(parcela["spread"])
        mes = int(parcela["mes"])

        dia = data_ini
        while dia <= data_fim:
            amort = amort_parcela if dia == data_ini else 0.0
            subsidio = saldo_fiscal * (taxa_selic_diaria - taxa_contrato_diaria)

            if selic_serie is not None:
                impacto = selic_serie.capitalizar(
                    subsidio, dia.to_pydatetime(), data_impacto
                )
            else:
                meses = meses_ate_impacto(dia.to_pydatetime(), data_impacto)
                impacto = round(subsidio * ((1.0 + taxa_selic_mensal) ** meses), 2)

            diarios.append(
                {
                    "contrato": contrato_id,
                    "Instituição Financeira": instituicao,
                    "mes": mes,
                    "data_fluxo": dia.date(),
                    "saldo_fiscal": round(saldo_fiscal, 2),
                    "saldo": round(saldo_fiscal, 2),  # alias compat
                    "amortizacao": round(amort, 2),
                    "taxa_selic_diaria": round(taxa_selic_diaria, 10),
                    "taxa_contrato_diaria": round(taxa_contrato_diaria, 10),
                    "taxa_selic_mensal": round(taxa_selic_mensal, 8),
                    "taxa_contrato_mensal": round(taxa_contrato_mensal, 8),
                    "spread": round(spread, 6),
                    "subsidio": round(subsidio, 4),
                    "impacto_fiscal": impacto,
                    "em_carencia": em_carencia,
                    "dia_parcela": dia == data_ini,
                }
            )
            dia += timedelta(days=1)

    return diarios


def salvar_fluxos_diarios(
    fluxos_diarios: list[dict],
    path: Path | None = None,
) -> Path:
    """Grava a tabela dia a dia em Excel (e CSV espelho se muito grande)."""
    out = Path(path) if path is not None else OUTPUT_DIR / "fluxos_diarios_detalhados.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(fluxos_diarios)

    # Limite prático do Excel (~1.048.576 linhas)
    excel_limit = 1_000_000
    if len(df) > excel_limit:
        csv_path = out.with_suffix(".csv")
        df.to_csv(csv_path, index=False)
        amostra = df.head(excel_limit)
        amostra.to_excel(out, index=False)
        print(
            f"⚠️  Fluxos diários: {len(df):,} linhas > limite Excel; "
            f"CSV completo em {csv_path} e amostra Excel em {out}"
        )
    else:
        df.to_excel(out, index=False)
        print(f"✅ Fluxos diários: {out} ({len(df):,} linhas)")
    return out


_OPS_COL_MARKERS = {
    "data_contratacao",
    "valor_desembolsado",
    "juros",
    "prazo_amortizacao",
    "data_da_contratacao",
    "valor_desembolsado_reais",
    "prazo_amortizacao_meses",
    "Data da contratação",
    "Valor Desembolsado R$ (*)",
    "Juros",
    "Prazo - Amortização (meses)",
    "Custo financeiro",
    "Instituição Financeira Credenciada",
}


def _parece_dataframe_operacoes(df: pd.DataFrame) -> bool:
    """True se o DataFrame parece massa de operações (não SELIC STP)."""
    cols = {str(c) for c in df.columns}
    if cols & _OPS_COL_MARKERS:
        return True
    norm = {_normalize_nome_coluna(c) for c in df.columns}
    return bool(
        norm
        & {
            "data_da_contratacao",
            "data_contratacao",
            "valor_desembolsado_reais",
            "valor_desembolsado_r",
            "juros",
            "prazo_amortizacao_meses",
            "prazo_amortizacao",
            "custo_financeiro",
        }
    )


def _parece_dataframe_selic(df: pd.DataFrame) -> bool:
    """True se parece série de fatores SELIC ContAgil/Bacen."""
    if _parece_dataframe_operacoes(df):
        return False
    cols_lower = {str(c).strip().lower() for c in df.columns}
    if "fator_acumulado" in cols_lower or "fator" in cols_lower:
        return True
    # Layout STP ContAgil: várias colunas, sem campos de contrato
    return df.shape[1] >= 5


def _as_contratos(df: pd.DataFrame) -> pd.DataFrame:
    """Aceita contratos já preparados ou planilha/CSV brutos ContAgil."""
    if {"data_contratacao", "valor_desembolsado", "juros", "prazo_amortizacao"}.issubset(
        df.columns
    ):
        out = df.copy()
        if "contrato" not in out.columns:
            out = out.reset_index(drop=True)
            out["contrato"] = out.index
        if "agente" not in out.columns:
            out["agente"] = AGENTE_NAO_INFORMADO
        if "custo_financeiro" not in out.columns:
            out["custo_financeiro"] = ""
        if "prazo_carencia" not in out.columns:
            out["prazo_carencia"] = 0
        return out

    prepared, _ = _mapear_colunas_contratos(df)
    return _prepare_contracts(prepared)


def _instituicao_de_original(
    df_original: pd.DataFrame | None, idx: int, fallback: str
) -> str:
    """Espelha ContAgil: instituição a partir de df_original.iloc[idx]."""
    if df_original is None or idx >= len(df_original):
        return fallback
    row = df_original.iloc[idx]
    for col in (
        "Instituição Financeira Credenciada",
        "Instituicao Financeira Credenciada",
        "instituicao_financeira_credenciada",
        "agente",
    ):
        if col in df_original.columns:
            val = row[col]
            if pd.notna(val) and str(val).strip():
                return str(val).strip()
    # tenta por nome normalizado
    for col in df_original.columns:
        if _normalize_nome_coluna(col) in {
            "instituicao_financeira_credenciada",
            "agente",
        }:
            val = row[col]
            if pd.notna(val) and str(val).strip():
                return str(val).strip()
    return fallback


def _resolver_segundo_arg(
    segundo: float | SelicSerie | pd.DataFrame,
    selic_serie: SelicSerie | None,
) -> tuple[float, SelicSerie | None, pd.DataFrame | None]:
    """Resolve 2º argumento ContAgil: SELIC, série, taxa ou df_original.

    Aceita o rascunho ContAgil ``gerar_fluxos(df, df)`` (df_original para
    Instituição Financeira) e também ``gerar_fluxos(df, selic_df)``.
    """
    if isinstance(segundo, SelicSerie):
        return TAXA_SELIC_ANUAL, segundo, None
    if isinstance(segundo, pd.DataFrame):
        if _parece_dataframe_selic(segundo):
            return (
                TAXA_SELIC_ANUAL,
                SelicSerie.from_dataframe(segundo, origem="dataframe"),
                None,
            )
        if _parece_dataframe_operacoes(segundo):
            # ContAgil paste: gerar_fluxos(df, df_original)
            return TAXA_SELIC_ANUAL, selic_serie, segundo
        # Ambíguo: assume SELIC só se tiver ≥5 colunas (layout STP)
        if segundo.shape[1] >= 5:
            return (
                TAXA_SELIC_ANUAL,
                SelicSerie.from_dataframe(segundo, origem="dataframe"),
                None,
            )
        return TAXA_SELIC_ANUAL, selic_serie, segundo
    return float(segundo), selic_serie, None


def _progresso_intervalo(n: int, progress_every: int | None) -> int | None:
    """Define intervalo de log de progresso (None = sem log intermediário)."""
    if progress_every is not None:
        return max(1, int(progress_every)) if progress_every > 0 else None
    if n >= 50_000:
        return 2_000
    if n >= 5_000:
        return 1_000
    if n >= 500:
        return 100
    return None


def gerar_fluxos(
    df: pd.DataFrame,
    selic_aa: float | SelicSerie | pd.DataFrame = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
    fluxo_diario: bool = False,
    saida_diario: Path | str | None = None,
    progress_every: int | None = None,
    quiet: bool = False,
) -> pd.DataFrame:
    """Gera DataFrame completo de fluxos (parcelas mensais).

    Compatível com o rascunho ContAgil (lógica corrigida)::

        df_fluxos = gerar_fluxos(df, df)          # df_original = instituições
        df_fluxos = gerar_fluxos(df, selic_df)    # fatores STP/Bacen
        df_fluxos = gerar_fluxos(contratos)       # SELIC 14,5% composta

    Aceita planilha bruta (header=5, colunas em português) ou contratos
    já preparados. Mantém carência+n (corrige o bug ``p=1..n`` com offset).

    Com fluxo_diario=True, também gera a tabela dia a dia em
    output/fluxos_diarios_detalhados.xlsx (ou saida_diario).

    Para massas grandes (>~5k contratos), prefira ``gerar_e_gravar_fluxos``
    (grava CSV em lotes e evita estourar memória).
    """
    selic_aa, selic_serie, df_original = _resolver_segundo_arg(selic_aa, selic_serie)
    contratos = _as_contratos(df)
    n = len(contratos)
    step = None if quiet else _progresso_intervalo(n, progress_every)
    if not quiet:
        print(f"🚀 Gerando fluxos com lógica corrigida... ({n:,} contratos)")
        sys.stdout.flush()

    records: list[dict] = []
    fluxos_diarios: list[dict] = []
    skipped = 0
    t0 = time.time()

    for pos, row in enumerate(contratos.itertuples(index=False), start=1):
        try:
            data_contr = pd.Timestamp(row.data_contratacao)
            if pd.isna(data_contr):
                skipped += 1
                continue
            fallback_agente = str(
                getattr(row, "agente", AGENTE_NAO_INFORMADO) or AGENTE_NAO_INFORMADO
            )
            instituicao = _instituicao_de_original(
                df_original, pos - 1, fallback_agente
            )
            custo = getattr(row, "custo_financeiro", "")
            juros_pct = float(row.juros)
            kwargs = dict(
                data_contr=data_contr,
                valor=float(row.valor_desembolsado),
                taxa_juros_aa=juros_pct / 100.0,
                carencia=int(float(row.prazo_carencia or 0)),
                n=int(float(row.prazo_amortizacao)),
                contrato_id=int(row.contrato),
                instituicao=instituicao,
                selic_aa=selic_aa,
                data_impacto=data_impacto,
                selic_serie=selic_serie,
                custo_financeiro=custo,
                juros_pct=juros_pct,
            )
            records.extend(gerar_fluxos_contrato(**kwargs))
            if fluxo_diario:
                fluxos_diarios.extend(gerar_fluxos_diarios_contrato(**kwargs))
        except (TypeError, ValueError, OverflowError):
            skipped += 1
            continue

        if step is not None and (pos % step == 0 or pos == n):
            elapsed = max(time.time() - t0, 1e-6)
            rate = pos / elapsed
            eta = (n - pos) / rate if rate > 0 else 0.0
            print(
                f"  progresso {pos:,}/{n:,} ({100.0 * pos / n:.1f}%) "
                f"| {rate:,.0f} contr/s | ETA ~{eta / 60:.1f} min "
                f"| parcelas={len(records):,}"
            )
            sys.stdout.flush()

    if skipped and not quiet:
        print(f"Contratos ignorados por erro: {skipped:,}")

    if fluxo_diario:
        path = Path(saida_diario) if saida_diario is not None else None
        salvar_fluxos_diarios(fluxos_diarios, path)

    return pd.DataFrame(records)


def gerar_e_gravar_fluxos(
    df: pd.DataFrame,
    selic_aa: float | SelicSerie | pd.DataFrame = TAXA_SELIC_ANUAL,
    *,
    saida_xlsx: Path | str,
    lote: int = 2_000,
    excel_max_linhas: int = 1_000_000,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
) -> dict:
    """Gera fluxos em lotes e grava CSV completo (+ Excel amostra se >1M linhas).

    Evita manter dezenas de milhões de parcelas em memória (massa BNDES ~100k–1M
    contratos). Retorna estatísticas do processamento.
    """
    saida_xlsx = Path(saida_xlsx)
    saida_xlsx.parent.mkdir(parents=True, exist_ok=True)
    csv_path = saida_xlsx.with_suffix(".csv")

    contratos = _as_contratos(df)
    n = len(contratos)
    lote = max(1, int(lote))
    print(
        f"🚀 Gerando fluxos com lógica corrigida... "
        f"({n:,} contratos, lote={lote:,}, grava CSV em streaming)"
    )
    sys.stdout.flush()

    if csv_path.exists():
        csv_path.unlink()

    total_parcelas = 0
    skipped = 0
    amostra: list[pd.DataFrame] = []
    amostra_linhas = 0
    t0 = time.time()
    header = True

    for start in range(0, n, lote):
        chunk = contratos.iloc[start : start + lote]
        try:
            fluxos = gerar_fluxos(
                chunk,
                selic_aa,
                data_impacto=data_impacto,
                selic_serie=selic_serie,
                quiet=True,
            )
        except Exception:  # noqa: BLE001 — lote isolado não derruba a massa
            skipped += len(chunk)
            fluxos = pd.DataFrame()

        if not fluxos.empty:
            fluxos.to_csv(csv_path, mode="a", index=False, header=header)
            header = False
            total_parcelas += len(fluxos)
            if amostra_linhas < excel_max_linhas:
                falta = excel_max_linhas - amostra_linhas
                amostra.append(fluxos.head(falta))
                amostra_linhas += min(len(fluxos), falta)

        feitos = min(start + lote, n)
        elapsed = max(time.time() - t0, 1e-6)
        rate = feitos / elapsed
        eta = (n - feitos) / rate if rate > 0 else 0.0
        print(
            f"  lote {start:,}-{feitos:,}/{n:,} ({100.0 * feitos / n:.1f}%) "
            f"| {rate:,.0f} contr/s | ETA ~{eta / 60:.1f} min "
            f"| parcelas={total_parcelas:,}"
        )
        sys.stdout.flush()

    if total_parcelas == 0:
        raise ValueError("Nenhuma parcela gerada (todos os contratos falharam?).")

    if amostra:
        amostra_df = pd.concat(amostra, ignore_index=True)
    else:
        amostra_df = pd.read_csv(csv_path, nrows=excel_max_linhas)

    amostra_df.to_excel(saida_xlsx, index=False)

    stats = {
        "contratos": n,
        "parcelas": total_parcelas,
        "skipped": skipped,
        "csv": str(csv_path),
        "xlsx": str(saida_xlsx),
        "xlsx_linhas": len(amostra_df),
        "segundos": round(time.time() - t0, 1),
    }
    if total_parcelas > excel_max_linhas:
        print(
            f"    → CSV completo: {csv_path} ({total_parcelas:,} parcelas)"
        )
        print(
            f"    → Excel (amostra {len(amostra_df):,}): {saida_xlsx}"
        )
    else:
        # Massa cabe no Excel: CSV auxiliar pode ser removido pelo usuário
        print(f"    → Salvo: {saida_xlsx} ({total_parcelas:,} parcelas)")
        print(f"    → CSV: {csv_path}")
    if skipped:
        print(f"    Contratos/lotes com erro: {skipped:,}")
    return stats


def processar_em_lotes(
    df: pd.DataFrame,
    csv_path: Path,
    lote: int = 2000,
    selic_aa: float = TAXA_SELIC_ANUAL,
    data_impacto: datetime = DATA_IMPACTO,
    selic_serie: SelicSerie | None = None,
    fluxo_diario: bool = False,
    saida_diario: Path | None = None,
) -> dict:
    """Processa em lotes, grava CSV detalhado e acumula estatísticas (+ por agente)."""
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    if csv_path.exists():
        csv_path.unlink()

    diario_xlsx = (
        Path(saida_diario)
        if saida_diario is not None
        else OUTPUT_DIR / "fluxos_diarios_detalhados.xlsx"
    )
    diario_csv = diario_xlsx.with_suffix(".csv")
    if fluxo_diario:
        diario_csv.parent.mkdir(parents=True, exist_ok=True)
        if diario_csv.exists():
            diario_csv.unlink()
        if diario_xlsx.exists():
            diario_xlsx.unlink()

    monthly: dict[str, float] = {}
    agent_agg: dict[str, dict[str, float]] = defaultdict(
        lambda: {"contratos": 0, "subsidio": 0.0, "impacto": 0.0}
    )
    total_impacto = 0.0
    total_subsidio = 0.0
    total_amort = 0.0
    n_parcelas = 0
    n_contratos_ok = 0
    n_em_carencia = 0
    n_dias = 0
    wrote_header = False
    wrote_diario_header = False

    n = len(df)
    if selic_serie is not None:
        modo = f"fator ContAgil ({selic_serie.origem})"
    else:
        modo = "SELIC composta constante"
    print(f"Processando {n:,} contratos (lote={lote:,}, impacto={modo})...")

    for start in range(0, n, lote):
        chunk = df.iloc[start : start + lote]
        records: list[dict] = []
        diarios_lote: list[dict] = []

        for row in chunk.itertuples(index=False):
            try:
                data_contr = pd.Timestamp(row.data_contratacao)
                if pd.isna(data_contr):
                    continue
                instituicao = str(
                    getattr(row, "agente", AGENTE_NAO_INFORMADO) or AGENTE_NAO_INFORMADO
                )
                custo = getattr(row, "custo_financeiro", "")
                juros_pct = float(row.juros)
                kwargs = dict(
                    data_contr=data_contr,
                    valor=float(row.valor_desembolsado),
                    taxa_juros_aa=juros_pct / 100.0,
                    carencia=int(float(row.prazo_carencia or 0)),
                    n=int(float(row.prazo_amortizacao)),
                    contrato_id=int(row.contrato),
                    instituicao=instituicao,
                    selic_aa=selic_aa,
                    data_impacto=data_impacto,
                    selic_serie=selic_serie,
                    custo_financeiro=custo,
                    juros_pct=juros_pct,
                )
                fluxos = gerar_fluxos_contrato(**kwargs)
                if fluxos:
                    n_contratos_ok += 1
                    records.extend(fluxos)
                    agent_agg[instituicao]["contratos"] += 1
                    agent_agg[instituicao]["subsidio"] += float(
                        sum(f["subsidio"] for f in fluxos)
                    )
                    agent_agg[instituicao]["impacto"] += float(
                        sum(f["impacto_fiscal"] for f in fluxos)
                    )
                    if fluxo_diario:
                        diarios_lote.extend(gerar_fluxos_diarios_contrato(**kwargs))
            except (TypeError, ValueError, OverflowError):
                continue

        if fluxo_diario and diarios_lote:
            pd.DataFrame(diarios_lote).to_csv(
                diario_csv, mode="a", index=False, header=not wrote_diario_header
            )
            wrote_diario_header = True
            n_dias += len(diarios_lote)

        if not records:
            print(f"  lote {start:,}-{start + len(chunk):,}: 0 fluxos")
            continue

        fluxos_df = pd.DataFrame(records)
        n_parcelas += len(fluxos_df)
        total_impacto += float(fluxos_df["impacto_fiscal"].sum())
        total_subsidio += float(fluxos_df["subsidio"].sum())
        total_amort += float(fluxos_df["amortizacao"].sum())
        n_em_carencia += int(fluxos_df["em_carencia"].sum())

        keys = pd.to_datetime(fluxos_df["data_fluxo"]).dt.to_period("M").astype(str)
        for k, v in fluxos_df.groupby(keys, sort=False)["impacto_fiscal"].sum().items():
            monthly[k] = monthly.get(k, 0.0) + float(v)

        fluxos_df.to_csv(csv_path, mode="a", index=False, header=not wrote_header)
        wrote_header = True
        print(
            f"  lote {start:,}-{start + len(chunk):,} → +{len(fluxos_df):,} "
            f"(acum {n_parcelas:,})"
        )

    diario_path: str | None = None
    if fluxo_diario and wrote_diario_header and diario_csv.exists():
        # Monta o Excel a partir do CSV acumulado (lista de dicts → DataFrame)
        df_diario = pd.read_csv(diario_csv)
        salvar_fluxos_diarios(df_diario.to_dict(orient="records"), diario_xlsx)
        diario_path = str(diario_xlsx)

    return {
        "n_contratos_entrada": n,
        "n_contratos_ok": n_contratos_ok,
        "n_parcelas": n_parcelas,
        "n_parcelas_em_carencia": n_em_carencia,
        "n_dias": n_dias,
        "total_amortizacao": round(total_amort, 2),
        "total_subsidio": round(total_subsidio, 2),
        "total_impacto_fiscal_2026": round(total_impacto, 2),
        "n_agentes": len(agent_agg),
        "metodologia_impacto": modo,
        "taxa_selic_anual": selic_aa,
        "taxa_selic_mensal_composta": round(taxa_mensal_composta(selic_aa), 8),
        "fluxos_diarios": diario_path,
        "monthly": monthly,
        "por_agente": dict(agent_agg),
    }


def salvar_resumo_por_agente(resumo: pd.DataFrame, stem: str = "resumo_por_agente") -> tuple[Path, Path]:
    """Grava CSV + Excel do ranking por agente financeiro."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = OUTPUT_DIR / f"{stem}.csv"
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    resumo.to_csv(csv_path, index=False)
    resumo.to_excel(xlsx_path, index=False, sheet_name="Por_Agente")
    return csv_path, xlsx_path


def salvar_excel_resumo(
    stats: dict,
    xlsx_path: Path,
    sample_csv: Path | None = None,
    sample_rows: int = 50_000,
    resumo_agente: pd.DataFrame | None = None,
) -> None:
    """Excel legível: resumo + por agente + impacto mensal + amostra de parcelas."""
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    resumo = pd.DataFrame(
        [
            {"Indicador": "Taxa SELIC anual (config)", "Valor": TAXA_SELIC_ANUAL},
            {"Indicador": "Data de impacto", "Valor": DATA_IMPACTO.date().isoformat()},
            {"Indicador": "Contratos na entrada", "Valor": stats["n_contratos_entrada"]},
            {"Indicador": "Contratos processados", "Valor": stats["n_contratos_ok"]},
            {"Indicador": "Parcelas geradas", "Valor": stats["n_parcelas"]},
            {
                "Indicador": "Parcelas em carência",
                "Valor": stats["n_parcelas_em_carencia"],
            },
            {"Indicador": "Soma Amortização", "Valor": stats["total_amortizacao"]},
            {"Indicador": "Soma Subsídio (nominal)", "Valor": stats["total_subsidio"]},
            {
                "Indicador": "Soma Impacto Fiscal 2026",
                "Valor": stats["total_impacto_fiscal_2026"],
            },
            {"Indicador": "Agentes financeiros", "Valor": stats.get("n_agentes", 0)},
            {
                "Indicador": "Metodologia impacto",
                "Valor": stats.get("metodologia_impacto", "SELIC composta constante"),
            },
            {
                "Indicador": "Arquivo detalhado",
                "Valor": str(xlsx_path.with_suffix(".csv").name),
            },
        ]
    )

    mensal = (
        pd.DataFrame(
            [
                {"Ano_Mes": k, "Impacto_Fiscal_2026": round(v, 2)}
                for k, v in sorted(stats["monthly"].items())
            ]
        )
        if stats["monthly"]
        else pd.DataFrame(columns=["Ano_Mes", "Impacto_Fiscal_2026"])
    )

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        if resumo_agente is not None and not resumo_agente.empty:
            resumo_agente.to_excel(writer, sheet_name="Por_Agente", index=False)
        mensal.to_excel(writer, sheet_name="Impacto_Mensal", index=False)
        if sample_csv and sample_csv.exists():
            sample = pd.read_csv(sample_csv, nrows=sample_rows)
            sample.to_excel(writer, sheet_name="Amostra_Parcelas", index=False)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--excel",
        type=Path,
        help="Excel local (header=5), ex.: operacoes_indiretas_automaticas_2009-01-01_ate_2010-12-31.xlsx",
    )
    p.add_argument(
        "--input",
        type=Path,
        help="CSV filtrado já baixado (sep=';).",
    )
    p.add_argument(
        "--download",
        action="store_true",
        help="Baixa/filtra CSV aberto do BNDES (2009–2010 por padrão).",
    )
    p.add_argument("--start", default="2009-01-01")
    p.add_argument("--end", default="2010-12-31")
    p.add_argument("--lote", type=int, default=2000)
    p.add_argument(
        "--max-contratos",
        type=int,
        default=None,
        help="Limita contratos (útil para testes).",
    )
    p.add_argument(
        "--stem",
        default="fluxos_completos_final",
        help="Prefixo dos arquivos de saída.",
    )
    p.add_argument(
        "--arquivo-selic",
        type=Path,
        default=None,
        help=(
            "Excel STP ContAgil/Bacen (col A=data, col D=fator acumulado). "
            "Se omitido, tenta auto-descobrir (ContAgil path, data/STP*.xlsx, "
            "data/selic_fatores_bacen.xlsx)."
        ),
    )
    p.add_argument(
        "--baixar-selic",
        action="store_true",
        help=(
            "Baixa SELIC diária do Bacen (SGS 11), monta fatores acumulados ContAgil "
            "e salva em data/selic_fatores_bacen.xlsx. Usado se nenhum STP for encontrado."
        ),
    )
    p.add_argument(
        "--sem-selic-fatores",
        action="store_true",
        help="Força impacto por SELIC composta constante (ignora STP/Bacen).",
    )
    p.add_argument(
        "--fluxo-diario",
        action="store_true",
        help="Gera tabela detalhada dia a dia (output/fluxos_diarios_detalhados.xlsx).",
    )
    p.add_argument(
        "--saida-diario",
        type=Path,
        default=None,
        help="Caminho do Excel dia a dia (default: output/fluxos_diarios_detalhados.xlsx).",
    )
    return p.parse_args(argv)


def carregar_selic_serie(args: argparse.Namespace) -> SelicSerie | None:
    """Resolve série de fatores: STP → cache → Bacen (padrão ContAgil)."""
    if args.sem_selic_fatores:
        print("Impacto fiscal: SELIC composta constante (14,5% a.a.)")
        return None

    path = resolver_arquivo_selic(args.arquivo_selic)
    if path is not None:
        print(f"Lendo SELIC fatores: {path}")
        serie = SelicSerie.from_excel(path)
        print(f"  {len(serie.datas):,} pontos ({serie.origem})")
        return serie

    if args.arquivo_selic is not None:
        raise FileNotFoundError(f"Arquivo SELIC não encontrado: {args.arquivo_selic}")

    # ContAgil exige fatores acumulados: baixa Bacen se não houver STP local.
    # --baixar-selic permanece como alias explícito (comportamento padrão).
    motivo = (
        "flag --baixar-selic"
        if args.baixar_selic
        else "nenhum STP ContAgil encontrado (auto Bacen)"
    )
    print(f"Baixando fatores SELIC via Bacen SGS 11 ({motivo})...")
    serie = SelicSerie.from_bacen(cache_path=SELIC_BACEN_CACHE)
    print(f"  {len(serie.datas):,} pontos ({serie.origem})")
    return serie


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    selic_serie = carregar_selic_serie(args)

    if args.excel:
        print(f"Lendo Excel: {args.excel}")
        df = load_from_excel(args.excel)
    elif args.input:
        print(f"Lendo CSV: {args.input}")
        df = load_from_csv(args.input)
    else:
        excel_auto = resolver_excel_operacoes()
        if excel_auto is not None and not args.download:
            print(f"Lendo Excel (auto): {excel_auto}")
            df = load_from_excel(excel_auto)
        elif FILTERED_CSV.exists() and not args.download:
            print(f"Lendo cache: {FILTERED_CSV}")
            df = load_from_csv(FILTERED_CSV)
        else:
            path = download_and_filter_csv(start=args.start, end=args.end)
            df = load_from_csv(path)

    if args.max_contratos is not None:
        df = df.head(args.max_contratos).copy()
        df["contrato"] = df.index
        print(f"Limitado a {len(df):,} contratos (--max-contratos)")

    csv_path = OUTPUT_DIR / f"{args.stem}.csv"
    xlsx_path = OUTPUT_DIR / f"{args.stem}.xlsx"
    stats_path = OUTPUT_DIR / f"{args.stem}_stats.json"
    saida_diario = args.saida_diario or (OUTPUT_DIR / "fluxos_diarios_detalhados.xlsx")

    stats = processar_em_lotes(
        df,
        csv_path,
        lote=args.lote,
        selic_serie=selic_serie,
        fluxo_diario=args.fluxo_diario,
        saida_diario=saida_diario if args.fluxo_diario else None,
    )
    resumo_agente = resumo_from_agent_agg(stats.get("por_agente", {}))
    agente_csv, agente_xlsx = salvar_resumo_por_agente(resumo_agente)

    printable = {
        k: v for k, v in stats.items() if k not in {"monthly", "por_agente"}
    }
    print(json.dumps(printable, indent=2))
    print("\nResumo por Agente Financeiro (top 20):")
    print(resumo_agente.head(20).to_string(index=False))

    salvar_excel_resumo(
        stats, xlsx_path, sample_csv=csv_path, resumo_agente=resumo_agente
    )
    with stats_path.open("w", encoding="utf-8") as f:
        json.dump(printable, f, indent=2)

    print(f"✅ CSV detalhado: {csv_path}")
    print(f"✅ Excel resumo:  {xlsx_path}")
    print(f"✅ Resumo agente: {agente_csv}")
    print(f"✅ Resumo agente: {agente_xlsx}")
    print(f"✅ Stats JSON:    {stats_path}")
    if args.fluxo_diario and stats.get("fluxos_diarios"):
        print(f"✅ Fluxos diários: {stats['fluxos_diarios']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
