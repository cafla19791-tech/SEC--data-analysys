"""Análise de dados financeiros de empresas via SEC EDGAR."""

from .metrics import FinancialMetrics, compute_metrics
from .report import format_report
from .sec_client import SecClient

__all__ = ["SecClient", "FinancialMetrics", "compute_metrics", "format_report"]
