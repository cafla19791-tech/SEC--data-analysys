"""Extração e análise do lucro líquido da Petrobras a partir da SEC EDGAR."""

from .net_income import extract_net_income, build_table

__all__ = ["extract_net_income", "build_table"]
