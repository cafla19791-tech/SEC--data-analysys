#!/usr/bin/env python3
"""
Resumo por Agente Financeiro — Versão Web (Streamlit).

  streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
RESUMO_CSV = OUTPUT / "resumo_por_agente.csv"
RESUMO_XLSX = OUTPUT / "resumo_por_agente.xlsx"
FLUXOS_XLSX = OUTPUT / "fluxos_completos_corrigido.xlsx"

st.set_page_config(
    page_title="Resumo por Agente Financeiro",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Visual: teal institucional (evita tema roxo/creme genérico)
st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700&family=Source+Serif+4:opsz,wght@8..60,600&display=swap');
      html, body, [class*="css"] {
        font-family: "DM Sans", sans-serif;
      }
      .stApp {
        background:
          radial-gradient(ellipse 80% 50% at 10% 0%, #d7ebe6 0%, transparent 55%),
          radial-gradient(ellipse 60% 40% at 90% 10%, #e8f0f4 0%, transparent 50%),
          linear-gradient(180deg, #f3f7f6 0%, #eef2f1 100%);
      }
      h1, h2, h3 {
        font-family: "Source Serif 4", Georgia, serif !important;
        color: #0b3d3a !important;
      }
      div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.55);
        border: 1px solid rgba(11,61,58,0.12);
        padding: 0.75rem 1rem;
        border-radius: 4px;
      }
      .brand-mark {
        font-family: "Source Serif 4", Georgia, serif;
        font-size: 2.4rem;
        font-weight: 600;
        color: #0b3d3a;
        letter-spacing: -0.02em;
        margin: 0 0 0.25rem 0;
      }
      .brand-sub {
        color: #3d5c58;
        font-size: 1.05rem;
        margin: 0 0 1.5rem 0;
        max-width: 42rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def carregar_resumo() -> pd.DataFrame:
    if RESUMO_CSV.exists():
        return pd.read_csv(RESUMO_CSV)
    if RESUMO_XLSX.exists():
        return pd.read_excel(RESUMO_XLSX)
    if FLUXOS_XLSX.exists():
        xl = pd.ExcelFile(FLUXOS_XLSX)
        if "Por_Agente" in xl.sheet_names:
            return pd.read_excel(xl, "Por_Agente")
    return pd.DataFrame()


def fmt_brl(valor: float) -> str:
    return (
        f"R$ {valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def main() -> None:
    st.markdown('<p class="brand-mark">Resumo por Agente Financeiro</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="brand-sub">Subsídio nominal e impacto fiscal a valor de 30/06/2026 '
        "(SELIC 14,5% a.a.) das operações indiretas automáticas do BNDES.</p>",
        unsafe_allow_html=True,
    )

    resumo = carregar_resumo()
    if resumo.empty:
        st.warning(
            "Nenhum resumo encontrado. Gere os dados com:\n\n"
            "`python scripts/gerar_fluxos.py --download`\n\n"
            "ou, para uma amostra rápida:\n\n"
            "`python scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv`"
        )
        st.stop()

    # Normaliza nomes de colunas
    colmap = {
        "Agente": "Agente",
        "Qtd Contratos": "Qtd Contratos",
        "Total Subsídio (R$)": "Total Subsídio (R$)",
        "Impacto Fiscal 2026 (R$)": "Impacto Fiscal 2026 (R$)",
    }
    missing = [c for c in colmap if c not in resumo.columns]
    if missing:
        st.error(f"Colunas ausentes no resumo: {missing}")
        st.stop()

    with st.sidebar:
        st.header("Filtros")
        top_n = st.slider("Top N agentes", min_value=5, max_value=min(100, len(resumo)), value=min(20, len(resumo)))
        busca = st.text_input("Buscar agente", placeholder="ex.: BANCO DO BRASIL")
        ordenar = st.selectbox(
            "Ordenar por",
            ["Total Subsídio (R$)", "Impacto Fiscal 2026 (R$)", "Qtd Contratos"],
        )

    view = resumo.copy()
    if busca.strip():
        view = view[view["Agente"].str.contains(busca.strip(), case=False, na=False)]
    view = view.sort_values(ordenar, ascending=False).head(top_n)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Agentes", f"{resumo['Agente'].nunique():,}")
    c2.metric("Contratos", f"{int(resumo['Qtd Contratos'].sum()):,}")
    c3.metric("Subsídio total", fmt_brl(float(resumo["Total Subsídio (R$)"].sum())))
    c4.metric("Impacto fiscal 2026", fmt_brl(float(resumo["Impacto Fiscal 2026 (R$)"].sum())))

    st.subheader(f"Ranking — top {len(view)}")
    chart_df = view.set_index("Agente")[["Total Subsídio (R$)", "Impacto Fiscal 2026 (R$)"]]
    st.bar_chart(chart_df, height=360)

    display = view.copy()
    display["Total Subsídio (R$)"] = display["Total Subsídio (R$)"].map(fmt_brl)
    display["Impacto Fiscal 2026 (R$)"] = display["Impacto Fiscal 2026 (R$)"].map(fmt_brl)
    display["Qtd Contratos"] = display["Qtd Contratos"].map(lambda x: f"{int(x):,}".replace(",", "."))
    st.dataframe(display, use_container_width=True, hide_index=True)

    st.download_button(
        "Baixar CSV completo",
        data=resumo.to_csv(index=False).encode("utf-8"),
        file_name="resumo_por_agente.csv",
        mime="text/csv",
    )

    with st.expander("Metodologia"):
        st.markdown(
            """
            - **Agente** = Instituição Financeira Credenciada  
            - **Subsídio** mensal = saldo × (SELIC/12 − juros/12)  
            - **Impacto fiscal 2026** = subsídio × (1 + SELIC/12)^(meses até 30/06/2026)  
            - Agregação por contrato → agente (não por índice de linha do CSV de parcelas)
            """
        )


if __name__ == "__main__":
    main()
