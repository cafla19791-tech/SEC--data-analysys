"""Gráficos de evolução do lucro líquido (matplotlib e plotly)."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots


def plot_matplotlib(df: pd.DataFrame, path: str | Path) -> Path:
    """Gráfico de barras + linha YoY (matplotlib)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    years = df["year"].astype(int)
    usd_bi = df["net_income_usd"] / 1_000_000_000
    brl_bi = df["net_income_brl"] / 1_000_000_000
    yoy = df["yoy_usd_pct"] * 100

    fig, axes = plt.subplots(2, 1, figsize=(11, 8), sharex=True, gridspec_kw={"height_ratios": [3, 1.2]})

    colors = ["#c0392b" if v < 0 else "#1f6f4a" for v in usd_bi]
    axes[0].bar(years - 0.18, usd_bi, width=0.36, color=colors, label="US$ bi", alpha=0.9)
    axes[0].bar(years + 0.18, brl_bi / 5, width=0.36, color="#2c3e50", label="R$ bi ÷ 5 (escala)", alpha=0.55)
    axes[0].axhline(0, color="#333", linewidth=0.8)
    axes[0].set_ylabel("Lucro líquido (US$ bilhões)")
    axes[0].set_title("Petrobras — Lucro líquido anual (SEC 20-F / IFRS)")
    axes[0].legend(loc="upper left")
    axes[0].grid(axis="y", linestyle="--", alpha=0.35)

    # Anotações dos valores USD
    for x, y in zip(years, usd_bi):
        axes[0].annotate(
            f"{y:.1f}",
            (x - 0.18, y),
            textcoords="offset points",
            xytext=(0, 6 if y >= 0 else -12),
            ha="center",
            fontsize=8,
        )

    axes[1].plot(years, yoy, marker="o", color="#0b3d5c", linewidth=2)
    axes[1].axhline(0, color="#333", linewidth=0.8)
    axes[1].set_ylabel("YoY USD (%)")
    axes[1].set_xlabel("Ano")
    axes[1].grid(True, linestyle="--", alpha=0.35)

    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def plot_plotly(df: pd.DataFrame, path: str | Path) -> Path:
    """Gráfico interativo dual-eixo (plotly)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    years = df["year"].astype(int)
    usd_bi = df["net_income_usd"] / 1_000_000_000
    brl_bi = df["net_income_brl"] / 1_000_000_000

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=years,
            y=usd_bi,
            name="Lucro líquido (US$ bi)",
            marker_color=["#c0392b" if v < 0 else "#1f6f4a" for v in usd_bi],
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=years,
            y=brl_bi,
            name="Lucro líquido (R$ bi)",
            mode="lines+markers",
            line=dict(color="#0b3d5c", width=3),
        ),
        secondary_y=True,
    )
    fig.update_layout(
        title="Petrobras — Evolução do Lucro Líquido (SEC EDGAR)",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=80),
    )
    fig.update_yaxes(title_text="US$ bilhões", secondary_y=False)
    fig.update_yaxes(title_text="R$ bilhões", secondary_y=True)
    fig.write_html(str(path), include_plotlyjs="cdn")
    return path


# --- Código sugerido (também disponível como módulo) ---
CHART_SNIPPET = '''
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from plotly.subplots import make_subplots

df = pd.read_csv("data/petrobras_net_income.csv")
years = df["year"]
usd_bi = df["net_income_usd"] / 1e9
brl_bi = df["net_income_brl"] / 1e9

# Matplotlib
fig, ax = plt.subplots(figsize=(10, 5))
colors = ["#c0392b" if v < 0 else "#1f6f4a" for v in usd_bi]
ax.bar(years, usd_bi, color=colors, label="US$ bi")
ax.plot(years, brl_bi / 5, color="#0b3d5c", marker="o", label="R$ bi ÷ 5")
ax.axhline(0, color="black", lw=0.8)
ax.set_title("Petrobras — Lucro líquido anual")
ax.set_xlabel("Ano")
ax.set_ylabel("US$ bilhões")
ax.legend()
ax.grid(axis="y", ls="--", alpha=0.4)
plt.tight_layout()
plt.savefig("reports/petrobras_net_income_chart.png", dpi=150)

# Plotly
fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(go.Bar(x=years, y=usd_bi, name="US$ bi"), secondary_y=False)
fig.add_trace(go.Scatter(x=years, y=brl_bi, name="R$ bi", mode="lines+markers"), secondary_y=True)
fig.update_layout(title="Petrobras — Lucro líquido", template="plotly_white")
fig.write_html("reports/petrobras_net_income_chart.html")
'''
