"""Imprime em PDF todos os relatórios de evolução SGS já gerados.

Uso:
  python3 scripts/imprimir_evolucoes_pdf.py
  python3 scripts/imprimir_evolucoes_pdf.py --output-dir output
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.evolucao_balanca_reservas import exportar_pdf_relatorio
from scripts.evolucao_fatores_base_monetaria import (
    agregar_anual as ag_base,
    cabecalhos_estoque as cab_base_est,
    cabecalhos_fases as cab_base_fases,
    cabecalhos_var as cab_base_var,
    carregar_series as car_base,
    fases_historicas as fases_base,
    gerar_pdf as pdf_base,
    linhas_estoque as lin_base_est,
    linhas_fases as lin_base_fases,
    linhas_var as lin_base_var,
    mes_referencia as mes_base,
)
from scripts.evolucao_fatores_dbgg import (
    agregar_anual as ag_dbgg,
    cabecalhos_estoque as cab_dbgg_est,
    cabecalhos_fatores as cab_dbgg_fat,
    cabecalhos_fases as cab_dbgg_fases,
    carregar_series as car_dbgg,
    fases_historicas as fases_dbgg,
    gerar_pdf as pdf_dbgg,
    linhas_estoque as lin_dbgg_est,
    linhas_fatores as lin_dbgg_fat,
    linhas_fases as lin_dbgg_fases,
    mes_referencia as mes_dbgg,
)

DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

# Relatórios já impressos em PNG (demais evoluções desta série).
PACOTES_IMAGEM = (
    {
        "pdf": "evolucao_agregados_monetarios_1995_2026.pdf",
        "titulo": "Agregados monetários M1–M4 (1995–2026)",
        "notas": [
            "Saldo de fim de período em R$ bilhões. Metodologia nova SGS 27791/27810/27813/27815 desde 2001.",
            "1995–2000 usa as séries descontinuadas 1827/1837/1840/1843. 2026* = último mês publicado.",
        ],
        "imagens": (
            "tabela_agregados_anual_1995_2026.png",
            "tabela_agregados_fases_1995_2026.png",
            "grafico_agregados_monetarios_1995_2026.png",
            "grafico_composicao_agregados_1995_2026.png",
            "grafico_share_agregados_m4_1995_2026.png",
        ),
    },
    {
        "pdf": "evolucao_balanca_reservas_1995_2025.pdf",
        "titulo": "Balança comercial e reservas internacionais (1995–2025)",
        "notas": [
            "Fluxos BPM6 em US$ bilhões (SGS 22707/22708/22709) e estoque de reservas (3546).",
        ],
        "imagens": (
            "tabela_anual_1995_2025.png",
            "tabela_fases_1995_2025.png",
            "grafico_saldo_comercial_1995_2025.png",
            "grafico_exportacoes_importacoes_1995_2025.png",
            "grafico_reservas_1995_2025.png",
            "grafico_saldo_e_reservas_1995_2025.png",
        ),
    },
    {
        "pdf": "evolucao_recursos_livres_direcionados_2002_2026.pdf",
        "titulo": "Recursos livres e direcionados no SFN (2002–2026)",
        "notas": [
            "Saldo de crédito do SFN. 2026* = último mês publicado.",
        ],
        "imagens": (
            "tabela_recursos_anual_2002_2026.png",
            "tabela_recursos_fases_2002_2026.png",
            "grafico_recursos_livres_direcionados_2002_2026.png",
            "grafico_composicao_credito_sfn_2002_2026.png",
            "grafico_credito_pib_livres_direcionados_2002_2026.png",
            "grafico_share_livres_direcionados_2002_2026.png",
        ),
    },
    {
        "pdf": "evolucao_consignado_cdc_cartao_2002_2016.pdf",
        "titulo": "Consignado, CDC e cartão (2002–2016)",
        "notas": [
            "Carteiras de crédito livre a pessoas físicas até a quebra de metodologia de 2016/2017.",
        ],
        "imagens": (
            "tabela_consignado_cdc_cartao_anual.png",
            "tabela_consignado_cdc_cartao_fases.png",
            "grafico_consignado_cdc_cartao_2007_2016.png",
            "grafico_composicao_consignado_cdc_cartao_2007_2016.png",
            "grafico_share_consignado_cdc_cartao_2007_2016.png",
        ),
    },
    {
        "pdf": "evolucao_top5_credito_livres_2002_2026.pdf",
        "titulo": "Crédito das 5 maiores instituições (2002–2026)",
        "notas": [
            "Totais de crédito e proxy de recursos livres das cinco maiores IFs (IF.data).",
        ],
        "imagens": (
            "tabela_top5_carteira_2002_2026.png",
            "tabela_top5_livres_2014_2026.png",
            "grafico_top5_carteira_credito_2002_2026.png",
            "grafico_top5_soma_carteira_livres_2002_2026.png",
        ),
    },
)


def _pdf_base(cache_dir: Path, output_dir: Path) -> Path:
    series = car_base(cache_dir=cache_dir, baixar=False)
    anual = ag_base(series)
    mes = mes_base(series, int(anual["ano"].max()))
    return pdf_base(anual, output_dir, mes)


def _pdf_dbgg(cache_dir: Path, output_dir: Path) -> Path:
    series = car_dbgg(cache_dir=cache_dir, baixar=False)
    anual = ag_dbgg(series)
    mes = mes_dbgg(series, int(anual["ano"].max()))
    return pdf_dbgg(anual, output_dir, mes)


def _pdf_pacote(spec: dict, output_dir: Path) -> Path | None:
    imagens = [output_dir / nome for nome in spec["imagens"] if (output_dir / nome).exists()]
    if not imagens:
        return None
    return exportar_pdf_relatorio(
        output_dir / spec["pdf"],
        spec["titulo"],
        spec["notas"],
        tabelas=None,
        imagens=imagens,
    )


def _pdf_completo(cache_dir: Path, output_dir: Path) -> Path:
    """Capa + tabelas vetoriais da base/DBGG + PNGs de todos os relatórios."""
    sb = car_base(cache_dir=cache_dir, baixar=False)
    ab = ag_base(sb)
    mb = mes_base(sb, int(ab["ano"].max()))
    sd = car_dbgg(cache_dir=cache_dir, baixar=False)
    ad = ag_dbgg(sd)
    md = mes_dbgg(sd, int(ad["ano"].max()))
    notas = [
        f"Compilação gerada em {datetime.now().strftime('%d/%m/%Y')}.",
        "Inclui: fatores da base monetária, fatores da DBGG, agregados M1–M4, "
        "balança e reservas, recursos livres/direcionados, consignado/CDC/cartão "
        "e crédito das 5 maiores IFs.",
        "Valores em R$ bilhões ou % do PIB, conforme a tabela. Grade contínua.",
        f"Último ponto da base: {mb or 'n/d'}. Último ponto da DBGG/Selic: {md or 'n/d'}.",
    ]
    tabelas = [
        (
            "Base monetária — fases (R$ bilhões)",
            cab_base_fases(),
            lin_base_fases(fases_base(ab)),
            [0.10, 0.26, 0.10, 0.10, 0.11, 0.11, 0.11, 0.11],
        ),
        (
            "Base monetária — variação anual (R$ bilhões)",
            cab_base_var(),
            lin_base_var(ab, mb),
            None,
        ),
        (
            "Base monetária — estoque de fim de período (R$ bilhões)",
            cab_base_est(),
            lin_base_est(ab, mb),
            None,
        ),
        (
            "DBGG — fases",
            cab_dbgg_fases(),
            lin_dbgg_fases(fases_dbgg(ad)),
            [0.10, 0.24, 0.11, 0.10, 0.12, 0.11, 0.11, 0.11],
        ),
        (
            "DBGG — fatores anuais (primário/juros/NFSP/efeito PIB em % PIB; ajustes em R$ bi)",
            cab_dbgg_fat(),
            lin_dbgg_fat(ad, md),
            None,
        ),
        (
            "DBGG — estoque",
            cab_dbgg_est(),
            lin_dbgg_est(ad, md),
            None,
        ),
    ]
    imagens: list[Path] = []
    for spec in PACOTES_IMAGEM:
        imagens.extend(output_dir / n for n in spec["imagens"] if (output_dir / n).exists())
    for nome in (
        "grafico_fatores_base_estoque_1995_2026.png",
        "grafico_fatores_base_variacao_1995_2026.png",
        "grafico_dbgg_dlsp_pib_1995_2026.png",
        "grafico_dbgg_primario_juros_1995_2026.png",
        "grafico_dbgg_selic_cambio_1995_2026.png",
    ):
        p = output_dir / nome
        if p.exists():
            imagens.append(p)
    return exportar_pdf_relatorio(
        output_dir / "evolucoes_sgs_completo.pdf",
        "Evoluções SGS — compilado (1995–hoje)",
        notas,
        tabelas=tabelas,
        imagens=imagens,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cache-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args(argv)

    caminhos: list[Path] = []
    caminhos.append(_pdf_base(args.cache_dir, args.output_dir))
    caminhos.append(_pdf_dbgg(args.cache_dir, args.output_dir))
    for spec in PACOTES_IMAGEM:
        gerado = _pdf_pacote(spec, args.output_dir)
        if gerado is not None:
            caminhos.append(gerado)
    caminhos.append(_pdf_completo(args.cache_dir, args.output_dir))
    for p in caminhos:
        print(f"  {p} ({p.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
