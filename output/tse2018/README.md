# Eleições 2018 — Presidente por urna / seção

Duas camadas, ambas oficiais do TSE:

## Nacional por seção (28 UFs, sem número de série)

Arquivo `votacao_secao_2018_BR` (Dados Abertos). Uma linha por zona/seção.
Não traz `NR_URNA_EFETIVADA` (uma seção = uma urna na prática).

| Turno | Arquivo | Seções | Totais (conferem com o TSE) |
|---|---|---:|---|
| 1º | `secoes_1t_presidente.csv.gz` | 454.450 | Bolsonaro 49.277.010 · Haddad 31.342.051 · Ciro 13.344.371 · Amoêdo 2.679.745 |
| 2º | `secoes_2t_presidente.csv.gz` | 454.448 | Bolsonaro 57.797.847 · Haddad 47.040.906 |

## Por urna com série e modelo (parcial)

ZIPs `BWEB_*` no Archive.org. Só as UFs que o Archive capturou (TSE CDN = 403).

- 1º turno (`urnas_1t_presidente.csv.gz`): AC AL AM AP BA CE DF ES GO MA MG RR (164.221 urnas)
- 2º turno (`urnas_2t_presidente.csv.gz`): AC AL AM AP BA CE DF ES GO MA (114.883 urnas)

Em 2018 Bolsonaro era o **17** (não o 22). Haddad **13**. Amoêdo **30**. João Goulart Filho **54**.

```bat
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 1
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 2
```
