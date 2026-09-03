# Eleições 2014 — Presidente por urna

Boletins de Urna oficiais (`bweb_*`), com `NR_URNA_EFETIVADA`. Fonte: pasta do Drive
https://drive.google.com/drive/folders/1EpLQeAQlwvTR9bCY7XDtBxjqIM-5Kiea
(IDs em `data/tse_catalog/drive_bweb_2014.json`).

## 2º turno (`urnas_2t_presidente.csv.gz`)

- 393.493 seções · 393.451 com série
- Dilma 48.502.249 · Aécio 49.305.163
- **Faltam CE e MA** nesta pasta do Drive (TSE oficial: Dilma 54.501.118 · Aécio 51.041.155)

## 1º turno (`urnas_1t_presidente.csv.gz`)

- 425.377 seções
- Dilma 42.899.237 · Aécio 34.694.028 · Marina 22.025.929
- **Falta TO** nesta pasta do Drive (TSE oficial: Dilma 43.267.668 · Aécio 34.897.211 · Marina 22.176.619)

```bat
python baixar_boletins_urna.py --ano 2014 --turno 2 --somente-processar
python planilha_resultados_presidente.py --ano 2014 --turno 2
```
