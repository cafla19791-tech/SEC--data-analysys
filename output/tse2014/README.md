# Eleições 2014 — Presidente por urna

Boletins de Urna oficiais (`bweb_*`), com `NR_URNA_EFETIVADA`.

Fontes no Drive:

- UFs em ZIP: https://drive.google.com/drive/folders/1EpLQeAQlwvTR9bCY7XDtBxjqIM-5Kiea
- 1º turno TO (TXT): https://drive.google.com/drive/folders/1OVHfD0n1AnGQoaqks5Y5mF1KtgjAjY3A

IDs em `data/tse_catalog/drive_bweb_2014.json`.

## 1º turno (`urnas_1t_presidente.csv.gz`)

- 428.894 seções · 428.852 com série · 27 UFs + Exterior
- Dilma 43.267.668 · Aécio 34.897.211 · Marina 22.176.619
- **Totais iguais ao TSE** (inclui Tocantins e Exterior)

## 2º turno (`urnas_2t_presidente.csv.gz`)

- 393.493 seções · 393.451 com série
- Dilma 48.502.249 · Aécio 49.305.163
- **Faltam CE e MA** nas pastas do Drive (TSE oficial: Dilma 54.501.118 · Aécio 51.041.155)

```bat
python baixar_boletins_urna.py --ano 2014 --turno 1 --somente-processar
python planilha_resultados_presidente.py
```
