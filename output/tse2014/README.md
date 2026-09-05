# Eleições 2014 — Presidente por urna

Boletins de Urna oficiais (`bweb_*`), com `NR_URNA_EFETIVADA`.

Fontes no Drive:

- UFs em ZIP: https://drive.google.com/drive/folders/1EpLQeAQlwvTR9bCY7XDtBxjqIM-5Kiea
- 1º turno TO (TXT): https://drive.google.com/drive/folders/1OVHfD0n1AnGQoaqks5Y5mF1KtgjAjY3A
- 2º turno MA (TXT): https://drive.google.com/drive/folders/1wnxROCh2mUvFUtr0zuUoX9dc7WlFYh5k
- 2º turno CE (TXT): https://drive.google.com/file/d/1NxnSusxyROLC-lhoJ7Xd9aC5OoF3VB0S/view

IDs em `data/tse_catalog/drive_bweb_2014.json`.

## 1º turno (`urnas_1t_presidente.csv.gz`)

- 428.894 seções · 428.852 com série · 27 UFs + Exterior
- Dilma 43.267.668 · Aécio 34.897.211 · Marina 22.176.619
- **Totais iguais ao TSE**

## 2º turno (`urnas_2t_presidente.csv.gz`)

- 428.894 seções · 428.852 com série · 27 UFs + Exterior
- Dilma 54.501.118 · Aécio 51.041.155
- **Totais iguais ao TSE** (inclui Ceará e Maranhão)

```bat
python baixar_boletins_urna.py --somente-resultado-github --ano 2014 --turno 2
python planilha_resultados_presidente.py
```
