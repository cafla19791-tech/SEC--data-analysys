# Boletins de Urna 2022 (Presidente)

CSV nacional: uma linha por urna/seção, com número de série (`NR_URNA_EFETIVADA`), modelo (UE2009–UE2020) e votos.

- 2º turno: `urnas_2t_presidente.csv.gz` — ZIPs `bweb_2t_{UF}_311020221535.zip`
- 1º turno: `urnas_1t_presidente.csv.gz` — ZIPs `bweb_1t_{UF}_051020221321.zip`

Fonte: Dados Abertos do TSE / captura Archive.org, processados por `scripts/baixar_boletins_urna.py` (1º turno) e `scripts/baixar_boletins_urna_2022.py` (2º turno).

Totais 1º turno (conferem com o TSE):
- 472.028 urnas/seções com série preenchida
- Lula 57.259.504 · Bolsonaro 51.072.345 · Tebet 4.915.423 · Ciro 3.599.287

Totais 2º turno (conferem com o TSE):
- 472.028 urnas/seções com série preenchida
- Lula 60.345.999
- Bolsonaro 58.206.354

Discriminativo municipal (UE2020 vs modelos anteriores a 2020):
- `discriminativo_municipio_ue2020.xlsx`
- `discriminativo_municipio_ue2020.csv`

No ContAgil, com o CSV de urnas já baixado:

```bat
python discriminativo_urnas_municipio.py
```

Na RFB (TSE 403, Archive.org TLS quebrado), baixe estes arquivos pelo GitHub:

```bat
python baixar_boletins_urna_2022.py --somente-resultado-github
python baixar_boletins_urna.py --somente-resultado-github --ano 2022 --turno 1
```

ou `baixar_resultado_urna_github.bat`.
