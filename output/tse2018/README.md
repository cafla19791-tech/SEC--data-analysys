# Boletins de Urna 2018 (Presidente)

Uma linha por urna/seção, com série (`NR_URNA_EFETIVADA`), modelo e votos.

Fonte: ZIPs oficiais `BWEB_{1t|2t}_{UF}_*.zip` (TSE / Archive.org), processados por `scripts/baixar_boletins_urna.py`.

No ContAgil (RFB), baixe o CSV já consolidado:

```bat
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 1
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 2
```
