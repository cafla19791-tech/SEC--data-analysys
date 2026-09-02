# Boletins de Urna 2014 (Presidente)

Uma linha por urna/seção, com série (`NR_URNA_EFETIVADA`), modelo e votos.

Fonte: ZIPs oficiais `bweb_{1t|2t}_{UF}_*.zip` (TXT posicional do TSE), processados por `scripts/baixar_boletins_urna.py`.

No ContAgil (RFB), baixe o CSV já consolidado:

```bat
python baixar_boletins_urna.py --somente-resultado-github --ano 2014 --turno 1
python baixar_boletins_urna.py --somente-resultado-github --ano 2014 --turno 2
```
