# Boletins de Urna 2022 (2º turno, Presidente)

CSV nacional: uma linha por urna/seção, com número de série (`NR_URNA_EFETIVADA`), modelo (UE2009–UE2020) e votos.

Fonte: ZIPs oficiais `bweb_2t_{UF}_311020221535.zip` (TSE / captura Archive.org 2022-11-08), processados por `scripts/baixar_boletins_urna_2022.py`.

Totais (conferem com o resultado oficial do 2º turno):
- 472.028 urnas/seções com série preenchida
- Lula 60.345.999
- Bolsonaro 58.206.354

Na RFB (TSE 403, Archive.org TLS quebrado), baixe estes arquivos pelo GitHub:

```bat
python baixar_boletins_urna_2022.py --somente-resultado-github
```

ou `baixar_resultado_urna_github.bat`.
