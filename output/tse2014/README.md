# Eleições 2014 — Presidente por urna

O TSE publica Boletins de Urna por UF (TXT posicional **com** `NR_URNA_EFETIVADA`).
O script `scripts/baixar_boletins_urna.py --ano 2014 --turno 1` já lê esse layout.

Desta rede o CDN do TSE devolve **HTTP 403** e o Internet Archive **não guardou** os ZIPs
de 2014 (nem o `votacao_secao_2014`). Por isso o CSV nacional ainda não está neste repositório.

URLs oficiais (catálogo em `data/tse_catalog/boletins_urna_urls.json`):

- 1º turno: `bweb_1t_{UF}_14102014*.zip`
- 2º turno: `bweb_2t_{UF}_28102014*.zip`

Se os ZIPs forem salvos em `dados\tse2014\raw` (Edge/rede que alcance o TSE):

```bat
python baixar_boletins_urna.py --ano 2014 --turno 1 --somente-processar
python baixar_boletins_urna.py --ano 2014 --turno 2 --somente-processar
```

Totais oficiais para conferência: Dilma 43.267.668 / Aécio 34.897.211 / Marina 22.176.619 (1º);
Dilma 54.501.118 / Aécio 51.041.155 (2º).
