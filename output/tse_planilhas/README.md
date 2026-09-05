# Planilhas — Presidente por região, UF, município, zona e urna

Resultados oficiais do cargo de Presidente, 1º e 2º turnos, nas eleições
de **2014**, **2018** e **2022**. Cada arquivo `resultados_presidente_{ano}_{turno}t.xlsx`
tem as abas:

| Aba | Recorte |
|---|---|
| Leia-me | Fonte, totais e conferência com o TSE |
| Regiao | Norte, Nordeste, Centro-Oeste, Sudeste, Sul, Exterior (ZZ) |
| UF | Unidade da Federação |
| Municipio | Município (código TSE + nome) |
| Zona | Zona eleitoral dentro do município |
| Urna | Uma linha por urna/seção (número de série quando existir) |

O consolidado `resultados_agregados_regiao_uf_municipio_zona.xlsx` junta
região / UF / município / zona de todos os pleitos gerados (sem a aba Urna,
que fica nas planilhas por ano/turno).

## Cobertura

| Pleito | Fonte | Série da urna | Totais vs TSE |
|---|---|---|---|
| 2022 1º e 2º | Boletins de Urna (`bweb_*`) | Completa (`NR_URNA_EFETIVADA`) | Conferem |
| 2018 1º e 2º | `votacao_secao` nacional | Parcial (só UFs com BU no Archive.org) | Conferem |
| 2014 1º | Boletins de Urna (Drive, incl. TO) | Completa (27 UFs + Exterior) | Conferem |
| 2014 2º | Boletins de Urna (Drive, incl. MA e CE) | Completa (27 UFs + Exterior) | Conferem |

2014 veio das pastas do Drive com os ZIPs `bweb_*2014*` (número de série da urna)
e os TXT de Tocantins (1º turno), Maranhão e Ceará (2º turno). Os dois turnos
nacionais conferem com o TSE.

2018: uma seção = uma urna na prática. A coluna `NR_URNA_EFETIVADA` é preenchida
só nas UFs cujo BU o Archive.org capturou (1º: AC AL AM AP BA CE DF ES GO MA MG RR;
2º: AC AL AM AP BA CE DF ES GO MA).

## Como gerar (ContAgil / WinPython)

```bat
python baixar_boletins_urna.py --somente-resultado-github --ano 2022 --turno 1
python baixar_boletins_urna.py --somente-resultado-github --ano 2022 --turno 2
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 1
python baixar_boletins_urna.py --somente-resultado-github --ano 2018 --turno 2
python planilha_resultados_presidente.py
```

No repositório:

```bash
python3 scripts/planilha_resultados_presidente.py
python3 scripts/discriminativo_resultados_presidente.py
```

## Discriminativo 2014 × 2018 × 2022

`discriminativo_presidente_2014_2018_2022.xlsx` compara os três 2º turnos
no mesmo recorte político (lado PT = Dilma / Haddad / Lula; oposição =
Aécio / Bolsonaro / Bolsonaro), com abas Brasil, região, UF, município e
zona. Há também UF e região do 1º turno.

```bat
python discriminativo_resultados_presidente.py
```

## Totais oficiais usados na conferência

- 2022 1º: Lula 57.259.504 · Bolsonaro 51.072.345 · Tebet 4.915.423 · Ciro 3.599.287
- 2022 2º: Lula 60.345.999 · Bolsonaro 58.206.354
- 2018 1º: Bolsonaro 49.277.010 · Haddad 31.342.051 · Ciro 13.344.371 · Amoêdo 2.679.745
- 2018 2º: Bolsonaro 57.797.847 · Haddad 47.040.906
- 2014 1º: Dilma 43.267.668 · Aécio 34.897.211 · Marina 22.176.619 (OK)
- 2014 2º: Dilma 54.501.118 · Aécio 51.041.155 (OK)
