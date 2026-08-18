# Simulação: Conta Única e fatores da DBGG, 2003 a 2015

Reservas internacionais ficam de lado. A pergunta é outra: **como
evoluiriam os fatores condicionantes da dívida bruta do governo geral**
se três fluxos tivessem ido para a Conta Única do Tesouro, em vez de
sair como desembolso do BNDES ou como renúncia tributária.

Os três fluxos, de 2003 a 2015, somam **R$ 1.112,5 bi** correntes:

1. orçamento do BNDES em **operações indiretas** — R$ 658,4 bi;
2. orçamento do BNDES em **participações acionárias** (BNDESPAR, 2007–2015)
   — R$ 44,4 bi;
3. receita que a RFB teria lançado **sem** os benefícios de desenvolvimento
   regional (função + ZFM/ALC), **sem** o gasto tributário de inovação
   (Lei 10.973/2004 operacionalizada pela Lei do Bem 11.196/2005) e **sem**
   os benefícios a entidades imunes e isentas — R$ 409,8 bi.

Em 2015 a DBGG oficial (SGS 13761) era R$ 3.927,5 bi
(65,5% do PIB). No contrafactual com o
caixa usado para não emitir / resgatar DPF e o saldo capitalizado à
Selic efetiva, a DBGG cairia para R$ 2.049,4 bi
(34,2% do PIB) — cerca de
**31,3 p.p. do PIB** a menos. Sem capitalizar juros, o
corte é o estoque acumulado dos fluxos: R$ 1.112,5 bi, e a
DBGG ficaria em R$ 2.815,0 bi.

A Selic entre as quatro mais altas do BIS em cada ano do período não
entra como causa das reservas. Entra aqui só como **preço do estoque
poupado**: cada real que tivesse ficado na Conta Única e evitado DPF
deixava de carregar esse juro.

## Identidade que a simulação usa

A DBGG é **bruta**. Depósito na Conta Única (haver no Bacen) reduz a
**DLSP**, não a DBGG. Para o fator “emissões líquidas” da Nota de
Política Fiscal mudar, o caixa tem de **não emitir ou resgatar** título.

Identidade anual (R$):

```
Δ DBGG_cf = Δ DBGG_oficial + emissão_evitada + (− juros_evitados)
emissão_evitada_t = − fluxo_t
juros_evitados_t  = Selic_t × saldo_selic_{t−1}
saldo_selic_t     = saldo_selic_{t−1} × (1+Selic_t) + fluxo_t
```

O fluxo entra no **fim** do ano (convenção conservadora: menos juros no
próprio exercício). Não há efeito-atividade: a RFB também calcula gasto
tributário com base estática.

Dois caminhos, de propósito distintos:

* **CF estoque** — o Tesouro só deixa de emitir o fluxo do ano. A DBGG
  cai 1 a 1 com o acumulado. É o fator “emissões líquidas”.
* **CF Selic** — o estoque poupado deixa de pagar Selic. Soma o fator
  “juros nominais”. É o canal em que a Selic alta do período (e o
  ranking no BIS) multiplica o resultado.

## 1. Indiretas do BNDES

Fonte: *Desembolsos Mensais* do portal de dados abertos, campo
`forma_de_apoio = INDIRETA`. Finame, Automático, Finem indireto, Exim
indireto e Cartão BNDES entram. O CSV “indiretas e produto” **não**
serve: ele zera máquinas/serviços.

Soma 2003–2015: **R$ 658,4 bi**. Pico em 2013
(R$ 97,2 bi). Em 2010,
R$ 81,5 bi — o ano do gráfico TCU da base
monetária e do salto Tesouro→BNDES.

Isso é **orçamento/desembolso do Sistema BNDES**, não “recurso do
Tesouro”. Até 2007 a fonte dominante é FAT/PIS/próprios. As captações
Tesouro→BNDES (página oficial) só aparecem em 2008–2014 e somam
**R$ 440,8 bi** (incluindo R$ 24,7 bi da capitalização da
Petrobras em 2010). São fungíveis entre direta, indireta e renda
variável. Redirecionar *toda* a indireta de 2003–2007 à Conta Única
exige mudar a destinação legal do FAT/PIS, não só “deixar de emitir
DPF”.

## 2. Participações acionárias

Fonte: BNDESPAR, desembolsos via renda variável, apenas
`PARTICIPAÇÃO ACIONÁRIA`. Debêntures (R$ 17,9 bi em 2007–2015) e cotas
de fundo (R$ 3,0 bi) ficam de fora — não são capital acionário.

A base começa em **2007**. 2003–2006 não têm microdado comparável; a
simulação trata esses anos como zero e o total de R$ 44,4 bi é
portanto um **piso**. O salto de 2010 (R$ 25,4 bi)
é a capitalização da Petrobras / ofertas daquela janela.

## 3. Três famílias de renúncia

Lei 10.973/2004 é o marco da inovação; o item que a RFB mensura no DGT
é sobretudo a **Lei do Bem** (Lei 11.196/2005). Informática (Lei 8.248)
**não** entra. Desenvolvimento regional na acepção da pergunta inclui a
função RFB *e* ZFM/ALC — no DGT elas vêm em linhas separadas; a tabela
mostra a soma (“regional ampla”) e as duas partes no código.

Âncora oficial de 2015 (IFI NT 17, bases efetivas, R$ milhões):
Desenvolvimento Regional 5.899, ZFM/ALC 23.232, Imunes/isentas 19.505,
Pesquisa científica e inovação 3.392.

A série 2003–2014 **não** é o DGT item a item de cada ano (os
demonstrativos anuais por modalidade não estão reproduzidos no
repositório). É uma reconstrução: a participação de cada família no PIB
de 2015 é aplicada ao PIB de cada ano (SGS 4382). Inovação = 0 em
2003–2005 (Lei do Bem ainda não opera). Em 2015 usam-se os valores
exatos da IFI. Isso preserva a ordem de grandeza e **não inventa** um
DGT anual que não foi transcrito.

Cheque de consistência: a isenção patronal das filantrópicas no TCU
2006–2010 (R$ 3,8 a 6,4 bi) é um *subconjunto* previdenciário das
imunes/isentas, e fica abaixo da reconstrução — como deve.

DGT PLOA 2015 (projeção, não efetiva): regional 7.274, ZFM 27.812,
imunes 22.323, inovação 3.403. A simulação usa a base efetiva.

## Fatores condicionantes — série anual

Valores da tabela em **R$ bilhões** correntes, salvo % PIB.

| Ano | Indiretas | Participações | Regional ampla | Inovação | Imunes/isentas | Fluxo total | % PIB |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 2003 | 14,5 | — | 8,3 | 0,0 | 5,6 | 28,5 | 1,66 |
| 2004 | 19,9 | — | 9,5 | 0,0 | 6,4 | 35,7 | 1,83 |
| 2005 | 21,0 | — | 10,5 | 0,0 | 7,1 | 38,6 | 1,78 |
| 2006 | 25,3 | — | 11,7 | 1,4 | 7,8 | 46,2 | 1,92 |
| 2007 | 32,8 | 2,0 | 13,2 | 1,5 | 8,8 | 58,4 | 2,15 |
| 2008 | 42,8 | 7,9 | 15,1 | 1,8 | 10,1 | 77,7 | 2,50 |
| 2009 | 50,4 | 2,7 | 16,2 | 1,9 | 10,8 | 82,1 | 2,46 |
| 2010 | 81,5 | 25,4 | 18,9 | 2,2 | 12,6 | 140,7 | 3,62 |
| 2011 | 69,4 | 0,7 | 21,3 | 2,5 | 14,2 | 108,1 | 2,47 |
| 2012 | 67,1 | 1,6 | 23,4 | 2,7 | 15,7 | 110,5 | 2,29 |
| 2013 | 97,2 | 2,1 | 25,9 | 3,0 | 17,3 | 145,6 | 2,73 |
| 2014 | 85,5 | 1,3 | 28,1 | 3,3 | 18,8 | 136,9 | 2,37 |
| 2015 | 50,9 | 0,7 | 29,1 | 3,4 | 19,5 | 103,6 | 1,73 |

| Ano | Fluxo | Emissões evitadas | Juros evitados | Acum. fluxos | Saldo à Selic | DBGG oficial | CF estoque | CF Selic | Oficial % PIB | CF Selic % PIB |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 2003 | 28,5 | -28,5 | 0,0 | 28,5 | 28,5 | — | — | — | — | — |
| 2004 | 35,7 | -35,7 | 4,6 | 64,2 | 68,8 | — | — | — | — | — |
| 2005 | 38,6 | -38,6 | 13,1 | 102,8 | 120,6 | — | — | — | — | — |
| 2006 | 46,2 | -46,2 | 18,2 | 149,0 | 185,0 | R$ 1.336,6 bi | R$ 1.187,6 bi | R$ 1.151,7 bi | 55,5 | 47,8 |
| 2007 | 58,4 | -58,4 | 21,9 | 207,4 | 265,2 | R$ 1.542,9 bi | R$ 1.335,4 bi | R$ 1.277,6 bi | 56,7 | 47,0 |
| 2008 | 77,7 | -77,7 | 33,1 | 285,1 | 376,0 | R$ 1.740,9 bi | R$ 1.455,8 bi | R$ 1.364,9 bi | 56,0 | 43,9 |
| 2009 | 82,1 | -82,1 | 37,3 | 367,2 | 495,4 | R$ 1.973,4 bi | R$ 1.606,3 bi | R$ 1.478,0 bi | 59,2 | 44,3 |
| 2010 | 140,7 | -140,7 | 48,4 | 507,9 | 684,5 | R$ 2.011,5 bi | R$ 1.503,7 bi | R$ 1.327,0 bi | 51,8 | 34,1 |
| 2011 | 108,1 | -108,1 | 79,5 | 615,9 | 872,0 | R$ 2.243,6 bi | R$ 1.627,7 bi | R$ 1.371,6 bi | 51,3 | 31,3 |
| 2012 | 110,5 | -110,5 | 73,9 | 726,4 | 1.056,5 | R$ 2.583,9 bi | R$ 1.857,5 bi | R$ 1.527,5 bi | 53,7 | 31,7 |
| 2013 | 145,6 | -145,6 | 86,7 | 872,0 | 1.288,8 | R$ 2.748,0 bi | R$ 1.876,0 bi | R$ 1.459,2 bi | 51,5 | 27,4 |
| 2014 | 136,9 | -136,9 | 140,6 | 1.008,9 | 1.566,3 | R$ 3.252,4 bi | R$ 2.243,5 bi | R$ 1.686,1 bi | 56,3 | 29,2 |
| 2015 | 103,6 | -103,6 | 208,2 | 1.112,5 | 1.878,1 | R$ 3.927,5 bi | R$ 2.815,0 bi | R$ 2.049,4 bi | 65,5 | 34,2 |

A DBGG oficial em metodologia 2008 só existe a partir de **dezembro de
2006** (SGS 13761). 2003–2005 entram como fluxo e saldo acumulado, sem
estoque oficial para subtrair. Em 2006 o contrafactual Selic já abre um
buraco de R$ 185,0 bi no estoque.

O fator que muda no ano *t* é quase todo **emissão líquida** (−fluxo).
O fator **juros nominais** só fica grande depois que o saldo acumulou —
e é aí que a Selic de dois dígitos (2003–2006, 2008, 2011, 2014–2015)
pesa. Em 2015 os juros evitados no ano são R$ 208,2 bi;
o saldo capitalizado chega a R$ 1.878,1 bi.

## O que a simulação não é

Não é um modelo de equilíbrio geral. Sem indiretas e sem os três
benefícios, o PIB, a arrecadação residual e a própria Selic seriam
outros. A RFB também ignora essa reação no DGT.

Não substitui o crédito direcionado por nada: o exercício só devolve o
caixa à Conta Única e pergunta o que acontece com a **dívida bruta** se
esse caixa resgata DPF.

Não trata o crédito Tesouro→BNDES como se fosse igual às indiretas.
Aquele crédito ( R$ 440,8 bi em 2008–2014 ) é o fluxo que de fato
saiu da União e entrou no passivo do Tesouro. As indiretas de 2003–2015
(R$ 658,4 bi) misturam Tesouro, FAT e recursos próprios.

## Arquivos

| Arquivo | Conteúdo |
|---|---|
| `output/TCU_CG_2010_DBGG_CONTA_UNICA.md` | Esta análise |
| `output/grafico_dbgg_fluxos_conta_unica_2003_2015.png` | Empilhamento dos três fluxos |
| `output/grafico_dbgg_contrafactual_2006_2015.png` | DBGG oficial × contrafactuais |
| `scripts/sim_dbgg_conta_unica_dados.py` | Séries oficiais e reconstrução da renúncia |

```bash
python3 scripts/simular_dbgg_conta_unica.py
```
