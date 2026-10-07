# Cotações — Drogarias Rocha

Cada cotação fica numa pasta com a data (`AAAA-MM-DD/`), sempre com a mesma estrutura:

```
AAAA-MM-DD/
├── envio/                    planilhas enviadas aos distribuidores (COTACAO_<DIST>_<data>.xlsx)
├── respostas/                arquivos recebidos, sem alteração (lista de compra + 1 por distribuidor)
├── pedidos/
│   ├── OL/                   1 arquivo por OL  (OL_<NOME>_<data>.xlsx)
│   └── distribuidores/       1 arquivo por distribuidor  (PEDIDO_<DIST>_<data>.xlsx)
├── COMPARATIVO_<data>.xlsx   preços lado a lado, vencedor, sem cotação
└── *.py                      scripts que geram tudo acima
```

## Regras

1. **OL x Distribuidor.** Item com a coluna `OL` preenchida na Lista de Compra é
   comprado direto do laboratório: cada OL vira um pedido independente e o item
   **não** entra na cotação. Todo o resto vai para cotação com os distribuidores.
2. **Uma linha por EAN** na planilha de cotação (principal e adicionais), porque
   alguns distribuidores ignoram as colunas de EAN adicional.
3. **Imposto por fora:** Panpharma e Nazária +3,92% sobre o preço cotado.
4. **Teto, sem tolerância (0%):** não se compra nada acima do maior preço entre
   última compra, custo atual e melhores 3/6/12 meses (× Un/Cx). Itens acima do
   teto ficam em "Sem cotação" com o menor preço recebido, para negociar.
5. **Embalagem:** preço abaixo de 50% ou acima de 200% do custo atual é tratado
   como embalagem/EAN diferente: fica marcado para conferir e não concorre.
6. **Vencedor** = menor preço entre os que passam nas regras, com estoque.
7. **Campanha de OL** (PDF do laboratório/distribuidor, ex.: Painel O.L Procter):
   `ol_campanha.py` gera o pedido com quantidade para **60 dias** de estoque
   (2 × Dem./mês − Est. rede, arredondado para a embalagem da oferta), respeitando
   o teto. O que for comprado na campanha sai dos pedidos de distribuidor.

## Passo a passo

```bash
cd cotacoes/AAAA-MM-DD
# 1) planilhas para os distribuidores (+ pedidos de OL)
python gerar_cotacao.py respostas/Lista_de_Compra_*.xlsx
python pedidos.py respostas/Lista_de_Compra_*.xlsx -d DD-MM-AAAA
# 2) com as respostas em respostas/
python comparativo.py respostas/Lista_de_Compra_*.xlsx PANPHARMA=respostas/... NAZARIA=... SBLOG=... TAPAJOS=...
# 3) campanha de OL (antes de gerar os pedidos de distribuidor)
python ol_campanha.py respostas/Lista_de_Compra_*.xlsx respostas/campanhas/<campanha>.pdf PROCTER -d DD-MM-AAAA
python pedidos.py     respostas/Lista_de_Compra_*.xlsx PANPHARMA=respostas/... NAZARIA=... SBLOG=... TAPAJOS=... -d DD-MM-AAAA
```

Formatos de resposta reconhecidos: tabela ePan PRO (Panpharma), a própria planilha
COTACAO respondida (Tapajós; Nazária com colunas próprias de preço/desconto/estoque)
e a planilha "cotacao" do portal SB Log.
