# Pedidos

## O que é gerado

**1. `Controle_Compras.xlsx` — nossa planilha de controle (fica com a gente)**
Acumula todos os pedidos, rodada após rodada:
- **Resumo:** valor, caixas e economia por data e distribuidor.
- **Comprados:** tudo o que já foi pedido. É por esta aba que o script sabe o que já foi comprado e não repete o item nas próximas planilhas (30 dias).
- **Pendências:** faltas de estoque no distribuidor e itens a conferir.
- **Uma aba com a tabela de preços de cada distribuidor**, com o que foi comprado destacado em verde.

**2. Um arquivo por distribuidor (para enviar a eles)**
- `Pedido_COMPRE_MAIS_AAAAMMDD.xlsx`, com uma aba por filial (Marabá e Castanhal).
- `Pedido_PROMO_REDE_AAAAMMDD.xlsx`

## Como gerar
```bash
python3 pedidos/gerar_planilhas.py Lista_de_Compra.xlsx COMPRE_MAIS_PA.xlsx PROMO_REDE_PA.xls --desconto 5
```

## Regras
- Cruzamento pelo EAN principal e pelos EANs adicionais da Lista de Compra.
- Frações: preço da caixa ÷ Un/Cx, comparado com a última compra (por unidade).
- Só entra no pedido o que estiver **mais barato que a última compra**. A quantidade é a coluna **Caixas**, limitada ao estoque do distribuidor.
- PROMO REDE com **5% de desconto adicional** sobre o PRECO_FINAL.
- O que vai para o Compre Mais não entra de novo na PROMO REDE, e o que já está em **Comprados** não se repete.

## 30/09/2026
| Distribuidor | Itens | Caixas | Pedido | Economia |
|---|---|---|---|---|
| COMPRE MAIS Marabá | 47 | 515 | R$ 3.965,03 | R$ 511,92 |
| COMPRE MAIS Castanhal | 27 | 379 | R$ 2.349,07 | R$ 249,81 |
| PROMO REDE (−5%) | 159 | 1.018 | R$ 10.367,62 | R$ 1.534,07 |
| **Total** | **233** | **1.912** | **R$ 16.681,72** | **R$ 2.295,79** |
