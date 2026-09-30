# Pedidos

Cada rodada gera **só duas planilhas**:

| Planilha | O que tem |
|---|---|
| `Controle_Compras_AAAAMMDD.xlsx` | **Resumo** (valor e economia por distribuidor) · uma aba com a **tabela de preços de cada distribuidor**, com o que foi comprado destacado em verde · **Comprados** · **Pendências** (faltas de estoque e itens a conferir) |
| `Pedido_AAAAMMDD.xlsx` | Uma aba por distribuidor, pronta para envio (código, EAN, descrição, quantidade em caixas, preço e total) |

## Como gerar
```bash
python3 pedidos/gerar_planilhas.py Lista_de_Compra.xlsx COMPRE_MAIS_PA.xlsx PROMO_REDE_PA.xls --desconto 5
```

## Regras
- Cruzamento pelo EAN principal e pelos EANs adicionais da Lista de Compra.
- Frações: preço da caixa ÷ Un/Cx, comparado com a última compra (por unidade).
- Só entra no pedido o que estiver **mais barato que a última compra**; quantidade = coluna **Caixas**, limitada ao estoque do distribuidor.
- PROMO REDE com **5% de desconto adicional** sobre o PRECO_FINAL.
- Itens já pedidos nos últimos 30 dias não se repetem (`historico_pedidos.csv`, uso interno do script).

## 30/09/2026
| Distribuidor | Itens | Caixas | Pedido | Economia |
|---|---|---|---|---|
| COMPRE MAIS Marabá | 47 | 515 | R$ 3.965,03 | R$ 511,92 |
| COMPRE MAIS Castanhal | 27 | 379 | R$ 2.349,07 | R$ 249,81 |
| PROMO REDE (−5%) | 159 | 1.018 | R$ 10.367,62 | R$ 1.534,07 |
| **Total** | **233** | **1.912** | **R$ 16.681,72** | **R$ 2.295,79** |
