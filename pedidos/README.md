# Pedidos

## Pedido COMPRE MAIS (PA) — 30/09/2026

Arquivo: `Pedido_COMPRE_MAIS_20260930.xlsx`, gerado por `montar_pedido_compre_mais.py`
a partir da *Lista de Compra* (30/09/2026 14:01) e da planilha *COMPRE MAIS PA* (filiais Marabá e Castanhal).

### Regras aplicadas
- Cruzamento pelo **EAN principal e pelos EANs adicionais 1 a 5** da Lista de Compra.
- **Frações:** o Compre Mais vende a caixa fechada e a Lista registra o preço por unidade.
  Preço comparável = `PVENDA ÷ Un/Cx`.
- Entra no pedido **somente o que está mais barato que a última compra**.
- Quantidade = coluna **Caixas**, respeitando `QTD_MIN` (múltiplo), `QTD_MAX` e o estoque do Compre Mais.
- Escolhe a filial mais barata com estoque; se faltar, completa na outra filial (se também for mais barata).

### Abas
| Aba | Conteúdo |
|---|---|
| Resumo | Totais, valor do pedido e economia |
| Pedido | Pedido completo (as duas filiais) |
| Pedido Marabá / Pedido Castanhal | Pedido separado por filial |
| Sem estoque CM | Itens mais baratos, mas sem estoque no Compre Mais |
| Comparativo | Todos os itens cruzados, mais baratos ou não |

### Como gerar de novo
```bash
python3 pedidos/montar_pedido_compre_mais.py Lista_de_Compra.xlsx COMPRE_MAIS_PA.xlsx pedidos/Pedido_COMPRE_MAIS_AAAAMMDD.xlsx
```
