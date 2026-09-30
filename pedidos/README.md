# Pedidos

## Pedido COMPRE MAIS (PA) — 30/09/2026

Arquivos:
- `Pedido_COMPRE_MAIS_20260930.xlsx`: análise completa (resumo, pedido, comparativo).
- `Pedido_COMPRE_MAIS_20260930_Maraba.xlsx` e `Pedido_COMPRE_MAIS_20260930_Castanhal.xlsx`: **pedidos para envio**
  (COD_PROD, CODAUXILIAR, DESCRICAO, QTD em caixas, PVENDA, TOTAL).

Tudo é gerado por `montar_pedido_compre_mais.py`
a partir da *Lista de Compra* (30/09/2026 14:01) e da planilha *COMPRE MAIS PA* (filiais Marabá e Castanhal).

### Regras aplicadas
- Cruzamento pelo **EAN principal e pelos EANs adicionais 1 a 5** da Lista de Compra.
- **Frações:** o Compre Mais vende a caixa fechada e a Lista registra o preço por unidade.
  Preço comparável = `PVENDA ÷ Un/Cx`.
- Entra no pedido **somente o que está mais barato que a última compra**.
- Quantidade = coluna **Caixas**, respeitando `QTD_MIN` (múltiplo), `QTD_MAX` e o estoque do Compre Mais.
- Escolhe a filial mais barata com estoque; se faltar, completa na outra filial (se também for mais barata).

- **Não repete itens já pedidos:** cada pedido é registrado em `historico_pedidos.csv`.
  Nas próximas planilhas, os produtos pedidos nos últimos 30 dias (reconhecidos pelo código interno
  ou por qualquer EAN) ficam fora do pedido e aparecem na aba **Já pedidos**. O prazo muda com `--dias N`.

### Pedido 30/09/2026
| Filial | Itens | Caixas | Valor | Economia |
|---|---|---|---|---|
| Marabá | 47 | 515 | R$ 3.965,03 | R$ 511,92 |
| Castanhal | 27 | 379 | R$ 2.349,07 | R$ 249,81 |
| **Total** | **74 linhas (71 produtos)** | **894** | **R$ 6.314,10** | **R$ 761,72** |

### Abas
| Aba | Conteúdo |
|---|---|
| Resumo | Totais, valor do pedido e economia |
| Pedido | Pedido completo (as duas filiais) |
| Pedido Marabá / Pedido Castanhal | Pedido separado por filial |
| Já pedidos | Itens mais baratos que ficaram fora por já estarem em pedido recente |
| Sem estoque CM | Itens mais baratos, mas sem estoque no Compre Mais |
| Comparativo | Todos os itens cruzados, mais baratos ou não |

### Como gerar de novo
```bash
python3 pedidos/montar_pedido_compre_mais.py Lista_de_Compra.xlsx COMPRE_MAIS_PA.xlsx pedidos/Pedido_COMPRE_MAIS_AAAAMMDD.xlsx [--dias 30]
```
Mantenha o `historico_pedidos.csv` no repositório: é ele que impede a repetição.
Rodar de novo com o mesmo nome de saída substitui o registro daquele pedido (não duplica).
