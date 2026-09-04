# Cotação 04/09/2026 — Verificação e geração de pedidos

Cruzamento da **LISTA DE COMPRA (controle)** com 4 cotações recebidas em 04/09/2026:

| Fonte | Arquivo original | Casamento | Preço usado |
|---|---|---|---|
| ePan (painel de preços) | `ePan_Precos_20260904_1843.xlsx` | EAN (principal + adicionais 1–3) | Preço líquido c/ ST |
| SB LOG (pedido nº 5518) | `pedido_N__5518_04_09_2026_13_27_00.xlsx` | EAN (principal + adicionais 1–3) | Valor un. + ST un. |
| Tapajós | `LISTA_DE_COMPRA_04092026__COTA__O.xlsx` | Cód. interno | VALOR (por caixa) |
| Nazária | `LISTA_DE_COMPRA_04092026__COTA__O_1.xlsx` | Cód. interno | PREÇO FINAL (por caixa) |

## Regras aplicadas

1. **OL separado**: todo item com valor na coluna `OL` do controle sai da cotação e vai
   para um arquivo próprio por OL (`OL_*.xlsx`), sem preços de cotação.
2. **Demais itens**: alocados ao distribuidor de **menor preço unitário equivalente**
   com estoque (ePan exige `Estoque = SIM`; Nazária exige estoque > 0).
3. **Validação contra histórico**: status `OK` quando o preço unitário equivalente é
   ≤ menor histórico (melhor 3/6/12 meses) **ou** ≤ última compra (compra atual);
   caso contrário `ACIMA +x%` para revisão.
4. **Normalização caixa/unidade**: Tapajós e Nazária cotam por caixa (preço ÷ Un/Cx);
   para ePan/SB LOG em itens com Un/Cx > 1 a embalagem é detectada pelo histórico e
   marcada com observação `conferir emb.`.

## Arquivos gerados

- `PEDIDO_EPAN_*.xlsx`, `PEDIDO_SBLOG_*.xlsx`, `PEDIDO_TAPAJOS_*.xlsx`, `PEDIDO_NAZARIA_*.xlsx`
  — pedido de cada distribuidor (itens vencedores, qtd, preço, total e comparação com histórico).
- `OL_*.xlsx` — 1 arquivo por OL, sem cotação (código, EANs, produto, quantidades).
- `COMPARATIVO_COTACAO_*.xlsx` — resumo, comparativo completo dos 4 distribuidores por item
  e aba `Sem_Cotacao` (itens que nenhum distribuidor cotou ou só havia fornecedor sem estoque).

Gerado por `gerar_pedidos.py` (os arquivos de entrada não estão versionados).
