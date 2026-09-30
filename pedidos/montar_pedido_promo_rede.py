"""Monta o pedido PROMO REDE (PA) a partir da Lista de Compra.

Regras:
- Preço da PROMO REDE com desconto adicional (--desconto, padrão 5%):
  PRECO_FINAL x (1 - desconto).
- Cruza pelo EAN principal e pelos EANs adicionais (1 a 5) da Lista de Compra.
- Frações: o preço da PROMO REDE é da caixa; a Lista registra o preço por unidade.
  Preço comparável = preço com desconto / Un/Cx.
- Entra no pedido somente o que estiver MAIS BARATO que a última compra.
- Quantidade = coluna "Caixas", limitada ao estoque da PROMO REDE.
- Itens com preço unitário abaixo de 35% da última compra vão para a aba "Conferir"
  (provável divergência de embalagem/fração) e não entram no pedido.
- Não repete itens já pedidos (aba "Comprados" do Controle_Compras.xlsx, últimos --dias).

Uso: python3 montar_pedido_promo_rede.py LISTA.xlsx PROMO_REDE.xls SAIDA.xlsx [--desconto 5] [--dias 30]
"""
import argparse
from pathlib import Path

import pandas as pd

from montar_pedido_compre_mais import (EAN_COLS, bloqueados, carregar_historico,
                                       norm_ean)

FORNECEDOR = "PROMO REDE PA"
LIMITE_CONFERIR = 0.35


def formatar(w):
    for ws in w.book.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            larg = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(larg + 2, 8), 50)
            for c in col[1:]:
                if isinstance(c.value, float):
                    c.number_format = "#,##0.00"


def main(lista_path, promo_path, saida, desconto=5.0, dias=30, bloquear=()):
    """`bloquear`: códigos internos já pedidos nesta mesma rodada (ex.: Compre Mais)."""
    saida = Path(saida)
    pedido_id = saida.stem
    fator = 1 - desconto / 100

    lista = pd.read_excel(lista_path, header=5, dtype={c: str for c in EAN_COLS})
    lista = lista[pd.to_numeric(lista["Cód. interno"], errors="coerce").notna()].copy()
    promo = pd.read_excel(promo_path, dtype={"EAN": str})
    promo["ean"] = promo["EAN"].map(norm_ean)

    eans = lista.melt(id_vars=["Cód. interno"], value_vars=EAN_COLS,
                      var_name="EAN usado", value_name="ean")
    eans["ean"] = eans["ean"].map(norm_ean)
    eans = eans.dropna(subset=["ean"]).drop_duplicates(["Cód. interno", "ean"])
    lista_eans = eans.groupby("Cód. interno")["ean"].apply(set).to_dict()

    m = (eans.merge(promo, on="ean").merge(lista, on="Cód. interno")
         .sort_values("PRECO_FINAL").drop_duplicates("Cód. interno"))
    m["Preço c/ desc (cx)"] = (m["PRECO_FINAL"] * fator).round(2)
    m["Preço c/ desc (un)"] = m["Preço c/ desc (cx)"] / m["Un/Cx"]
    m["Economia un"] = m["Últ. preço compra"] - m["Preço c/ desc (un)"]
    ref = m["Últ. preço compra"] > 0
    m["Razão"] = m["Preço c/ desc (un)"] / m["Últ. preço compra"].where(ref)
    mais_barato = ref & (m["Economia un"] > 0.005)
    conferir = mais_barato & (m["Razão"] < LIMITE_CONFERIR)

    hist = carregar_historico()
    h_rec, cods_bloq, eans_bloq = bloqueados(hist, pedido_id, dias)
    cods_bloq |= {str(c) for c in bloquear}
    ja = m["Cód. interno"].map(
        lambda c: str(c) in cods_bloq or bool(lista_eans.get(c, set()) & eans_bloq))

    cand = m[mais_barato & ~conferir & ~ja].copy()
    cand["Qtd pedido (cx)"] = cand[["Caixas", "ESTOQUE"]].min(axis=1).clip(lower=0).astype(int)
    cand["Obs"] = ""
    cand.loc[cand["ESTOQUE"] < cand["Caixas"], "Obs"] = (
        "Faltaram " + (cand["Caixas"] - cand["ESTOQUE"]).astype(str) + " cx por estoque na PROMO REDE")
    pedido = cand[cand["Qtd pedido (cx)"] > 0].copy()
    pedido["Un. pedido"] = pedido["Qtd pedido (cx)"] * pedido["Un/Cx"]
    pedido["Total"] = pedido["Qtd pedido (cx)"] * pedido["Preço c/ desc (cx)"]
    pedido["Total últ. compra"] = pedido["Un. pedido"] * pedido["Últ. preço compra"]
    pedido["Economia"] = pedido["Total últ. compra"] - pedido["Total"]

    cols = {"CODPROD": "Cód. PROMO", "EAN": "EAN PROMO", "DESCRICAO": "Descrição PROMO",
            "Cód. interno": "Cód. interno", "Produto": "Produto (lista)", "EAN usado": "EAN cruzado",
            "ABC": "ABC", "Un/Cx": "Un/Cx", "Qtd pedido (cx)": "Qtd pedido (cx)",
            "Un. pedido": "Un. pedido", "PRECO_FINAL": "Preço tabela (cx)",
            "Preço c/ desc (cx)": f"Preço -{desconto:g}% (cx)",
            "Preço c/ desc (un)": f"Preço -{desconto:g}% (un)",
            "Últ. preço compra": "Últ. compra (un)", "Fornecedor": "Fornecedor anterior",
            "Total": "Total PROMO", "Total últ. compra": "Total pela últ. compra",
            "Economia": "Economia", "Obs": "Obs"}
    ped = pedido[list(cols)].rename(columns=cols).sort_values("Produto (lista)")

    resumo = [
        ("Desconto adicional aplicado", f"{desconto:g}%"),
        ("Itens na Lista de Compra", len(lista)),
        ("Itens encontrados na PROMO REDE (EAN princ. + adicionais)", len(m)),
        ("Itens mais baratos que a última compra (com desconto)", int(mais_barato.sum())),
        (f"  - já pedidos nos últimos {dias} dias (fora deste pedido)", int((mais_barato & ~conferir & ja).sum())),
        ("  - para conferir (possível divergência de fração)", int(conferir.sum())),
        ("Itens no pedido", len(pedido)),
        ("Caixas no pedido", int(pedido["Qtd pedido (cx)"].sum())),
        ("Valor tabela, sem o desconto (R$)", round((pedido["Qtd pedido (cx)"] * pedido["PRECO_FINAL"]).sum(), 2)),
        (f"Valor do pedido com -{desconto:g}% (R$)", round(pedido["Total"].sum(), 2)),
        ("Mesmas unidades pelo preço da última compra (R$)", round(pedido["Total últ. compra"].sum(), 2)),
        ("Economia (R$)", round(pedido["Economia"].sum(), 2)),
    ]

    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        pd.DataFrame(resumo, columns=["Indicador", "Valor"]).to_excel(w, sheet_name="Resumo", index=False)
        ped.to_excel(w, sheet_name="Pedido", index=False)
        jp = m[mais_barato & ~conferir & ja]
        if len(jp):
            ult = h_rec.sort_values("Data pedido").drop_duplicates("Cód. interno", keep="last")
            jp = jp[["Cód. interno", "Produto", "Caixas", "Últ. preço compra", "Preço c/ desc (un)"]].copy()
            jp["Cód. interno"] = jp["Cód. interno"].astype(str)
            jp = jp.merge(ult[["Cód. interno", "Pedido", "Data pedido", "Fornecedor", "Qtd (cx)"]],
                          on="Cód. interno", how="left")
            jp["Fornecedor"] = jp["Fornecedor"].fillna("pedido desta rodada")
            jp.to_excel(w, sheet_name="Já pedidos", index=False)
        if conferir.any():
            m[conferir][["Cód. interno", "Produto", "DESCRICAO", "Un/Cx", "Caixas", "PRECO_FINAL",
                         "Preço c/ desc (cx)", "Preço c/ desc (un)", "Últ. preço compra"]].to_excel(
                w, sheet_name="Conferir", index=False)
        comp = m.copy()
        comp["Situação"] = "igual/mais caro"
        comp.loc[mais_barato, "Situação"] = "MAIS BARATO"
        comp.loc[conferir, "Situação"] = "CONFERIR"
        comp.loc[~ref, "Situação"] = "sem preço de referência"
        comp[["Situação", "Cód. interno", "Produto", "DESCRICAO", "EAN usado", "ean", "Un/Cx", "Caixas",
              "PRECO_FINAL", "Preço c/ desc (cx)", "Preço c/ desc (un)", "Últ. preço compra",
              "Economia un", "ESTOQUE", "Fornecedor"]].sort_values(["Situação", "Produto"]).to_excel(
            w, sheet_name="Comparativo", index=False)
        formatar(w)

    for linha in resumo:
        print(f"{linha[0]}: {linha[1]}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("lista")
    ap.add_argument("promo")
    ap.add_argument("saida")
    ap.add_argument("--desconto", type=float, default=5.0, help="desconto adicional em %% (padrão 5)")
    ap.add_argument("--dias", type=int, default=30, help="bloqueia itens pedidos nos últimos N dias")
    a = ap.parse_args()
    main(a.lista, a.promo, a.saida, a.desconto, a.dias)
