"""Monta o pedido COMPRE MAIS (PA) a partir da Lista de Compra.

Regras:
- Cruza os produtos pelo EAN principal e pelos EANs adicionais (1 a 5).
- O Compre Mais vende a caixa fechada; a Lista de Compra registra o preço por
  unidade (fração). Preço comparável = PVENDA / Un/Cx.
- Entra no pedido somente o que estiver MAIS BARATO que a última compra.
- Quantidade pedida = coluna "Caixas" (respeitando QTD_MIN, QTD_MAX e estoque).
- Entre as filiais (Marabá e Castanhal) escolhe a mais barata com estoque;
  se faltar estoque, completa na outra filial desde que também seja mais barata.

Uso: python3 montar_pedido_compre_mais.py LISTA.xlsx COMPRE_MAIS.xlsx SAIDA.xlsx
"""
import math
import sys

import pandas as pd

EAN_COLS = ["EAN princ."] + [f"EAN adic. {i}" for i in range(1, 6)]
FILIAIS = {"MARABA": "Marabá", "CASTA": "Castanhal"}


def norm_ean(v):
    if pd.isna(v):
        return None
    s = str(v).strip().split(".")[0].lstrip("0")
    return s or None


def carregar(lista_path, cm_path):
    lista = pd.read_excel(lista_path, header=5, dtype={c: str for c in EAN_COLS})
    lista = lista[pd.to_numeric(lista["Cód. interno"], errors="coerce").notna()].copy()

    abas = pd.read_excel(cm_path, sheet_name=None)
    cm = pd.concat([df.assign(FILIAL_NOME=FILIAIS.get(aba, aba)) for aba, df in abas.items()])
    cm["ean"] = cm["CODAUXILIAR"].map(norm_ean)

    eans = lista.melt(id_vars=["Cód. interno"], value_vars=EAN_COLS,
                      var_name="EAN usado", value_name="ean")
    eans["ean"] = eans["ean"].map(norm_ean)
    eans = eans.dropna(subset=["ean"]).drop_duplicates(["Cód. interno", "ean"])

    m = eans.merge(cm, on="ean").merge(lista, on="Cód. interno")
    m["Preço CM un"] = m["PVENDA"] / m["Un/Cx"]
    m["Economia un"] = m["Últ. preço compra"] - m["Preço CM un"]
    m["Mais barato"] = m["Economia un"] > 0.005
    return lista, m


def alocar(m):
    linhas = []
    for cod, g in m[m["Mais barato"]].groupby("Cód. interno"):
        g = g.sort_values(["Preço CM un", "ESTOQUE"], ascending=[True, False])
        falta = int(g["Caixas"].iloc[0])
        for _, r in g.iterrows():
            if falta <= 0:
                break
            qmin, qmax, est = int(r["QTD_MIN"]), int(r["QTD_MAX"]), int(r["ESTOQUE"])
            qtd = max(falta, qmin)
            qtd = math.ceil(qtd / qmin) * qmin if qmin > 1 else qtd
            qtd = min(qtd, qmax, est)
            if qmin > 1:
                qtd = (qtd // qmin) * qmin
            if qtd <= 0:
                continue
            linhas.append({**r.to_dict(), "Qtd pedido (cx)": qtd})
            falta -= qtd
        if falta > 0 and linhas and linhas[-1]["Cód. interno"] == cod:
            linhas[-1]["Obs"] = f"Faltaram {falta} cx por estoque no Compre Mais"
        elif falta > 0:
            r = g.iloc[0]
            linhas.append({**r.to_dict(), "Qtd pedido (cx)": 0,
                           "Obs": "Sem estoque no Compre Mais"})
    p = pd.DataFrame(linhas)
    p["Un. pedido"] = p["Qtd pedido (cx)"] * p["Un/Cx"]
    p["Total CM"] = p["Qtd pedido (cx)"] * p["PVENDA"]
    p["Total últ. compra"] = p["Un. pedido"] * p["Últ. preço compra"]
    p["Economia"] = p["Total últ. compra"] - p["Total CM"]
    return p


COLS_PEDIDO = {
    "FILIAL_NOME": "Filial CM", "COD_PROD": "Cód. CM", "CODAUXILIAR": "EAN CM",
    "DESCRICAO": "Descrição CM", "Cód. interno": "Cód. interno", "Produto": "Produto (lista)",
    "EAN usado": "EAN cruzado", "ABC": "ABC", "Un/Cx": "Un/Cx",
    "Qtd pedido (cx)": "Qtd pedido (cx)", "Un. pedido": "Un. pedido",
    "PVENDA": "Preço CM (cx)", "Preço CM un": "Preço CM (un)",
    "Últ. preço compra": "Últ. compra (un)", "Fornecedor": "Fornecedor anterior",
    "Total CM": "Total CM", "Total últ. compra": "Total pela últ. compra",
    "Economia": "Economia", "Obs": "Obs",
}


def main(lista_path, cm_path, saida):
    lista, m = carregar(lista_path, cm_path)
    p = alocar(m)
    if "Obs" not in p:
        p["Obs"] = ""
    pedido = p[p["Qtd pedido (cx)"] > 0]
    sem_est = p[p["Qtd pedido (cx)"] == 0]

    comp = m.copy()
    comp["Situação"] = comp["Mais barato"].map({True: "MAIS BARATO", False: "igual/mais caro"})
    comp = comp.sort_values(["Produto", "FILIAL_NOME"])

    resumo = [
        ("Itens na Lista de Compra", len(lista)),
        ("Itens encontrados no Compre Mais (EAN princ. + adicionais)", m["Cód. interno"].nunique()),
        ("  - cruzados só por EAN adicional", m.groupby("Cód. interno")["EAN usado"].apply(lambda s: (s != "EAN princ.").all()).sum()),
        ("Itens mais baratos que a última compra", m.loc[m["Mais barato"], "Cód. interno"].nunique()),
        ("Itens no pedido", pedido["Cód. interno"].nunique()),
        ("Itens mais baratos mas sem estoque no CM", sem_est["Cód. interno"].nunique()),
        ("Caixas no pedido", int(pedido["Qtd pedido (cx)"].sum())),
    ]
    for fil in FILIAIS.values():
        f = pedido[pedido["FILIAL_NOME"] == fil]
        resumo.append((f"Pedido {fil} — itens / valor (R$)", f"{len(f)} / {f['Total CM'].sum():.2f}"))
    resumo += [
        ("Valor do pedido Compre Mais (R$)", round(pedido["Total CM"].sum(), 2)),
        ("Mesmas unidades pelo preço da última compra (R$)", round(pedido["Total últ. compra"].sum(), 2)),
        ("Economia (R$)", round(pedido["Economia"].sum(), 2)),
    ]

    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        pd.DataFrame(resumo, columns=["Indicador", "Valor"]).to_excel(w, sheet_name="Resumo", index=False)
        ped = pedido[list(COLS_PEDIDO)].rename(columns=COLS_PEDIDO).sort_values(["Filial CM", "Produto (lista)"])
        ped.to_excel(w, sheet_name="Pedido", index=False)
        for fil in FILIAIS.values():
            ped[ped["Filial CM"] == fil].to_excel(w, sheet_name=f"Pedido {fil}", index=False)
        if len(sem_est):
            sem_est[list(COLS_PEDIDO)].rename(columns=COLS_PEDIDO).to_excel(w, sheet_name="Sem estoque CM", index=False)
        comp[["Situação", "FILIAL_NOME", "Cód. interno", "Produto", "DESCRICAO", "EAN usado", "ean",
              "Un/Cx", "Caixas", "PVENDA", "Preço CM un", "Últ. preço compra", "Economia un",
              "ESTOQUE", "QTD_MIN", "Fornecedor"]].rename(columns={
                  "FILIAL_NOME": "Filial CM", "DESCRICAO": "Descrição CM", "ean": "EAN",
                  "PVENDA": "Preço CM (cx)", "Preço CM un": "Preço CM (un)",
                  "Últ. preço compra": "Últ. compra (un)", "Economia un": "Economia (un)",
                  "ESTOQUE": "Estoque CM", "QTD_MIN": "Qtd mín CM"}).to_excel(w, sheet_name="Comparativo", index=False)

        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                larg = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(larg + 2, 8), 50)
                for c in col[1:]:
                    if isinstance(c.value, float):
                        c.number_format = "#,##0.00"

    for linha in resumo:
        print(f"{linha[0]}: {linha[1]}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
