"""Monta o pedido COMPRE MAIS (PA) a partir da Lista de Compra.

Regras:
- Cruza os produtos pelo EAN principal e pelos EANs adicionais (1 a 5).
- O Compre Mais vende a caixa fechada; a Lista de Compra registra o preço por
  unidade (fração). Preço comparável = PVENDA / Un/Cx.
- Entra no pedido somente o que estiver MAIS BARATO que a última compra.
- Quantidade pedida = coluna "Caixas" (respeitando QTD_MIN, QTD_MAX e estoque).
- Entre as filiais (Marabá e Castanhal) escolhe a mais barata com estoque;
  se faltar estoque, completa na outra filial desde que também seja mais barata.
- Não repete itens já pedidos: tudo o que entra no pedido é gravado em
  historico_pedidos.csv, e os produtos pedidos nos últimos --dias (padrão 30),
  identificados pelo código interno ou por qualquer EAN, ficam fora das próximas
  planilhas (aba "Já pedidos"). Gerar de novo o mesmo pedido substitui o registro
  dele, sem se bloquear.
- Gera também um arquivo de envio por filial (SAIDA_Maraba.xlsx, SAIDA_Castanhal.xlsx).

Uso: python3 montar_pedido_compre_mais.py LISTA.xlsx COMPRE_MAIS.xlsx SAIDA.xlsx [--dias 30]
"""
import argparse
import datetime as dt
import math
import unicodedata
from pathlib import Path

import pandas as pd

HISTORICO = Path(__file__).with_name("historico_pedidos.csv")
FORNECEDOR = "COMPRE MAIS PA"

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


def carregar_historico():
    if not HISTORICO.exists():
        return pd.DataFrame(columns=["Pedido", "Data pedido", "Fornecedor", "Filial", "Cód. CM",
                                     "EAN CM", "Descrição CM", "Cód. interno", "Produto", "EANs",
                                     "Qtd (cx)", "Preço (cx)", "Total"])
    return pd.read_csv(HISTORICO, sep=";", dtype=str)


def bloqueados(hist, pedido_id, dias):
    """Códigos internos e EANs pedidos nos últimos `dias` (exceto o próprio pedido)."""
    h = hist[hist["Pedido"] != pedido_id]
    limite = dt.date.today() - dt.timedelta(days=dias)
    h = h[pd.to_datetime(h["Data pedido"]).dt.date >= limite]
    cods = set(h["Cód. interno"].dropna())
    eans = {e for v in h["EANs"].dropna() for e in v.split("|") if e}
    return h, cods, eans


def sem_acento(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


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


def main(lista_path, cm_path, saida, dias=30):
    saida = Path(saida)
    pedido_id = saida.stem
    lista, m = carregar(lista_path, cm_path)

    hist = carregar_historico()
    h_rec, cods_bloq, eans_bloq = bloqueados(hist, pedido_id, dias)
    eans_prod = (lista.set_index("Cód. interno")[EAN_COLS].apply(
        lambda r: {e for e in map(norm_ean, r) if e}, axis=1))
    lista_eans = eans_prod.to_dict()
    ja = m["Cód. interno"].map(
        lambda c: str(c) in cods_bloq or bool(lista_eans.get(c, set()) & eans_bloq))
    ja_pedidos = m[ja & m["Mais barato"]].drop_duplicates("Cód. interno")
    m = m[~ja]

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
        ("Itens mais baratos que a última compra", m.loc[m["Mais barato"], "Cód. interno"].nunique() + len(ja_pedidos)),
        (f"  - já pedidos nos últimos {dias} dias (fora deste pedido)", len(ja_pedidos)),
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
        if len(ja_pedidos):
            ult = h_rec.sort_values("Data pedido").drop_duplicates("Cód. interno", keep="last")
            jp = ja_pedidos[["Cód. interno", "Produto", "Caixas", "Últ. preço compra"]].copy()
            jp["Cód. interno"] = jp["Cód. interno"].astype(str)
            jp = jp.merge(ult[["Cód. interno", "Pedido", "Data pedido", "Filial", "Qtd (cx)"]],
                          on="Cód. interno", how="left")
            jp.to_excel(w, sheet_name="Já pedidos", index=False)
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

    # Arquivo de envio por filial
    for fil in FILIAIS.values():
        f = pedido[pedido["FILIAL_NOME"] == fil].sort_values("DESCRICAO")
        if f.empty:
            continue
        env = pd.DataFrame({
            "COD_PROD": f["COD_PROD"], "CODAUXILIAR": f["CODAUXILIAR"].astype(str),
            "DESCRICAO": f["DESCRICAO"], "QTD": f["Qtd pedido (cx)"].astype(int),
            "PVENDA": f["PVENDA"], "TOTAL": f["Total CM"].round(2)})
        total = pd.DataFrame([{"DESCRICAO": "TOTAL", "QTD": env["QTD"].sum(),
                               "TOTAL": round(env["TOTAL"].sum(), 2)}])
        env = pd.concat([env.astype({"COD_PROD": object}), total], ignore_index=True)
        arq = saida.with_name(f"{saida.stem}_{sem_acento(fil)}.xlsx")
        with pd.ExcelWriter(arq, engine="openpyxl") as w:
            env.to_excel(w, sheet_name=f"Pedido {fil}", index=False)
            ws = w.book.active
            for col, larg in zip("ABCDEF", (10, 16, 45, 8, 10, 12)):
                ws.column_dimensions[col].width = larg
            for row in ws.iter_rows(min_row=2, min_col=5, max_col=6):
                for c in row:
                    c.number_format = "#,##0.00"
        print(f"Arquivo de envio: {arq.name}")

    # Registra o pedido no histórico (substitui o registro anterior do mesmo pedido)
    novos = pd.DataFrame({
        "Pedido": pedido_id, "Data pedido": dt.date.today().isoformat(), "Fornecedor": FORNECEDOR,
        "Filial": pedido["FILIAL_NOME"], "Cód. CM": pedido["COD_PROD"].astype(str),
        "EAN CM": pedido["CODAUXILIAR"].astype(str), "Descrição CM": pedido["DESCRICAO"],
        "Cód. interno": pedido["Cód. interno"].astype(str), "Produto": pedido["Produto"],
        "EANs": pedido["Cód. interno"].map(lambda c: "|".join(sorted(lista_eans.get(c, set())))),
        "Qtd (cx)": pedido["Qtd pedido (cx)"].astype(int).astype(str),
        "Preço (cx)": pedido["PVENDA"].map("{:.2f}".format),
        "Total": pedido["Total CM"].map("{:.2f}".format),
    })
    hist = pd.concat([hist[hist["Pedido"] != pedido_id], novos], ignore_index=True)
    hist.to_csv(HISTORICO, sep=";", index=False)
    print(f"Histórico atualizado: {len(novos)} linhas do pedido {pedido_id} em {HISTORICO.name}")

    for linha in resumo:
        print(f"{linha[0]}: {linha[1]}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("lista")
    ap.add_argument("compre_mais")
    ap.add_argument("saida")
    ap.add_argument("--dias", type=int, default=30,
                    help="bloqueia itens pedidos nos últimos N dias (padrão 30)")
    a = ap.parse_args()
    main(a.lista, a.compre_mais, a.saida, a.dias)
