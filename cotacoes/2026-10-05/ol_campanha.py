"""Pedido de OL a partir de uma campanha em PDF (ex.: Painel O.L Procter / SB Log).

Quantidade = cobertura de 60 dias: 2 × Dem./mês − Est. rede (em unidades),
arredondada para cima na embalagem da campanha (display = Un/Cx da lista
quando o preço da oferta bate com caixa fechada). Mesmas regras da cotação:
não compra acima do teto (maior preço do histórico, 0% de tolerância).
Itens da campanha que não estão na Lista de Compra ficam no arquivo sem
quantidade (sem demanda/estoque para calcular).

Uso: python ol_campanha.py <Lista_de_Compra.xlsx> <campanha.pdf> <NOME> [-d dd-mm-aaaa] [--dias 60]
"""
import math
import re
import subprocess
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook

from comparativo import HIST, M, cabecalho, celula
from gerar_cotacao import COLS_EAN
from pedidos import AQUI, nome_arquivo, titulo


def le_campanha(pdf):
    txt = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True, check=True).stdout
    num = lambda s: float(s.replace(".", "").replace(",", "."))
    rows = []
    for l in txt.splitlines():
        m = re.match(r"\s*(\d{8,14})\s+(\d+)\s+(.+?)\s+R\$\s+([\d.,]+)\s+(\d+)%\s+R\$\s+([\d.,]+)\s+(\d\d/\d\d/\d{4})", l)
        if m:
            rows.append({"ean": m[1], "cod_forn": m[2], "desc": m[3].strip(), "venda": num(m[4]),
                         "desc_pct": int(m[5]) / 100, "oferta": num(m[6]), "validade": m[7]})
    return pd.DataFrame(rows)


def gera(lista, pdf, nome, data, dias):
    camp = le_campanha(pdf)
    ctl = pd.read_excel(lista, header=5, dtype={c: str for c in COLS_EAN})
    ctl = ctl[pd.to_numeric(ctl["Cód. interno"], errors="coerce").notna()]
    idx = {}
    for _, r in ctl.iterrows():
        for c in COLS_EAN:
            e = str(r[c] or "").strip().lstrip("0")
            if e and e != "nan":
                idx.setdefault(e, r)

    linhas = []
    for _, c in camp.iterrows():
        r = idx.get(c["ean"].lstrip("0"))
        d = {**c, "cod": None, "produto": None, "est": None, "dem": None, "pack": 1,
             "teto": None, "un_oferta": None, "qtde": 0, "obs": ""}
        if r is None:
            d["obs"] = "Fora da Lista de Compra: sem demanda/estoque para calcular"
            linhas.append(d)
            continue
        uncx = int(r["Un/Cx"])
        teto = pd.to_numeric(r[HIST], errors="coerce").max()
        ref = teto if pd.notna(teto) and teto > 0 else None
        # embalagem da oferta: unidade ou caixa fechada (o que deixa o preço unitário mais perto do histórico)
        pack = 1
        if uncx > 1 and ref and abs(c["oferta"] / uncx - ref) < abs(c["oferta"] - ref):
            pack = uncx
        un = c["oferta"] / pack
        est, dem = float(r["Est. rede"] or 0), float(r["Dem./mês"] or 0)
        nec = max(0.0, dem * dias / 30 - est)
        d.update(cod=int(r["Cód. interno"]), produto=r["Produto"], est=est, dem=dem, pack=pack,
                 teto=ref, un_oferta=round(un, 4))
        if ref is not None and un > ref + 0.005:
            d["obs"] = f"Acima do teto ({un / ref - 1:+.1%}): não comprar"
        elif nec <= 0:
            d["obs"] = f"Estoque cobre {dias} dias"
        else:
            d["qtde"] = math.ceil(nec / pack - 1e-9)
            if ref is None:
                d["obs"] = "Sem histórico de preço"
        linhas.append(d)
    t = pd.DataFrame(linhas)
    t["ordem"] = t["cod"].isna()
    t = t.sort_values(["ordem", "qtde", "desc"], ascending=[True, False, True])

    wb = Workbook()
    ws = wb.active
    ws.title = "Pedido"
    cab = ["EAN", "Cód. forn.", "Produto (campanha)", "Cód. Rocha", "Produto (Rocha)", "Est. rede (un)",
           "Dem./mês (un)", f"Necess. {dias}d (un)", "Emb. (un)", "Qtde pedido (emb.)", "Preço venda", "Desc.",
           "Preço oferta (emb.)", "Preço oferta (un)", "Teto (un)", "Total", "Validade lote", "Obs."]
    cabecalho(ws, cab, [15, 9, 40, 10, 36, 9, 9, 10, 7, 10, 10, 6, 11, 10, 10, 11, 11, 40], "6A1B9A")
    for i, (_, r) in enumerate(t.iterrows(), start=2):
        nec = None if r["dem"] is None or pd.isna(r["dem"]) else round(max(0, r["dem"] * dias / 30 - r["est"]), 1)
        q = int(r["qtde"])
        vals = [(r["ean"], "@"), (r["cod_forn"], None), (r["desc"], None), (r["cod"], None), (r["produto"], None),
                (r["est"], "0"), (r["dem"], "0.0"), (nec, "0.0"), (r["pack"], None), (q or None, None),
                (r["venda"], M), (r["desc_pct"], "0%"), (r["oferta"], M), (r["un_oferta"], M), (r["teto"], M),
                (round(q * r["oferta"], 2) if q else None, M), (r["validade"], None), (r["obs"], None)]
        fill = "C8E6C9" if q else ("FFCDD2" if "teto" in r["obs"] else None)
        for j, (v, f) in enumerate(vals, start=1):
            celula(ws, i, j, v, f, fill if j in (10, 18) else None, bold=(j == 10 and q > 0))
    n_ped = int((t["qtde"] > 0).sum())
    titulo(ws, [f"Pedido OL — {nome}",
                f"Drogarias Rocha • {data} • campanha válida até a data do arquivo • {n_ped} itens com pedido • "
                f"quantidade para {dias} dias de estoque (2 × demanda mensal − estoque)"],
           len(t), 15, 16)
    saida = AQUI / "pedidos" / "OL" / f"OL_{nome_arquivo(nome)}_{data}.xlsx"
    saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(saida)
    return t, saida


if __name__ == "__main__":
    args = sys.argv[1:]
    data, dias = f"{date.today():%d-%m-%Y}", 60
    for flag in ("-d", "--dias"):
        if flag in args:
            k = args.index(flag)
            v = args[k + 1]
            del args[k:k + 2]
            data, dias = (v, dias) if flag == "-d" else (data, int(v))
    t, saida = gera(args[0], args[1], args[2], data, dias)
    p = t[t["qtde"] > 0]
    print(f"{len(t)} itens na campanha, {t['cod'].notna().sum()} na lista, {len(p)} com pedido, "
          f"total R$ {(p['qtde'] * p['oferta']).sum():,.2f} -> {saida.name}")
    print(t[["desc", "cod", "est", "dem", "pack", "qtde", "oferta", "un_oferta", "teto", "obs"]].to_string())
