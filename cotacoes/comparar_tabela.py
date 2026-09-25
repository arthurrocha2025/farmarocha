"""Compara a tabela de preços de um distribuidor (CSV) com a Lista de Compra.

A comparação é feita por embalagem de compra: custo atual = Custo un x Un/Cx
(ex.: display de Aspirina com 50 cartelas), quantidade = Caixas.
Quando o produto tem mais de um EAN, vale o EAN com menor preço líquido.
Diferença muito grande (preço < 50% ou > 200% do custo) costuma ser EAN de
embalagem diferente e fica marcada para conferir, fora dos totais.

Uso: python comparar_tabela.py <Lista_de_Compra.xlsx> <precos.csv> "<Nome distribuidor>" [saida.xlsx]
"""
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from gerar_cotacao import carrega

LIM_BAIXO, LIM_ALTO = 0.5, 2.0


def num(s):
    return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", "."), errors="coerce")


def monta(lista, csv):
    p = pd.read_csv(csv, sep=";", dtype=str, encoding="utf-8-sig")
    for c in ["ESTOQUE", "COMPRA_AVISTA_7d", "DESC_AVISTA_%"]:
        p[c] = num(p[c])
    p["k"] = p["EAN"].str.strip().str.lstrip("0")

    d = pd.DataFrame(carrega(lista)[0])
    d["k"] = d["ean"].str.lstrip("0")
    ctl = pd.read_excel(lista, header=5)
    ctl = ctl[pd.to_numeric(ctl["Cód. interno"], errors="coerce").notna()].copy()
    ctl["cod"] = ctl["Cód. interno"].astype(int)

    m = d[d["k"] != ""].merge(p, on="k", how="inner")
    m = m.merge(ctl[["cod", "Custo un", "Un/Cx"]], on="cod")
    m = m[m["COMPRA_AVISTA_7d"] > 0].copy()
    m["liq"] = m["COMPRA_AVISTA_7d"] * (1 - m["DESC_AVISTA_%"].fillna(0) / 100)
    m["r"] = m["liq"] / (m["Custo un"] * m["Un/Cx"])
    m["ok"] = m["r"].between(LIM_BAIXO, LIM_ALTO)
    # melhor EAN por produto: primeiro os comparáveis, depois menor preço
    m = m.sort_values(["cod", "ok", "liq"], ascending=[True, False, True])
    melhor = m.drop_duplicates("cod").set_index("cod")
    achados = set(d[d["k"] != ""].merge(p, on="k")["cod"])

    rows = []
    for _, c in ctl.sort_values("Produto").iterrows():
        cod = c["cod"]
        r = {"cod": cod, "produto": c["Produto"], "grupo": c["Grupo"], "abc": c["ABC"],
             "qtde": int(c["Caixas"]), "uncx": int(c["Un/Cx"]),
             "custo": round(float(c["Custo un"]) * int(c["Un/Cx"]), 4), "forn": c["Fornecedor"],
             "ean": None, "desc": None, "tab": None, "descp": None, "est": None, "obs": ""}
        if cod in melhor.index:
            b = melhor.loc[cod]
            r.update(ean=b["EAN"], desc=b["DESCRICAO"], tab=b["COMPRA_AVISTA_7d"],
                     descp=(b["DESC_AVISTA_%"] or 0) / 100,
                     est=None if pd.isna(b["ESTOQUE"]) else int(b["ESTOQUE"]))
            if not b["ok"]:
                r["obs"] = "Conferir embalagem/EAN"
        elif cod in achados:
            r["obs"] = "Sem preço"
        else:
            r["obs"] = "Não encontrado"
        rows.append(r)
    return rows


def gera(rows, nome, saida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo"
    F = "Arial"
    azul = "1565C0"
    fina = Side(style="thin", color="BFBFBF")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)
    fonte = Font(name=F, size=10)
    moeda = '#,##0.00;[Red]-#,##0.00;"-"'

    ws["A1"] = f"COMPARATIVO DE PREÇOS — {nome}"
    ws["A1"].font = Font(name=F, size=16, bold=True, color=azul)
    ws["A2"] = ("Base: embalagem de compra (custo atual = Custo un × Un/Cx do controle). "
                "Preço líquido = preço tabela à vista × (1 − desconto). Tabela sem ST informada (LIQ_COM_ST vazio).")
    ws["A2"].font = Font(name=F, size=9, italic=True, color="595959")

    H = 13
    ini, fim = H + 1, H + len(rows)
    rng = lambda col: f"{col}{ini}:{col}{fim}"
    resumo = [
        ("Produtos na lista", f"=COUNTA({rng('A')})", "0"),
        ("Com preço comparável", f'=COUNTIFS({rng("M")},"<>",{rng("R")},"")', "0"),
        ("  • mais baratos que o custo atual", f'=COUNTIF({rng("Q")},"Mais barato")', "0"),
        ("  • mais caros que o custo atual", f'=COUNTIF({rng("Q")},"Mais caro")', "0"),
        ("Conferir embalagem / sem preço / não encontrado",
         f'=COUNTIF({rng("R")},"Conferir*")&" / "&COUNTIF({rng("R")},"Sem preço")&" / "&COUNTIF({rng("R")},"Não encontrado")', "@"),
        ("Economia potencial comprando só os mais baratos (R$)", f'=SUMIF({rng("P")},">0")', moeda),
        ("Saldo se comprar todos os comparáveis aqui (R$)", f"=SUM({rng('P')})", moeda),
    ]
    for i, (lbl, f, fmt) in enumerate(resumo, start=4):
        ws[f"A{i}"] = lbl
        ws[f"A{i}"].font = Font(name=F, size=10, bold=not lbl.startswith("  "))
        ws[f"E{i}"] = f
        ws[f"E{i}"].font = Font(name=F, size=10, bold=True)
        ws[f"E{i}"].number_format = fmt
        ws[f"E{i}"].alignment = Alignment(horizontal="right")

    cab = ["Cód. Rocha", "Produto", "Grupo", "ABC", "Qtde (emb.)", "Un/emb.", "Custo atual (emb.)",
           "Fornecedor atual", "EAN cotado", "Descrição no distribuidor", "Preço tabela", "Desc. %",
           "Preço líquido", "Estoque distrib.", "Dif. % vs custo", "Economia (R$)", "Situação", "Obs."]
    larg = [10, 44, 14, 5, 8, 7, 11, 26, 15, 44, 10, 7, 10, 9, 9, 11, 12, 22]
    for j, (t, w) in enumerate(zip(cab, larg), start=1):
        c = ws.cell(row=H, column=j, value=t)
        c.font = Font(name=F, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=azul)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[H].height = 32

    for k, r in enumerate(rows):
        n = ini + k
        vals = [r["cod"], r["produto"], r["grupo"], r["abc"], r["qtde"], r["uncx"], r["custo"], r["forn"],
                r["ean"], r["desc"], r["tab"], r["descp"],
                f'=IF(K{n}="","",ROUND(K{n}*(1-L{n}),2))', r["est"],
                f'=IF(OR(M{n}="",R{n}<>""),"",M{n}/G{n}-1)',
                f'=IF(OR(M{n}="",R{n}<>""),"",(G{n}-M{n})*E{n})',
                f'=IF(R{n}<>"","—",IF(M{n}<G{n},"Mais barato",IF(M{n}>G{n},"Mais caro","Igual")))',
                r["obs"] or None]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=n, column=j, value=v)
            c.font = fonte
            c.border = borda
        ws.cell(row=n, column=9).number_format = "@"
        for j in (7, 11, 13, 16):
            ws.cell(row=n, column=j).number_format = moeda
        ws.cell(row=n, column=12).number_format = "0.0%"
        ws.cell(row=n, column=15).number_format = '+0.0%;[Red]-0.0%;0.0%'
        for j in (1, 4, 5, 6, 9, 14, 17):
            ws.cell(row=n, column=j).alignment = Alignment(horizontal="center")

    verde, vermelho = PatternFill("solid", fgColor="C8E6C9"), PatternFill("solid", fgColor="FFCDD2")
    ws.conditional_formatting.add(rng("Q"), CellIsRule(operator="equal", formula=['"Mais barato"'], fill=verde))
    ws.conditional_formatting.add(rng("Q"), CellIsRule(operator="equal", formula=['"Mais caro"'], fill=vermelho))
    ws.conditional_formatting.add(rng("R"), FormulaRule(formula=[f'LEFT(R{ini},8)="Conferir"'],
                                                        fill=PatternFill("solid", fgColor="FFF59D")))
    ws.conditional_formatting.add(rng("N"), FormulaRule(formula=[f'AND(N{ini}<>"",N{ini}<E{ini})'],
                                                        font=Font(color="C62828", bold=True)))
    # a % de diferença fica vermelha quando o distribuidor é mais caro (sinal positivo)
    ws.conditional_formatting.add(rng("O"), CellIsRule(operator="greaterThan", formula=["0"],
                                                       font=Font(color="C62828")))
    ws.conditional_formatting.add(rng("O"), CellIsRule(operator="lessThan", formula=["0"],
                                                       font=Font(color="2E7D32")))

    ws.freeze_panes = f"C{ini}"
    ws.auto_filter.ref = f"A{H}:R{fim}"
    ws.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.save(saida)
    print(f"{len(rows)} produtos. Salvo em {saida}")


if __name__ == "__main__":
    lista, csv, nome = sys.argv[1:4]
    saida = sys.argv[4] if len(sys.argv) > 4 else f"Comparativo_{nome.replace(' ', '_')}.xlsx"
    gera(monta(lista, csv), nome, saida)
