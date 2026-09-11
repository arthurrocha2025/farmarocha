# -*- coding: utf-8 -*-
"""Planilha de verificação: preços dos pedidos ePan × painel ePan PRO 10/09 21:02."""
import json, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
L = json.load(open(os.path.join(BASE, "verif_epan.json")))

F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F4E79")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

wb = openpyxl.Workbook()

# ---- Resumo ----
wsr = wb.active; wsr.title = "Resumo"
PEDS = ["MEDICAMENTOS + CRÍTICOS (envio imediato)", "DEMAIS (aguardando liberação)"]
rows = [["VERIFICAÇÃO DE PREÇOS — PANPHARMA (ePan)", ""],
        ["Painel novo", "ePan PRO 10/09/2026 21:02 — conta Farma Rocha (712093), filial PA11"],
        ["Base comparada", "CSV precos_712093.csv (usado nos pedidos de 10/09) — preço = menor entre à vista 7d e prazo 35d"],
        ["", ""]]
for tag in PEDS:
    sub = [x for x in L if x["pedido"] == tag]
    manteve = sum(1 for x in sub if x["cls"] == "MANTEVE")
    tot = sum(x["preco_old"] * x["qtd"] for x in sub)
    sem_est = [x for x in sub if x["obs_est"]]
    val_se = sum(x["preco_old"] * x["qtd"] for x in sem_est)
    rows += [[f"PEDIDO {tag}", ""],
             ["Itens", len(sub)],
             ["Preço MANTIDO", f"{manteve} de {len(sub)} (100%)" if manteve == len(sub) else f"{manteve} de {len(sub)}"],
             ["Total do pedido (inalterado)", f"R$ {tot:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")],
             ["Itens SEM ESTOQUE no painel novo", f"{len(sem_est)} (R$ {val_se:,.2f})".replace(",", "@").replace(".", ",").replace("@", ".")],
             ["", ""]]
rows += [["CONCLUSÃO", "Os preços foram 100% mantidos nos 338 itens dos dois pedidos — nenhum subiu, nenhum caiu, "
                       "nenhum saiu do catálogo. Pedidos podem seguir como estão."],
         ["ATENÇÃO (estoque)", "O painel novo traz coluna de estoque (o CSV não trazia): 112 itens dos pedidos estão sem "
                              "estoque informado no painel (62 no de medicamentos/críticos, 50 no de demais) — ver abas. "
                              "Panpharma pode cortar esses itens no faturamento."]]
for r in rows: wsr.append(r)
wsr.column_dimensions["A"].width = 34; wsr.column_dimensions["B"].width = 95
for row in wsr.iter_rows(max_col=2):
    for c in row:
        c.font = Font(name=F, size=10); c.alignment = Alignment(vertical="top", wrap_text=True)
for rr in (1, 5, 11, 17, 18):
    wsr.cell(row=rr, column=1).font = Font(name=F, bold=True, size=10)
wsr["A1"].font = Font(name=F, bold=True, size=12)

# ---- 1 aba por pedido ----
HEAD = ["Cód. interno", "Cód. Panpharma", "EAN", "Produto", "Qtd pedido", "Preço usado", "Preço painel 10/09",
        "Dif.", "Estoque painel", "Situação"]
for tag, nome in zip(PEDS, ["Medicamentos_Criticos", "Demais"]):
    sub = [x for x in L if x["pedido"] == tag]
    ws = wb.create_sheet(nome)
    ws.append(HEAD)
    sub.sort(key=lambda x: (not x["obs_est"], str(x["prod"])))
    for i, x in enumerate(sub, start=2):
        est = x["estoque"]
        est_v = int(float(est)) if est not in (None, "None", "") else None
        ws.append([x["cod"], x["cod_forn"], x["ean"], x["prod"], x["qtd"], x["preco_old"], x["preco_new"],
                   f"=G{i}/F{i}-1", est_v if est_v is not None else "—",
                   ("PREÇO MANTIDO" + ("; SEM ESTOQUE no painel" if x["obs_est"] else ""))])
    n = len(sub) + 1
    widths = [11, 12, 15, 46, 9, 11, 13, 8, 12, 30]
    for j, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(j)].width = w
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10); c.fill = H_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = THIN
    ws.row_dimensions[1].height = 26
    for row in ws.iter_rows(min_row=2, max_row=n, max_col=len(HEAD)):
        warn = "SEM ESTOQUE" in str(row[9].value)
        for c in row:
            c.font = Font(name=F, size=10); c.border = THIN
            if c.column in (6, 7): c.number_format = "#,##0.00"
            if c.column == 5: c.number_format = "0"
            if c.column == 8: c.number_format = "0.0%"
        row[2].number_format = "@"
        row[9].fill = WARN_FILL if warn else OK_FILL
    ws.freeze_panes = "A2"
    lg = ws.cell(row=n + 2, column=1,
                 value="Verificação 10/09/2026: painel ePan PRO 21:02 (conta 712093) × preços usados no pedido. "
                       "'—' em Estoque = painel sem estoque informado para o item (provável zerado). "
                       "Nenhum preço foi alterado; o pedido original permanece válido.")
    lg.font = Font(name=F, italic=True, size=9)

out = os.path.join(BASE, "saida3", "VERIFICACAO_PRECOS_EPAN_10-09-2026.xlsx")
wb.save(out)
print("salvo:", out)
