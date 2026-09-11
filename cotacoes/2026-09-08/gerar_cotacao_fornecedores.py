# -*- coding: utf-8 -*-
"""
Cotação 08/09/2026 - planilha para enviar aos fornecedores (RFQ).
Fonte: Radar Pareto 08/09/2026. Somente itens SEM OL (OL é comprada direto).
Sem histórico de compras. EANs adicionais viram LINHAS repetidas do item
(a maioria dos fornecedores ignora colunas extras); na volta, juntamos por cód.
EANs completados com o controle de 04/09 pelo cód. interno.
"""
import openpyxl, re, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

UP = "/root/.claude/uploads/506daea3-7102-5cb6-9758-517151fad432/"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saida2")
os.makedirs(OUT, exist_ok=True)

def norm_ean(v):
    if v is None:
        return None
    s = re.sub(r"\D", "", str(v)).lstrip("0")
    return s if len(s) >= 7 else None

def raw_ean(v):
    if v in (None, ""):
        return None
    s = str(v).strip()
    return s if norm_ean(s) else None

# ---- radar (controle novo) ----
wb = openpyxl.load_workbook(UP + "f22e9ec6-Radar_Pareto_20260908_102517.xlsx", data_only=True)
ws = wb["Radar_Pareto"]
HDR = [c.value for c in ws[6]]
radar = []
for r in ws.iter_rows(min_row=7, values_only=True):
    if r[1] is None:
        continue
    radar.append(dict(zip(HDR, r)))

# ---- base de EANs: SOMENTE a tabela Necessidade 08/09 (EAN princ. + adic. 1-5) ----
wbn = openpyxl.load_workbook(UP + "a3f3a40b-Necessidade_20260908_162026.xlsx", data_only=True)
wsn = wbn["Necessidade"]
hn = [c.value for c in wsn[6]]
nec_eans = {}
for r in wsn.iter_rows(min_row=7, values_only=True):
    if r[0] is None:
        continue
    d = dict(zip(hn, r))
    lst = nec_eans.setdefault(d["Cód. interno"], [])
    for e in [raw_ean(d.get("EAN princ."))] + [raw_ean(d.get(f"EAN adic. {i}")) for i in range(1, 6)]:
        if e and norm_ean(e) not in {norm_ean(x) for x in lst}:
            lst.append(e)

# ---- monta linhas: 1 por EAN (adicionais como linhas) ----
rows = []
n_itens = n_sem_ean = n_alt = 0
for d in radar:
    if d.get("OL") not in (None, ""):
        continue  # OL sai da cotação
    n_itens += 1
    eans = list(nec_eans.get(d["Cód"], []))
    if not eans:
        eans = [""]
        n_sem_ean += 1
    for j, e in enumerate(eans):
        obs = "" if j == 0 else "MESMO ITEM (EAN alternativo) — cotar 1x"
        rows.append([d["Cód"], e, d["Produto"], d.get("Fabricante") or "", d.get("Grupo") or "",
                     d.get("ABC") or "", d.get("Caixas"), d.get("Necessidade"), None, None, obs])
        if j > 0:
            n_alt += 1

# ---- planilha ----
wbo = openpyxl.Workbook()
wso = wbo.active
wso.title = "Cotacao"
HEAD = ["Cód. interno", "EAN", "Produto", "Fabricante", "Grupo", "ABC",
        "Caixas", "Unidades", "VALOR (R$/caixa)", "TOTAL", "OBS"]
wso.append(HEAD)
for row in rows:
    wso.append(row)

F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F4E79")
Y_FILL = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
widths = [12, 16, 52, 22, 18, 6, 8, 9, 14, 12, 34]
for i, w in enumerate(widths, 1):
    wso.column_dimensions[get_column_letter(i)].width = w
for c in wso[1]:
    c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
    c.fill = H_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = THIN
wso.row_dimensions[1].height = 26
n = len(rows) + 1
for r_ in wso.iter_rows(min_row=2, max_row=n, max_col=len(HEAD)):
    alt = bool(r_[10].value)
    for c in r_:
        c.font = Font(name=F, size=10)
        c.border = THIN
        if c.column in (7, 8):
            c.number_format = "0"
        if c.column in (9, 10):
            c.number_format = "#,##0.00"
            c.fill = Y_FILL
        elif alt:
            c.fill = ALT_FILL
    r_[1].number_format = "@"  # EAN como texto
wso.freeze_panes = "A2"

lg = wso.cell(row=n + 2, column=1,
              value="Preencher as colunas amarelas: VALOR (preço unitário por caixa) e TOTAL (VALOR × Caixas). "
                    "Linhas cinzas marcadas 'MESMO ITEM (EAN alternativo)' repetem o produto da linha acima com outro código de barras — "
                    "cotar apenas uma vez, no EAN que seu sistema reconhecer. Exemplo: item cód. 25403, VALOR 17,95 → TOTAL 35,90 (2 caixas).")
lg.font = Font(name=F, italic=True, size=9)

out = os.path.join(OUT, "COTACAO_FORNECEDORES_08-09-2026.xlsx")
wbo.save(out)
print(f"itens sem OL: {n_itens} | linhas geradas: {len(rows)} (EANs alternativos: {n_alt}) | sem EAN: {n_sem_ean}")
print("salvo em:", out)
