# -*- coding: utf-8 -*-
"""
Cotação 04/09/2026 - DROGARIA ROCHA
Cruza a LISTA DE COMPRA (controle) com as cotações ePan, SB LOG, Tapajós e Nazária.
- Itens com OL: 1 arquivo por OL (sem cotação).
- Demais itens: alocados ao distribuidor de menor preço unitário (com estoque),
  validados contra o menor preço histórico (últ. compra / melhor 3m/6m/12m).
Casamento ePan/SB LOG por EAN (principal + adicionais); Tapajós/Nazária por cód. interno.
"""
import openpyxl, re, os, unicodedata
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

UP = "/root/.claude/uploads/506daea3-7102-5cb6-9758-517151fad432/"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saida")
os.makedirs(OUT, exist_ok=True)

TOL = 1e-4  # tolerância de comparação "igual ao menor preço"

def norm_ean(v):
    if v is None:
        return None
    s = re.sub(r"\D", "", str(v)).lstrip("0")
    return s if len(s) >= 7 else None

def money(s):
    if s is None:
        return None
    s = str(s).replace("R$", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

# ---------------------------------------------------------------- controle
wb = openpyxl.load_workbook(UP + "b26d7f19-LISTA_DE_COMPRA_04092026__CONTROLE.xlsx", data_only=True)
ws = wb["Lista_de_Compra"]
HDR = [c.value for c in ws[1]]
itens = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    d = dict(zip(HDR, r))
    d["_eans"] = [e for e in (norm_ean(d.get("EAN princ.")), norm_ean(d.get("EAN adic. 1")),
                              norm_ean(d.get("EAN adic. 2")), norm_ean(d.get("EAN adic. 3"))) if e]
    refs = [float(d[k]) for k in ("Últ. preço compra", "Melhor 3m", "Melhor 6m", "Melhor 12m")
            if isinstance(d.get(k), (int, float)) and d[k] > 0]
    d["_hist"] = min(refs) if refs else None
    ult = d.get("Últ. preço compra")
    d["_ult"] = float(ult) if isinstance(ult, (int, float)) and ult > 0 else None
    d["_uncx"] = int(d.get("Un/Cx") or 1)
    d["_cx"] = float(d.get("Caixas") or 0)
    d["_un"] = float(d.get("Un. compradas") or 0)
    ol = d.get("OL")
    d["_ol"] = str(ol).strip() if ol not in (None, "") else None
    # Aché não vem marcada na coluna OL do controle, mas é OL: identifica pelo
    # prefixo GS1 da Aché/Labofarma (7896658) em qualquer EAN do item.
    if d["_ol"] is None:
        eans_full = [str(d.get(k) or "") for k in ("EAN princ.", "EAN adic. 1", "EAN adic. 2", "EAN adic. 3")]
        if any(e.startswith("7896658") for e in eans_full):
            d["_ol"] = "ACHE"
    itens.append(d)

# ---------------------------------------------------------------- ePan (por EAN)
wbp = openpyxl.load_workbook(UP + "4b5193a4-ePan_Precos_20260904_1843.xlsx", data_only=True)
epan = {}
for r in wbp["Precos"].iter_rows(min_row=2, values_only=True):
    e = norm_ean(r[1])
    if not e or not isinstance(r[7], (int, float)):
        continue
    epan[e] = dict(codigo=str(r[0]), ean=str(r[1]), desc=r[2], lab=r[3],
                   preco=float(r[7]), estoque=str(r[8]).upper() == "SIM", entrega=r[9], cond=r[10])

# ---------------------------------------------------------------- SB LOG (por EAN)
wbs = openpyxl.load_workbook(UP + "3bcc1a1e-pedido_N__5518_04_09_2026_13_27_00.xlsx", data_only=True)
sblog, started = {}, False
for r in wbs["Dados"].iter_rows(min_row=1, values_only=True):
    if r[0] == "Código" and r[4] == "Cód. de Barra":
        started = True
        continue
    if not started or r[0] is None:
        continue
    e = norm_ean(r[4])
    if not e:
        continue
    unit = (money(r[7]) or 0) + (money(r[8]) or 0)
    if unit > 0:
        sblog[e] = dict(codigo=str(r[0]), ean=str(r[4]), desc=r[2], lab=r[3], preco=unit)

# ---------------------------------------------------------------- Tapajós / Nazária (por cód. interno)
wbt = openpyxl.load_workbook(UP + "7a604e05-LISTA_DE_COMPRA_04092026__COTA__O.xlsx", data_only=True)
tapajos = {}
for r in wbt.active.iter_rows(min_row=2, values_only=True):
    if r[0] is not None and isinstance(r[11], (int, float)) and r[11] > 0:
        tapajos[r[0]] = dict(preco_cx=float(r[11]))

wbn = openpyxl.load_workbook(UP + "5804846f-LISTA_DE_COMPRA_04092026__COTA__O_1.xlsx", data_only=True)
nazaria = {}
for r in wbn.active.iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    est = float(r[14]) if isinstance(r[14], (int, float)) else 0.0
    pf = float(r[17]) if isinstance(r[17], (int, float)) else 0.0
    if pf > 0:
        nazaria[r[0]] = dict(preco_cx=pf, estoque=est)

# ---------------------------------------------------------------- cotações por item
def cotacoes(d):
    """Retorna lista de dicts: fornecedor, preco_un (equivalente por unidade),
    qtd_pedido e preco_pedido (na base em que o fornecedor fatura), obs, elegível."""
    out = []
    uncx, hist = d["_uncx"], d["_hist"]

    anchor = hist if hist is not None else d["_ult"]

    def base_epan_sblog(p):
        # ePan/SB LOG cotam pela embalagem do EAN casado; p/ Un/Cx>1 detecta se o
        # preço é da caixa (≈ hist*uncx) ou da unidade (≈ hist), usando o histórico.
        if uncx > 1 and anchor:
            if abs(p / uncx - anchor) < abs(p - anchor):
                return p / uncx, "caixa", "conferir emb. (caixa c/ %d)" % uncx
        return p, "unidade", ("conferir embalagem" if uncx > 1 else "")

    ep = next((epan[e] for e in d["_eans"] if e in epan), None)
    if ep:
        pu, base, obs = base_epan_sblog(ep["preco"])
        eleg = ep["estoque"]
        if not ep["estoque"]:
            obs = (obs + "; " if obs else "") + "sem estoque no ePan"
        qtd = d["_cx"] if base == "caixa" else d["_un"]
        out.append(dict(forn="EPAN", preco_un=pu, qtd=qtd, preco=ep["preco"], eleg=eleg,
                        obs=obs, cod=ep["codigo"], ean=ep["ean"], desc=ep["desc"], extra=ep["cond"]))
    sb = next((sblog[e] for e in d["_eans"] if e in sblog), None)
    if sb:
        pu, base, obs = base_epan_sblog(sb["preco"])
        qtd = d["_cx"] if base == "caixa" else d["_un"]
        out.append(dict(forn="SB LOG", preco_un=pu, qtd=qtd, preco=sb["preco"], eleg=True,
                        obs=obs, cod=sb["codigo"], ean=sb["ean"], desc=sb["desc"], extra=""))
    tp = tapajos.get(d["Cód. interno"])
    if tp:
        out.append(dict(forn="TAPAJOS", preco_un=tp["preco_cx"] / uncx, qtd=d["_cx"],
                        preco=tp["preco_cx"], eleg=True, obs="", cod="", ean=str(d.get("EAN princ.") or ""),
                        desc=d["Produto"], extra=""))
    nz = nazaria.get(d["Cód. interno"])
    if nz:
        obs = ""
        eleg = nz["estoque"] > 0
        if not eleg:
            obs = "sem estoque na Nazária"
        elif nz["estoque"] < d["_un"]:
            obs = "estoque parcial (%d un)" % int(nz["estoque"])
        out.append(dict(forn="NAZARIA", preco_un=nz["preco_cx"] / uncx, qtd=d["_cx"],
                        preco=nz["preco_cx"], eleg=eleg, obs=obs, cod="",
                        ean=str(d.get("EAN princ.") or ""), desc=d["Produto"], extra=""))
    return out

# ---------------------------------------------------------------- alocação
FORN_ORD = {"EPAN": 0, "SB LOG": 1, "TAPAJOS": 2, "NAZARIA": 3}
aloc = {"EPAN": [], "SB LOG": [], "TAPAJOS": [], "NAZARIA": []}
comparativo, sem_cotacao, ols = [], [], {}

for d in itens:
    if d["_un"] <= 0:
        continue
    if d["_ol"]:
        ols.setdefault(d["_ol"], []).append(d)
        continue
    cots = cotacoes(d)
    d["_cots"] = cots
    eleg = sorted([c for c in cots if c["eleg"]], key=lambda c: (c["preco_un"], FORN_ORD[c["forn"]]))
    if not eleg:
        sem_cotacao.append(d)
        comparativo.append((d, cots, None))
        continue
    win = eleg[0]
    hist, ult = d["_hist"], d["_ult"]
    ref = ult if ult is not None else hist  # referência principal: compra atual
    if ref is None:
        win["status"] = "SEM HISTÓRICO"
    elif hist is not None and win["preco_un"] <= hist + TOL:
        win["status"] = "OK (≤ menor hist.)"
    elif win["preco_un"] <= ref + TOL:
        win["status"] = "OK (≤ últ. compra)"
    else:
        win["status"] = "ACIMA +%.1f%% (vs últ. compra)" % ((win["preco_un"] / ref - 1) * 100)
    aloc[win["forn"]].append((d, win))
    comparativo.append((d, cots, win))

# ---------------------------------------------------------------- formatação
F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F4E79")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
BRL = '#,##0.00'

def style_sheet(wsx, widths, nrows, money_cols=(), int_cols=()):
    for i, w in enumerate(widths, 1):
        wsx.column_dimensions[get_column_letter(i)].width = w
    for c in wsx[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
        c.fill = H_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN
    wsx.row_dimensions[1].height = 28
    for row in wsx.iter_rows(min_row=2, max_row=nrows, max_col=len(widths)):
        for c in row:
            c.font = Font(name=F, size=10)
            c.border = THIN
            if c.column in money_cols:
                c.number_format = BRL
            if c.column in int_cols:
                c.number_format = '0'
    wsx.freeze_panes = "A2"

def total_row(wsx, row, label_col, sum_cols, last_data_row):
    wsx.cell(row=row, column=label_col, value="TOTAL").font = Font(name=F, bold=True, size=10)
    for col in sum_cols:
        L = get_column_letter(col)
        c = wsx.cell(row=row, column=col, value="=SUM(%s2:%s%d)" % (L, L, last_data_row))
        c.font = Font(name=F, bold=True, size=10)
        c.number_format = BRL if col not in (0,) else BRL

def sanitize(name):
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")

# ---------------------------------------------------------------- arquivos por OL
for ol, lst in sorted(ols.items()):
    wbo = openpyxl.Workbook()
    wso = wbo.active
    wso.title = "Pedido_OL"
    hdr = ["Cód. interno", "EAN princ.", "EAN adic. 1", "EAN adic. 2", "EAN adic. 3",
           "Produto", "Grupo", "ABC", "Un/Cx", "Caixas", "Un. compradas",
           "Últ. preço compra", "Data últ. compra"]
    wso.append(hdr)
    for d in sorted(lst, key=lambda x: str(x["Produto"])):
        dt = d.get("Data últ. compra")
        wso.append([d["Cód. interno"], str(d.get("EAN princ.") or ""),
                    str(d.get("EAN adic. 1") or ""), str(d.get("EAN adic. 2") or ""),
                    str(d.get("EAN adic. 3") or ""), d["Produto"], d.get("Grupo"),
                    d.get("ABC"), d["_uncx"], d["_cx"], d["_un"],
                    d.get("Últ. preço compra") or None,
                    dt.strftime("%d/%m/%Y") if hasattr(dt, "strftime") else (dt or "")])
    n = len(lst) + 1
    style_sheet(wso, [11, 15, 13, 13, 13, 52, 12, 6, 7, 8, 12, 13, 12], n,
                money_cols=(12,), int_cols=(9, 10, 11))
    total_row(wso, n + 1, 6, (10, 11), n)
    for col in (10, 11):
        wso.cell(row=n + 1, column=col).number_format = '0'
    wbo.save(os.path.join(OUT, "OL_%s_04-09-2026.xlsx" % sanitize(ol)))

# ---------------------------------------------------------------- pedidos por distribuidor
PED_HDR = ["Cód. distribuidor", "Cód. interno", "EAN", "Produto (distribuidor)",
           "Produto (controle)", "Un/Cx", "Qtd pedido", "Un. totais", "Preço cotado",
           "Total", "Preço un. equiv.", "Últ. compra (un)", "Menor hist. (un)", "Dif. %",
           "Status", "Obs"]

def write_pedido(fname, rows, forn):
    wbo = openpyxl.Workbook()
    wso = wbo.active
    wso.title = "Pedido"
    wso.append(PED_HDR)
    rows = sorted(rows, key=lambda t: str(t[0]["Produto"]))
    for i, (d, w) in enumerate(rows, start=2):
        hist, ult = d["_hist"], d["_ult"]
        dif = ("=K%d/L%d-1" % (i, i)) if ult else (("=K%d/M%d-1" % (i, i)) if hist else None)
        wso.append([w["cod"], d["Cód. interno"], w["ean"], w["desc"] if w["desc"] != d["Produto"] else "",
                    d["Produto"], d["_uncx"], w["qtd"], d["_un"], w["preco"],
                    "=G%d*I%d" % (i, i), w["preco_un"], ult, hist,
                    dif, w["status"], w["obs"]])
    n = len(rows) + 1
    style_sheet(wso, [13, 11, 15, 42, 42, 7, 9, 9, 11, 12, 11, 11, 11, 9, 20, 26], n,
                money_cols=(9, 10, 11, 12, 13), int_cols=(6, 7, 8))
    for row in wso.iter_rows(min_row=2, max_row=n, min_col=14, max_col=14):
        row[0].number_format = '0.0%'
    for row in wso.iter_rows(min_row=2, max_row=n, min_col=15, max_col=15):
        v = str(row[0].value or "")
        row[0].fill = OK_FILL if v.startswith("OK") else (WARN_FILL if v.startswith("ACIMA") else PatternFill())
    total_row(wso, n + 1, 5, (10,), n)
    # legenda
    lg = wso.cell(row=n + 3, column=1, value=("Preços da cotação %s de 04/09/2026. 'Menor hist. (un)' = menor entre melhores 3/6/12 meses; 'Últ. compra (un)' = compra atual. "
                                              "Status OK quando o preço unit. equivalente é ≤ menor histórico ou ≤ última compra. "
                                              "'Qtd pedido' na base de faturamento do distribuidor (caixa ou unidade); 'Un. totais' = unidades da lista de compra.") % forn)
    lg.font = Font(name=F, italic=True, size=9)
    wbo.save(os.path.join(OUT, fname))

write_pedido("PEDIDO_EPAN_04-09-2026.xlsx", aloc["EPAN"], "ePan")
write_pedido("PEDIDO_SBLOG_04-09-2026.xlsx", aloc["SB LOG"], "SB LOG")
write_pedido("PEDIDO_TAPAJOS_04-09-2026.xlsx", aloc["TAPAJOS"], "Tapajós")
write_pedido("PEDIDO_NAZARIA_04-09-2026.xlsx", aloc["NAZARIA"], "Nazária")

# ---------------------------------------------------------------- comparativo geral
wbc = openpyxl.Workbook()
wsr = wbc.active
wsr.title = "Resumo"
wsc = wbc.create_sheet("Comparativo")
wsx = wbc.create_sheet("Sem_Cotacao")

CMP_HDR = ["Cód. interno", "EAN princ.", "Produto", "Un/Cx", "Caixas", "Un. compradas",
           "Últ. compra (un)", "Menor hist. (un)", "ePan (un)", "SB LOG (un)", "Tapajós (un)",
           "Nazária (un)", "Vencedor", "Preço venc. (un)", "Dif. vs últ. compra", "Status", "Obs"]
wsc.append(CMP_HDR)
r = 1
for d, cots, win in sorted(comparativo, key=lambda t: str(t[0]["Produto"])):
    r += 1
    p = {c["forn"]: c for c in cots}
    obs = "; ".join(sorted({c["obs"] for c in cots if c["obs"]}))
    hist, ult = d["_hist"], d["_ult"]
    dif = None
    if win and ult:
        dif = "=N%d/G%d-1" % (r, r)
    elif win and hist:
        dif = "=N%d/H%d-1" % (r, r)
    wsc.append([d["Cód. interno"], str(d.get("EAN princ.") or ""), d["Produto"], d["_uncx"],
                d["_cx"], d["_un"], ult, hist,
                p.get("EPAN", {}).get("preco_un"), p.get("SB LOG", {}).get("preco_un"),
                p.get("TAPAJOS", {}).get("preco_un"), p.get("NAZARIA", {}).get("preco_un"),
                win["forn"] if win else "SEM COTAÇÃO",
                win["preco_un"] if win else None,
                dif,
                win["status"] if win else "", obs])
n = r
style_sheet(wsc, [11, 15, 46, 7, 7, 9, 11, 11, 10, 10, 10, 10, 13, 11, 10, 20, 30], n,
            money_cols=(7, 8, 9, 10, 11, 12, 14), int_cols=(4, 5, 6))
for row in wsc.iter_rows(min_row=2, max_row=n, min_col=15, max_col=15):
    row[0].number_format = '0.0%'
for row in wsc.iter_rows(min_row=2, max_row=n, min_col=16, max_col=16):
    v = str(row[0].value or "")
    row[0].fill = OK_FILL if v.startswith("OK") else (WARN_FILL if v.startswith("ACIMA") else PatternFill())

wsx.append(["Cód. interno", "EAN princ.", "Produto", "Un/Cx", "Caixas", "Un. compradas",
            "Menor hist. (un)", "Motivo"])
for d in sorted(sem_cotacao, key=lambda x: str(x["Produto"])):
    reasons = [c["obs"] for c in d.get("_cots", []) if c["obs"]]
    wsx.append([d["Cód. interno"], str(d.get("EAN princ.") or ""), d["Produto"], d["_uncx"],
                d["_cx"], d["_un"], d["_hist"],
                "; ".join(reasons) if reasons else "nenhum distribuidor cotou"])
style_sheet(wsx, [11, 15, 46, 7, 7, 9, 11, 40], len(sem_cotacao) + 1,
            money_cols=(7,), int_cols=(4, 5, 6))

# Resumo
wsr["A1"] = "COTAÇÃO 04/09/2026 — DROGARIA ROCHA — RESUMO"
wsr["A1"].font = Font(name=F, bold=True, size=13)
wsr["A3"] = "Distribuidor"; wsr["B3"] = "Itens"; wsr["C3"] = "Valor total (R$)"
for c in ("A3", "B3", "C3"):
    wsr[c].font = Font(name=F, bold=True, color="FFFFFF"); wsr[c].fill = H_FILL; wsr[c].border = THIN
row = 4
for forn, nome in (("EPAN", "ePan"), ("SB LOG", "SB LOG"), ("TAPAJOS", "Tapajós"), ("NAZARIA", "Nazária")):
    wsr.cell(row=row, column=1, value=nome)
    wsr.cell(row=row, column=2, value='=COUNTIF(Comparativo!M:M,"%s")' % forn)
    wsr.cell(row=row, column=3, value=sum(w["preco"] * w["qtd"] for _, w in aloc[forn]))
    row += 1
wsr.cell(row=row, column=1, value="Sem cotação")
wsr.cell(row=row, column=2, value='=COUNTIF(Comparativo!M:M,"SEM COTAÇÃO")')
row += 2
wsr.cell(row=row, column=1, value="Pedidos OL (arquivos separados, sem cotação)").font = Font(name=F, bold=True)
row += 1
for ol, lst in sorted(ols.items()):
    wsr.cell(row=row, column=1, value=ol)
    wsr.cell(row=row, column=2, value=len(lst))
    row += 1
for rw in wsr.iter_rows(min_row=3, max_row=row, max_col=3):
    for c in rw:
        if c.font.name != F or not c.font.bold:
            c.font = Font(name=F, size=10, bold=c.font.bold, color=c.font.color)
        c.border = THIN if c.row <= row else c.border
wsr.column_dimensions["A"].width = 42; wsr.column_dimensions["B"].width = 10; wsr.column_dimensions["C"].width = 16
for rw in wsr.iter_rows(min_row=4, max_row=7, min_col=3, max_col=3):
    rw[0].number_format = BRL
wsr.cell(row=row + 1, column=1,
         value="Obs: valores 'Valor total' calculados sobre preço cotado × qtd na base de faturamento de cada distribuidor.").font = Font(name=F, italic=True, size=9)

wbc.save(os.path.join(OUT, "COMPARATIVO_COTACAO_04-09-2026.xlsx"))

# ---------------------------------------------------------------- log
print("== RESUMO ==")
for forn in ("EPAN", "SB LOG", "TAPAJOS", "NAZARIA"):
    tot = sum(w["preco"] * w["qtd"] for _, w in aloc[forn])
    ok = sum(1 for _, w in aloc[forn] if w["status"].startswith("OK"))
    ac = sum(1 for _, w in aloc[forn] if w["status"].startswith("ACIMA"))
    print(f"{forn:8s} itens={len(aloc[forn]):3d} (OK={ok}, acima hist.={ac})  total=R$ {tot:,.2f}")
print(f"SEM COTAÇÃO: {len(sem_cotacao)} itens")
for ol, lst in sorted(ols.items()):
    print(f"OL {ol}: {len(lst)} itens")
print("Arquivos em:", OUT)
for f in sorted(os.listdir(OUT)):
    print("  ", f)
