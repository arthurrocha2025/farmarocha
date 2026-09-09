# -*- coding: utf-8 -*-
"""
COTAÇÃO PARCIAL 08/09/2026 — controle + pedido-rascunho dos fornecedores já cotados.
Cotações de 04/09 (ePan, SB LOG, Tapajós, Nazária) DESCARTADAS a pedido (preços mudaram).
Fornecedor cotado até agora: DIMEC (08/09).
EANs: SOMENTE base Necessidade 08/09 (princ. + adic. 1-5), com DUN-14<->EAN-13.
Sem limiar de preços: menor preço unitário com estoque vence; histórico é só referência.
"""
import openpyxl, re, os, math
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

UP = "/root/.claude/uploads/506daea3-7102-5cb6-9758-517151fad432/"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saida2")
os.makedirs(OUT, exist_ok=True)
FORNS = ["DIMEC", "PLUSFARMA", "EPAN", "CENTROFARMA", "TAPAJOS", "NAZARIA", "MEDCENTRO", "BRASFARMA",
         "AT_SUPLEMENTOS"]
FORN_META = [("DIMEC", "08/09/2026", "EAN (todos, c/ DUN-14)"),
             ("PLUSFARMA", "08/09/2026", "EAN (pedido nº 36)"),
             ("EPAN", "09/09/2026", "EAN (catálogo Panpharma CSV)"),
             ("CENTROFARMA", "08/09/2026", "EAN (promo rede)"),
             ("TAPAJOS", "08/09/2026", "Cód. interno (planilha RFQ)"),
             ("NAZARIA", "08/09/2026", "Cód. interno (RFQ, PREÇO FINAL)"),
             ("MEDCENTRO", "08/09/2026", "Cód. interno (planilha RFQ)"),
             ("BRASFARMA", "09/09/2026", "Cód. interno (planilha RFQ)"),
             ("AT_SUPLEMENTOS", "09/09/2026", "Cód. interno (planilha RFQ)")]

def digits(v):
    if v is None: return None
    if isinstance(v, float) and v.is_integer(): v = int(v)
    s = re.sub(r"\D", "", str(v))
    return s if len(s) >= 7 else None

def ean13_check(b):
    return str((10 - sum(int(d)*(3 if i%2 else 1) for i,d in enumerate(b)) % 10) % 10)

def keys(v):
    s = digits(v)
    if not s: return set()
    out = {s.lstrip("0")}
    if len(s) == 14 and s[0] != "0":
        out.add((s[1:13] + ean13_check(s[1:13])).lstrip("0"))
    return out

# ---------- bases ----------
wbn = openpyxl.load_workbook(UP+"a3f3a40b-Necessidade_20260908_162026.xlsx", data_only=True)
wsn = wbn["Necessidade"]; hn = [c.value for c in wsn[6]]
nec_eans = {}
for r in wsn.iter_rows(min_row=7, values_only=True):
    if r[0] is None: continue
    dn = dict(zip(hn, r))
    lst = nec_eans.setdefault(dn["Cód. interno"], [])
    for e in [dn.get("EAN princ.")]+[dn.get(f"EAN adic. {i}") for i in range(1,6)]:
        if e not in (None,"") and e not in lst: lst.append(e)

wbc = openpyxl.load_workbook(UP+"b26d7f19-LISTA_DE_COMPRA_04092026__CONTROLE.xlsx", data_only=True)
wsc = wbc["Lista_de_Compra"]; hc = [c.value for c in wsc[1]]
old_uncx = {}
for r in wsc.iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    d = dict(zip(hc, r))
    old_uncx[d["Cód. interno"]] = int(d.get("Un/Cx") or 1)

# ---------- cotações recebidas ----------
wbd = openpyxl.load_workbook(UP+"f01e600f-pre_o_dimec.xlsx", data_only=True)
dimec = {}
for r in wbd["geral"].iter_rows(min_row=2, values_only=True):
    if not isinstance(r[4],(int,float)) or r[4] <= 0: continue
    rec = dict(cod=str(r[1]), prod=r[2], preco=float(r[4]), estoque=r[5] or 0)
    for k in keys(r[0]): dimec.setdefault(k, rec)

def money_br(s):
    """Converte 'R$ 2,55' / 'R$ 1.234,56' / 'R$ 18.3' em float."""
    if s is None: return None
    t = re.sub(r"[^\d,.\-]", "", str(s))
    if not t: return None
    if "," in t and "." in t: t = t.replace(".", "").replace(",", ".")
    elif "," in t: t = t.replace(",", ".")
    try: return float(t)
    except ValueError: return None

# PLUSFARMA (pedido nº 36, formato tipo SB LOG; xlsx enviado como .csv)
wbpf = openpyxl.load_workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)), "plusfarma.xlsx"), data_only=True)
plusfarma, started = {}, False
for r in wbpf["in"].iter_rows(min_row=1, values_only=True):
    if r[0] is not None and str(r[0]).strip() == "Codigo":
        started = True; continue
    if not started or r[0] is None: continue
    unit = (money_br(r[8]) or 0) + (money_br(r[9]) or 0)
    if unit <= 0: continue
    rec = dict(cod=str(r[0]), prod=str(r[2] or "").strip(), preco=unit, estoque=True)
    for k in keys(r[4]): plusfarma.setdefault(k, rec)

# ePan/Panpharma (atualização 09/09: CSV precos_712093, catálogo completo; menor entre à vista 7d e prazo 35d)
import csv
epan = {}
def _num_br(s):
    try: return float(str(s).replace(".", "").replace(",", ".")) if s not in (None, "") else None
    except ValueError: return None
with open(UP+"ddb963c0-precos_712093.csv", encoding="utf-8-sig", newline="") as fh:
    for row in csv.DictReader(fh, delimiter=";"):
        avista, prazo = _num_br(row.get("LIQ_AVISTA_7d")), _num_br(row.get("LIQ_PRAZO_35d"))
        precos = [p for p in (avista, prazo) if p and p > 0]
        if not precos: continue
        rec = dict(cod=str(row.get("CODIGO") or ""), prod=row.get("DESCRICAO"),
                   preco=min(precos), estoque=True)
        for k in keys(row.get("EAN")): epan.setdefault(k, rec)

# Centro Farma (PROMO REDE 08/09, .xls legado; exige estoque > 0)
import xlrd
wbcf = xlrd.open_workbook(UP+"06bb362f-PROMO_REDE_0809_CENTRO_FARMA.xls")
wscf = wbcf.sheet_by_index(0)
centrofarma = {}
for i in range(1, wscf.nrows):
    row = [wscf.cell_value(i, j) for j in range(7)]
    preco = row[6] if isinstance(row[6], (int, float)) else None
    if not preco or preco <= 0: continue
    est = row[5] if isinstance(row[5], (int, float)) else 0
    cod = str(int(row[0])) if isinstance(row[0], float) and row[0].is_integer() else str(row[0])
    rec = dict(cod=cod, prod=row[2], preco=float(preco), estoque=est)
    for k in keys(row[1]): centrofarma.setdefault(k, rec)

# Tapajós (RFQ 08/09 preenchida; por cód. interno; linhas de EAN alternativo unificadas pelo menor VALOR)
wbtj = openpyxl.load_workbook(UP+"cd1c3e05-TAPAJOS.xlsx", data_only=True)
_tj = {}
for r in wbtj["Cotacao"].iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    v = r[8]
    if isinstance(v, (int, float)) and v > 0:
        _tj.setdefault(r[0], []).append(float(v))
tapajos = {cod: dict(preco=min(vs), prod="", cod="",
                     obs_extra="preços divergentes entre EANs (usado menor)" if len(set(vs)) > 1 else "")
           for cod, vs in _tj.items()}

# Nazária (RFQ 08/09 devolvida com colunas próprias: ESTOQUE-UNIDADES e PREÇO FINAL por caixa)
wbnz = openpyxl.load_workbook(UP+"4a9f8d23-NAZARIA.xlsx", data_only=True)
_nz = {}
for r in wbnz["Cotacao"].iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    pf, est = r[14], r[11]
    if isinstance(pf, (int, float)) and pf > 0:
        _nz.setdefault(r[0], []).append((float(pf), float(est) if isinstance(est, (int, float)) else 0.0))
nazaria = {cod: dict(preco=min(p for p, _ in vs), estoque=max(e for _, e in vs),
                     obs_extra="preços divergentes entre EANs (usado menor)" if len({p for p, _ in vs}) > 1 else "")
           for cod, vs in _nz.items()}

# Medcentro MTF (RFQ 08/09 preenchida; por cód. interno; linhas de EAN alternativo unificadas pelo menor VALOR)
wbmc = openpyxl.load_workbook(UP+"2b14c06f-MEDCENTRO_MTF.xlsx", data_only=True)
_mc = {}
for r in wbmc["Cotacao"].iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    v = r[8]
    if isinstance(v, (int, float)) and v > 0:
        _mc.setdefault(r[0], []).append(float(v))
medcentro = {cod: dict(preco=min(vs),
                       obs_extra="preços divergentes entre EANs (usado menor)" if len(set(vs)) > 1 else "")
             for cod, vs in _mc.items()}

# Brasfarma (RFQ 09/09 preenchida; por cód. interno; linhas de EAN alternativo unificadas pelo menor VALOR)
wbbf = openpyxl.load_workbook(UP+"66f233d3-BRASFARMA.xlsx", data_only=True)
_bf = {}
for r in wbbf["Cotacao"].iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    v = r[8]
    if isinstance(v, (int, float)) and v > 0:
        _bf.setdefault(r[0], []).append(float(v))
brasfarma = {cod: dict(preco=min(vs),
                       obs_extra="preços divergentes entre EANs (usado menor)" if len(set(vs)) > 1 else "")
             for cod, vs in _bf.items()}

# AT Suplementos (RFQ 09/09 preenchida; por cód. interno; EAN alternativo unificado pelo menor VALOR)
wbat = openpyxl.load_workbook(UP+"b4653808-at_suplementos.xlsx", data_only=True)
_at = {}
for r in wbat["Cotacao"].iter_rows(min_row=2, values_only=True):
    if r[0] is None: continue
    v = r[8]
    if isinstance(v, (int, float)) and v > 0:
        _at.setdefault(r[0], []).append(float(v))
at_supl = {cod: dict(preco=min(vs),
                     obs_extra="preços divergentes entre EANs (usado menor)" if len(set(vs)) > 1 else "")
           for cod, vs in _at.items()}

def pack_candidates(nome, extra):
    cands = {1} | set(extra)
    for m in re.finditer(r"\((\d+)\s*X\s*\d+\)", str(nome), re.I): cands.add(int(m.group(1)))
    for m in re.finditer(r"C/\s*(\d+)\s*(UND|FLAC|BL)", str(nome), re.I): cands.add(int(m.group(1)))
    return cands

# ---------- radar × cotações ----------
wbr = openpyxl.load_workbook(UP+"f22e9ec6-Radar_Pareto_20260908_102517.xlsx", data_only=True)
wsr = wbr["Radar_Pareto"]; hr = [c.value for c in wsr[6]]
itens = []
for r in wsr.iter_rows(min_row=7, values_only=True):
    if r[1] is None: continue
    d = dict(zip(hr, r))
    if d["OL"] not in (None,""): continue
    ks = set()
    for src in nec_eans.get(d["Cód"], []): ks |= keys(src)
    ult = d.get("Últ. preço compra") if isinstance(d.get("Últ. preço compra"),(int,float)) and d["Últ. preço compra"]>0 else None
    refs = [d[k] for k in ("Melhor 3m","Melhor 6m","Melhor 12m") if isinstance(d.get(k),(int,float)) and d[k]>0]
    hist = min(refs) if refs else None
    med = sum(refs)/len(refs) if refs else None
    ref = ult if ult is not None else hist
    cx = d.get("Caixas") or 0
    uncx = old_uncx.get(d["Cód"], 1)
    q = {}
    def add(forn, preco, nome, cod, eleg, obs0=""):
        cands = pack_candidates(nome, [uncx] if uncx > 1 else [])
        mult = min(cands, key=lambda c: abs(preco/c - ref)) if ref else 1
        pun = preco / mult
        qtd = math.ceil(cx / mult) if mult > 1 else cx
        obs = [obs0] if obs0 else []
        if mult > 1: obs.append(f"1 emb = {mult} un")
        if ref and not (0.5 <= pun/ref <= 2.0): obs.append("conferir emb.")
        q[forn] = dict(pun=pun, preco=preco, mult=mult, qtd=qtd, cod=cod, eleg=eleg, obs="; ".join(obs))
    hit = next((dimec[k] for k in ks if k in dimec), None)
    if hit: add("DIMEC", hit["preco"], hit["prod"], hit["cod"], (hit["estoque"] or 0) > 0,
                "" if (hit["estoque"] or 0) > 0 else "sem estoque")
    hit = next((plusfarma[k] for k in ks if k in plusfarma), None)
    if hit: add("PLUSFARMA", hit["preco"], hit["prod"], hit["cod"], True)
    hit = next((epan[k] for k in ks if k in epan), None)
    if hit: add("EPAN", hit["preco"], hit["prod"], hit["cod"], True)
    hit = next((centrofarma[k] for k in ks if k in centrofarma), None)
    if hit: add("CENTROFARMA", hit["preco"], hit["prod"], hit["cod"], (hit["estoque"] or 0) > 0,
                "" if (hit["estoque"] or 0) > 0 else "sem estoque")
    if d["Cód"] in tapajos:
        tj = tapajos[d["Cód"]]
        add("TAPAJOS", tj["preco"], "", "", True, tj["obs_extra"])
    if d["Cód"] in nazaria:
        nz = nazaria[d["Cód"]]
        un_need = d.get("Necessidade") or 0
        if nz["estoque"] <= 0:
            obs0 = "sem estoque"
        elif nz["estoque"] < un_need:
            obs0 = f"estoque parcial ({int(nz['estoque'])} un)"
        else:
            obs0 = ""
        if nz["obs_extra"]:
            obs0 = (obs0 + "; " if obs0 else "") + nz["obs_extra"]
        add("NAZARIA", nz["preco"], "", "", nz["estoque"] > 0, obs0)
    if d["Cód"] in medcentro:
        mc = medcentro[d["Cód"]]
        add("MEDCENTRO", mc["preco"], "", "", True, mc["obs_extra"])
    if d["Cód"] in brasfarma:
        bf = brasfarma[d["Cód"]]
        add("BRASFARMA", bf["preco"], "", "", True, bf["obs_extra"])
    if d["Cód"] in at_supl:
        at = at_supl[d["Cód"]]
        add("AT_SUPLEMENTOS", at["preco"], "", "", True, at["obs_extra"])
    eleg = sorted([(v["pun"], f) for f, v in q.items() if v["eleg"]])
    win = eleg[0][1] if eleg else None
    cob = d.get("Cobertura(d)")
    zerado = isinstance(cob, (int, float)) and cob <= 0
    lim = 0.10 if zerado else 0.05  # reajuste tolerado: 5%; 10% se item zerado
    aceito = False
    if win:
        w = q[win]
        # referência do reajuste: última compra; sem ela, menor histórico
        if ref is None:
            status, aceito = "SEM HISTÓRICO", True
        elif hist is not None and w["pun"] <= hist + 1e-4:
            status, aceito = "OK (≤ menor hist.)", True
        elif ult is not None and w["pun"] <= ult + 1e-4:
            status, aceito = "OK (≤ últ. compra)", True
        elif w["pun"] <= ref * (1 + lim) + 1e-4:
            status, aceito = "REAJUSTE +%.1f%% (aceito, lim %d%%%s)" % (
                (w["pun"]/ref - 1) * 100, int(lim*100), " zerado" if zerado else ""), True
        else:
            status = "DESCARTADO +%.1f%% (> lim %d%%%s)" % (
                (w["pun"]/ref - 1) * 100, int(lim*100), " zerado" if zerado else "")
    else:
        status = ""
    itens.append(dict(d=d, q=q, win=win, status=status, aceito=aceito, zerado=zerado, lim=lim,
                      ult=ult, hist=hist, med=med, cx=cx,
                      ean=(nec_eans.get(d["Cód"]) or [""])[0]))

# ---------- estilos ----------
F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F4E79")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
BAD_FILL = PatternFill("solid", fgColor="F8CBAD")
WIN_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Border(*[Side(style="thin", color="BFBFBF")]*4)

def style(ws, widths, nrows, money_cols=(), int_cols=(), pct_cols=()):
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10); c.fill = H_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = THIN
    ws.row_dimensions[1].height = 26
    for row in ws.iter_rows(min_row=2, max_row=nrows, max_col=len(widths)):
        for c in row:
            c.font = Font(name=F, size=10); c.border = THIN
            if c.column in money_cols: c.number_format = "#,##0.00"
            if c.column in int_cols: c.number_format = "0"
            if c.column in pct_cols: c.number_format = "0.0%"
    ws.freeze_panes = "A2"

wbo = openpyxl.Workbook()
wsR = wbo.active; wsR.title = "Resumo"
wsC = wbo.create_sheet("Controle")

# ---------- Controle ----------
# A..O: fornecedores começam em J; K=Melhor forn., L=Melhor(un), M=Dif.%, N=Status, O=OBS
nF = len(FORNS)
FCOL0 = 11               # 1ª coluna de fornecedor (K): A..J = dados fixos (J = Média hist.)
colMF = FCOL0 + nF       # Melhor forn.
colMP = colMF + 1        # Melhor (un)
colDMED = colMP + 1      # % vs média hist.
colDULT = colDMED + 1    # % vs últ. compra
colST = colDULT + 1
colOBS = colST + 1
L = get_column_letter
CH = (["Cód. interno","EAN princ.","Produto","Grupo","ABC","Caixas","Unid.","Últ. compra","Menor hist.","Média hist."]
      + [f + " (un)" for f in FORNS]
      + ["Melhor forn.","Melhor (un)","% vs média hist.","% vs últ. compra","Status","OBS"])
wsC.append(CH)
itens.sort(key=lambda x: str(x["d"]["Produto"]))
for i, it in enumerate(itens, start=2):
    d, q = it["d"], it["q"]
    obs = "; ".join(f"{f}: {v['obs']}" for f, v in q.items() if v["obs"])
    dmed = f"={L(colMP)}{i}/J{i}-1" if (it["win"] and it["med"]) else None
    dult = f"={L(colMP)}{i}/H{i}-1" if (it["win"] and it["ult"]) else None
    wsC.append([d["Cód"], str(it["ean"]), d["Produto"], d.get("Grupo"), d.get("ABC"), it["cx"],
                d.get("Necessidade"), it["ult"], it["hist"], it["med"]]
               + [q.get(f, {}).get("pun") for f in FORNS]
               + [it["win"] or "SEM COTAÇÃO",
                  q[it["win"]]["pun"] if it["win"] else None, dmed, dult, it["status"], obs])
nC = len(itens) + 1
style(wsC, [11,15,46,16,6,8,8,10,10,10] + [10]*nF + [13,10,10,10,17,30], nC,
      money_cols=tuple([8,9,10] + list(range(FCOL0, FCOL0+nF)) + [colMP]), int_cols=(6,7),
      pct_cols=(colDMED, colDULT))
for row in wsC.iter_rows(min_row=2, max_row=nC, min_col=2, max_col=2): row[0].number_format = "@"
for i, it in enumerate(itens, start=2):
    if it["win"]:
        wsC.cell(row=i, column=FCOL0 + FORNS.index(it["win"])).fill = WIN_FILL
    s = it["status"]
    c = wsC.cell(row=i, column=colST)
    if s.startswith("OK"): c.fill = OK_FILL
    elif s.startswith("REAJUSTE"): c.fill = WARN_FILL
    elif s.startswith("DESCARTADO"): c.fill = BAD_FILL

# ---------- Pedido-rascunho por fornecedor ----------
PH = ["Cód. interno","Cód. forn.","EAN","Produto","Un/Emb","Caixas","Qtd pedido","Preço (emb.)","Total",
      "Preço un. eq.","Últ. compra","Menor hist.","Média hist.","% vs média hist.","% vs últ. compra",
      "Status","OBS"]
ped_rows = {}
for f in FORNS:
    wsP = wbo.create_sheet("Ped_"+f)
    wsP.append(PH)
    sel = [it for it in itens if it["win"] == f and it["aceito"]]
    for i, it in enumerate(sel, start=2):
        w = it["q"][f]
        dmed = f"=J{i}/M{i}-1" if it["med"] else None
        dult = f"=J{i}/K{i}-1" if it["ult"] else None
        wsP.append([it["d"]["Cód"], w["cod"], str(it["ean"]), it["d"]["Produto"], w["mult"], it["cx"],
                    w["qtd"], w["preco"], f"=G{i}*H{i}", f"=H{i}/E{i}", it["ult"], it["hist"], it["med"],
                    dmed, dult, it["status"], w["obs"]])
    n = len(sel) + 1
    style(wsP, [11,10,15,46,8,8,9,11,11,11,10,10,10,10,10,17,26], n,
          money_cols=(8,9,10,11,12,13), int_cols=(5,6,7), pct_cols=(14,15))
    for row in wsP.iter_rows(min_row=2, max_row=n, min_col=3, max_col=3): row[0].number_format = "@"
    wsP.cell(row=n+1, column=4, value="TOTAL").font = Font(name=F, bold=True, size=10)
    tc = wsP.cell(row=n+1, column=9, value=f"=SUM(I2:I{n})")
    tc.font = Font(name=F, bold=True, size=10); tc.number_format = "#,##0.00"
    ped_rows[f] = n + 1

# ---------- Descartados por preço demasiado ----------
wsD = wbo.create_sheet("Descartados_Preco", 2)
DH = ["Cód. interno","EAN","Produto","Grupo","Caixas","Zerado?","Limite","Melhor forn.","Melhor (un)",
      "Últ. compra","Menor hist.","Média hist.","% vs últ. compra","% vs média hist.","OBS"]
wsD.append(DH)
desc = [it for it in itens if it["win"] and not it["aceito"]]
desc.sort(key=lambda x: str(x["d"]["Produto"]))
for i, it in enumerate(desc, start=2):
    w = it["q"][it["win"]]
    dult = f"=I{i}/J{i}-1" if it["ult"] else None
    dmed = f"=I{i}/L{i}-1" if it["med"] else None
    wsD.append([it["d"]["Cód"], str(it["ean"]), it["d"]["Produto"], it["d"].get("Grupo"), it["cx"],
                "SIM" if it["zerado"] else "NÃO", f"{int(it['lim']*100)}%", it["win"], w["pun"],
                it["ult"], it["hist"], it["med"], dult, dmed, w["obs"]])
nD = len(desc) + 1
style(wsD, [11,15,46,16,8,8,8,13,10,10,10,10,10,10,26], nD,
      money_cols=(9,10,11,12), int_cols=(5,), pct_cols=(13,14))
for row in wsD.iter_rows(min_row=2, max_row=nD, min_col=2, max_col=2): row[0].number_format = "@"
lgD = wsD.cell(row=nD+2, column=1,
               value="Itens com cotação vencedora ACIMA do limite de reajuste (5% sobre a referência; 10% quando o item "
                     "está zerado — Cobertura = 0 no Radar). Referência = última compra; sem ela, menor histórico. "
                     "Ficam FORA dos pedidos-rascunho até renegociação ou nova cotação.")
lgD.font = Font(name=F, italic=True, size=9)

# ---------- Resumo ----------
wsR["A1"] = "COTAÇÃO 08/09/2026 — DROGARIAS ROCHA — PARCIAL (em andamento)"
wsR["A1"].font = Font(name=F, bold=True, size=13)
wsR["A2"] = ("Base: Radar Pareto 08/09 (itens sem OL) • EANs: tabela Necessidade 08/09 • Sem limiar de preços — "
             "pedidos abaixo são RASCUNHO. Cotações de 04/09 (ePan, SB LOG, Tapajós, Nazária) descartadas: preços mudaram.")
wsR["A2"].font = Font(name=F, italic=True, size=9)

wsR["A4"] = "Visão geral"; wsR["A4"].font = Font(name=F, bold=True, size=11)
cMF, cST = L(colMF), L(colST)
geral = [
    ("Itens em cotação (Radar sem OL)", len(itens)),
    ("Itens com ≥ 1 cotação", f'=COUNTIF(Controle!{cMF}2:{cMF}{nC},"<>SEM COTAÇÃO")'),
    ("Itens ainda sem cotação", f'=COUNTIF(Controle!{cMF}2:{cMF}{nC},"SEM COTAÇÃO")'),
    ("Itens OK (≤ menor hist. ou últ. compra)", f'=COUNTIF(Controle!{cST}2:{cST}{nC},"OK*")'),
    ("Reajuste aceito (até 5%; 10% se zerado)", f'=COUNTIF(Controle!{cST}2:{cST}{nC},"REAJUSTE*")'),
    ("Descartados por preço (> limite)", f'=COUNTIF(Controle!{cST}2:{cST}{nC},"DESCARTADO*")'),
]
r0 = 5
for j, (lab, val) in enumerate(geral):
    wsR.cell(row=r0+j, column=1, value=lab).font = Font(name=F, size=10)
    wsR.cell(row=r0+j, column=2, value=val).font = Font(name=F, bold=True, size=10)

r1 = r0 + len(geral) + 2
wsR.cell(row=r1, column=1, value="Detalhe por fornecedor").font = Font(name=F, bold=True, size=11)
hdr2 = ["Fornecedor","Data cotação","Base de casamento","Itens cotados","No pedido","Valor rascunho (R$)"]
for j, h in enumerate(hdr2, 1):
    c = wsR.cell(row=r1+1, column=j, value=h)
    c.font = Font(name=F, bold=True, color="FFFFFF", size=10); c.fill = H_FILL; c.border = THIN
for j, (f, dt, base) in enumerate(FORN_META):
    rr = r1 + 2 + j
    col = L(FCOL0 + FORNS.index(f))
    wsR.cell(row=rr, column=1, value=f)
    wsR.cell(row=rr, column=2, value=dt)
    wsR.cell(row=rr, column=3, value=base)
    wsR.cell(row=rr, column=4, value=f"=COUNT(Controle!{col}2:{col}{nC})")
    wsR.cell(row=rr, column=5,
             value=f'=COUNTIFS(Controle!{cMF}2:{cMF}{nC},"{f}",Controle!{cST}2:{cST}{nC},"<>DESCARTADO*")')
    wsR.cell(row=rr, column=6, value=f"=Ped_{f}!I{ped_rows[f]}")
    for cc in range(1, 7):
        cell = wsR.cell(row=rr, column=cc); cell.font = Font(name=F, size=10); cell.border = THIN
        if cc == 6: cell.number_format = "#,##0.00"
rr = r1 + 1 + len(FORN_META)
rr += 1
wsR.cell(row=rr, column=1, value="TOTAL").font = Font(name=F, bold=True, size=10)
for cc, colL2 in ((4, "D"), (5, "E"), (6, "F")):
    c = wsR.cell(row=rr, column=cc, value=f"=SUM({colL2}{r1+2}:{colL2}{rr-1})")
    c.font = Font(name=F, bold=True, size=10); c.border = THIN
    if cc == 6: c.number_format = "#,##0.00"
wsR.cell(row=rr+1, column=1, value="Aguardando cotação").font = Font(name=F, bold=True, size=11)
wsR.cell(row=rr+2, column=1,
         value="Demais fornecedores: planilha COTACAO_FORNECEDORES_08-09-2026.xlsx enviada; conforme voltarem preenchidas, "
               "entram neste comparativo.").font = Font(name=F, italic=True, size=9)
wsR.cell(row=rr+4, column=1,
         value="Abas: Controle = todos os itens sem OL, preço unitário equivalente por fornecedor (célula azul = vencendo); "
               "Ped_* = rascunho do pedido de cada fornecedor cotado. Nada será fechado sem sua ordem.").font = Font(name=F, italic=True, size=9)
wsR.column_dimensions["A"].width = 42
for L, w in (("B",14),("C",22),("D",13),("E",11),("F",16)): wsR.column_dimensions[L].width = w

out = os.path.join(OUT, "COTACAO_PARCIAL_08-09-2026.xlsx")
wbo.save(out)

from collections import Counter
cnt = Counter(it["win"] or "SEM COTAÇÃO" for it in itens)
print("itens:", len(itens))
for f in FORNS:
    sel = [it for it in itens if it["win"] == f and it["aceito"]]
    tot = sum(it["q"][f]["preco"] * it["q"][f]["qtd"] for it in sel)
    print(f"  {f:8s} cotados={sum(1 for it in itens if f in it['q']):4d} no_pedido={len(sel):4d} valor=R$ {tot:,.2f}")
print("  sem cotação:", cnt.get("SEM COTAÇÃO", 0))
ok = sum(1 for it in itens if it["status"].startswith("OK"))
rj = sum(1 for it in itens if it["status"].startswith("REAJUSTE"))
dc = sum(1 for it in itens if it["status"].startswith("DESCARTADO"))
print(f"  status: OK={ok} REAJUSTE={rj} DESCARTADO={dc}")
print("salvo:", out)
