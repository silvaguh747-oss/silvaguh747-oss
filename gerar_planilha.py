import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.hyperlink import Hyperlink

OUT = "/home/user/silvaguh747-oss/Minha-Planilha-Financeira.xlsx"
FONT = "Arial"
TEAL, TEAL2, MINT, SOFT, LINE, ORANGE, YEL = "1F6F66", "2E8B7E", "E3F1EE", "F4F7F6", "C9DAD6", "E8622C", "FFF8DC"
RED, REDBG, GRN, GRNBG, AMB, AMBBG = "C0392B", "FADBD8", "1E8449", "D5F5E3", "9A6B00", "FCF3CF"
BRL = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"-"'
BRL0 = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"R$" 0,00'
DATE = "dd/mm/yyyy"
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

def q(n): return "'" + n + "'"
def ft(b=False, s=10, c="222222", i=False): return Font(name=FONT, bold=b, size=s, color=c, italic=i)
def fl(c): return PatternFill("solid", start_color=c, end_color=c)
th = Side(style="thin", color=LINE)
BOX = Border(left=th, right=th, top=th, bottom=th)
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)

def band(ws, rng, text, size=16, color=TEAL, fc="FFFFFF", h=32, align=None):
    a = rng.split(":")[0]
    if ":" in rng: ws.merge_cells(rng)
    ws[a] = text; ws[a].font = ft(True, size, fc); ws[a].fill = fl(color)
    ws[a].alignment = align or Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[ws[a].row].height = h

def head(ws, row, col, labels, color=TEAL2):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = ft(True, 9, "FFFFFF"); c.fill = fl(color); c.alignment = AL_C; c.border = BOX
    ws.row_dimensions[row].height = 30

def inp(c, fmt=None, center=False):
    c.font = ft(False, 10, "0000FF"); c.fill = fl(YEL); c.border = BOX
    if fmt: c.number_format = fmt
    if center: c.alignment = Alignment(horizontal="center", vertical="center")

def out(c, fmt=None, b=False, color="222222", fill=None, center=False):
    c.font = ft(b, 10, color); c.border = BOX
    if fill: c.fill = fl(fill)
    if fmt: c.number_format = fmt
    if center: c.alignment = Alignment(horizontal="center", vertical="center")

def note(ws, cell, text, size=8):
    ws[cell] = text; ws[cell].font = ft(False, size, "6B7B78", True)
    ws[cell].alignment = Alignment(wrap_text=False, vertical="center")

def widths(ws, d):
    for k, v in d.items(): ws.column_dimensions[k].width = v

def link(cell, sheet, text=None):
    if text: cell.value = text
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"{q(sheet)}!A1", display=str(cell.value))
    cell.font = ft(True, 9, "FFFFFF"); cell.fill = fl(TEAL2); cell.alignment = AL_C; cell.border = BOX

def dv(ws, formula, rng, err=None, typ="list", **kw):
    d = DataValidation(type=typ, formula1=formula, allow_blank=True, **kw)
    if err: d.showErrorMessage = True; d.errorTitle = "Valor inválido"; d.error = err
    else: d.showErrorMessage = False
    ws.add_data_validation(d); d.add(rng); return d

wb = Workbook()
names = ["Início", "Dashboard", "Diário", "Receitas", "Despesas"] + MESES + ["Cartão e Dívidas", "Reserva", "Investimentos", "Metas", "Planejamento", "Simulador", "Config"]
S = {}
S["Início"] = wb.active; S["Início"].title = "Início"
for n in names[1:]: S[n] = wb.create_sheet(n)
tabc = {"Início": TEAL, "Dashboard": ORANGE, "Diário": "2E86C1", "Receitas": GRN, "Despesas": RED, "Cartão e Dívidas": "8E44AD",
        "Reserva": ORANGE, "Investimentos": "117A65", "Metas": "D4AC0D", "Planejamento": "5D6D7E", "Simulador": "AF7AC5", "Config": "7F8C8D"}
for n, ws in S.items():
    ws.sheet_properties.tabColor = tabc.get(n, TEAL2)
    ws.sheet_view.showGridLines = False

# ======================= CONFIG =======================
C = S["Config"]
band(C, "A1:L1", "⚙️  CONFIGURAÇÕES", 16)
note(C, "A2", "Edite só as células amarelas (texto azul). Nada aqui exige mexer em fórmulas.", 9)
rows = [(3, "Seu nome", "Seu Nome"), (4, "Ano de referência", "Atual"), (5, "Mês de referência", "Atual"),
        (8, "Meta de economia (% da renda em dinheiro)", 0.20)]
for r, l, v in rows:
    C[f"A{r}"] = l; C[f"A{r}"].font = ft(True); C[f"B{r}"] = v; inp(C[f"B{r}"], "0%" if r == 8 else None, True)
C["A6"] = "Ano em uso (automático)"; C["B6"] = '=IF(B4="Atual",YEAR(TODAY()),B4)'
C["A7"] = "Mês em uso (automático)"; C["B7"] = '=IF(B5="Atual",MONTH(TODAY()),MATCH(B5,$L$11:$L$22,0))'
for r in (6, 7): C[f"A{r}"].font = ft(True, 10, "6B7B78"); out(C[f"B{r}"], "0", True, center=True)
note(C, "C4", "“Atual” = ano de hoje. Ou escolha 2025, 2026, 2027…")
note(C, "C5", "“Atual” = mês de hoje. O Dashboard mostra este mês.")
dv(C, '"Atual,2025,2026,2027,2028,2029,2030"', "B4")
dv(C, '"Atual,' + ",".join(MESES) + '"', "B5")

# receitas recorrentes A10:D18
head(C, 10, 1, ["Receitas recorrentes (todo mês)", "Tipo", "Valor mensal", "Ativa?"])
rec = [("Salário", "Dinheiro", 1500), ("Vale-alimentação", "Benefício", 890), ("Vale-refeição", "Benefício", 400)]
for i in range(8):
    r = 11 + i
    if i < len(rec):
        C[f"A{r}"], C[f"B{r}"], C[f"C{r}"], C[f"D{r}"] = rec[i][0], rec[i][1], rec[i][2], "Sim"
    for col in "ABCD": inp(C[f"{col}{r}"], BRL if col == "C" else None, col in "BD")
dv(C, '"Dinheiro,Benefício"', "B11:B18"); dv(C, '"Sim,Não"', "D11:D18")
note(C, "A19", "Salário = dinheiro em conta. Benefícios (VA/VR) ficam separados e NÃO entram no dinheiro disponível.")
SAL = "Config!$C$11"
RR_V, RR_T, RR_A = "Config!$C$11:$C$18", "Config!$B$11:$B$18", "Config!$D$11:$D$18"

# categorias A22:C50 (receita 23-34, despesa 35-50)
head(C, 22, 1, ["Categorias", "Tipo", "Orçamento mensal (limite)"])
rcats = ["Salário", "Vale-alimentação", "Vale-refeição", "Hora extra", "Adiantamento salarial", "Décimo terceiro", "Férias",
         "Renda extra", "Bonificação", "Outras receitas"]
dcats = [("Moradia", 0), ("Alimentação", 600), ("Transporte", 200), ("Educação", 600), ("Saúde", 150), ("Assinaturas", 120),
         ("Lazer", 150), ("Compras", 150), ("Cartão de crédito", 0), ("Dívidas e financiamentos", 0), ("Outros", 100)]
for i in range(12):
    r = 23 + i; C[f"A{r}"] = rcats[i] if i < len(rcats) else None; C[f"B{r}"] = "Receita"
    inp(C[f"A{r}"]); out(C[f"B{r}"], color="6B7B78", fill=SOFT, center=True); out(C[f"C{r}"], fill=SOFT)
for i in range(16):
    r = 35 + i
    if i < len(dcats): C[f"A{r}"], C[f"C{r}"] = dcats[i]
    C[f"B{r}"] = "Despesa"; inp(C[f"A{r}"]); inp(C[f"C{r}"], BRL)
    out(C[f"B{r}"], color="6B7B78", fill=SOFT, center=True)
note(C, "A51", "Orçamento = limite mensal da categoria (0 = sem limite). Serve para os alertas de gasto excessivo.")
note(C, "A52", "Atenção: “Adiantamento salarial” (linha 27) é tratado como parte do salário, para não contar duas vezes.")
ADIANT = "Config!$A$27"
CATS_ALL = "Config!$A$23:$A$50"; CATS_REC = "Config!$A$23:$A$34"; CATS_DES = "Config!$A$35:$A$50"

# formas F10:F22
head(C, 10, 6, ["Formas de pagamento"])
formas = ["Cartão de crédito", "Vale-alimentação", "Vale-refeição", "Pix", "Cartão de débito", "Dinheiro", "Boleto", "Transferência", "Débito automático"]
for i in range(12):
    r = 11 + i; C[f"F{r}"] = formas[i] if i < len(formas) else None; inp(C[f"F{r}"])
note(C, "F23", "As 3 primeiras têm função especial: crédito vai para a fatura; VA/VR descontam dos benefícios.")
CRED, VA, VR = "Config!$F$11", "Config!$F$12", "Config!$F$13"
FORMAS = "Config!$F$11:$F$22"
head(C, 10, 8, ["Tipos de investimento"])
inv_t = ["CDB", "Tesouro Direto", "Renda fixa", "Ações", "Fundos imobiliários", "Outros"]
for i in range(10):
    r = 11 + i; C[f"H{r}"] = inv_t[i] if i < len(inv_t) else None; inp(C[f"H{r}"])
INVT = "Config!$H$11:$H$20"
head(C, 10, 10, ["Contas recorrentes"])
C["J11"] = "→ cadastradas na aba Despesas"; C["J11"].font = ft(False, 9, "6B7B78", True)
C["J11"].hyperlink = Hyperlink(ref="J11", location="'Despesas'!A1")
head(C, 10, 12, ["Meses"])
for i, m in enumerate(MESES): C[f"L{11+i}"] = m; out(C[f"L{11+i}"], color="6B7B78", fill=SOFT)
MESL = "Config!$L$11:$L$22"
widths(C, {"A": 40, "B": 14, "C": 22, "D": 10, "E": 3, "F": 22, "G": 3, "H": 22, "I": 3, "J": 26, "K": 3, "L": 14})
YR, MM = "Config!$B$6", "Config!$B$7"

# ======================= DIÁRIO =======================
D = S["Diário"]
band(D, "A1:G1", "📝  CONTROLE DIÁRIO", 16)
note(D, "A2", "Registre aqui cada gasto do dia a dia e entradas avulsas. Tudo alimenta o mês, o Dashboard e os relatórios.", 9)
note(D, "A4", "Salário e VA/VR já entram sozinhos (aba Config): não lance aqui. Escolha sempre uma categoria da lista.", 9)
note(D, "A3", "NÃO lance aqui: contas fixas (marque ☑ na aba do mês), parcelas (aba Cartão e Dívidas) nem pagamento de fatura. Assim nada é contado 2 vezes.", 9)
head(D, 5, 1, ["Data", "Descrição", "Categoria", "Valor (R$)", "Entrada ou Saída", "Forma de pagamento", "Observações"])
DF_, DL_ = 6, 2005
ex = [((2026, 11, 2), "Mercado do mês", "Alimentação", 320.40, "Saída", "Vale-alimentação"),
      ((2026, 11, 3), "Almoço", "Alimentação", 32.90, "Saída", "Vale-refeição"),
      ((2026, 11, 4), "Uber", "Transporte", 18.50, "Saída", "Pix"),
      ((2026, 11, 6), "Farmácia", "Saúde", 46.80, "Saída", "Cartão de débito"),
      ((2026, 11, 8), "Cinema", "Lazer", 60.00, "Saída", "Cartão de crédito"),
      ((2026, 11, 9), "Venda de item usado", "Renda extra", 120.00, "Entrada", "Pix")]
for r in range(DF_, DL_ + 1):
    i = r - DF_
    if i < len(ex):
        d, de, ca, v, t, f = ex[i]
        D[f"A{r}"], D[f"B{r}"], D[f"C{r}"], D[f"D{r}"], D[f"E{r}"], D[f"F{r}"], D[f"G{r}"] = dt.date(*d), de, ca, v, t, f, "EXEMPLO – pode apagar"
    for col in "ABCDEFG":
        c = D[f"{col}{r}"]; c.font = ft(False, 10, "0000FF"); c.border = Border(bottom=th)
    D[f"A{r}"].number_format = DATE; D[f"D{r}"].number_format = BRL
dv(D, f"={CATS_ALL}", f"C{DF_}:C{DL_}")
dv(D, '"Entrada,Saída"', f"E{DF_}:E{DL_}")
dv(D, f"={FORMAS}", f"F{DF_}:F{DL_}")
dv(D, "0", f"D{DF_}:D{DL_}", "Use valor positivo. Diga se é Entrada ou Saída na coluna E.", typ="decimal", operator="greaterThanOrEqual")
dv(D, "36526", f"A{DF_}:A{DL_}", "Digite uma data. Ex.: 11/11/2026", typ="date", operator="greaterThan")
D.conditional_formatting.add(f"E{DF_}:E{DL_}", FormulaRule(formula=[f'E{DF_}="Entrada"'], font=Font(name=FONT, color=GRN, bold=True), fill=fl(GRNBG)))
D.conditional_formatting.add(f"E{DF_}:E{DL_}", FormulaRule(formula=[f'E{DF_}="Saída"'], font=Font(name=FONT, color=RED, bold=True), fill=fl(REDBG)))
D.auto_filter.ref = f"A5:G{DL_}"; D.freeze_panes = "A6"
widths(D, {"A": 12, "B": 28, "C": 22, "D": 14, "E": 14, "F": 20, "G": 24})
DA, DB, DC, DV_, DE, DFm = [f"'Diário'!${c}${DF_}:${c}${DL_}" for c in "ABCDEF"]

# ======================= RECEITAS =======================
R = S["Receitas"]
band(R, "A1:H1", "💰  MINHAS RECEITAS", 16)
note(R, "A2", "Receitas fixas vêm da aba Config. Registre abaixo horas extras, adiantamento, 13º, férias, bônus e rendas extras.", 9)
head(R, 4, 1, ["Receita recorrente", "Tipo", "Valor mensal", "Ativa?"], TEAL)
for i in range(8):
    r = 5 + i
    for j, col in enumerate("ABCD"):
        R[f"{col}{r}"] = f'=IF(Config!{col}{11+i}="","",Config!{col}{11+i})'
        out(R[f"{col}{r}"], BRL if col == "C" else None, color="008000", center=col in "BD")
R["F4"] = "Previsto todo mês"; R["F4"].font = ft(True, 9, "FFFFFF"); R["F4"].fill = fl(TEAL); R["F4"].alignment = AL_C
R.merge_cells("F4:G4")
R["F5"] = "Dinheiro em conta"; R["G5"] = f'=SUMIFS({RR_V},{RR_T},"Dinheiro",{RR_A},"Sim")'
R["F6"] = "Benefícios alimentares"; R["G6"] = f'=SUMIFS({RR_V},{RR_T},"Benefício",{RR_A},"Sim")'
R["F7"] = "Renda total"; R["G7"] = "=G5+G6"
for r in (5, 6, 7): R[f"F{r}"].font = ft(True); out(R[f"G{r}"], BRL0, r == 7)

# log de extras
RL0, RL1 = 31, 230
band(R, f"A{RL0-2}:F{RL0-2}", "Registro de receitas extras (com data)", 11, TEAL, h=22, align=AL_L)
head(R, RL0 - 1, 1, ["Data", "Tipo", "Descrição", "Valor (R$)", "Vai para", "Observações"])
R[f"A{RL0}"] = dt.date(2026, 11, 10); R[f"B{RL0}"] = "Hora extra"; R[f"C{RL0}"] = "Horas extras de outubro"; R[f"D{RL0}"] = 150
R[f"E{RL0}"] = "Dinheiro"; R[f"F{RL0}"] = "EXEMPLO – pode apagar"
for r in range(RL0, RL1 + 1):
    for col in "ABCDEF":
        c = R[f"{col}{r}"]; c.font = ft(False, 10, "0000FF"); c.border = Border(bottom=th)
    R[f"A{r}"].number_format = DATE; R[f"D{r}"].number_format = BRL
dv(R, f"={CATS_REC}", f"B{RL0}:B{RL1}"); dv(R, '"Dinheiro,Benefício"', f"E{RL0}:E{RL1}")
RA, RB, RD, RE = [f"'Receitas'!${c}${RL0}:${c}${RL1}" for c in "ABDE"]

def mrange(m):
    X = f"DATE({YR},{m},1)"; NX = f"DATE({YR},{m+1},1)"
    return X, NX
def crit(rng, m):
    X, NX = mrange(m); return f'{rng},">="&{X},{rng},"<"&{NX}'

def rec_din(m):
    return (f'SUMIFS({RR_V},{RR_T},"Dinheiro",{RR_A},"Sim")'
            f'+SUMIFS({RD},{RE},"Dinheiro",{RB},"<>"&{ADIANT},{crit(RA, m)})'
            f'+SUMIFS({DV_},{DE},"Entrada",{DFm},"<>"&{VA},{DFm},"<>"&{VR},{DC},"<>"&{ADIANT},{crit(DA, m)})')
def rec_ben(m):
    return (f'SUMIFS({RR_V},{RR_T},"Benefício",{RR_A},"Sim")+SUMIFS({RD},{RE},"Benefício",{crit(RA, m)})'
            f'+SUMIFS({DV_},{DE},"Entrada",{DFm},{VA},{crit(DA, m)})+SUMIFS({DV_},{DE},"Entrada",{DFm},{VR},{crit(DA, m)})')
def adiant(m):
    return f'SUMIFS({RD},{RB},{ADIANT},{crit(RA, m)})+SUMIFS({DV_},{DE},"Entrada",{DC},{ADIANT},{crit(DA, m)})'

band(R, "A14:H14", "Resumo mensal das receitas", 11, TEAL, h=22, align=AL_L)
head(R, 15, 1, ["Mês", "Salário-base", "Adiantamentos recebidos", "Salário restante a receber", "Extras em dinheiro", "Total em dinheiro", "Benefícios alimentares", "Renda total"])
for i, m in enumerate(MESES, start=1):
    r = 15 + i
    R[f"A{r}"] = m
    R[f"B{r}"] = f'=IF(Config!$D$11="Sim",{SAL},0)'
    R[f"C{r}"] = "=" + adiant(i)
    R[f"D{r}"] = f"=MAX(0,B{r}-C{r})"
    R[f"F{r}"] = "=" + rec_din(i)
    R[f"E{r}"] = f'=F{r}-SUMIFS({RR_V},{RR_T},"Dinheiro",{RR_A},"Sim")'
    R[f"G{r}"] = "=" + rec_ben(i)
    R[f"H{r}"] = f"=F{r}+G{r}"
    for col in "ABCDEFGH": out(R[f"{col}{r}"], None if col == "A" else BRL, col in "AH")
R["A28"] = "TOTAL"
for col in "BCDEFGH": R[f"{col}28"] = f"=SUM({col}16:{col}27)"
for col in "ABCDEFGH": out(R[f"{col}28"], None if col == "A" else BRL, True, fill=MINT)
note(R, "A29", "Adiantamento é parte do salário: ele aparece separado, mas o total em dinheiro conta o salário uma única vez.")
R.freeze_panes = "A5"
widths(R, {"A": 22, "B": 16, "C": 18, "D": 18, "E": 16, "F": 20, "G": 18, "H": 16})

# ======================= DESPESAS (cadastro) =======================
E = S["Despesas"]
band(E, "A1:I1", "🧾  MINHAS DESPESAS FIXAS (cadastro)", 16)
note(E, "A2", "Cadastre aqui as contas que se repetem. Elas aparecem sozinhas na aba de cada mês, com checkbox de pago.", 9)
note(E, "A3", "Para adicionar, use a próxima linha vazia. Para remover, apague a linha inteira de A a G (não arraste linhas, para não desalinhar os meses já marcados).", 9)
head(E, 5, 1, ["Despesa", "Categoria", "Valor previsto", "Dia do vencimento", "Forma de pagamento", "Aparece em", "Observações", "Valor real (mês do Dashboard)", "Status (mês do Dashboard)"])
EF, EL = 6, 35
desp = [("Inglês", "Educação", 370, 5, "Boleto"), ("Faculdade", "Educação", 210, 10, "Boleto"), ("ChatGPT", "Assinaturas", 100, 15, "Cartão de crédito"),
        ("Água", "Moradia", 280, 20, "Pix"), ("Internet", "Moradia", 100, 12, "Débito automático")]
for k in range(30):
    r = EF + k
    if k < len(desp):
        E[f"A{r}"], E[f"B{r}"], E[f"C{r}"], E[f"D{r}"], E[f"E{r}"] = desp[k]; E[f"F{r}"] = "Todos os meses"
        E[f"G{r}"] = "Vencimento de exemplo – ajuste"
    for col in "ABCDEFG": inp(E[f"{col}{r}"], BRL if col == "C" else None, col in "DF")
    mr = 32 + k
    E[f"H{r}"] = f'=IF(A{r}="","",CHOOSE({MM},' + ",".join(f"{q(m)}!$L${mr}" for m in MESES) + "))"
    E[f"I{r}"] = f'=IF(A{r}="","",CHOOSE({MM},' + ",".join(f"{q(m)}!$K${mr}" for m in MESES) + "))"
    out(E[f"H{r}"], BRL); out(E[f"I{r}"], center=True)
dv(E, f"={CATS_DES}", f"B{EF}:B{EL}"); dv(E, f"={FORMAS}", f"E{EF}:E{EL}")
dv(E, '"Todos os meses,' + ",".join(MESES) + '"', f"F{EF}:F{EL}")
dv(E, "1", f"D{EF}:D{EL}", "Dia entre 1 e 31", typ="whole", operator="between", formula2="31")
E[f"A{EL+1}"] = "TOTAL PREVISTO (todo mês)"; E[f"A{EL+1}"].font = ft(True)
E[f"C{EL+1}"] = f'=SUMIFS(C{EF}:C{EL},F{EF}:F{EL},"Todos os meses")'; out(E[f"C{EL+1}"], BRL, True, fill=MINT)
for txt, fc, bg in [("ATRASADA", RED, REDBG), ("Pago", GRN, GRNBG), ("Pendente", AMB, AMBBG)]:
    E.conditional_formatting.add(f"I{EF}:I{EL}", FormulaRule(formula=[f'I{EF}="{txt}"'], font=Font(name=FONT, color=fc, bold=True), fill=fl(bg)))
E.freeze_panes = "B6"
widths(E, {"A": 24, "B": 18, "C": 15, "D": 12, "E": 20, "F": 18, "G": 28, "H": 18, "I": 16})

# ======================= CARTÃO E DÍVIDAS =======================
K = S["Cartão e Dívidas"]
CD = q("Cartão e Dívidas")
band(K, "A1:J1", "💳  CARTÃO DE CRÉDITO E DÍVIDAS", 16)
note(K, "A2", "Parcelas lançadas aqui entram sozinhas em cada mês. Não registre a mesma compra no Diário nem o pagamento da fatura.", 9)
K["A3"] = "Limite do cartão"; K["A3"].font = ft(True); K["B3"] = 3000; inp(K["B3"], BRL)
K["A4"] = "Fatura do mês do Dashboard"; K["A4"].font = ft(True)
K["A5"] = "Limite disponível (estimado)"; K["A5"].font = ft(True)
K["A6"] = "Uso do limite"; K["A6"].font = ft(True)
note(K, "C3", "Valor de exemplo – coloque o limite real do seu cartão.")
head(K, 8, 1, ["Tipo", "Descrição", "Categoria", "Valor da parcela", "Nº de parcelas", "Mês da 1ª parcela", "Parcelas pagas", "Parcelas restantes", "Saldo restante", "Quitação prevista"])
KF, KL = 9, 28
park = [("Compra no cartão", "Celular (EXEMPLO – pode apagar)", "Compras", 150, 10, dt.date(2026, 9, 1))]
for i in range(20):
    r = KF + i
    if i < len(park):
        for col, v in zip("ABCDEF", park[i]): K[f"{col}{r}"] = v
    for col in "ABCDEF": inp(K[f"{col}{r}"], BRL if col == "D" else DATE if col == "F" else None, col in "AEF")
    elapsed = f"((YEAR(TODAY())*12+MONTH(TODAY()))-(YEAR(F{r})*12+MONTH(F{r}))+1)"
    K[f"G{r}"] = f'=IF(OR(F{r}="",N(E{r})=0),"",MAX(0,MIN(E{r},{elapsed})))'
    K[f"H{r}"] = f'=IF(G{r}="","",E{r}-G{r})'
    K[f"I{r}"] = f'=IF(G{r}="","",H{r}*D{r})'
    K[f"J{r}"] = f'=IF(G{r}="","",EDATE(F{r},E{r}-1))'
    out(K[f"G{r}"], "0", center=True); out(K[f"H{r}"], "0", center=True); out(K[f"I{r}"], BRL); out(K[f"J{r}"], "mm/yyyy", center=True)
dv(K, '"Compra no cartão,Dívida pessoal"', f"A{KF}:A{KL}"); dv(K, f"={CATS_DES}", f"C{KF}:C{KL}")
K[f"A{KL+1}"] = "TOTAL"; K[f"A{KL+1}"].font = ft(True)
K[f"I{KL+1}"] = f"=SUM(I{KF}:I{KL})"; out(K[f"I{KL+1}"], BRL, True, fill=MINT)
note(K, f"A{KL+2}", "Mês da 1ª parcela: qualquer dia daquele mês serve. A parcela entra como gasto do mês (na categoria escolhida).")
KA, KC_, KD, KE, KF_ = [f"{CD}!${c}${KF}:${c}${KL}" for c in "ACDEF"]

def dm(m, off=0):
    return f"(({YR}*12+{m}+{off})-(YEAR({KF_})*12+MONTH({KF_})))"
def parc(m, kind=None, cat=None, off=0, mexpr=None):
    mm = mexpr if mexpr else str(m)
    d = f"(({YR}*12+{mm}+{off})-(YEAR({KF_})*12+MONTH({KF_})))"
    f = f"SUMPRODUCT(({d}>=0)*({d}<{KE})*{KD}"
    if kind: f += f'*({KA}="{kind}")'
    if cat: f += f"*({KC_}={cat})"
    return f + ")"

def cred_diario(mexpr):
    X = f"DATE({YR},{mexpr},1)"; NX = f"DATE({YR},{mexpr}+1,1)"
    return f'SUMIFS({DV_},{DE},"Saída",{DFm},{CRED},{DA},">="&{X},{DA},"<"&{NX})'

K["B4"] = f"=CHOOSE({MM}," + ",".join(f"{q(m)}!$G$17" for m in MESES) + ")"
K["B5"] = f'=B3-B4-SUMIFS(I{KF}:I{KL},A{KF}:A{KL},"Compra no cartão")'
K["B6"] = "=IF(B3=0,0,1-B5/B3)"
out(K["B4"], BRL0, True); out(K["B5"], BRL0, True); out(K["B6"], "0%", True)
K.conditional_formatting.add("B6", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="8E44AD"))
note(K, "C5", "Limite − fatura do mês − parcelas futuras ainda não faturadas.")
# próximas faturas
band(K, "L8:N8", "Próximas faturas (parcelas já assumidas)", 10, TEAL, h=30)
head(K, 9, 12, ["Mês", "Parcelas no cartão", "Outras dívidas"])
for j in range(1, 7):
    r = 9 + j
    mexpr = f"({MM}+{j})"
    K[f"L{r}"] = f'=DATE({YR},{MM}+{j},1)'
    K[f"M{r}"] = "=" + parc(None, "Compra no cartão", mexpr=mexpr)
    K[f"N{r}"] = "=" + parc(None, "Dívida pessoal", mexpr=mexpr)
    out(K[f"L{r}"], "mm/yyyy", center=True); out(K[f"M{r}"], BRL0); out(K[f"N{r}"], BRL0)
note(K, "L16", "Compras à vista no cartão feitas no futuro ainda não aparecem aqui.")
K.freeze_panes = "A9"
widths(K, {"A": 18, "B": 26, "C": 22, "D": 15, "E": 11, "F": 13, "G": 11, "H": 11, "I": 15, "J": 13, "K": 3, "L": 12, "M": 18, "N": 16})

# ======================= RESERVA =======================
V = S["Reserva"]
band(V, "A1:H1", "🛟  RESERVA DE EMERGÊNCIA", 16, ORANGE)
labels = [(3, "Meta da reserva", 15000, True), (4, "Valor já guardado antes da planilha", 0, True), (5, "Depósito mensal planejado", 300, True)]
for r, l, v, _ in labels:
    V[f"A{r}"] = l; V[f"A{r}"].font = ft(True); V[f"B{r}"] = v; inp(V[f"B{r}"], BRL)
note(V, "C3", "Meta editável. Sugestão comum: 3 a 6 meses das suas despesas.")
note(V, "C5", "Usado só para prever a data de conclusão.")
VF, VL = 22, 221
VA_, VB_, VC_ = [f"'Reserva'!${c}${VF}:${c}${VL}" for c in "ABC"]
calc = [(7, "Total guardado", f'=B4+SUMIFS({VC_},{VB_},"Depósito")-SUMIFS({VC_},{VB_},"Retirada")', BRL0),
        (8, "Quanto falta", "=MAX(0,B3-B7)", BRL0),
        (9, "Percentual concluído", "=IF(B3=0,0,MIN(1,B7/B3))", "0.0%"),
        (10, "Progresso", '=REPT("█",ROUND(B9*20,0))&REPT("░",20-ROUND(B9*20,0))', None),
        (11, "Meses para concluir", '=IF(B8=0,0,IF(B5<=0,"Defina o depósito mensal",ROUNDUP(B8/B5,0)))', "0"),
        (12, "Previsão de conclusão", '=IF(ISNUMBER(B11),EDATE(TODAY(),B11),"—")', "mm/yyyy"),
        (13, "Depósitos no mês do Dashboard", f'=SUMIFS({VC_},{VB_},"Depósito",{VA_},">="&DATE({YR},{MM},1),{VA_},"<"&DATE({YR},{MM}+1,1))', BRL0),
        (14, "Retiradas no mês do Dashboard", f'=SUMIFS({VC_},{VB_},"Retirada",{VA_},">="&DATE({YR},{MM},1),{VA_},"<"&DATE({YR},{MM}+1,1))', BRL0)]
for r, l, f, fmt in calc:
    V[f"A{r}"] = l; V[f"A{r}"].font = ft(True); V[f"B{r}"] = f; out(V[f"B{r}"], fmt, r in (7, 9))
V["B10"].font = ft(True, 11, ORANGE)
V.conditional_formatting.add("B9", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F39C12"))
band(V, f"A{VF-2}:D{VF-2}", "Movimentações da reserva", 11, ORANGE, h=22, align=AL_L)
head(V, VF - 1, 1, ["Data", "Tipo", "Valor (R$)", "Observações"], ORANGE)
V[f"A{VF}"], V[f"B{VF}"], V[f"C{VF}"], V[f"D{VF}"] = dt.date(2026, 11, 5), "Depósito", 100, "EXEMPLO – pode apagar"
for r in range(VF, VL + 1):
    for col in "ABCD":
        c = V[f"{col}{r}"]; c.font = ft(False, 10, "0000FF"); c.border = Border(bottom=th)
    V[f"A{r}"].number_format = DATE; V[f"C{r}"].number_format = BRL
dv(V, '"Depósito,Retirada"', f"B{VF}:B{VL}")
# mensal
head(V, 3, 6, ["Mês", "Depósitos", "Retiradas", "Saldo acumulado"], ORANGE)
for i, m in enumerate(MESES, start=1):
    r = 3 + i
    V[f"F{r}"] = m
    V[f"G{r}"] = f'=SUMIFS({VC_},{VB_},"Depósito",{crit(VA_, i)})'
    V[f"H{r}"] = f'=SUMIFS({VC_},{VB_},"Retirada",{crit(VA_, i)})'
    V[f"I{r}"] = f"=$B$4+SUM($G$4:G{r})-SUM($H$4:H{r})"
    for col in "FGHI": out(V[f"{col}{r}"], None if col == "F" else BRL)
note(V, "F16", "Saldo acumulado considera as movimentações do ano em uso.")
lc = LineChart(); lc.title = "Evolução da reserva"; lc.height = 7; lc.width = 15; lc.style = 12
lc.add_data(Reference(V, min_col=9, min_row=3, max_row=15), titles_from_data=True)
lc.set_categories(Reference(V, min_col=6, min_row=4, max_row=15)); lc.legend = None
lc.series[0].graphicalProperties.line.solidFill = ORANGE; lc.series[0].graphicalProperties.line.width = 28000
lc.y_axis.numFmt = "#,##0"; lc.y_axis.delete = False; lc.x_axis.delete = False
V.add_chart(lc, "K3")
widths(V, {"A": 34, "B": 22, "C": 14, "D": 26, "E": 3, "F": 12, "G": 14, "H": 14, "I": 16})

# ======================= INVESTIMENTOS =======================
I = S["Investimentos"]
band(I, "A1:I1", "📈  INVESTIMENTOS", 16, "117A65")
note(I, "A2", "Registre aportes, retiradas e os rendimentos REAIS que aparecerem no seu extrato. A planilha não estima rentabilidade.", 9)
IF0, IL0 = 22, 321
IA, IB, IC, ID = [f"'Investimentos'!${c}${IF0}:${c}${IL0}" for c in "ABCD"]
head(I, 4, 1, ["Tipo", "Aportado", "Retirado", "Rendimentos", "Saldo atual", "% da carteira"], "117A65")
for i in range(10):
    r = 5 + i
    I[f"A{r}"] = f'=IF(Config!H{11+i}="","",Config!H{11+i})'
    for col, op in zip("BCD", ["Aporte", "Retirada", "Rendimento"]):
        I[f"{col}{r}"] = f'=IF($A{r}="",0,SUMIFS({ID},{IB},$A{r},{IC},"{op}"))'
    I[f"E{r}"] = f"=B{r}-C{r}+D{r}"
    I[f"F{r}"] = f'=IF(OR($A{r}="",$E$15=0),"",E{r}/$E$15)'
    for col in "ABCDEF": out(I[f"{col}{r}"], None if col == "A" else "0%" if col == "F" else BRL)
I["A15"] = "TOTAL"
for col in "BCDE": I[f"{col}15"] = f"=SUM({col}5:{col}14)"
for col in "ABCDEF": out(I[f"{col}15"], None if col == "A" else BRL, True, fill=MINT)
band(I, f"A{IF0-2}:E{IF0-2}", "Movimentações", 11, "117A65", h=22, align=AL_L)
head(I, IF0 - 1, 1, ["Data", "Tipo de investimento", "Operação", "Valor (R$)", "Observações"], "117A65")
I[f"A{IF0}"], I[f"B{IF0}"], I[f"C{IF0}"], I[f"D{IF0}"], I[f"E{IF0}"] = dt.date(2026, 11, 7), "CDB", "Aporte", 50, "EXEMPLO – pode apagar"
for r in range(IF0, IL0 + 1):
    for col in "ABCDE":
        c = I[f"{col}{r}"]; c.font = ft(False, 10, "0000FF"); c.border = Border(bottom=th)
    I[f"A{r}"].number_format = DATE; I[f"D{r}"].number_format = BRL
dv(I, f"={INVT}", f"B{IF0}:B{IL0}"); dv(I, '"Aporte,Retirada,Rendimento"', f"C{IF0}:C{IL0}")
head(I, 4, 8, ["Mês", "Aportes", "Retiradas", "Rendimentos", "Saldo acumulado"], "117A65")
for i, m in enumerate(MESES, start=1):
    r = 4 + i
    I[f"H{r}"] = m
    for col, op in zip("IJK", ["Aporte", "Retirada", "Rendimento"]):
        I[f"{col}{r}"] = f'=SUMIFS({ID},{IC},"{op}",{crit(IA, i)})'
    I[f"L{r}"] = f"=SUM($I$5:I{r})-SUM($J$5:J{r})+SUM($K$5:K{r})"
    for col in "HIJKL": out(I[f"{col}{r}"], None if col == "H" else BRL)
note(I, "H18", "Aportes (dinheiro que você colocou) e rendimentos aparecem separados.")
bc = BarChart(); bc.type = "col"; bc.grouping = "stacked"; bc.overlap = 100; bc.title = "Aportes x Rendimentos por mês"
bc.add_data(Reference(I, min_col=9, min_row=4, max_row=16), titles_from_data=True)
bc.add_data(Reference(I, min_col=11, min_row=4, max_row=16), titles_from_data=True)
bc.set_categories(Reference(I, min_col=8, min_row=5, max_row=16)); bc.height = 7; bc.width = 15
bc.series[0].graphicalProperties.solidFill = "117A65"; bc.series[1].graphicalProperties.solidFill = "82E0AA"
bc.y_axis.numFmt = "#,##0"; bc.y_axis.delete = False; bc.x_axis.delete = False
I.add_chart(bc, "N3")
lc2 = LineChart(); lc2.title = "Evolução do saldo investido"; lc2.height = 7; lc2.width = 15
lc2.add_data(Reference(I, min_col=12, min_row=4, max_row=16), titles_from_data=True)
lc2.set_categories(Reference(I, min_col=8, min_row=5, max_row=16)); lc2.legend = None
lc2.series[0].graphicalProperties.line.solidFill = "117A65"; lc2.series[0].graphicalProperties.line.width = 28000
lc2.y_axis.numFmt = "#,##0"; lc2.y_axis.delete = False; lc2.x_axis.delete = False
I.add_chart(lc2, "N18")
widths(I, {"A": 20, "B": 20, "C": 14, "D": 14, "E": 24, "F": 12, "G": 3, "H": 12, "I": 13, "J": 13, "K": 13, "L": 16})

# ======================= MESES =======================
CROW0 = 32; NCONT = 30; CROW1 = CROW0 + NCONT - 1
CAT0 = 66; NCAT = 16; CAT1 = CAT0 + NCAT - 1; CATT = CAT1 + 1
def paid(c): return f'OR({c}="☑",{c}=TRUE)'

for mi, mname in enumerate(MESES, start=1):
    M = S[mname]
    X, NX = mrange(mi)
    band(M, "B1:L1", f"ORGANIZAÇÃO DE {mname.upper()}", 18, TEAL, h=38)
    M.merge_cells("B2:L2")
    M["B2"] = '="Olá, "&Config!$B$3&"!  Pequenas ações geram grandes resultados."'
    M["B2"].font = ft(True, 10, TEAL, True); M["B2"].alignment = Alignment(horizontal="center")
    prev = MESES[mi - 2] if mi > 1 else None
    cards = [("B", "C", "Entradas (dinheiro)", "G11"), ("D", "E", "Benefícios recebidos", "G12"), ("F", "G", "Saídas realizadas", "G16"),
             ("H", "I", "Reservas & Investimentos", "G28"), ("J", "K", "Uso do crédito", "G17")]
    for a, b, lab, ref in cards:
        M.merge_cells(f"{a}4:{b}4"); M.merge_cells(f"{a}5:{b}5")
        M[f"{a}4"] = lab; M[f"{a}4"].font = ft(True, 9, "FFFFFF"); M[f"{a}4"].fill = fl(TEAL2); M[f"{a}4"].alignment = AL_C
        M[f"{a}5"] = f"={ref}"; M[f"{a}5"].font = ft(True, 14, TEAL); M[f"{a}5"].number_format = BRL0; M[f"{a}5"].alignment = AL_C
        M[f"{a}6"] = "Mês anterior:"; M[f"{a}6"].font = ft(False, 8, "6B7B78")
        M[f"{b}6"] = f"={q(prev)}!{ref}" if prev else "=0"; M[f"{b}6"].font = ft(False, 8, "6B7B78"); M[f"{b}6"].number_format = BRL0
        for col in (a, b):
            for r in (4, 5, 6): M[f"{col}{r}"].border = BOX
            M[f"{col}5"].fill = fl("FFFFFF"); M[f"{col}6"].fill = fl(SOFT)
    M.row_dimensions[5].height = 30
    # Restante
    M.merge_cells("B8:F8"); M["B8"] = "RESTANTE PARA GASTAR (depois de contas pendentes e fatura)"
    M["B8"].font = ft(True, 10, "FFFFFF"); M["B8"].fill = fl(TEAL); M["B8"].alignment = AL_C
    M.merge_cells("B9:F9"); M["B9"] = "=G23"; M["B9"].font = ft(True, 22, TEAL); M["B9"].number_format = BRL0
    M["B9"].alignment = AL_C; M.row_dimensions[9].height = 40
    for col in "BCDEF": M[f"{col}9"].border = BOX; M[f"{col}8"].border = BOX
    M.conditional_formatting.add("B9", FormulaRule(formula=["$G$23<0"], font=Font(name=FONT, bold=True, size=22, color=RED)))
    # Resumo
    resumo = [
        (11, "Renda em dinheiro", "=" + rec_din(mi)),
        (12, "Benefícios alimentares recebidos (VA/VR)", "=" + rec_ben(mi)),
        (13, "Contas fixas planejadas", f"=SUM(G{CROW0}:G{CROW1})"),
        (14, "Contas fixas pagas", f"=SUM(L{CROW0}:L{CROW1})"),
        (15, "Contas fixas pendentes", f'=SUMIFS(G{CROW0}:G{CROW1},K{CROW0}:K{CROW1},"Pendente")+SUMIFS(G{CROW0}:G{CROW1},K{CROW0}:K{CROW1},"ATRASADA")'),
        (16, "Saídas realizadas (contas pagas + gastos + parcelas)", f"=E{CATT}"),
        (17, "   • no cartão de crédito (vai para a fatura)", f'=SUMIFS({DV_},{DE},"Saída",{DFm},{CRED},{crit(DA, mi)})+SUMIFS(L{CROW0}:L{CROW1},E{CROW0}:E{CROW1},{CRED})+' + parc(mi, "Compra no cartão")),
        (18, "   • com benefícios (VA/VR)", f'=SUMIFS({DV_},{DE},"Saída",{DFm},{VA},{crit(DA, mi)})+SUMIFS({DV_},{DE},"Saída",{DFm},{VR},{crit(DA, mi)})+SUMIFS(L{CROW0}:L{CROW1},E{CROW0}:E{CROW1},{VA})+SUMIFS(L{CROW0}:L{CROW1},E{CROW0}:E{CROW1},{VR})'),
        (19, "   • à vista (Pix, débito, dinheiro, boleto…)", "=G16-G17-G18"),
        (20, "Reserva de emergência (depósitos − retiradas)", f'=SUMIFS({VC_},{VB_},"Depósito",{crit(VA_, mi)})-SUMIFS({VC_},{VB_},"Retirada",{crit(VA_, mi)})'),
        (21, "Investimentos (aportes − retiradas)", f'=SUMIFS({ID},{IC},"Aporte",{crit(IA, mi)})-SUMIFS({ID},{IC},"Retirada",{crit(IA, mi)})'),
        (22, "Saldo atual em conta (o que já saiu de fato)", "=G11-G19-G20-G21"),
        (23, "Restante para gastar", "=G22-G15-G17"),
        (24, "Saldo dos benefícios (VA/VR)", "=G12-G18"),
        (25, "Contas atrasadas", f'=COUNTIF(K{CROW0}:K{CROW1},"ATRASADA")'),
        (26, "Parcelas do mês (cartão + dívidas)", "=" + parc(mi)),
        (27, "Gastos variáveis (Diário)", "=G16-G14-G26"),
        (28, "Reservas + investimentos", "=G20+G21"),
    ]
    for r, lab, f in resumo:
        M.merge_cells(f"B{r}:F{r}"); M[f"B{r}"] = lab
        M[f"B{r}"].font = ft(r in (11, 16, 22, 23), 10); M[f"B{r}"].alignment = AL_L
        M[f"G{r}"] = f
        out(M[f"G{r}"], "0" if r == 25 else BRL0, r in (11, 16, 22, 23), fill=MINT if r in (23,) else None)
        for col in "BCDEF": M[f"{col}{r}"].border = Border(bottom=th)
    M["M11"] = f'=COUNTIF(K{CROW0}:K{CROW1},"Pago")'; M["M12"] = f'=COUNTIF(C{CROW0}:C{CROW1},"?*")'
    M.column_dimensions["M"].hidden = True
    M.conditional_formatting.add("G25", FormulaRule(formula=["$G$25>0"], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=fl(RED)))
    # Contas
    band(M, f"B{CROW0-2}:L{CROW0-2}", "CONTAS FIXAS DO MÊS  —  marque ☑ quando pagar", 11, TEAL, h=24)
    head(M, CROW0 - 1, 2, ["Pago", "Conta", "Categoria", "Forma de pagamento", "Vencimento", "Planejado", "Valor pago", "Data do pagamento", "Diferença", "Status", "Total pago"])
    mref = f"Config!$L${10+mi}"
    for k in range(NCONT):
        r = CROW0 + k; dr = EF + k
        M[f"C{r}"] = f'=IF(AND(Despesas!$A${dr}<>"",OR(Despesas!$F${dr}="Todos os meses",Despesas!$F${dr}="",Despesas!$F${dr}={mref})),Despesas!$A${dr},"")'
        M[f"D{r}"] = f'=IF(C{r}="","",Despesas!$B${dr})'
        M[f"E{r}"] = f'=IF(C{r}="","",Despesas!$E${dr})'
        M[f"F{r}"] = f'=IF(OR(C{r}="",N(Despesas!$D${dr})=0),"",DATE({YR},{mi},MIN(Despesas!$D${dr},DAY(EOMONTH({X},0)))))'
        M[f"G{r}"] = f'=IF(C{r}="",0,N(Despesas!$C${dr}))'
        M[f"J{r}"] = f'=IF(AND(C{r}<>"",{paid(f"B{r}")}),L{r}-G{r},"")'
        M[f"K{r}"] = f'=IF(C{r}="","",IF({paid(f"B{r}")},"Pago",IF(AND(F{r}<>"",F{r}<TODAY()),"ATRASADA","Pendente")))'
        M[f"L{r}"] = f'=IF(AND(C{r}<>"",{paid(f"B{r}")}),IF(N(H{r})=0,G{r},H{r}),0)'
        inp(M[f"B{r}"], center=True); M[f"B{r}"].font = ft(True, 12, TEAL)
        inp(M[f"H{r}"], BRL); inp(M[f"I{r}"], DATE, True)
        for col, fmt in zip("CDEFGJKL", [None, None, None, DATE, BRL, BRL, None, BRL]):
            out(M[f"{col}{r}"], fmt, col == "C", center=col in "FK")
    M[f"C{CROW1+1}"] = "TOTAL"; M[f"C{CROW1+1}"].font = ft(True)
    for col in "GHJL": M[f"{col}{CROW1+1}"] = f"=SUM({col}{CROW0}:{col}{CROW1})"; out(M[f"{col}{CROW1+1}"], BRL, True, fill=MINT)
    dv(M, '"☐,☑"', f"B{CROW0}:B{CROW1}")
    rng = f"B{CROW0}:L{CROW1}"
    M.conditional_formatting.add(rng, FormulaRule(formula=[f'$K{CROW0}="Pago"'], fill=fl(GRNBG)))
    M.conditional_formatting.add(f"K{CROW0}:K{CROW1}", FormulaRule(formula=[f'$K{CROW0}="ATRASADA"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=fl(RED)))
    M.conditional_formatting.add(f"K{CROW0}:K{CROW1}", FormulaRule(formula=[f'$K{CROW0}="Pendente"'], font=Font(name=FONT, bold=True, color=AMB), fill=fl(AMBBG)))
    M.conditional_formatting.add(f"J{CROW0}:J{CROW1}", FormulaRule(formula=[f'AND(ISNUMBER($J{CROW0}),$J{CROW0}>0)'], font=Font(name=FONT, bold=True, color=RED)))
    # Categorias
    band(M, f"B{CAT0-2}:L{CAT0-2}", "GASTOS POR CATEGORIA  —  orçamento x realizado", 11, TEAL, h=24)
    head(M, CAT0 - 1, 2, ["", "Categoria", "Orçamento", "Realizado", "% usado", "Situação"])
    for i in range(NCAT):
        r = CAT0 + i; cr = 35 + i
        M[f"C{r}"] = f'=IF(Config!$A${cr}="","",Config!$A${cr})'
        M[f"D{r}"] = f'=IF(C{r}="",0,N(Config!$C${cr}))'
        M[f"E{r}"] = (f'=IF(C{r}="",0,SUMIFS({DV_},{DC},C{r},{DE},"Saída",{crit(DA, mi)})'
                      f'+SUMIFS($L${CROW0}:$L${CROW1},$D${CROW0}:$D${CROW1},C{r})+{parc(mi, cat=f"C{r}")})')
        M[f"F{r}"] = f'=IF(OR(C{r}="",D{r}=0),"",E{r}/D{r})'
        M[f"G{r}"] = f'=IF(C{r}="","",IF(D{r}=0,IF(E{r}>0,"Sem limite","—"),IF(E{r}>D{r},"Acima do limite",IF(E{r}>=0.8*D{r},"Atenção","OK"))))'
        for col, fmt in zip("CDEFG", [None, BRL, BRL, "0%", None]): out(M[f"{col}{r}"], fmt, col == "C", center=col == "G")
    M[f"C{CATT}"] = "TOTAL"; M[f"D{CATT}"] = f"=SUM(D{CAT0}:D{CAT1})"; M[f"E{CATT}"] = f"=SUM(E{CAT0}:E{CAT1})"
    for col in "CDEFG": out(M[f"{col}{CATT}"], BRL if col in "DE" else None, True, fill=MINT)
    M.conditional_formatting.add(f"F{CAT0}:F{CAT1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.2, color="2E8B7E"))
    for txt, fc, bg in [("Acima do limite", RED, REDBG), ("Atenção", AMB, AMBBG), ("OK", GRN, GRNBG)]:
        M.conditional_formatting.add(f"G{CAT0}:G{CAT1}", FormulaRule(formula=[f'$G{CAT0}="{txt}"'], font=Font(name=FONT, color=fc, bold=True), fill=fl(bg)))
    # charts
    dn = DoughnutChart(); dn.title = "Para onde foi o dinheiro"; dn.holeSize = 60
    dn.add_data(Reference(M, min_col=5, min_row=CAT0, max_row=CAT1), titles_from_data=False)
    dn.set_categories(Reference(M, min_col=3, min_row=CAT0, max_row=CAT1))
    dn.height = 10; dn.width = 12
    M.add_chart(dn, "H8")
    bc = BarChart(); bc.type = "bar"; bc.title = "Contas: planejado x pago"; bc.style = 10
    bc.add_data(Reference(M, min_col=7, min_row=CROW0 - 1, max_row=CROW0 + 9), titles_from_data=True)
    bc.add_data(Reference(M, min_col=12, min_row=CROW0 - 1, max_row=CROW0 + 9), titles_from_data=True)
    bc.set_categories(Reference(M, min_col=3, min_row=CROW0, max_row=CROW0 + 9))
    bc.series[0].graphicalProperties.solidFill = "A9CCE3"; bc.series[1].graphicalProperties.solidFill = TEAL
    bc.height = 9; bc.width = 13; bc.y_axis.delete = False; bc.x_axis.delete = False; bc.y_axis.numFmt = "#,##0"
    M.add_chart(bc, f"I{CAT0-1}")
    widths(M, {"A": 2, "B": 9, "C": 26, "D": 18, "E": 18, "F": 12, "G": 15, "H": 14, "I": 14, "J": 12, "K": 12, "L": 13})
    M.freeze_panes = "A3"

# ======================= PLANEJAMENTO =======================
P = S["Planejamento"]
band(P, "A1:T1", "📅  PLANEJAMENTO MENSAL E RESUMO ANUAL", 16)
P.merge_cells("A2:T2"); P["A2"] = '="Ano em uso: "&' + YR + '&"   •   Tudo aqui é automático (vem das abas de cada mês e dos registros)."'
P["A2"].font = ft(False, 9, TEAL, True)
cols = ["Mês", "Receita prevista", "Receita realizada (dinheiro)", "Benefícios (VA/VR)", "Contas fixas planejadas", "Contas fixas pagas",
        "Gastos variáveis", "Cartão de crédito (fatura)", "Parcelas (cartão + dívidas)", "Saídas realizadas", "Reserva", "Investimentos",
        "Saldo em conta", "Restante p/ gastar", "Economizado", "Taxa de economia", "% renda comprometida", "Patrimônio acumulado",
        "Economizado vs mês anterior", "Projeção de gastos até o fim do mês"]
head(P, 4, 1, cols)
P.row_dimensions[4].height = 44
for i, m in enumerate(MESES, start=1):
    r = 4 + i; s = q(m)
    P[f"A{r}"] = m
    P[f"B{r}"] = f'=SUMIFS({RR_V},{RR_T},"Dinheiro",{RR_A},"Sim")'
    for col, cell in zip("CDEFGHIJKLMN", ["G11", "G12", "G13", "G14", "G27", "G17", "G26", "G16", "G20", "G21", "G22", "G23"]):
        P[f"{col}{r}"] = f"={s}!{cell}"
    P[f"O{r}"] = f"=N{r}+K{r}+L{r}"
    P[f"P{r}"] = f'=IF(C{r}=0,"",O{r}/C{r})'
    P[f"Q{r}"] = f'=IF(C{r}=0,"",(E{r}+I{r})/C{r})'
    P[f"R{r}"] = f"='Reserva'!$I${3+i}+'Investimentos'!$L${4+i}"
    P[f"S{r}"] = f"=O{r}-O{r-1}" if i > 1 else '=""'
    P[f"T{r}"] = (f'=IF(AND({YR}=YEAR(TODAY()),{i}=MONTH(TODAY())),MAX(0,G{r}-{s}!$G$18)/DAY(TODAY())*(DAY(EOMONTH(TODAY(),0))-DAY(TODAY())),0)')
    for col in "ABCDEFGHIJKLMNOPQRST":
        out(P[f"{col}{r}"], None if col == "A" else "0%" if col in "PQ" else BRL, col in "AN")
P["A17"] = "TOTAL / MÉDIA"
for col in "BCDEFGHIJKLMNO": P[f"{col}17"] = f"=SUM({col}5:{col}16)"
P["P17"] = '=IF(C17=0,"",O17/C17)'; P["Q17"] = '=IF(C17=0,"",(E17+I17)/C17)'; P["R17"] = "=R16"
for col in "ABCDEFGHIJKLMNOPQRST": out(P[f"{col}17"], None if col == "A" else "0%" if col in "PQ" else BRL, True, fill=MINT)
P.conditional_formatting.add("N5:N16", FormulaRule(formula=["N5<0"], font=Font(name=FONT, bold=True, color=RED)))
P.conditional_formatting.add("S5:S16", FormulaRule(formula=["AND(ISNUMBER(S5),S5<0)"], font=Font(name=FONT, color=RED)))
P.conditional_formatting.add("S5:S16", FormulaRule(formula=["AND(ISNUMBER(S5),S5>0)"], font=Font(name=FONT, color=GRN)))
# matriz categorias
band(P, "A20:N20", "Gastos por categoria × mês", 11, TEAL, h=22, align=AL_L)
head(P, 21, 1, ["Categoria"] + [m[:3] for m in MESES] + ["Total"])
for i in range(NCAT):
    r = 22 + i; cr = 35 + i
    P[f"A{r}"] = f'=IF(Config!$A${cr}="","",Config!$A${cr})'
    for j, m in enumerate(MESES):
        col = chr(ord("B") + j); P[f"{col}{r}"] = f"={q(m)}!$E${CAT0+i}"
        out(P[f"{col}{r}"], '#,##0.00;-#,##0.00;"-"')
    P[f"N{r}"] = f"=SUM(B{r}:M{r})"; out(P[f"A{r}"]); out(P[f"N{r}"], '#,##0.00;-#,##0.00;"-"', True)
PCAT0, PCAT1 = 22, 22 + NCAT - 1
# gráficos
ch = BarChart(); ch.type = "col"; ch.title = "Receitas x Saídas por mês"; ch.style = 10
ch.add_data(Reference(P, min_col=3, min_row=4, max_row=16), titles_from_data=True)
ch.add_data(Reference(P, min_col=10, min_row=4, max_row=16), titles_from_data=True)
ch.set_categories(Reference(P, min_col=1, min_row=5, max_row=16))
ch.series[0].graphicalProperties.solidFill = TEAL; ch.series[1].graphicalProperties.solidFill = ORANGE
ch.height = 8; ch.width = 18; ch.y_axis.numFmt = "#,##0"; ch.y_axis.delete = False; ch.x_axis.delete = False
P.add_chart(ch, "A40")
lp = LineChart(); lp.title = "Evolução do patrimônio (reserva + investimentos)"
lp.add_data(Reference(P, min_col=18, min_row=4, max_row=16), titles_from_data=True)
lp.add_data(Reference(P, min_col=15, min_row=4, max_row=16), titles_from_data=True)
lp.set_categories(Reference(P, min_col=1, min_row=5, max_row=16))
lp.series[0].graphicalProperties.line.solidFill = TEAL; lp.series[0].graphicalProperties.line.width = 30000
lp.series[1].graphicalProperties.line.solidFill = ORANGE
lp.height = 8; lp.width = 18; lp.y_axis.numFmt = "#,##0"; lp.y_axis.delete = False; lp.x_axis.delete = False
P.add_chart(lp, "H40")
P.freeze_panes = "B5"
widths(P, {"A": 22, **{chr(ord("B") + j): 14 for j in range(19)}})

def PI(col): return f"INDEX(Planejamento!${col}$5:${col}$16,{MM})"

# ======================= METAS =======================
T = S["Metas"]
band(T, "A1:I1", "🎯  METAS FINANCEIRAS", 16, "B7950B")
note(T, "A2", "Cadastre seus objetivos. A planilha calcula quanto guardar por mês até o prazo.", 9)
head(T, 4, 1, ["Meta", "Valor desejado", "Valor acumulado", "Prazo", "Meses até o prazo", "Contribuição mensal necessária", "Falta", "Progresso", "Situação"], "B7950B")
goals = [("Reserva financeira", 15000, "='Reserva'!B7", dt.date(2027, 12, 31)), ("Viagem (exemplo)", 4000, 500, dt.date(2027, 7, 1)),
         ("Comprar equipamentos (exemplo)", 2500, 0, dt.date(2027, 3, 1))]
for i in range(10):
    r = 5 + i
    if i < len(goals):
        for col, v in zip("ABCD", goals[i]): T[f"{col}{r}"] = v
    for col in "ABCD": inp(T[f"{col}{r}"], BRL if col in "BC" else DATE if col == "D" else None, col == "D")
    T[f"E{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",MAX(0,(YEAR(D{r})*12+MONTH(D{r}))-(YEAR(TODAY())*12+MONTH(TODAY()))))'
    T[f"F{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",IF(N(C{r})>=B{r},0,IF(OR(E{r}="",E{r}=0),B{r}-N(C{r}),(B{r}-N(C{r}))/E{r})))'
    T[f"G{r}"] = f'=IF(A{r}="","",MAX(0,N(B{r})-N(C{r})))'
    T[f"H{r}"] = f'=IF(OR(A{r}="",N(B{r})=0),"",MIN(1,N(C{r})/B{r}))'
    T[f"I{r}"] = f'=IF(H{r}="","",IF(H{r}>=1,"Concluída 🎉",IF(E{r}=0,"Prazo vencido",REPT("█",ROUND(H{r}*10,0))&REPT("░",10-ROUND(H{r}*10,0)))))'
    for col, fmt in zip("EFGHI", ["0", BRL, BRL, "0%", None]): out(T[f"{col}{r}"], fmt, col == "F", center=col in "EHI")
T["C5"].font = ft(False, 10, "008000")
note(T, "A16", "A meta “Reserva financeira” já puxa o valor guardado da aba Reserva. As outras, atualize o valor acumulado quando guardar.")
T.conditional_formatting.add("H5:H14", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="D4AC0D"))
widths(T, {"A": 30, "B": 16, "C": 16, "D": 13, "E": 11, "F": 18, "G": 15, "H": 14, "I": 18})
gb = BarChart(); gb.type = "bar"; gb.title = "Progresso das metas"; gb.height = 7; gb.width = 15
gb.add_data(Reference(T, min_col=8, min_row=4, max_row=9), titles_from_data=True)
gb.set_categories(Reference(T, min_col=1, min_row=5, max_row=9)); gb.legend = None
gb.series[0].graphicalProperties.solidFill = "D4AC0D"; gb.y_axis.numFmt = "0%"; gb.y_axis.scaling.max = 1; gb.y_axis.scaling.min = 0
gb.y_axis.delete = False; gb.x_axis.delete = False
T.add_chart(gb, "A18")

# ======================= SIMULADOR =======================
Z = S["Simulador"]
band(Z, "A1:F1", "🧮  SIMULADOR FINANCEIRO", 16, "7D3C98")
note(Z, "A2", "Mude os números amarelos à vontade: as simulações NÃO alteram seus dados reais.", 9)
Z.column_dimensions["A"].width = 52; Z.column_dimensions["B"].width = 18; Z.column_dimensions["C"].width = 4; Z.column_dimensions["D"].width = 60
row = 4
def sim_block(title, inputs, outputs, notes=None):
    global row
    band(Z, f"A{row}:B{row}", title, 11, "7D3C98", h=22, align=AL_L); row += 1
    refs = {}
    for key, lab, val, fmt in inputs:
        Z[f"A{row}"] = lab; Z[f"B{row}"] = val; inp(Z[f"B{row}"], fmt); refs[key] = f"B{row}"; row += 1
    for key, lab, f, fmt in outputs:
        Z[f"A{row}"] = lab; Z[f"A{row}"].font = ft(True)
        Z[f"B{row}"] = f.format(**refs); out(Z[f"B{row}"], fmt, True, fill="EBDEF0"); refs[key] = f"B{row}"; row += 1
    if notes: note(Z, f"A{row}", notes); row += 1
    row += 1
sim_block("1) Quanto consigo guardar com essa renda?",
          [("r", "Renda mensal em dinheiro", 1500, BRL), ("d", "Despesas mensais", "=Despesas!C36", BRL)],
          [("s", "Sobra por mês", "={r}-{d}", BRL0), ("p", "% da renda", "=IF({r}=0,0,({r}-{d})/{r})", "0%"), ("a", "Em 12 meses", "=MAX(0,{s})*12", BRL0)],
          "Despesas começam com o total das suas contas fixas (você pode mudar aqui sem afetar nada).")
sim_block("2) Quanto economizo cortando um gasto?",
          [("c", "Valor cortado por mês", 200, BRL), ("m", "Por quantos meses", 12, "0")],
          [("t", "Economia total", "={c}*{m}", BRL0)])
sim_block("3) Quanto vou juntar guardando todo mês?",
          [("a", "Valor guardado por mês", 300, BRL), ("m", "Meses", 24, "0"), ("j", "Rendimento ao mês (%) – opcional", 0, "0.00%")],
          [("v", "Total acumulado", "=IF({j}=0,{a}*{m},FV({j},{m},-{a},0))", BRL0), ("ap", "Só aportes", "={a}*{m}", BRL0)],
          "Rendimento começa em 0%. Coloque a taxa real do seu investimento se quiser.")
sim_block("4) Quanto tempo para juntar um valor?",
          [("o", "Objetivo", 15000, BRL), ("t", "Já tenho", "='Reserva'!B7", BRL), ("a", "Guardo por mês", 300, BRL), ("j", "Rendimento ao mês (%) – opcional", 0, "0.00%")],
          [("n", "Meses necessários", '=IF({t}>={o},0,IF({a}<=0,"Defina um valor mensal",IF({j}=0,ROUNDUP(({o}-{t})/{a},0),ROUNDUP(NPER({j},-{a},-{t},{o}),0))))', "0"),
           ("dd", "Data prevista", '=IF(ISNUMBER({n}),EDATE(TODAY(),{n}),"—")', "mm/yyyy")])
sim_block("5) Impacto de uma compra parcelada",
          [("pv", "Valor da parcela", 150, BRL), ("np", "Nº de parcelas", 10, "0"), ("rn", "Renda mensal em dinheiro", 1500, BRL),
           ("cc", "Já comprometido por mês (contas + parcelas)", "=Despesas!C36+" + parc(None, mexpr=MM), BRL)],
          [("tt", "Custo total da compra", "={pv}*{np}", BRL0), ("na", "Comprometimento atual", "=IF({rn}=0,0,{cc}/{rn})", "0%"),
           ("nd", "Comprometimento com a compra", "=IF({rn}=0,0,({cc}+{pv})/{rn})", "0%"),
           ("ok", "Avaliação", '=IF({nd}>0.7,"🔴 Arriscado: passa de 70% da renda",IF({nd}>0.5,"⚠️ Apertado: acima de 50% da renda","✅ Cabe no orçamento"))', None)])
sim_block("6) Quanto posso gastar sem comprometer as contas?",
          [("rn", "Renda em dinheiro do mês", f"={PI('C')}", BRL), ("cf", "Contas fixas do mês", f"={PI('E')}", BRL),
           ("pa", "Parcelas do mês", f"={PI('I')}", BRL), ("me", "Quanto quero guardar (%)", "=Config!B8", "0%")],
          [("li", "Pode gastar no mês", "=MAX(0,{rn}-{cf}-{pa}-{rn}*{me})", BRL0), ("se", "Por semana", "={li}/4.33", BRL0), ("di", "Por dia", "={li}/30", BRL0)],
          "Os valores iniciais vêm do mês do Dashboard, mas aqui você pode mudar sem afetar nada.")

# ======================= DASHBOARD =======================
B = S["Dashboard"]
band(B, "B1:M1", "📊  MEU DASHBOARD FINANCEIRO", 18, TEAL, h=40)
B.merge_cells("B2:M2")
B["B2"] = f'="Olá, "&Config!$B$3&"!   Mês: "&INDEX({MESL},{MM})&" de "&{YR}&"   •   Pequenas ações geram grandes resultados."'
B["B2"].font = ft(True, 10, TEAL, True); B["B2"].alignment = Alignment(horizontal="center")
navs = [("B", "Diário"), ("C", "Receitas"), ("D", "Despesas"), ("E", "Cartão e Dívidas"), ("F", "Reserva"), ("G", "Investimentos"),
        ("H", "Metas"), ("I", "Planejamento"), ("J", "Simulador"), ("K", "Config")]
for col, sh in navs: link(B[f"{col}3"], sh, sh if sh != "Cartão e Dívidas" else "Cartão")
B["L3"] = "Trocar mês:"; B["L3"].font = ft(True, 9, TEAL); B["L3"].alignment = Alignment(horizontal="right")
link(B["M3"], "Config", "Config ▸")
for j, m in enumerate(MESES):
    c = B.cell(row=4, column=2 + (j % 12), value=m[:3])
    link(c, m, m[:3]); c.fill = fl(MINT); c.font = ft(True, 9, TEAL)
B.row_dimensions[3].height = 22; B.row_dimensions[4].height = 20
BM = lambda cell: f"CHOOSE({MM}," + ",".join(f"{q(m)}!{cell}" for m in MESES) + ")"
kpis = [
    ("Meu salário do mês", f"=IF(Config!$D$11=\"Sim\",{SAL},0)", '="Adiantamentos: R$ "&FIXED(INDEX(Receitas!$C$16:$C$27,' + MM + '),2)'),
    ("Renda total (dinheiro + benefícios)", f"={PI('C')}+{PI('D')}", f'="Em dinheiro: R$ "&FIXED({PI("C")},2)'),
    ("Benefícios alimentares", f"={PI('D')}", f'="Saldo VA/VR: R$ "&FIXED({BM("$G$24")},2)'),
    ("Total de despesas (realizadas)", f"={PI('J')}", f'="Variáveis: R$ "&FIXED({PI("G")},2)'),
    ("Contas pagas", f"={PI('F')}", f'={BM("$M$11")}&" de "&{BM("$M$12")}&" contas pagas"'),
    ("Contas pendentes", f"={BM('$G$15')}", f'={BM("$G$25")}&" atrasada(s)"'),
    ("Disponível para gastar", f"={PI('N')}", '="Após contas pendentes e fatura"'),
    ("Saldo previsto no fim do mês", f"={PI('N')}-{PI('T')}", '="Se mantiver o ritmo de gastos"'),
    ("Valor economizado", f"={PI('O')}", f'="Taxa: "&IF({PI("P")}="","0",ROUND({PI("P")}*100,0))&"% (meta "&ROUND(Config!$B$8*100,0)&"%)"'),
    ("Total investido", "='Investimentos'!$E$15", '="Rendimentos: R$ "&FIXED(\'Investimentos\'!$D$15,2)'),
    ("Patrimônio (reserva + investimentos)", f"={PI('R')}", f'="Mês anterior: R$ "&FIXED(IF({MM}=1,0,INDEX(Planejamento!$R$5:$R$16,{MM}-1)),2)'),
    ("Renda comprometida", f'=IF({PI("Q")}="",0,{PI("Q")})', '="Contas fixas + parcelas"'),
]
pos = [("B", "D"), ("E", "G"), ("H", "J"), ("K", "M")]
for idx, (lab, val, sub) in enumerate(kpis):
    a, b = pos[idx % 4]; r0 = 6 + (idx // 4) * 4
    for rr in (r0, r0 + 1, r0 + 2): B.merge_cells(f"{a}{rr}:{b}{rr}")
    B[f"{a}{r0}"] = lab; B[f"{a}{r0}"].font = ft(True, 9, "FFFFFF"); B[f"{a}{r0}"].fill = fl(TEAL2); B[f"{a}{r0}"].alignment = AL_C
    B[f"{a}{r0+1}"] = val; B[f"{a}{r0+1}"].font = ft(True, 16, TEAL); B[f"{a}{r0+1}"].alignment = AL_C
    B[f"{a}{r0+1}"].number_format = "0%" if "comprometida" in lab else BRL0
    B[f"{a}{r0+2}"] = sub; B[f"{a}{r0+2}"].font = ft(False, 8, "6B7B78", True); B[f"{a}{r0+2}"].alignment = AL_C
    for col in [chr(c) for c in range(ord(a), ord(b) + 1)]:
        for rr in (r0, r0 + 1, r0 + 2): B[f"{col}{rr}"].border = BOX
        B[f"{col}{r0+2}"].fill = fl(SOFT)
    B.row_dimensions[r0 + 1].height = 30
# destaques condicionais
B.conditional_formatting.add("H11", FormulaRule(formula=[f"{PI('N')}<0"], font=Font(name=FONT, bold=True, size=16, color=RED)))
B.conditional_formatting.add("E11", FormulaRule(formula=[f"{BM('$G$25')}>0"], font=Font(name=FONT, bold=True, size=16, color=RED)))
B.conditional_formatting.add("K15", FormulaRule(formula=["$K$15>0.5"], font=Font(name=FONT, bold=True, size=16, color=RED)))
# reserva progresso
band(B, "B18:M18", "🛟  Reserva de emergência", 11, ORANGE, h=22, align=AL_L)
B.merge_cells("B19:F19"); B["B19"] = "='Reserva'!B10"; B["B19"].font = ft(True, 14, ORANGE); B["B19"].alignment = Alignment(vertical="center")
B.merge_cells("G19:M19")
B["G19"] = '=ROUND(\'Reserva\'!B9*100,0)&"% concluído  •  R$ "&FIXED(\'Reserva\'!B7,2)&" de R$ "&FIXED(\'Reserva\'!B3,2)&"  •  previsão: "&IF(ISNUMBER(\'Reserva\'!B12),MONTH(\'Reserva\'!B12)&"/"&YEAR(\'Reserva\'!B12),"—")'
B["G19"].font = ft(True, 10); B["G19"].alignment = Alignment(vertical="center")
B.row_dimensions[19].height = 26
# alertas
band(B, "B21:M21", "💡  Alertas e recomendações automáticas", 11, TEAL, h=22, align=AL_L)
DT0 = 80  # tabela de dados dos gráficos
alerts = [
    f'=IF({BM("$G$25")}>0,"🔴 Você tem "&{BM("$G$25")}&" conta(s) atrasada(s). Pague primeiro para evitar multa e juros.","✅ Nenhuma conta atrasada neste mês.")',
    f'=IF({PI("N")}<0,"🔴 Suas obrigações passam a renda em R$ "&FIXED(-{PI("N")},2)&". Corte gastos variáveis ou adie compras.","✅ Depois de pagar tudo, ainda sobram R$ "&FIXED({PI("N")},2)&" para gastar ("&"≈ R$ "&FIXED(MAX(0,{PI("N")})/MAX(1,DAY(EOMONTH(DATE({YR},{MM},1),0))-IF(AND({YR}=YEAR(TODAY()),{MM}=MONTH(TODAY())),DAY(TODAY()),0)),2)&" por dia).")',
    f'=IF({PI("C")}=0,"ℹ️ Sem renda em dinheiro neste mês.",IF({PI("P")}>=Config!$B$8,"✅ Você está economizando "&ROUND({PI("P")}*100,0)&"% da renda (meta: "&ROUND(Config!$B$8*100,0)&"%).","⚠️ Economia de "&ROUND(MAX(0,{PI("P")})*100,0)&"% — faltam R$ "&FIXED(Config!$B$8*{PI("C")}-{PI("O")},2)&" para bater a meta de "&ROUND(Config!$B$8*100,0)&"%."))',
    f'=IF({PI("Q")}="","",IF({PI("Q")}>0.7,"🔴 Contas fixas e parcelas comprometem "&ROUND({PI("Q")}*100,0)&"% da renda. Evite novas parcelas.",IF({PI("Q")}>0.5,"⚠️ "&ROUND({PI("Q")}*100,0)&"% da renda já está comprometida com contas e parcelas.","✅ Renda comprometida em "&ROUND({PI("Q")}*100,0)&"%: saudável.")))',
    f'=IF(COUNTIF($F${DT0+1}:$F${DT0+NCAT},"Acima do limite")=0,"✅ Nenhuma categoria passou do limite.","🔴 "&COUNTIF($F${DT0+1}:$F${DT0+NCAT},"Acima do limite")&" categoria(s) acima do limite. Maior gasto: "&INDEX($B${DT0+1}:$B${DT0+NCAT},MATCH(MAX($C${DT0+1}:$C${DT0+NCAT}),$C${DT0+1}:$C${DT0+NCAT},0))&".")',
    f'="🍽️ Benefícios: recebidos R$ "&FIXED({PI("D")},2)&", saldo R$ "&FIXED({BM("$G$24")},2)&". Eles não entram no dinheiro disponível."',
    f'=IF({PI("H")}>0,"💳 Fatura do cartão deste mês: R$ "&FIXED({PI("H")},2)&" (já descontada do disponível). Uso do limite: "&ROUND(\'Cartão e Dívidas\'!B6*100,0)&"%.","💳 Nenhum gasto no cartão de crédito neste mês.")',
]
for i, f in enumerate(alerts):
    r = 22 + i
    B.merge_cells(f"B{r}:M{r}"); B[f"B{r}"] = f
    B[f"B{r}"].font = ft(False, 10); B[f"B{r}"].alignment = AL_L
    for col in "BCDEFGHIJKLM":
        B[f"{col}{r}"].fill = fl(SOFT if i % 2 == 0 else "FFFFFF"); B[f"{col}{r}"].border = Border(bottom=th)
    B.row_dimensions[r].height = 24
B.conditional_formatting.add("B22:B28", FormulaRule(formula=['LEFT($B22,1)="🔴"'], font=Font(name=FONT, bold=True, color=RED)))
# tabela dados
band(B, f"B{DT0-1}:F{DT0-1}", "Dados dos gráficos (automático — não editar)", 9, "95A5A6", h=18, align=AL_L)
head(B, DT0, 2, ["Categoria", "Realizado", "Limite", "% usado", "Situação"])
for i in range(NCAT):
    r = DT0 + 1 + i
    B[f"B{r}"] = f"=Planejamento!A{PCAT0+i}"
    B[f"C{r}"] = f"=INDEX(Planejamento!$B${PCAT0+i}:$M${PCAT0+i},{MM})"
    B[f"D{r}"] = f"=IF(B{r}=\"\",0,N(Config!$C${35+i}))"
    B[f"E{r}"] = f'=IF(OR(B{r}="",D{r}=0),"",C{r}/D{r})'
    B[f"F{r}"] = f'=IF(B{r}="","",IF(D{r}=0,IF(C{r}>0,"Sem limite","—"),IF(C{r}>D{r},"Acima do limite",IF(C{r}>=0.8*D{r},"Atenção","OK"))))'
    for col, fmt in zip("BCDEF", [None, BRL, BRL, "0%", None]): out(B[f"{col}{r}"], fmt)
# gráficos
d1 = DoughnutChart(); d1.title = "Gastos por categoria (mês)"; d1.holeSize = 60
d1.add_data(Reference(B, min_col=3, min_row=DT0 + 1, max_row=DT0 + NCAT), titles_from_data=False)
d1.set_categories(Reference(B, min_col=2, min_row=DT0 + 1, max_row=DT0 + NCAT)); d1.height = 9; d1.width = 13
B.add_chart(d1, "B30")
b1 = BarChart(); b1.type = "col"; b1.title = "Receitas x Saídas (ano)"; b1.style = 10
b1.add_data(Reference(P, min_col=3, min_row=4, max_row=16), titles_from_data=True)
b1.add_data(Reference(P, min_col=10, min_row=4, max_row=16), titles_from_data=True)
b1.set_categories(Reference(P, min_col=1, min_row=5, max_row=16))
b1.series[0].graphicalProperties.solidFill = TEAL; b1.series[1].graphicalProperties.solidFill = ORANGE
b1.height = 9; b1.width = 15; b1.y_axis.numFmt = "#,##0"; b1.y_axis.delete = False; b1.x_axis.delete = False
B.add_chart(b1, "H30")
b2 = BarChart(); b2.type = "bar"; b2.title = "Limite x realizado por categoria"; b2.style = 10
b2.add_data(Reference(B, min_col=4, min_row=DT0, max_row=DT0 + NCAT), titles_from_data=True)
b2.add_data(Reference(B, min_col=3, min_row=DT0, max_row=DT0 + NCAT), titles_from_data=True)
b2.set_categories(Reference(B, min_col=2, min_row=DT0 + 1, max_row=DT0 + NCAT))
b2.series[0].graphicalProperties.solidFill = "C8D6D3"; b2.series[1].graphicalProperties.solidFill = TEAL
b2.height = 10; b2.width = 13; b2.y_axis.numFmt = "#,##0"; b2.y_axis.delete = False; b2.x_axis.delete = False
B.add_chart(b2, "B50")
l1 = LineChart(); l1.title = "Evolução do patrimônio"
l1.add_data(Reference(P, min_col=18, min_row=4, max_row=16), titles_from_data=True)
l1.set_categories(Reference(P, min_col=1, min_row=5, max_row=16)); l1.legend = None
l1.series[0].graphicalProperties.line.solidFill = TEAL; l1.series[0].graphicalProperties.line.width = 30000
l1.height = 10; l1.width = 15; l1.y_axis.numFmt = "#,##0"; l1.y_axis.delete = False; l1.x_axis.delete = False
B.add_chart(l1, "H50")
widths(B, {"A": 2, **{chr(c): 11.5 for c in range(ord("B"), ord("N"))}})

# ======================= INÍCIO =======================
N_ = S["Início"]
band(N_, "B1:H1", "💚  MINHA PLANILHA FINANCEIRA", 20, TEAL, h=46)
N_.merge_cells("B2:H2"); N_["B2"] = "Controle cada real que você recebe, gasta, guarda ou investe."
N_["B2"].font = ft(True, 11, TEAL, True); N_["B2"].alignment = Alignment(horizontal="center")
navs2 = ["Dashboard", "Diário", "Receitas", "Despesas", "Cartão e Dívidas", "Reserva", "Investimentos", "Metas", "Planejamento", "Simulador", "Config"]
for j, sh in enumerate(navs2):
    c = N_.cell(row=4 + j // 7, column=2 + j % 7); link(c, sh, sh)
N_.row_dimensions[4].height = 26; N_.row_dimensions[5].height = 26
for j, m in enumerate(MESES):
    c = N_.cell(row=6 + j // 7, column=2 + j % 7); link(c, m, m); c.fill = fl(MINT); c.font = ft(True, 9, TEAL)
blocks = [
    ("🚀 COMEÇAR EM 4 PASSOS", [
        "1) Config: coloque seu nome e confira salário e benefícios (já vêm R$ 1.500, VA R$ 890 e VR R$ 400 — edite se mudar).",
        "2) Despesas: confira suas contas fixas (Inglês, Faculdade, ChatGPT, Água, Internet). Ajuste o dia de vencimento de cada uma.",
        "3) Durante o mês: na aba do mês (ex.: Novembro) marque ☑ quando pagar uma conta. Gastos do dia a dia vão no Diário.",
        "4) Abra o Dashboard: ele mostra quanto você ainda pode gastar, alertas e gráficos. O mês exibido é o atual (mude em Config)."]),
    ("🧠 REGRAS QUE EVITAM CONTAR EM DOBRO", [
        "• Benefícios (VA/VR) ficam separados: só pagam gastos lançados com forma “Vale-alimentação” ou “Vale-refeição” e nunca viram dinheiro disponível.",
        "• Adiantamento salarial é parte do salário: aparece separado, mas o salário é contado uma vez só.",
        "• Contas fixas: só marque ☑ na aba do mês. Não lance de novo no Diário.",
        "• Compras parceladas e dívidas: cadastre uma vez em Cartão e Dívidas. Cada parcela entra sozinha no mês certo.",
        "• Pagamento da fatura do cartão NÃO é lançado: as compras no crédito já foram contadas quando você comprou.",
        "• Reserva e investimentos não são despesa: entram como dinheiro guardado e no seu patrimônio."]),
    ("🎨 CÓDIGO DE CORES", [
        "Amarelo com texto azul = você preenche.   Branco/cinza com texto preto = automático (não mexer).   Texto verde = vem de outra aba.",
        "Linhas com “EXEMPLO – pode apagar” mostram o formato. Apague esses exemplos quando começar a usar (só as colunas amarelas)."]),
    ("📱💻 USAR NO GOOGLE PLANILHAS (celular e computador)", [
        "1) No computador, abra drive.google.com → Novo → Upload de arquivo → escolha este .xlsx.",
        "2) Clique no arquivo → Abrir com → Planilhas Google. Depois: Arquivo → Salvar como Planilhas Google (recomendado).",
        "3) No celular, instale o app “Planilhas Google” e abra o mesmo arquivo. Tudo sincroniza sozinho entre celular e PC.",
        "4) Opcional: na aba do mês, selecione a coluna Pago (B32:B61) → Inserir → Caixa de seleção. As fórmulas aceitam ☑ e a caixa nativa.",
        "Também funciona no Excel (PC e celular) e no Excel Online/OneDrive."]),
    ("📌 OBSERVAÇÕES", [
        "• Valores de salário, benefícios e despesas são referências editáveis. Vencimentos e limites de categoria são exemplos: ajuste aos seus.",
        "• A planilha não inventa rentabilidade: rendimentos só entram quando você registra o valor real do extrato.",
        "• Planilha de uso pessoal, para as suas finanças. Não substitui orientação financeira profissional."]),
]
r = 9
for ttl, lines in blocks:
    band(N_, f"B{r}:H{r}", ttl, 11, TEAL, h=24, align=AL_L); r += 1
    for ln in lines:
        N_.merge_cells(f"B{r}:H{r}"); N_[f"B{r}"] = ln; N_[f"B{r}"].font = ft(False, 10)
        N_[f"B{r}"].alignment = AL_L; N_.row_dimensions[r].height = 32 if len(ln) > 105 else 20; r += 1
    r += 1
widths(N_, {"A": 2, **{chr(c): 17 for c in range(ord("B"), ord("I"))}})

# exemplo: contas pagas em novembro
NOV = S["Novembro"]
NOV[f"B{CROW0}"] = "☑"; NOV[f"H{CROW0}"] = 370; NOV[f"I{CROW0}"] = dt.date(2026, 11, 5)
NOV[f"B{CROW0+1}"] = "☑"; NOV[f"H{CROW0+1}"] = 210; NOV[f"I{CROW0+1}"] = dt.date(2026, 11, 10)

for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
from openpyxl.chart.line_chart import LineChart as _LC
for ws in wb.worksheets:
    for chx in ws._charts:
        if isinstance(chx, _LC):
            for se in chx.series: se.smooth = False
        if isinstance(chx, DoughnutChart) and chx.legend is not None:
            chx.legend.position = "b"
wb.active = 1
wb.save(OUT)
print("ok")
