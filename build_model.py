"""
Builds the Excel model: Construction-Finance Appraisal of a G+5 Commercial
Building (lease-rental project) in Mangaluru. Every output is an Excel formula
driven by the Inputs sheet; stress scenarios are separate model blocks so the
sensitivity table recalculates live.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

wb = Workbook()
F = "Arial"
BLUE, BLACK, GREEN = Font(name=F, color="0000FF"), Font(name=F), Font(name=F, color="008000")
B = Font(name=F, bold=True)
H = Font(name=F, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
YEL = PatternFill("solid", fgColor="FFFF00")
thin = Border(top=Side(style="thin"))
CR = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
X = '0.00"x"'

# ------------------------------------------------------------------ Inputs
ws = wb.active; ws.title = "Inputs"
ws["A1"] = "Construction-Finance Appraisal - G+5 Commercial Building, Mangaluru"; ws["A1"].font = Font(name=F, bold=True, size=14)
ws["A2"] = "Rs crore unless stated. Blue = input, black = formula, green = link. Yellow = key assumption."; ws["A2"].font = Font(name=F, italic=True, size=9)
inputs = [
    ("PROJECT COST", None, None, None),
    ("Built-up area", 60000, "sq ft", "G+5 commercial block (ties to G+5 STAAD/ETABS design project)"),
    ("Construction cost", 2600, "Rs / sq ft", "RCC frame, finishes, MEP - Tier-2 Karnataka benchmark (assumption)"),
    ("Land cost", 6.0, "Rs crore", "Assumed market value, city-centre plot"),
    ("Soft costs (design, approvals, contingency)", 0.10, "% of hard cost", "Assumption"),
    ("FINANCING", None, None, None),
    ("Debt share of project cost (ex-land)", 0.70, "%", "Land funded by promoter equity; bank funds 70% of construction + soft cost"),
    ("Interest rate", 0.105, "% p.a.", "Assumed lending rate for commercial real estate"),
    ("Construction period", 2, "years", "Loan drawn 50% in each year; interest during construction paid by promoter"),
    ("REVENUE & OPERATIONS", None, None, None),
    ("Leasable efficiency", 0.80, "% of built-up", "Carpet/leasable ratio"),
    ("Lease rent (Year 3)", 75, "Rs / sq ft / month", "Grade-B office/retail rent, Mangaluru (assumption)"),
    ("Annual rent escalation", 0.045, "% p.a.", "~15% every 3 years, smoothed"),
    ("Occupancy - ops year 1", 0.60, "%", "Lease-up ramp"),
    ("Occupancy - ops year 2", 0.80, "%", ""),
    ("Occupancy - stabilised (year 3+)", 0.90, "%", ""),
    ("Opex + property tax", 0.15, "% of rent", "Maintenance, insurance, property tax"),
    ("Exit cap rate (terminal value)", 0.085, "%", "Yield at which the asset could be sold in final year"),
    ("CREDIT POLICY", None, None, None),
    ("Minimum DSCR (any year)", 1.20, "x", "Typical bank covenant for lease-rental discounting"),
    ("Minimum average DSCR", 1.40, "x", ""),
    ("STRESS SHOCKS", None, None, None),
    ("Rent shock", -0.10, "%", "Scenario 2"),
    ("Construction cost overrun", 0.15, "%", "Scenario 3"),
    ("Interest rate shock", 0.01, "abs", "Scenario 4 (+100 bps)"),
    ("Occupancy shock", -0.10, "abs", "Scenario 5 (-10 pp on every year)"),
]
ws.column_dimensions["A"].width = 44; ws.column_dimensions["B"].width = 12; ws.column_dimensions["C"].width = 18; ws.column_dimensions["D"].width = 70
name = {}
r = 4
for lab, val, unit, note in inputs:
    ws.cell(r, 1, lab)
    if val is None:
        ws.cell(r, 1).font = B; ws.cell(r, 1).fill = SUB
        for c in range(2, 5): ws.cell(r, c).fill = SUB
    else:
        c = ws.cell(r, 2, val); c.font = BLUE
        c.number_format = PCT if isinstance(val, float) and abs(val) < 1 and unit not in ("Rs crore",) else ('0.00"x"' if unit == "x" else '#,##0.##')
        ws.cell(r, 1).font = BLACK; ws.cell(r, 3, unit).font = Font(name=F, size=9); ws.cell(r, 4, note).font = Font(name=F, size=9, italic=True)
        name[lab] = f"Inputs!$B${r}"
    r += 1
for k in ["Lease rent (Year 3)", "Interest rate", "Construction cost", "Occupancy - stabilised (year 3+)"]:
    ws[name[k].split("!")[1].replace("$", "")].fill = YEL
I = name

# ------------------------------------------------------------------ Model blocks
m = wb.create_sheet("Model")
m.column_dimensions["A"].width = 38; m.column_dimensions["B"].width = 12
YRS = 12; C0 = 3  # column C = year 1
for j in range(YRS): m.column_dimensions[L(C0 + j)].width = 10
m["A1"] = "Year-wise model (Rs crore). Years 1-2 construction, 3-12 operations."; m["A1"].font = B

scen = [
    ("1. Base case", {}),
    ("2. Rent -10%", {"rent": I["Rent shock"]}),
    ("3. Construction cost +15%", {"cost": I["Construction cost overrun"]}),
    ("4. Interest rate +100 bps", {"rate": I["Interest rate shock"]}),
    ("5. Occupancy -10 pp", {"occ": I["Occupancy shock"]}),
    ("6. Combined stress (rent -10%, cost +15%)", {"rent": I["Rent shock"], "cost": I["Construction cost overrun"]}),
]
results = []
row = 3
for title, sh in scen:
    top = row
    m.cell(row, 1, title).font = H
    for c in range(1, C0 + YRS): m.cell(row, c).fill = HFILL
    for j in range(YRS):
        cc = m.cell(row, C0 + j, f"Y{j+1}"); cc.font = H; cc.alignment = Alignment(horizontal="center")
    row += 1
    # block-level drivers in column B
    def drv(lbl, f, fmt=CR):
        global row
        m.cell(row, 1, lbl); c = m.cell(row, 2, f); c.number_format = fmt; c.font = GREEN if f.startswith("=Inputs") else BLACK
        row += 1; return f"$B${row-1}"
    cost_mult = f"(1+{sh['cost']})" if "cost" in sh else "1"
    hard = drv("Hard cost", f"={I['Built-up area']}*{I['Construction cost']}*{cost_mult}/10^7")
    soft = drv("Soft cost", f"={hard}*{I['Soft costs (design, approvals, contingency)']}")
    land = drv("Land", f"={I['Land cost']}")
    loan = drv("Bank loan", f"=({hard}+{soft})*{I['Debt share of project cost (ex-land)']}")
    rate = drv("Interest rate", f"={I['Interest rate']}" + (f"+{sh['rate']}" if "rate" in sh else ""), PCT)
    idc = drv("Interest during construction (promoter)", f"={loan}*0.5*{rate}+{loan}*{rate}")
    equity = drv("Promoter equity incl. land & IDC", f"={hard}+{soft}+{land}-{loan}+{idc}")
    rent0 = drv("Rent, Y3 (Rs/sqft/month)", f"={I['Lease rent (Year 3)']}" + (f"*(1+{sh['rent']})" if "rent" in sh else ""), '#,##0.0')
    occ_sh = f"+{sh['occ']}" if "occ" in sh else ""
    lines = {}
    def line(lbl, fn, fmt=CR, bold=False, border=False):
        global row
        m.cell(row, 1, lbl).font = B if bold else BLACK
        for j in range(YRS):
            col = L(C0 + j); c = m.cell(row, C0 + j, fn(j, col)); c.number_format = fmt
            if bold: c.font = B
            if border: c.border = thin
        lines[lbl] = row; row += 1
    yr = lambda j: j + 1
    line("Occupancy", lambda j, c: 0 if j < 2 else f"=MAX(0,{I['Occupancy - ops year 1'] if j==2 else I['Occupancy - ops year 2'] if j==3 else I['Occupancy - stabilised (year 3+)']}{occ_sh})", PCT)
    line("Rent (Rs/sqft/month)", lambda j, c: 0 if j < 2 else f"={rent0}*(1+{I['Annual rent escalation']})^({j}-2)", '#,##0.0;;"-"')
    line("Rental income", lambda j, c: f"={I['Built-up area']}*{I['Leasable efficiency']}*{c}{lines['Occupancy']}*{c}{lines['Rent (Rs/sqft/month)']}*12/10^7")
    line("Opex + property tax", lambda j, c: f"=-{c}{lines['Rental income']}*{I['Opex + property tax']}")
    line("Net operating income (CFADS)", lambda j, c: f"={c}{lines['Rental income']}+{c}{lines['Opex + property tax']}", bold=True, border=True)
    line("Principal repayment schedule", lambda j, c: 0 if j < 2 else f"={L(C0+j)}$2", PCT)
    line("Opening loan", lambda j, c: f"={loan}*0.5" if j == 0 else (f"={L(C0+j-1)}{row+3}+{loan}*0.5" if j == 1 else f"={L(C0+j-1)}{row+3}"))
    line("Interest", lambda j, c: 0 if j < 2 else f"={c}{lines['Opening loan']}*{rate}")
    line("Principal", lambda j, c: 0 if j < 2 else f"={loan}*{c}{lines['Principal repayment schedule']}")
    line("Closing loan", lambda j, c: f"={c}{lines['Opening loan']}-{c}{lines['Principal']}")
    line("Debt service", lambda j, c: f"={c}{lines['Interest']}+{c}{lines['Principal']}", bold=True)
    line("DSCR", lambda j, c: '=""' if j < 2 else f"=IFERROR({c}{lines['Net operating income (CFADS)']}/{c}{lines['Debt service']},0)", X, bold=True)
    line("Terminal value (exit)", lambda j, c: 0 if j < YRS - 1 else f"={c}{lines['Net operating income (CFADS)']}*(1+{I['Annual rent escalation']})/{I['Exit cap rate (terminal value)']}")
    line("Project cash flow", lambda j, c: (f"=-({hard}+{soft}+{land})/2" if j < 2 else f"={c}{lines['Net operating income (CFADS)']}+{c}{lines['Terminal value (exit)']}"))
    line("Equity cash flow", lambda j, c: (f"=-{equity}/2" if j < 2 else f"={c}{lines['Project cash flow']}-{c}{lines['Debt service']}-{c}{lines['Closing loan']}*{1 if j==YRS-1 else 0}"))
    first, last = L(C0), L(C0 + YRS - 1); ops0 = L(C0 + 2)
    dsr = lines["DSCR"]
    res = {
        "title": title,
        "loan": f"=Model!{loan}", "cost": f"=Model!{hard}+Model!{soft}+Model!{land}",
        "min": f"=MIN(Model!{ops0}{dsr}:{last}{dsr})",
        "avg": f"=AVERAGE(Model!{ops0}{dsr}:{last}{dsr})",
        "pirr": f"=IRR(Model!{first}{lines['Project cash flow']}:{last}{lines['Project cash flow']})",
        "eirr": f"=IRR(Model!{first}{lines['Equity cash flow']}:{last}{lines['Equity cash flow']})",
        "stab_dscr": f"=Model!{L(C0+4)}{dsr}",
        "ltv": f"=Model!{loan}/(Model!{L(C0+2)}{lines['Net operating income (CFADS)']}*(1+{I['Annual rent escalation']})^2/{I['Exit cap rate (terminal value)']})",
    }
    results.append(res)
    row += 1

# repayment schedule row 2 (shared, step-up structure), inputs in blue
m["A2"] = "Principal repayment % by year (step-up, sums to 100%)"; m["A2"].font = B
sched = [0, 0, 0, 0.07, 0.08, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14, 0.16]  # 1-yr principal moratorium after completion, then step-up
for j, v in enumerate(sched):
    c = m.cell(2, C0 + j, v); c.font = BLUE; c.number_format = PCT
m.cell(2, 2, f"=SUM({L(C0)}2:{L(C0+YRS-1)}2)").number_format = PCT
m.freeze_panes = "C3"

# ------------------------------------------------------------------ Summary
s = wb.create_sheet("Summary", 0)
s["A1"] = "Credit Appraisal Summary - Stress Test"; s["A1"].font = Font(name=F, bold=True, size=14)
s["A2"] = "All values link to the Model sheet; change any blue input in Inputs and this table updates."; s["A2"].font = Font(name=F, italic=True, size=9)
hdr = ["Scenario", "Project cost (Rs cr)", "Bank loan (Rs cr)", "Min DSCR", "Avg DSCR", "Stabilised DSCR (Y5)", "LTV at stabilisation", "Project IRR", "Equity IRR", "Covenant check"]
for j, h in enumerate(hdr):
    c = s.cell(4, j + 1, h); c.font = H; c.fill = HFILL; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
s.row_dimensions[4].height = 32
for i, rs in enumerate(results):
    r = 5 + i
    s.cell(r, 1, rs["title"])
    for j, (k, fmt) in enumerate([("cost", CR), ("loan", CR), ("min", X), ("avg", X), ("stab_dscr", X), ("ltv", PCT), ("pirr", PCT), ("eirr", PCT)]):
        c = s.cell(r, j + 2, rs[k]); c.number_format = fmt; c.font = GREEN
    s.cell(r, 10, f'=IF(AND(D{r}>={I["Minimum DSCR (any year)"]},E{r}>={I["Minimum average DSCR"]}),"PASS","FAIL")').font = B
s.column_dimensions["A"].width = 42
for j in range(2, 11): s.column_dimensions[L(j)].width = 14
s["A13"] = "How to read this"; s["A13"].font = B
notes = [
    "DSCR = Net operating income / (interest + principal). Banks usually need >= 1.20x in every year and ~1.40x on average.",
    "Lease-up years (Y3-Y4) are the weakest: occupancy is still ramping while interest is highest, so a flat EMI would break the covenant.",
    "The step-up repayment schedule (Model row 2) back-ends principal to match the rent ramp; change it to test other structures.",
    "Recommendation and risks: see README.md in the repository.",
]
for i, t in enumerate(notes):
    s.cell(14 + i, 1, f"- {t}").font = Font(name=F, size=10)
for wsx in wb.worksheets:
    for rr in wsx.iter_rows():
        for c in rr:
            if c.font.name != F: c.font = Font(name=F, bold=c.font.bold, italic=c.font.italic, color=c.font.color, size=c.font.size)
wb.save("Project_Loan_Appraisal_Model.xlsx")
print("saved")
