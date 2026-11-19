"""
Fake files for the claims macros (fictional company: Acme Pumps).

1. Claims_Export_Oct2026.xlsx   - the month's claims cost export, grouped outline view
     sheet RAW DATA: Settl. Per. | Order | Plant | Descr. Deb. | Order Type | Description | Value
       detail rows   : Settl. Per. filled (order details only on the first line, blanks below)
       order subtotal: Settl. Per. blank, Order filled
       grand total   : last row, Settl. Per. and Order blank
     sheet Claim No lookup: order -> claim number (stands in for the sales order list)
2. Claims_and_Goodwill_2026.xlsx - year-to-date register, January to September already posted
     CLAIMS 2026 (warranty, order type WA01), GOODWILL 2026 (order type GW01), GRAPH
"""
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

random.seed(7)
HERE = Path(__file__).parent
CUSTOMERS = ["Riverbend Water Board", "Karoo Mining Co", "Northgate Municipality", "Bluecrest Chemicals",
             "Summit Agri Co-op", "Harbourview Refinery", "Ironveld Smelters", "Greenfield Estates",
             "Coastline Desalination", "Highveld Power Station"]
DESCRIPTIONS = ["Mechanical seal leaking", "Bearing failure after 3 months", "Impeller worn - slurry",
                "Motor overheating", "Casing crack on delivery", "Vibration on site", "Coupling misaligned",
                "Valve seat damaged", "Paint failure - coastal site", "Wrong flange drilling"]
NAVY = PatternFill("solid", fgColor="1F3A5F")
WHITE = Font(bold=True, color="FFFFFF")

# ---------------------------------------------------------------- existing register (Jan-Sep)
existing = {"WA01": [], "GW01": []}
order = 2001100
for otype, n in (("WA01", 22), ("GW01", 9)):
    for _ in range(n):
        order += random.randint(3, 40)
        existing[otype].append({
            "order": order, "cust": random.choice(CUSTOMERS), "desc": random.choice(DESCRIPTIONS),
            "claim": f"CL{10000 + order % 9000}",
            "accum": round(random.uniform(0, 30000), 2) if otype == "WA01" and random.random() < 0.3 else 0,
            "months": {m: round(random.uniform(1500, 45000), 2)
                       for m in random.sample(range(1, 10), random.randint(1, 4))},
        })

# ---------------------------------------------------------------- October export
oct_orders = []
for otype in ("WA01", "GW01"):
    oct_orders += [(otype, e) for e in random.sample(existing[otype], 4 if otype == "WA01" else 2)]
for otype, n in (("WA01", 8), ("GW01", 4)):
    for _ in range(n):
        order += random.randint(3, 40)
        oct_orders.append((otype, {"order": order, "cust": random.choice(CUSTOMERS),
                                    "desc": random.choice(DESCRIPTIONS), "claim": f"CL{10000 + order % 9000}"}))
random.shuffle(oct_orders)

wb = Workbook()
ws = wb.active
ws.title = "RAW DATA"
ws.append(["Settl. Per.", "Order", "Plant", "Descr. Deb.", "Order Type", "Description",
           "Value (Warranties/goodwill)"])
for c in ws[1]:
    c.font, c.fill = WHITE, NAVY
grand = 0.0
for otype, e in oct_orders:
    lines = [round(random.uniform(400, 25000), 2) for _ in range(random.randint(1, 4))]
    for i, amt in enumerate(lines):
        ws.append(["010.2026", e["order"] if i == 0 else "", random.choice(["1100", "1200"]) if i == 0 else "",
                   e["cust"] if i == 0 else "", otype if i == 0 else "", e["desc"] if i == 0 else "", amt])
        ws.row_dimensions[ws.max_row].outline_level = 2
    ws.append(["", e["order"], "", "", "", "", round(sum(lines), 2)])        # order subtotal
    ws.row_dimensions[ws.max_row].outline_level = 1
    ws.cell(row=ws.max_row, column=7).font = Font(bold=True)
    grand += sum(lines)
ws.append(["", "", "", "", "", "", round(grand, 2)])                          # grand total
ws.cell(row=ws.max_row, column=7).font = Font(bold=True)
for col, w in zip("ABCDEFG", (11, 10, 7, 26, 11, 30, 16)):
    ws.column_dimensions[col].width = w
lk = wb.create_sheet("Claim No lookup")
lk.append(["Order", "Claim No"])
for _, e in oct_orders:
    lk.append([e["order"], e["claim"]])
wb.save(HERE / "Claims_Export_Oct2026.xlsx")
print(f"Claims_Export_Oct2026.xlsx: {len(oct_orders)} orders, total R{grand:,.2f}")

# ---------------------------------------------------------------- YTD register
MONTHS = [date(2026, m, 1) for m in range(1, 13)]
reg = Workbook()
totals = {}


def build_ledger(ws, title, items, claims_layout):
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14, color="1F3A5F")
    hr = 3
    heads = (["Order No", "Customer", "DESCRIPTION", "Claim No", "Accum to Dec 25"] if claims_layout
             else ["Order No", "Customer", "Description"])
    first_m = len(heads) + 1
    heads += MONTHS + (["YTD 2026", "Total Costs to Date"] if claims_layout else ["YTD 2026"])
    for c, h in enumerate(heads, 1):
        cell = ws.cell(row=hr, column=c, value=h)
        cell.font, cell.fill = WHITE, NAVY
        if isinstance(h, date):
            cell.number_format = "mmm-yy"
    last_m = first_m + 11
    ytd = last_m + 1
    r = hr + 1
    for e in items:
        ws.cell(row=r, column=1, value=e["order"])
        ws.cell(row=r, column=2, value=e["cust"])
        ws.cell(row=r, column=3, value=e["desc"])
        if claims_layout:
            ws.cell(row=r, column=4, value=e["claim"])
            ws.cell(row=r, column=5, value=e["accum"] or None)
        for m, v in e["months"].items():
            ws.cell(row=r, column=first_m + m - 1, value=v)
        r += 1
    r += 3                                                      # spare rows, ready for new orders
    first, last = hr + 1, r - 1
    total_row = r
    ws.cell(row=total_row, column=1,
            value=f"Total {'Claims' if claims_layout else 'Goodwill'} From January 2026 to December 2026")
    for rr in range(first, last + 1):
        ws.cell(row=rr, column=ytd,
                value=f"=SUM({get_column_letter(first_m)}{rr}:{get_column_letter(last_m)}{rr})")
        if claims_layout:
            ws.cell(row=rr, column=ytd + 1, value=f"={get_column_letter(ytd)}{rr}+E{rr}")
    start_col = 5 if claims_layout else first_m
    end_col = ytd + 1 if claims_layout else ytd
    for c in range(start_col, end_col + 1):
        L = get_column_letter(c)
        ws.cell(row=total_row, column=c, value=f"=SUM({L}{first}:{L}{last})").font = Font(bold=True)
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    for rr in range(first, total_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=rr, column=c).number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 30
    for c in range(4, end_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.freeze_panes = ws.cell(row=hr + 1, column=4)
    return first_m, total_row


cl = reg.active
cl.title = "CLAIMS 2026"
c_first, c_total = build_ledger(cl, "Warranty Claims Register 2026", existing["WA01"], True)
gw = reg.create_sheet("GOODWILL 2026")
g_first, g_total = build_ledger(gw, "Goodwill Register 2026", existing["GW01"], False)

gr = reg.create_sheet("GRAPH")
gr.append(["Month", "Pure Claim Cost", "Goodwill Claim Cost", "Total Claim Cost", "Turnover", "% of Turnover"])
for c in gr[1]:
    c.font, c.fill = WHITE, NAVY
for i, m in enumerate(MONTHS):
    r = i + 2
    cL, gL = get_column_letter(c_first + i), get_column_letter(g_first + i)
    gr.cell(row=r, column=1, value=m).number_format = "mmm-yy"
    gr.cell(row=r, column=2, value=f"='CLAIMS 2026'!{cL}{c_total}")
    gr.cell(row=r, column=3, value=f"='GOODWILL 2026'!{gL}{g_total}")
    gr.cell(row=r, column=4, value=f"=B{r}+C{r}")
    gr.cell(row=r, column=5, value=round(random.uniform(38e6, 52e6), -3) if i < 9 else None)
    gr.cell(row=r, column=6, value=f'=IF(E{r}="","",D{r}/E{r})').number_format = "0.00%"
    for c in (2, 3, 4, 5):
        gr.cell(row=r, column=c).number_format = "#,##0"
gr["O1"], gr["P1"], gr["Q1"] = "Pure Claim Cost (current month)", "Goodwill Claim Cost", "Total Claim Cost"
sepC, sepG = get_column_letter(c_first + 8), get_column_letter(g_first + 8)
gr["O2"] = f"='CLAIMS 2026'!{sepC}{c_total}"
gr["P2"] = f"='GOODWILL 2026'!{sepG}{g_total}"
gr["Q2"] = "=SUM(O2:P2)"
for c in ("O2", "P2", "Q2"):
    gr[c].number_format = "#,##0"
gr["O4"] = "O2/P2 point at the latest month posted; the macro moves them each month."
gr["O4"].font = Font(italic=True, color="777777")
ch = BarChart()
ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
ch.title = "Claims cost by month (ZAR)"
ch.add_data(Reference(gr, min_col=2, max_col=3, min_row=1, max_row=13), titles_from_data=True)
ch.set_categories(Reference(gr, min_col=1, min_row=2, max_row=13))
ch.height, ch.width = 8, 18
gr.add_chart(ch, "H6")
for col, w in zip("ABCDEF", (10, 16, 18, 16, 14, 13)):
    gr.column_dimensions[col].width = w
for col in "OPQ":
    gr.column_dimensions[col].width = 22
reg.save(HERE / "Claims_and_Goodwill_2026.xlsx")
print("Claims_and_Goodwill_2026.xlsx: register Jan-Sep,",
      len(existing["WA01"]), "warranty +", len(existing["GW01"]), "goodwill orders")
n_new_wa = sum(1 for t, e in oct_orders if t == "WA01" and "months" not in e)
n_new_gw = sum(1 for t, e in oct_orders if t == "GW01" and "months" not in e)
print(f"Expected October run: CLAIMS 4 updated, {n_new_wa} new | GOODWILL 2 updated, {n_new_gw} new")
