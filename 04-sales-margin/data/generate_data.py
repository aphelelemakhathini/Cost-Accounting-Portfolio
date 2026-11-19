"""
Fake invoiced-sales export (grouped view) for modSalesMargin, fictional company Acme Pumps.

Row 1 headers, data from row 2. Each order has detail lines plus ONE total line;
the total line is the one with Plant (column R) blank.
  A Order | B Item | C Billing doc | D Billing date | E Net Sales Revenue
  F..N the 9 cost lines (material, variable prod., other special direct, reject & rework,
       fixed prod., production OH %, development costs x2, material OH %)
  O Cost of goods sold (ERP's own) | P CMII | Q CMII %   (ignored - the macro computes its own)
  R Plant | S Segment_Market Area | T Market Area | U Distr. Channel | V Descr. Deb. (customer)

Built-in cases: losses, zero-sales orders, no-cost orders, both-zero orders, partially
invoiced outliers (COGS 1000%+ of sales), internal adjustment orders with letters
(must be excluded), group-company customers ("Acme ...").
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

random.seed(3)
HERE = Path(__file__).parent
AREAS = ["MIN", "WAT", "IND", "CHM", "BLD", "PWR"]
CUSTOMERS = {
    "MIN": ["Karoo Mining Co", "Ironveld Smelters", "Deepshaft Gold"],
    "WAT": ["Riverbend Water Board", "Northgate Municipality", "Coastline Desalination"],
    "IND": ["Summit Agri Co-op", "Greenfield Estates", "Valley Paper Mill"],
    "CHM": ["Bluecrest Chemicals", "Harbourview Refinery"],
    "BLD": ["Skyline Property Group", "Metro Hospitals Trust"],
    "PWR": ["Highveld Power Station", "Sunridge Solar Thermal"],
}
GROUP = ["Acme Pumps Namibia", "Acme Pumps Botswana", "Acme Valves Europe"]

wb = Workbook()
ws = wb.active
ws.title = "Sales Invoiced"
heads = ["Order", "Item", "Billing Doc.", "Billing Date", "Net Sales Revenue", "Material for Sales",
         "Variable production costs", "Other special direct costs", "Reject and rework",
         "Fixed production costs", "Production overhead costs (%)", "Development Costs (WTE) from Sales",
         "Development Costs (WTE) R&D", "Material overhead costs (%)", "Cost of goods sold", "CMII", "CMII %",
         "Plant", "Segment_Market Area", "Market Area", "Distr. Channel", "Descr. Deb."]
ws.append(heads)
for c in ws[1]:
    c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3A5F")

order, billing = 30051000, 90120000
stats = {"orders": 0, "losses": 0}


def cost_split(cogs):
    w = [0.55, 0.12, 0.03, 0.01, 0.08, 0.12, 0.02, 0.01, 0.06]
    w = [x * random.uniform(0.7, 1.3) for x in w]
    s = sum(w)
    return [round(cogs * x / s, 2) for x in w]


def add_order(onum, sales, cogs, plant, segment, area, channel, cust):
    global billing
    n_lines = random.randint(1, 3)
    shares = [random.random() for _ in range(n_lines)]
    tot = sum(shares)
    lines = []
    for i in range(n_lines):
        f = shares[i] / tot
        costs = cost_split(cogs * f)
        lines.append([round(sales * f, 2)] + costs)
    # detail lines: Market Area / Segment can be blank on detail lines too
    for i, l in enumerate(lines):
        billing += random.randint(1, 4)
        ws.append([onum, (i + 1) * 10, billing, date(2026, 10, 1) + timedelta(days=random.randint(0, 30)),
                   *l, round(sum(l[1:]), 2), round(l[0] - sum(l[1:]), 2), None,
                   plant, segment if random.random() > 0.2 else None, area if random.random() > 0.2 else None,
                   channel, cust])
        ws.cell(row=ws.max_row, column=4).number_format = "dd/mm/yyyy"
    tot_line = [round(sum(l[k] for l in lines), 2) for k in range(10)]
    ws.append([onum, None, None, None, *tot_line, round(sum(tot_line[1:]), 2),
               round(tot_line[0] - sum(tot_line[1:]), 2), None, None, None, None, None, None])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    stats["orders"] += 1
    if sales and (sales - cogs) / sales < 0:
        stats["losses"] += 1


for _ in range(240):
    order += random.randint(1, 30)
    area = random.choice(AREAS)
    channel = random.choice([1, 1, 1, 2, 3])
    segment = {1: area, 2: area, 3: "SE"}[channel]
    if random.random() < 0.1:
        segment = "VA"
    cust = random.choice(GROUP) if random.random() < 0.08 else random.choice(CUSTOMERS[area])
    plant = random.choice(["1100", "1200", "1300"])
    sales = round(random.uniform(10_000, 420_000), 2)
    margin = random.gauss(0.30 if channel != 2 else 0.45, 0.17)
    cogs = round(sales * (1 - margin), 2)
    r = random.random()
    if r < 0.03:
        sales = 0.0                                  # cost booked, nothing invoiced yet
    elif r < 0.05:
        cogs = 0.0                                   # invoiced, no cost yet
    elif r < 0.06:
        sales, cogs = 0.0, 0.0                       # credit/debit netted to zero
    elif r < 0.08:
        sales = round(random.uniform(20, 900), 2)    # partial invoice: outlier
    add_order(order, sales, cogs, plant, segment, area, channel, cust)

for adj in ("ADJ0001", "ADJ0002"):                    # internal adjustments - must be excluded
    add_order(adj, round(random.uniform(-50_000, 50_000), 2), 0.0, "1200", "IND", "IND", 1, "Internal adjustment")

ws.append([None] * 4 + [round(sum((ws.cell(r, 5).value or 0) for r in range(2, ws.max_row + 1)
                                  if ws.cell(r, 18).value is None), 2)])   # grand total, blank order
for col, w in zip("ABCDEFGHIJKLMNOPQRSTUV", [11, 6, 12, 11] + [14] * 13 + [7, 10, 9, 8, 26]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B2"
wb.save(HERE / "Sales_Invoiced_Oct2026.xlsx")
print(f"Sales_Invoiced_Oct2026.xlsx: {stats['orders']} orders incl. 2 internal adjustments, "
      f"{ws.max_row} rows")
