"""
Fake ERP exports for the production-variance macros (fictional company: Acme Pumps).

1. Order_Cost_Report.xlsx    -> input for modVarianceFlags (stage 1)
   3-row block per production order, numbers as ERP text ("1.234,56"):
     row 1: B plant, E order, G material, I order type
     row 2: F actual labour/unit, I actual material/unit, K std labour/unit, L std material/unit
     row 3: B quantity, O labour % variance, P material variance (R), Q material % variance

2. Order_Operations_Export.xlsx -> input for modHoursVarianceReview (stage 2), sheet "Sheet1"
   one row per operation; order no., material and plant only on the order's first row.

Stories built into the data:
  * MCH19 has no proper setup standard (setups run 2-4 h over plan)
  * MCH22 got a new CNC machine: machine time well over plan while operators train
  * ASM03 finishes consistently under plan (favourable variances)
  * welding on BASEPLATE jobs overruns on large frames
  * plant 1300 (service workshop) works on repair / rewind / claim materials
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

random.seed(42)
HERE = Path(__file__).parent


def erp(x, dec=2):
    """12345.6 -> '12.345,60' (dot thousands, comma decimal), trailing minus like many ERP lists"""
    neg = x < 0
    t = f"{abs(x):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return t + "-" if neg else t


# ------------------------------------------------------------------ 1. cost report
wb = Workbook()
ws = wb.active
ws.title = "Order Cost Report"
ws["A1"] = "Production order cost report: actual vs standard (surcharges excluded)"
ws["A1"].font = Font(bold=True)
ws["A2"] = "Period 10/2026   Company: Acme Pumps (Pty) Ltd   Currency: ZAR"
hdr = {2: "Plant / Qty", 5: "Order", 6: "Act. Labour", 7: "Material", 9: "Type / Act. Mat.",
       11: "Std Labour", 12: "Std Material", 15: "Labour %", 16: "Mat. Var (R)", 17: "Mat. %"}
for c, h in hdr.items():
    ws.cell(row=4, column=c, value=h).font = Font(bold=True)

MATS = {"1100": [f"EP-{n}" for n in range(4100, 4160)], "1200": [f"SP-{n}" for n in range(2200, 2290)]}
r = 6
order = 1004500
for _ in range(900):
    plant = random.choice(["1100", "1100", "1200", "1200", "1200"])
    order += random.randint(1, 5)
    qty = random.choice([1, 1, 2, 3, 4, 6, 10]) if plant == "1100" else random.choice([2, 4, 6, 10, 12, 20, 40])
    otype = "PP01" if plant == "1100" else random.choice(["PP02", "PP02", "PP01"])
    std_lab = round(random.uniform(800, 9000) * (2.2 if plant == "1100" else 1), 2)
    noise = random.gauss(0, 0.32 if plant == "1100" else 0.14)
    if random.random() < 0.04:
        noise += random.uniform(0.6, 1.5)            # big overrun
    act_lab = round(std_lab * max(1 + noise, 0.3), 2)
    std_mat = round(random.uniform(3000, 60000) * (2.5 if plant == "1100" else 1), 2)
    mnoise = random.gauss(0, 0.12 if plant == "1100" else 0.06)
    if random.random() < 0.05:
        mnoise += random.uniform(0.25, 0.9)          # scrap / extra components
    act_mat = round(std_mat * max(1 + mnoise, 0.4), 2)
    lab_pct = (act_lab - std_lab) / std_lab * 100
    mat_var = (act_mat - std_mat) * qty
    mat_pct = (act_mat - std_mat) / std_mat * 100

    ws.cell(row=r, column=2, value=plant)
    ws.cell(row=r, column=5, value=str(order))
    ws.cell(row=r, column=7, value=random.choice(MATS[plant]))
    ws.cell(row=r, column=9, value=otype)
    ws.cell(row=r + 1, column=6, value=erp(act_lab))
    ws.cell(row=r + 1, column=9, value=erp(act_mat))
    ws.cell(row=r + 1, column=11, value=erp(std_lab))
    ws.cell(row=r + 1, column=12, value=erp(std_mat))
    ws.cell(row=r + 2, column=2, value=erp(qty, 0))
    ws.cell(row=r + 2, column=15, value=erp(lab_pct, 1))
    ws.cell(row=r + 2, column=16, value=erp(mat_var))
    ws.cell(row=r + 2, column=17, value=erp(mat_pct, 1))
    r += 3
wb.save(HERE / "Order_Cost_Report.xlsx")
print(f"Order_Cost_Report.xlsx: 900 orders, rows 6-{r - 1}")

# ------------------------------------------------------------------ 2. operations export
END = date(2026, 11, 14)        # a Saturday: last confirmation in the export
START = END - timedelta(days=60)


def workday_between(a, b):
    d = a + timedelta(days=random.randint(0, (b - a).days))
    while d.weekday() == 6:      # no Sundays
        d -= timedelta(days=1)
    return d


ROUTES = {
    "1100": [("WLD01", "Weld baseplate frame", 1.5, 3.0), ("MCH11", "Turn impeller hub", 0.8, 0.6),
             ("MCH19", "Bore bearing housing", 1.0, 0.9), ("MCH22", "CNC mill volute casing", 1.2, 1.8),
             ("PNT01", "Prime and paint", 0.5, 0.4), ("ASM01", "Assemble rotating unit", 0.4, 2.0),
             ("QAL01", "Hydrostatic test", 0.3, 0.8)],
    "1200": [("MCH14", "Turn shaft", 0.5, 0.25), ("MCH19", "Bore bearing housing", 1.0, 0.3),
             ("MCH28", "Drill casing flanges", 0.4, 0.2), ("PNT02", "Spray paint", 0.3, 0.15),
             ("ASM03", "Assemble and mount", 0.3, 0.6), ("QAL01", "Pressure test", 0.2, 0.2)],
    "1300": [("ASM05", "Strip and inspect", 0.5, 2.5), ("MCH11", "Machine worn parts", 0.6, 1.5),
             ("ASM05", "Reassemble", 0.3, 2.0), ("QAL01", "Test run", 0.3, 1.0)],
}
SERVICE_MATS = ["SV_REPAIR", "SV_REPAIR", "SV_REWIND", "SV_CLAIM", "SV_COMM"]

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
heads = ["Order", "Activity", "Work Center", "Order Type", "Operation Short Text", "Control Key",
         "Confirmation", "Op. Quantity", "Unit", "Base Qty", "Std Value Unit", "Actual Start Date",
         "Setup (Std)", "Setup (Conf.)", "Unit", "Machine Std / Unit", "Machine (Conf.)", "Material",
         "Material Description", "Actual Finish Date", "System Status", "Plant", "MRP Controller"]
ws.append(heads)
for c in ws[1]:
    c.font = Font(bold=True)

order = 1006000
n_ops = 0
for _ in range(420):
    plant = random.choice(["1100", "1100", "1200", "1200", "1200", "1300"])
    order += random.randint(1, 9)
    qty = random.choice([1, 1, 2, 4]) if plant != "1200" else random.choice([4, 6, 10, 12, 20])
    otype = "PP03" if plant == "1300" else ("PP01" if plant == "1100" else "PP02")
    mat = random.choice(SERVICE_MATS) if plant == "1300" else random.choice(MATS[plant])
    finish0 = workday_between(START, END)
    route = list(ROUTES[plant])
    # goods issue line first, like a real export
    rows = [("GI-STAND", "Goods issue", 0, 0, 0, 0, None, None)]
    day = finish0 - timedelta(days=len(route) * 2)
    for wc, text, s_std, m_std in route:
        day = day + timedelta(days=random.randint(1, 3))
        while day.weekday() == 6:
            day += timedelta(days=1)
        s_conf = s_std * random.uniform(0.8, 1.25)
        m_conf = m_std * qty * random.uniform(0.85, 1.2)
        if wc == "MCH19":
            s_conf = s_std + random.uniform(2.0, 4.0)                 # no proper setup standard
        if wc == "MCH22" and day >= date(2026, 10, 15):
            m_conf = m_std * qty * random.uniform(1.6, 2.4)           # new CNC, training
        if wc == "ASM03":
            m_conf = m_std * qty * random.uniform(0.4, 0.7)           # favourable
        if wc == "WLD01" and random.random() < 0.4:
            m_conf = m_std * qty + random.uniform(2.5, 6)              # large frames
        if plant == "1300" and random.random() < 0.3:
            s_std = 0.0                                                # no plan on repairs
        done = day <= END
        if not done:
            s_conf = m_conf = 0
        start = day - timedelta(days=random.choice([0, 0, 1]))
        rows.append((wc, text, s_std, round(s_conf, 2) if done else 0, m_std, round(m_conf, 2) if done else 0,
                     start if done else None, day if done else None))
    rows.append(("GR01", "Goods receipt", 0, 0, 0, 0, None, None))

    for i, (wc, text, s_std, s_conf, m_std, m_conf, st, fin) in enumerate(rows):
        n_ops += 1
        ws.append([
            str(order) if i == 0 else None, f"{(i + 1) * 10:04d}", wc, otype, text, "PP01",
            50000000 + n_ops, qty, "PC", 1, "H", st, s_std, s_conf, "H", m_std, m_conf,
            mat if i == 0 else None, ("Repair job" if plant == "1300" else "Pump assembly") if i == 0 else None,
            fin, "CNF" if (s_conf + m_conf) > 0 else "REL", plant if i == 0 else None, "P10",
        ])
for row in ws.iter_rows(min_row=2):
    for c in (row[11], row[19]):
        c.number_format = "dd/mm/yyyy"
wb.save(HERE / "Order_Operations_Export.xlsx")
print(f"Order_Operations_Export.xlsx: 420 orders, {n_ops} operation rows")
