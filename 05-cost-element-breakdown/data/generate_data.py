"""
Fake "actual costs - expanded view" report for one sales order item (Acme Pumps),
as a tab-delimited text export, in the layout modCostElementBreakdown reads:

  boilerplate lines (Report group, Selections, Sales order, Item from / to, Controlling area,
  Period from, Values in, Cost Elements header) - skipped
  detail lines : "<item id>  <description>" | total | fixed | variable | "qty unit"
  element line : "*  <cost element>  <description>" closes the block of detail lines above it
  grand total  : "**" line
Numbers are SAP text: space thousands, comma decimal, trailing minus for credits.

Built in: a settlement element with credits, activity lines with hours, and one cost
element whose reported total is R312,40 more than its visible lines (the macro should add
a "(reconciling adjustment)" line).
"""
from pathlib import Path

HERE = Path(__file__).parent


def sap(x):
    if x is None:
        return ""
    t = f"{abs(x):,.2f}".replace(",", " ").replace(".", ",")
    return t + "-" if x < 0 else t


elements = [
    ("400000", "Raw materials", [
        ("00047105", "Volute casing GG25 casting", 15281.67, 0, "3 PC"),
        ("00047300", "Bar stock 1.4021 Ø60", 1953.60, 0, "6,6 KG"),
        ("00047313", "Casting CuSn10", 3556.80, 0, "19,5 KG"),
    ], 0),
    ("410000", "Purchased parts", [
        ("00047144", "Mechanical seal faces SiC", 22140.00, 0, "3 PC"),
        ("00047157", "Ball bearing 6312", 3864.00, 0, "6 PC"),
        ("00047170", "Electric motor 22 kW IE3", 118500.00, 0, "3 PC"),
        ("00047183", "Flexible coupling", 9870.00, 0, "3 PC"),
    ], 312.40),                                   # hidden lines: reported total is higher
    ("620100", "Labour - machining", [
        ("MCH14/LAB", "Turn shaft", 4110.00, 1644.00, "6 H"),
        ("MCH22/LAB", "Machine casing flanges", 9864.00, 3945.60, "14,4 H"),
        ("MCH28/LAB", "Grind bearing seats", 1644.00, 657.60, "2,4 H"),
    ], 0),
    ("620200", "Labour - assembly & welding", [
        ("ASM01/LAB", "Assemble rotating unit", 4860.00, 1944.00, "9 H"),
        ("WLD01/LAB", "Weld baseplate frame", 13725.00, 5490.00, "22,5 H"),
        ("ASM03/LAB", "Final assembly and test", 9720.00, 3888.00, "18 H"),
    ], 0),
    ("630000", "Manufacturing overhead", [
        ("OH-LAB", "Overhead surcharge 30 % on labour", 13176.90, 13176.90, ""),
    ], 0),
    ("655000", "Settlement - production orders", [
        ("1006012", "Prod. order 1006012 (settled)", -61234.55, -18300.00, ""),
        ("1006019", "Prod. order 1006019 (settled)", -58900.10, -17150.00, ""),
    ], 0),
]

out = [
    "Report group 6O01\tActual costs - expanded view",
    "Selections",
    "Sales order\t30051877",
    "Item from\t10",
    "     to\t10",
    "Controlling area\t1000 Acme Pumps",
    "Period from\t1.2026 to 10.2026",
    "Values in\tZAR",
    "",
    "Cost Elements\tTotal\tFixed\tVariable\tOutput Quantities",
]
grand_t = grand_f = 0.0
for code, desc, lines, hidden in elements:
    tot_t = tot_f = 0.0
    for item, idesc, total, fixed, qty in lines:
        var = round(total - fixed, 2)
        label = f"{item}  {idesc}" if item else f"  {idesc}"
        out.append("\t".join([label, sap(total), sap(fixed) if fixed else "", sap(var), qty]))
        tot_t += total
        tot_f += fixed
    tot_t = round(tot_t + hidden, 2)
    tot_v = round(tot_t - tot_f, 2)
    out.append("\t".join([f"*  {code}  {desc}", sap(tot_t), sap(tot_f) if tot_f else "", sap(tot_v), ""]))
    grand_t += tot_t
    grand_f += tot_f
grand_t = round(grand_t, 2)
out.append("\t".join(["**", sap(grand_t), sap(round(grand_f, 2)), sap(round(grand_t - grand_f, 2)), ""]))
(HERE / "Actual_Costs_SO30051877.txt").write_bytes(("\r\n".join(out) + "\r\n").encode("cp1252"))
print(f"Actual_Costs_SO30051877.txt: {len(elements)} cost elements, grand total R{grand_t:,.2f}")
