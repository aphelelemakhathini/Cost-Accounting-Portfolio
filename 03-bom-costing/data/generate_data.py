"""
Fake CK11N costed-BOM text exports (cost estimate with quantity structure) for the
fictional company Acme Pumps, in the layout modCostedBomBuilder reads:

  * header block: Material / Plant / Costing Variant / ... label-value lines
  * a "Costing Structure" header line
  * each line: the description sits in the column for its depth, the description area
    is padded with tabs to a fixed width, then the SAME fixed fields follow:
      Material, Price, Quantity, Total value, Manuf.material, Traded material,
      Manuf.Costs total, Mat. overhead, Manu overhead, Variable value, Fixed Value
    plus Work Ctr and ActTyp on activity (labour) lines
  * numbers as SAP text: space thousands, comma decimal ("1 234,56")
  * Windows-1252 encoding

Every parent line equals the sum of its children and every leaf is Price x Qty,
so the macro should report that every rollup matches.
"""
import random
from pathlib import Path

random.seed(11)
HERE = Path(__file__).parent
DESC_W = 5                                   # description columns (max depth 4)


def sap(x, dec=2):
    neg = x < 0
    t = f"{abs(x):,.{dec}f}".replace(",", " ").replace(".", ",")
    return t + "-" if neg else t


class Node:
    def __init__(self, desc, material="", kind="asm", qty=1.0, price=None, wc="", act=""):
        self.desc, self.material, self.kind, self.qty = desc, material, kind, qty
        self.price, self.wc, self.act = price, wc, act
        self.children = []
        self.vals = {}

    def add(self, *kids):
        self.children += kids
        return self


def leaf_values(n):
    """cost split for a leaf; components always add up exactly to the total"""
    total = round(n.price * n.qty, 2)
    v = dict.fromkeys(["total", "manuf", "traded", "mfgcost", "matoh", "manuoh", "var", "fix"], 0.0)
    v["total"] = total
    if n.kind == "buy":                      # bought-in part: traded material + 5 % material overhead
        v["traded"] = round(total / 1.05, 2)
        v["matoh"] = round(total - v["traded"], 2)
        v["var"] = total
    elif n.kind == "raw":                    # raw material for in-house parts
        v["manuf"] = round(total / 1.05, 2)
        v["matoh"] = round(total - v["manuf"], 2)
        v["var"] = total
    else:                                    # activity: labour + 30 % manufacturing overhead
        v["mfgcost"] = round(total / 1.3, 2)
        v["manuoh"] = round(total - v["mfgcost"], 2)
        v["var"] = round(total * 0.6, 2)
        v["fix"] = round(total - v["var"], 2)
    return v


def rollup(n):
    if not n.children:
        n.vals = leaf_values(n)
    else:
        for c in n.children:
            rollup(c)
        n.vals = {k: round(sum(c.vals[k] for c in n.children), 2) for k in n.children[0].vals}
        n.price = n.vals["total"]            # parent price shown per unit of 1


mat_no = iter(range(47001, 49999, 13))


def m():
    return f"{next(mat_no):08d}"


def buy(desc, qty, lo, hi):
    return Node(desc, m(), "buy", qty, round(random.uniform(lo, hi), 2))


def op(desc, wc, hours):
    return Node(desc, "", "act", round(hours, 3), {"MCH": 685.0, "ASM": 540.0, "WLD": 610.0, "PNT": 495.0}[wc[:3]],
                wc, "LAB")


def pump(model, size):
    impeller = Node(f"Impeller {model} bronze", m(), "asm").add(
        Node("Casting CuSn10, 120 °C rated", m(), "raw", round(6.5 * size, 3), 182.40),
        op("Turn and balance impeller", "MCH11", 1.2 * size))
    shaft = Node(f"Shaft dia {35 + 10 * size} 1.4021", m(), "asm").add(
        Node("Bar stock 1.4021 dia 60", m(), "raw", round(1.1 * size, 3), 296.00),
        op("Turn shaft", "MCH14", 0.9 * size), op("Grind bearing seats", "MCH28", 0.4 * size))
    rotating = Node("Rotating unit", m(), "asm").add(
        impeller, shaft, buy("Impeller nut M24", 1, 60, 140), buy("Key 8x7x40", 1, 15, 40),
        op("Assemble rotating unit", "ASM01", 1.0 * size))
    casing = Node("Casing assembly", m(), "asm").add(
        buy("Volute casing GG25 (casting)", 1, 3800 * size, 5200 * size), buy("Casing gasket", 1, 60, 180),
        buy("Casing stud M16", 8, 22, 38), op("Machine casing flanges", "MCH22", 1.6 * size))
    bearing = Node("Bearing bracket", m(), "asm").add(
        buy("Bearing housing GG25", 1, 2400, 3900), buy(f"Ball bearing 63{10 + size}", 2, 380, 720),
        buy("Lip seal", 2, 45, 95), op("Assemble bearing bracket", "ASM01", 0.7))
    seal = Node("Mechanical seal set", m(), "asm").add(
        buy("Seal faces SiC/carbon", 1, 5200, 9800), buy("O-ring Viton", 2, 25, 70))
    base = Node("Baseplate & coupling", m(), "asm").add(
        buy("Steel channel 100x50", round(3.2 * size, 3), 165, 210), op("Weld baseplate frame", "WLD01", 2.5 * size),
        op("Prime and paint", "PNT01", 0.8 * size), buy("Flexible coupling", 1, 2100, 4600))
    motor = buy(f"Electric motor {7.5 * size:g} kW 4-pole IE3", 1, 14000 * size, 19000 * size)
    root = Node(f"PUMP ASSEMBLY {model}", m(), "asm").add(
        casing, rotating, bearing, seal, base, motor, op("Final assembly and test", "ASM03", 2.0 * size))
    rollup(root)
    return root


def lines_of(n, depth=0):
    yield n, depth
    for c in n.children:
        yield from lines_of(c, depth + 1)


for model, size in (("AP-80", 1), ("MS-4", 2), ("SL-150", 3)):
    root = pump(model, size)
    out = [
        "Material\t" + root.material + " " + f"{model} complete pump",
        "Plant\t1200 Acme Pumps Standard",
        "Costing Variant\tPPC1 Standard cost est. (mat.)",
        "Costing Version\t1",
        "Costing Date from-to\t01.10.2026\t31.12.2026",
        "Lot Size\t1 PC",
        "Cost Comp. Structure\t01 Acme cost component structure",
        "Created By\tCOSTING01",
        "Entered On\t05.10.2026",
        "",
        "Costing Structure" + "\t" * DESC_W + "\t".join(
            ["Material", "Price", "Quantity", "Total value", "Manuf.material", "Traded material",
             "Manuf.Costs total", "Mat. overhead", "Manu overhead", "Variable value", "Fixed Value",
             "Work Ctr", "ActTyp"]),
    ]
    for n, d in lines_of(root):
        desc = [""] * DESC_W
        desc[d] = n.desc
        v = n.vals
        fields = [n.material, sap(n.price), sap(n.qty, 3), sap(v["total"]), sap(v["manuf"]), sap(v["traded"]),
                  sap(v["mfgcost"]), sap(v["matoh"]), sap(v["manuoh"]), sap(v["var"]), sap(v["fix"])]
        if n.kind == "act":
            fields += [n.wc, n.act]
        out.append("\t".join(desc + fields))
    path = HERE / f"CK11N_{model}.txt"
    path.write_bytes(("\r\n".join(out) + "\r\n").encode("cp1252"))
    print(f"{path.name}: {len(out) - 11} lines, total R{root.vals['total']:,.2f}")
