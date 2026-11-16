"""ตรวจว่าตัวเลขสำคัญตรงกันทุกไฟล์: CSV, JSON สรุป, Excel, HTML แดชบอร์ด, เอกสาร

รัน:  .venv/bin/python check_numbers.py
ออกด้วยรหัส 1 ถ้ามีตัวเลขไม่ตรง
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
fails = []


def check(name: str, got, want, tol: float = 0.005) -> None:
    ok = abs(float(got) - float(want)) <= tol
    print(f"{'OK  ' if ok else 'FAIL'} {name}: {got} (คาดว่า {want})")
    if not ok:
        fails.append(name)


# 1) ค่าอ้างอิง: คำนวณใหม่จากไฟล์ CSV
s = pd.read_csv(ROOT / "sales_clean.csv", dtype={"Invoice": str})
r = pd.read_csv(ROOT / "data" / "clean" / "returns.csv", dtype={"Invoice": str})
e = pd.read_csv(ROOT / "data" / "clean" / "excluded.csv", dtype={"Invoice": str})
raw_rows = json.loads((ROOT / "reports" / "clean_summary.json").read_text(encoding="utf-8"))["raw_rows"]
sales = round(float(s["Sales_GBP"].sum()), 2)
orders = int(s["Invoice"].nunique())
returns = round(float(-r["Return_GBP"].sum()), 2)
units = int(s["Quantity"].sum())
ref = {"sales": sales, "orders": orders, "aov": round(sales / orders, 2), "returns": returns,
       "rate": round(returns / sales * 100, 2), "net": round(sales - returns, 2), "units": units}
print("ค่าอ้างอิงจาก CSV:", ref)
check("จำนวนแถวครบ (ยอดขาย+คืน+คัดออก = ต้นฉบับ)", len(s) + len(r) + len(e), raw_rows, 0)
check("ไม่มีแถวซ้ำในยอดขาย", s.duplicated().sum(), 0, 0)

# 2) JSON สรุป
k = json.loads((ROOT / "reports" / "clean_summary.json").read_text(encoding="utf-8"))["kpi"]
check("clean_summary: ยอดขาย", k["total_sales_gbp"], ref["sales"])
check("clean_summary: ออร์เดอร์", k["orders"], ref["orders"], 0)
check("clean_summary: ยอดคืน", -k["returns_gbp"], ref["returns"])
check("clean_summary: สุทธิ", k["net_sales_gbp"], ref["net"])

# 3) แดชบอร์ด HTML (ตัวเลขที่ฝังในไฟล์)
html = (ROOT / "dashboard.html").read_text(encoding="utf-8")
d = json.loads(re.search(r"const DATA = (\{.*?\});\n", html, re.S).group(1).replace("<\\/", "</"))["kpi"]
check("dashboard: ยอดขาย", d["sales"], ref["sales"])
check("dashboard: ออร์เดอร์", d["orders"], ref["orders"], 0)
check("dashboard: เฉลี่ย", d["aov"], ref["aov"])
check("dashboard: ยอดคืน", d["returns"], ref["returns"])
check("dashboard: อัตราคืน %", d["return_rate"], ref["rate"])
check("dashboard: จำนวนชิ้น", d["units"], ref["units"], 0)

# 4) Excel (ค่าที่เก็บไว้ในไฟล์ ซึ่งตรวจสูตรซ้ำใน Excel จริงมาแล้ว)
ws = load_workbook(ROOT / "sales_report.xlsx", data_only=True, read_only=True)["สรุป"]
cells = {c.coordinate: c.value for row in ws.iter_rows(min_row=6, max_row=6, max_col=7) for c in row}
check("excel: ยอดขาย A6", cells["A6"], ref["sales"])
check("excel: ออร์เดอร์ B6", cells["B6"], ref["orders"], 0)
check("excel: เฉลี่ย C6", cells["C6"], ref["aov"])
check("excel: ยอดคืน D6", cells["D6"], ref["returns"])
check("excel: อัตราคืน E6 %", cells["E6"] * 100, ref["rate"])
check("excel: สุทธิ F6", cells["F6"], ref["net"])
check("excel: ชิ้น G6", cells["G6"], ref["units"], 0)

# 5) ไฟล์สรุปสำหรับ Looker Studio
ks = pd.read_csv(ROOT / "data" / "clean" / "kpi_summary.csv")
check("kpi_summary: ยอดขาย", ks["Sales_GBP"].sum(), ref["sales"], 0.01)
check("kpi_summary: ยอดคืน", ks["Returns_GBP"].sum(), ref["returns"], 0.01)
check("kpi_summary: สุทธิ", ks["Net_GBP"].sum(), ref["net"], 0.01)

# 6) เอกสาร: ตัวเลขที่เขียนไว้ต้องตรงกับค่าจริง
text = (ROOT / "README.md").read_text(encoding="utf-8")
if (ROOT / "portfolio_th.md").exists():  # ไฟล์นี้เก็บไว้ในเครื่องเท่านั้น ไม่ขึ้น GitHub
    text += (ROOT / "portfolio_th.md").read_text(encoding="utf-8")
want = [f"£{ref['sales']:,.2f}", f"{ref['orders']:,}", f"£{ref['aov']:,.2f}", f"£{ref['returns']:,.2f}",
        f"£{ref['net']:,.2f}", f"{ref['units']:,}", f"{raw_rows:,}", f"{len(s):,}", f"{len(r):,}", f"{len(e):,}",
        f"{ref['rate']:.2f}%"]
for w in want:
    ok = w in text
    print(f"{'OK  ' if ok else 'FAIL'} เอกสารมีตัวเลข {w}")
    if not ok:
        fails.append("doc " + w)

print("\nสรุป:", "ตัวเลขตรงกันทุกไฟล์" if not fails else f"ไม่ตรง {len(fails)} รายการ: {fails}")
sys.exit(1 if fails else 0)
