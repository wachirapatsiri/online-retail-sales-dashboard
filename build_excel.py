"""สร้างไฟล์ Excel สรุปยอดสำหรับเจ้าของร้าน -> sales_report.xlsx

รัน (หลัง clean.py):  .venv/bin/python build_excel.py

- ตัวเลขสรุปทั้งหมดเป็น "สูตร Excel จริง" (SUM, SUMIFS, INDEX/MATCH, RANK)
  อ้างทั้งคอลัมน์ของชีต "ข้อมูล" วางข้อมูลเดือนใหม่ต่อท้ายแล้วตัวเลขอัปเดตเอง
- ค่าที่แสดงตอนเปิดไฟล์ (cached value) คำนวณจาก pandas ชุดเดียวกัน
- หน่วยเงิน GBP (£)
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "sales_report.xlsx"
MONEY = "[$£-809]#,##0.00"  # รูปแบบเงินปอนด์
FONT = "Tahoma"  # มีทั้งบน Windows/Mac และรองรับภาษาไทย
BLUE, BLUE_DARK, BLUE_LIGHT, GREY = "#2563EB", "#1E3A8A", "#EFF6FF", "#64748B"

S_DATA, S_RET, S_PROD, S_CTRY, S_SUM, S_QA = "ข้อมูล", "คืนสินค้า", "สินค้า", "ประเทศ", "สรุป", "รายงานคุณภาพข้อมูล"
EPOCH = pd.Timestamp("1899-12-30")
MONTH_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


def serial(s: pd.Series) -> np.ndarray:
    """แปลงวันที่เป็นเลขวันที่ของ Excel"""
    return ((s - EPOCH) / pd.Timedelta(days=1)).to_numpy()


def load(path: Path, amount_col: str) -> pd.DataFrame:
    df = pd.read_csv(path, dtype={"Invoice": str, "StockCode": str, "CustomerID": str})
    df = df.rename(columns={amount_col: "Amount"})
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    df["Date"] = pd.to_datetime(df["Date"])
    df["MonthStart"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    # เรียงตามเลขที่บิล เพื่อให้สูตรนับออร์เดอร์ (คอลัมน์ O) ทำงานได้
    return df.sort_values(["Invoice"], kind="stable").reset_index(drop=True)


def write_data_sheet(wb, ws, df, f, amount_header, with_flag):
    headers = ["Invoice", "InvoiceDate", "Date", "Month", "WeekdayNum", "Weekday", "Hour",
               "StockCode", "Description", "Quantity", "Price", amount_header, "CustomerID", "Country"]
    if with_flag:
        headers.append("OrderFlag")
    widths = [10, 18, 12, 10, 12, 11, 7, 11, 40, 10, 9, 13, 12, 18, 10]
    fmts = [None, f["dt"], f["date"], f["month"], None, None, None, None, None, f["int"], f["money"], f["money"], None, None, None]
    for c, h in enumerate(headers):
        ws.set_column(c, c, widths[c], fmts[c])
        ws.write(0, c, h, f["head"])
    n = len(df)
    ws.write_column(1, 0, df["Invoice"].tolist())
    ws.write_column(1, 1, serial(df["InvoiceDate"]).tolist())
    ws.write_column(1, 2, serial(df["Date"]).tolist())
    ws.write_column(1, 3, serial(df["MonthStart"]).tolist())
    ws.write_column(1, 4, df["WeekdayNum"].tolist())
    ws.write_column(1, 5, df["Weekday"].tolist())
    ws.write_column(1, 6, df["Hour"].tolist())
    for r, v in enumerate(df["StockCode"].tolist(), start=1):
        ws.write_string(r, 7, v)  # บังคับเป็นข้อความ กันรหัสกลายเป็นตัวเลข
    ws.write_column(1, 8, df["Description"].tolist())
    ws.write_column(1, 9, df["Quantity"].tolist())
    ws.write_column(1, 10, df["Price"].tolist())
    ws.write_column(1, 11, df["Amount"].tolist())
    for r, v in enumerate(df["CustomerID"].tolist(), start=1):
        ws.write_string(r, 12, v)
    ws.write_column(1, 13, df["Country"].tolist())
    if with_flag:
        flags = df["OrderFlag"].tolist()
        for r in range(1, n + 1):
            ws.write_formula(r, 14, f"=IF(A{r + 1}<>A{r},1,0)", None, flags[r - 1])
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, n, len(headers) - 1)


def main() -> None:
    sales = load(ROOT / "sales_clean.csv", "Sales_GBP")
    rets = load(ROOT / "data" / "clean" / "returns.csv", "Return_GBP")
    audit = json.loads((ROOT / "reports" / "audit_summary.json").read_text(encoding="utf-8"))
    clean = json.loads((ROOT / "reports" / "clean_summary.json").read_text(encoding="utf-8"))
    sales["OrderFlag"] = (sales["Invoice"] != sales["Invoice"].shift()).astype(int)
    assert sales["OrderFlag"].sum() == sales["Invoice"].nunique()

    # ---------- ค่าที่ใช้เป็น cached value ของสูตร ----------
    total_sales, orders, units = sales["Amount"].sum(), int(sales["OrderFlag"].sum()), int(sales["Quantity"].sum())
    ret_total = -rets["Amount"].sum()
    months = sorted(sales["MonthStart"].unique())
    m_sales = sales.groupby("MonthStart")["Amount"].sum()
    m_orders = sales.groupby("MonthStart")["OrderFlag"].sum()
    m_ret = (-rets.groupby("MonthStart")["Amount"].sum()).reindex(months, fill_value=0.0)

    prod = sales.groupby("StockCode").agg(name=("Description", "first"), sales=("Amount", "sum"),
                                          qty=("Quantity", "sum")).reset_index()
    order = np.argsort(-prod["sales"].to_numpy(), kind="stable")
    prod["rank"] = 0
    prod.loc[order, "rank"] = np.arange(1, len(prod) + 1)
    ctry = sales.groupby("Country").agg(sales=("Amount", "sum"), orders=("OrderFlag", "sum")).reset_index()
    order = np.argsort(-ctry["sales"].to_numpy(), kind="stable")
    ctry["rank"] = 0
    ctry.loc[order, "rank"] = np.arange(1, len(ctry) + 1)

    wb = xlsxwriter.Workbook(OUT)
    base = {"font_name": FONT, "font_size": 10}
    wb.formats[0].set_font_name(FONT)
    wb.formats[0].set_font_size(10)
    F = lambda **k: wb.add_format({**base, **k})  # noqa: E731
    f = {
        "head": F(bold=True, font_color="white", bg_color=BLUE, border=1, border_color="#BFDBFE", align="center", valign="vcenter", text_wrap=True),
        "money": F(num_format=MONEY), "int": F(num_format="#,##0"), "int_c": F(num_format="#,##0", align="center"), "pct": F(num_format="0.00%"),
        "date": F(num_format="[$-409]yyyy-mm-dd"), "dt": F(num_format="[$-409]yyyy-mm-dd hh:mm"), "month": F(num_format="[$-409]yyyy-mm"), "month_c": F(num_format="[$-409]yyyy-mm", align="center"),
        "title": F(bold=True, font_size=18, font_color=BLUE_DARK), "sub": F(font_color=GREY),
        "h2": F(bold=True, font_size=12, font_color=BLUE_DARK),
        "kpi_l": F(font_color=GREY, bg_color=BLUE_LIGHT, top=2, top_color=BLUE, align="center", text_wrap=True, valign="vcenter"),
        "kpi_money": F(bold=True, font_size=14, font_color=BLUE_DARK, bg_color=BLUE_LIGHT, num_format=MONEY, align="center"),
        "kpi_int": F(bold=True, font_size=14, font_color=BLUE_DARK, bg_color=BLUE_LIGHT, num_format="#,##0", align="center"),
        "kpi_pct": F(bold=True, font_size=14, font_color=BLUE_DARK, bg_color=BLUE_LIGHT, num_format="0.00%", align="center"),
        "tot": F(bold=True, top=1, bg_color=BLUE_LIGHT), "tot_money": F(bold=True, top=1, bg_color=BLUE_LIGHT, num_format=MONEY),
        "tot_int": F(bold=True, top=1, bg_color=BLUE_LIGHT, num_format="#,##0", align="center"),
        "wrap": F(text_wrap=True, valign="top"), "wrap_b": F(text_wrap=True, valign="top", bold=True),
        "int_top": F(num_format="#,##0", valign="top"), "center": F(align="center"),
        "date_l": F(num_format="[$-409]d mmm yyyy", align="left", bold=True),
    }

    ws_sum = wb.add_worksheet(S_SUM)
    ws_data = wb.add_worksheet(S_DATA)
    ws_ret = wb.add_worksheet(S_RET)
    ws_prod = wb.add_worksheet(S_PROD)
    ws_ctry = wb.add_worksheet(S_CTRY)
    ws_qa = wb.add_worksheet(S_QA)
    for ws, color in [(ws_sum, BLUE_DARK), (ws_data, BLUE), (ws_ret, "#F59E0B"), (ws_prod, "#93C5FD"), (ws_ctry, "#93C5FD"), (ws_qa, "#10B981")]:
        ws.set_tab_color(color)

    write_data_sheet(wb, ws_data, sales, f, "Sales_GBP", with_flag=True)
    write_data_sheet(wb, ws_ret, rets, f, "Return_GBP", with_flag=False)

    D = f"'{S_DATA}'!"
    R = f"'{S_RET}'!"

    # ---------- ชีต สินค้า (ตัวช่วยจัดอันดับ) ----------
    for c, (h, w) in enumerate([("รหัสสินค้า", 12), ("ชื่อสินค้า", 42), ("ยอดขาย (£)", 16), ("อันดับ", 9)]):
        ws_prod.set_column(c, c, w)
        ws_prod.write(0, c, h, f["head"])
    for i, p in enumerate(prod.itertuples(index=False), start=2):
        ws_prod.write_string(i - 1, 0, p.StockCode)
        ws_prod.write_string(i - 1, 1, p.name)
        ws_prod.write_formula(i - 1, 2, f"=SUMIFS({D}L:L,{D}H:H,A{i})", f["money"], float(p.sales))
        ws_prod.write_formula(i - 1, 3, f"=RANK(C{i},C:C)+COUNTIF(C$2:C{i},C{i})-1", f["int"], int(p.rank))
    ws_prod.freeze_panes(1, 0)

    # ---------- ชีต ประเทศ ----------
    for c, (h, w) in enumerate([("ประเทศ", 22), ("ยอดขาย (£)", 16), ("ออร์เดอร์", 11), ("อันดับ", 9)]):
        ws_ctry.set_column(c, c, w)
        ws_ctry.write(0, c, h, f["head"])
    for i, p in enumerate(ctry.itertuples(index=False), start=2):
        ws_ctry.write_string(i - 1, 0, p.Country)
        ws_ctry.write_formula(i - 1, 1, f"=SUMIFS({D}L:L,{D}N:N,A{i})", f["money"], float(p.sales))
        ws_ctry.write_formula(i - 1, 2, f"=SUMIFS({D}O:O,{D}N:N,A{i})", f["int"], int(p.orders))
        ws_ctry.write_formula(i - 1, 3, f"=RANK(B{i},B:B)+COUNTIF(B$2:B{i},B{i})-1", f["int"], int(p.rank))
    ws_ctry.freeze_panes(1, 0)

    # ---------- ชีต สรุป ----------
    ws = ws_sum
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_column("A:A", 25)
    ws.set_column("B:B", 20)
    ws.set_column("C:C", 40)
    ws.set_column("D:G", 25)
    ws.set_column("H:H", 3)

    # โทนสี: น้ำเงิน = ยอดขาย, ส้ม = คืนสินค้า, เขียว = ยอดสุทธิ
    NAVY, ZEBRA, LINE = BLUE_DARK, "#F8FAFC", "#E2E8F0"
    TONE = {
        "sale": {"main": BLUE, "tint": "#EFF6FF", "text": NAVY, "bar": "#93C5FD"},
        "ret": {"main": "#EA580C", "tint": "#FFF7ED", "text": "#C2410C", "bar": "#FDBA74"},
        "net": {"main": "#059669", "tint": "#ECFDF5", "text": "#047857", "bar": "#6EE7B7"},
    }
    NUMFMT = {"money": MONEY, "int": "#,##0", "int_c": "#,##0", "pct": "0.00%", "month_c": "[$-409]yyyy-mm"}
    _cache = {}

    def cell(kind, z=False, tone="sale", bold=False):
        """รูปแบบเซลล์ในตาราง: สลับสีแถว + เส้นคั่นบาง + สีตัวอักษรตามโทน"""
        key = (kind, z, tone, bold)
        if key not in _cache:
            p = {"bottom": 1, "bottom_color": LINE, "valign": "vcenter"}
            if z:
                p["bg_color"] = ZEBRA
            if kind in NUMFMT:
                p["num_format"] = NUMFMT[kind]
            if kind in ("int_c", "month_c", "center"):
                p["align"] = "center"
            if tone != "sale":
                p["font_color"] = TONE[tone]["text"]
            if bold:
                p["bold"] = True
            if kind == "rank":
                p.update(align="center", bold=True, font_color=NAVY)
            if kind == "note":
                p["font_color"] = GREY
            _cache[key] = F(**p)
        return _cache[key]

    def head(tone="sale"):
        return F(bold=True, font_color="white", bg_color=TONE[tone]["main"], align="center", valign="vcenter",
                 text_wrap=True, border=1, border_color="white")

    def total(kind, tone="sale"):
        p = {"bold": True, "top": 2, "top_color": NAVY, "bg_color": "#DBEAFE", "valign": "vcenter",
             "font_color": TONE[tone]["text"]}
        if kind in NUMFMT:
            p["num_format"] = NUMFMT[kind]
        if kind == "int_c":
            p["align"] = "center"
        return F(**p)

    f_sec = F(bold=True, font_size=12, font_color=NAVY, bg_color="#DBEAFE", left=5, left_color=NAVY, valign="vcenter", indent=1)

    def section(row, text):
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, 6, text, f_sec)

    # แถบหัวรายงาน
    ws.set_row(0, 40)
    ws.set_row(1, 22)
    ws.merge_range("A1:G1", "รายงานสรุปยอดขาย", F(bold=True, font_size=20, font_color="white", bg_color=NAVY, valign="vcenter", indent=1))
    ws.merge_range("A2:G2", "หน่วยเงินทั้งหมด: ปอนด์สเตอร์ลิง (GBP, £)  ·  ตัวเลขทุกช่องเป็นสูตร คำนวณจากชีต \"ข้อมูล\" และ \"คืนสินค้า\"",
                   F(font_color="#DBEAFE", bg_color=BLUE, valign="vcenter", indent=1))
    ws.write("A3", "ช่วงข้อมูล", F(font_color=GREY, indent=1, valign="vcenter"))
    ws.write_formula("B3", f"=MIN({D}C:C)", f["date_l"], float(serial(pd.Series([sales['Date'].min()]))[0]))
    ws.write("C3", "ถึง", F(font_color=GREY, align="right"))
    ws.write_formula("D3", f"=MAX({D}C:C)", f["date_l"], float(serial(pd.Series([sales['Date'].max()]))[0]))
    ws.write("E3", "สี:", F(font_color=GREY, align="right"))
    ws.write("F3", "■ ยอดขาย   ■ คืนสินค้า   ■ ยอดสุทธิ", F(font_color=GREY))
    ws.write_rich_string("F3", F(font_color=BLUE, bold=True), "■ ยอดขาย   ", F(font_color=TONE["ret"]["main"], bold=True), "■ คืนสินค้า   ",
                         F(font_color=TONE["net"]["main"], bold=True), "■ ยอดสุทธิ")

    # การ์ดตัวเลขสำคัญ
    ws.set_row(4, 34)
    ws.set_row(5, 34)
    ws.set_row(6, 6)
    kpis = [
        ("ยอดขายรวม (£)\nก่อนหักคืนสินค้า", f"=SUM({D}L:L)", "money", total_sales, "sale"),
        ("จำนวนออร์เดอร์", f"=SUM({D}O:O)", "int", orders, "sale"),
        ("ยอดเฉลี่ยต่อออร์เดอร์ (£)", "=A6/B6", "money", total_sales / orders, "sale"),
        ("ยอดคืนสินค้า (£)", f"=-SUM({R}L:L)", "money", ret_total, "ret"),
        ("อัตราคืนสินค้า\n(ยอดคืน ÷ ยอดขาย)", "=D6/A6", "pct", ret_total / total_sales, "ret"),
        ("ยอดขายสุทธิ (£)\nหลังหักคืนสินค้า", "=A6-D6", "money", total_sales - ret_total, "net"),
        ("จำนวนสินค้าที่ขาย (ชิ้น)", f"=SUM({D}J:J)", "int", units, "sale"),
    ]
    for c, (label, formula, kind, val, tone) in enumerate(kpis):
        t = TONE[tone]
        side = {"left": 5, "left_color": "white", "right": 5, "right_color": "white"}  # ช่องไฟระหว่างการ์ด
        ws.write(4, c, label, F(font_color=GREY, bg_color=t["tint"], top=5, top_color=t["main"], align="center",
                                valign="vcenter", text_wrap=True, **side))
        ws.write_formula(5, c, formula, F(bold=True, font_size=14, font_color=t["text"], bg_color=t["tint"], align="center",
                                          valign="vcenter", num_format=NUMFMT[kind], **side), float(val))

    # ยอดขายรายเดือน
    r0 = 8  # แถวหัวตาราง (index)
    section(r0 - 1, "ยอดขายรายเดือน")
    ws.set_row(r0, 30)
    for c, (h, tone) in enumerate([("เดือน", "sale"), ("ยอดขาย (£)", "sale"), ("ออร์เดอร์", "sale"), ("ยอดเฉลี่ย/ออร์เดอร์ (£)", "sale"),
                                   ("ยอดคืนสินค้า (£)", "ret"), ("ยอดสุทธิ (£)", "net"), ("หมายเหตุ", "sale")]):
        ws.write(r0, c, h, head(tone))
    last_day = sales["Date"].max()
    m_first, m_last = r0 + 2, r0 + 1 + len(months)  # เลขแถว Excel
    best_month = m_sales.idxmax()
    for i, m in enumerate(months):
        r = r0 + 1 + i
        x = r + 1
        z = i % 2 == 1
        m = pd.Timestamp(m)
        ws.write_number(r, 0, float(serial(pd.Series([m]))[0]), cell("month_c", z))
        ws.write_formula(r, 1, f"=SUMIFS({D}L:L,{D}D:D,A{x})", cell("money", z), float(m_sales[m]))
        ws.write_formula(r, 2, f"=SUMIFS({D}O:O,{D}D:D,A{x})", cell("int_c", z), int(m_orders[m]))
        ws.write_formula(r, 3, f"=IF(C{x}=0,0,B{x}/C{x})", cell("money", z), float(m_sales[m] / m_orders[m]))
        ws.write_formula(r, 4, f"=-SUMIFS({R}L:L,{R}D:D,A{x})", cell("money", z, "ret"), float(m_ret[m]))
        ws.write_formula(r, 5, f"=B{x}-E{x}", cell("money", z, "net"), float(m_sales[m] - m_ret[m]))
        if m == last_day.to_period("M").to_timestamp() and last_day.day < last_day.days_in_month:
            ws.write(r, 6, f"ไม่ครบเดือน (ถึง {last_day.day} {MONTH_TH[last_day.month - 1]})", cell("note", z))
        else:
            ws.write_formula(r, 6, f'=IF(B{x}=MAX($B${m_first}:$B${m_last}),"★ เดือนที่ขายดีที่สุด","")', cell("note", z, bold=True),
                             "★ เดือนที่ขายดีที่สุด" if m == best_month else "")
    rt = m_last  # index ของแถวรวม
    ws.write(rt, 0, "รวม", total("text"))
    ws.write_formula(rt, 1, f"=SUM(B{m_first}:B{m_last})", total("money"), float(total_sales))
    ws.write_formula(rt, 2, f"=SUM(C{m_first}:C{m_last})", total("int_c"), orders)
    ws.write_formula(rt, 3, f"=B{rt + 1}/C{rt + 1}", total("money"), float(total_sales / orders))
    ws.write_formula(rt, 4, f"=SUM(E{m_first}:E{m_last})", total("money", "ret"), float(ret_total))
    ws.write_formula(rt, 5, f"=SUM(F{m_first}:F{m_last})", total("money", "net"), float(total_sales - ret_total))
    ws.write(rt, 6, "", total("text"))
    ws.ignore_errors({"formula_differs": f"D{rt + 1}"})

    def bars(first, last, col, tone):
        """แถบสีในช่อง ให้เห็นขนาดตัวเลขทันที (first/last เป็น index ของแถว)"""
        ws.conditional_format(first, col, last, col, {"type": "data_bar", "bar_color": TONE[tone]["bar"], "bar_solid": True,
                                                      "data_bar_2010": True, "min_type": "num", "min_value": 0,
                                                      "bar_border_color": TONE[tone]["bar"]})

    bars(r0 + 1, r0 + len(months), 1, "sale")
    bars(r0 + 1, r0 + len(months), 4, "ret")
    # ไฮไลต์เดือนที่ขายดีที่สุดทั้งแถว (สูตร จึงอัปเดตเองเมื่อข้อมูลเปลี่ยน)
    ws.conditional_format(r0 + 1, 0, r0 + len(months), 3, {"type": "formula", "criteria": f"=$B{m_first}=MAX($B${m_first}:$B${m_last})",
                                                           "format": F(bold=True, font_color=NAVY)})

    # สินค้าขายดี 10 อันดับ
    p0 = rt + 3
    section(p0 - 1, "สินค้าขายดี 10 อันดับ (ตามยอดขาย)")
    ws.set_row(p0, 30)
    for c, (h, tone) in enumerate([("อันดับ", "sale"), ("รหัสสินค้า", "sale"), ("ชื่อสินค้า", "sale"), ("ยอดขาย (£)", "sale"),
                                   ("จำนวนที่ขาย (ชิ้น)", "sale"), ("% ของยอดขายรวม", "sale"), ("ยอดคืนสินค้า (£)", "ret")]):
        ws.write(p0, c, h, head(tone))
    ret_by_code = -rets.groupby("StockCode")["Amount"].sum()
    ws.write(p0 + 11, 0, "หมายเหตุ: อันดับเรียงตามยอดขายก่อนหักคืนสินค้า คอลัมน์ขวาสุดคือยอดที่ลูกค้าคืนของสินค้านั้น", f["sub"])
    top_p = prod.sort_values("rank").head(10)
    P = f"'{S_PROD}'!"
    for k, p in enumerate(top_p.itertuples(index=False), start=1):
        r = p0 + k
        x = r + 1
        z = k % 2 == 0
        ws.write_number(r, 0, k, cell("rank", z))
        ws.write_formula(r, 1, f"=INDEX({P}A:A,MATCH(A{x},{P}D:D,0))", cell("text", z), p.StockCode)
        ws.write_formula(r, 2, f"=INDEX({P}B:B,MATCH(A{x},{P}D:D,0))", cell("text", z), p.name)
        ws.write_formula(r, 3, f"=INDEX({P}C:C,MATCH(A{x},{P}D:D,0))", cell("money", z), float(p.sales))
        ws.write_formula(r, 4, f"=SUMIFS({D}J:J,{D}H:H,B{x})", cell("int", z), int(p.qty))
        ws.write_formula(r, 5, f"=D{x}/$A$6", cell("pct", z), float(p.sales / total_sales))
        ws.write_formula(r, 6, f"=-SUMIFS({R}L:L,{R}H:H,B{x})", cell("money", z, "ret"), float(ret_by_code.get(p.StockCode, 0.0)))
    bars(p0 + 1, p0 + 10, 3, "sale")

    # ประเทศที่ขายได้มากสุด
    c0 = p0 + 14
    section(c0 - 1, "ประเทศที่ขายได้มากสุด 10 อันดับ")
    ws.set_row(c0, 30)
    for c, h in enumerate(["อันดับ", "ประเทศ", "", "ยอดขาย (£)", "ออร์เดอร์", "% ของยอดขายรวม", ""]):
        ws.write(c0, c, h, head())
    ws.merge_range(c0, 1, c0, 2, "ประเทศ", head())
    top_c = ctry.sort_values("rank").head(10)
    C = f"'{S_CTRY}'!"
    for k, p in enumerate(top_c.itertuples(index=False), start=1):
        r = c0 + k
        x = r + 1
        z = k % 2 == 0
        ws.write_number(r, 0, k, cell("rank", z))
        ws.write_formula(r, 1, f"=INDEX({C}A:A,MATCH(A{x},{C}D:D,0))", cell("text", z), p.Country)
        ws.write(r, 2, "", cell("text", z))
        ws.write_formula(r, 3, f"=INDEX({C}B:B,MATCH(A{x},{C}D:D,0))", cell("money", z), float(p.sales))
        ws.write_formula(r, 4, f"=INDEX({C}C:C,MATCH(A{x},{C}D:D,0))", cell("int", z), int(p.orders))
        ws.write_formula(r, 5, f"=D{x}/$A$6", cell("pct", z), float(p.sales / total_sales))
        ws.write(r, 6, "", cell("text", z))
    bars(c0 + 1, c0 + 10, 5, "sale")
    ws.write(c0 + 12, 0, "ที่มาข้อมูล: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D (CC BY 4.0)", f["sub"])

    # กราฟ
    cf = {"name": FONT, "size": 9, "color": GREY}
    grid = {"visible": True, "line": {"color": LINE, "width": 0.75}}
    no_line = {"line": {"none": True}}

    def style(ch, title, w=640, h=320):
        ch.set_title({"name": title, "name_font": {"name": FONT, "size": 13, "bold": True, "color": NAVY}, "overlay": False})
        ch.set_legend({"none": True})
        ch.set_size({"width": w, "height": h})
        ch.set_chartarea({"border": {"color": LINE, "width": 1}, "fill": {"color": "white"}})
        ch.set_plotarea({"border": {"none": True}, "fill": {"color": "white"}})

    def labels(fmt):
        return {"value": True, "num_format": fmt, "font": {"name": FONT, "size": 8, "color": NAVY}}

    # 1) ยอดขายรายเดือน: เดือนขายดีสุด = น้ำเงินเข้ม, เดือนไม่ครบ = น้ำเงินอ่อน
    pts = []
    for m in months:
        m = pd.Timestamp(m)
        partial = m == last_day.to_period("M").to_timestamp() and last_day.day < last_day.days_in_month
        pts.append({"fill": {"color": NAVY if m == best_month else ("#93C5FD" if partial else BLUE)}})
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "ยอดขาย (£)", "categories": [S_SUM, r0 + 1, 0, r0 + len(months), 0],
                   "values": [S_SUM, r0 + 1, 1, r0 + len(months), 1], "fill": {"color": BLUE}, "gap": 45, "points": pts,
                   "data_labels": labels('[$£-809]0.00,,"M"')})
    style(ch, "ยอดขายรายเดือน (ล้าน £)  ·  เข้ม = ขายดีที่สุด, อ่อน = ไม่ครบเดือน")
    ch.set_x_axis({"num_font": cf, "num_format": "[$-409]yyyy-mm", "text_axis": True, "line": {"color": "#CBD5E1"}})
    ch.set_y_axis({"num_font": cf, "num_format": '[$£-809]0.0,,"M"', "major_gridlines": grid, **no_line})
    ws.insert_chart(r0 - 1, 8, ch)

    # 2) สินค้าขายดี: อันดับ 1 = น้ำเงินเข้ม
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "ยอดขาย (£)", "categories": [S_SUM, p0 + 1, 2, p0 + 10, 2],
                   "values": [S_SUM, p0 + 1, 3, p0 + 10, 3], "fill": {"color": BLUE}, "gap": 40,
                   "points": [{"fill": {"color": NAVY}}] + [{"fill": {"color": BLUE}}] * 9,
                   "data_labels": labels("[$£-809]#,##0")})
    style(ch, "สินค้าขายดี 10 อันดับ (£)", h=290)
    ch.set_y_axis({"reverse": True, "num_font": {"name": FONT, "size": 8, "color": "#334155"}, "line": {"color": "#CBD5E1"}})
    ch.set_x_axis({"num_font": cf, "num_format": '[$£-809]0,"K"', "major_gridlines": grid, **no_line})
    ws.insert_chart(p0 - 1, 8, ch)

    # 3) ประเทศ: อันดับ 1 = น้ำเงินเข้ม
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "ยอดขาย (£)", "categories": [S_SUM, c0 + 1, 1, c0 + 10, 1],
                   "values": [S_SUM, c0 + 1, 3, c0 + 10, 3], "fill": {"color": BLUE}, "gap": 40,
                   "points": [{"fill": {"color": NAVY}}] + [{"fill": {"color": BLUE}}] * 9,
                   "data_labels": labels('[$£-809]0.00,,"M"')})
    style(ch, "ยอดขายตามประเทศ 10 อันดับ (ล้าน £)", h=290)
    ch.set_y_axis({"reverse": True, "num_font": {"name": FONT, "size": 9, "color": "#334155"}, "line": {"color": "#CBD5E1"}})
    ch.set_x_axis({"num_font": cf, "num_format": '[$£-809]0,,"M"', "major_unit": 2000000, "max": 10000000, "major_gridlines": grid, **no_line})
    ws.insert_chart(c0 - 1, 8, ch)

    # 4) ยอดคืนสินค้ารายเดือน (โทนส้ม) วางใต้กราฟประเทศ
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "ยอดคืนสินค้า (£)", "categories": [S_SUM, r0 + 1, 0, r0 + len(months), 0],
                   "values": [S_SUM, r0 + 1, 4, r0 + len(months), 4], "fill": {"color": TONE["ret"]["main"]}, "gap": 45,
                   "data_labels": {"value": True, "num_format": '[$£-809]0,"K"', "font": {"name": FONT, "size": 8, "color": TONE["ret"]["text"]}}})
    style(ch, "ยอดคืนสินค้ารายเดือน (พัน £)", h=260)
    ch.set_x_axis({"num_font": cf, "num_format": "[$-409]yyyy-mm", "text_axis": True, "line": {"color": "#CBD5E1"}})
    ch.set_y_axis({"num_font": cf, "num_format": '[$£-809]0,"K"', "major_gridlines": grid, **no_line})
    ws.insert_chart(c0 + 14, 8, ch)

    ws.set_landscape()
    ws.fit_to_pages(1, 0)

    # ---------- ชีต รายงานคุณภาพข้อมูล ----------
    ws = ws_qa
    ws.hide_gridlines(2)
    ws.set_column("A:A", 38)
    ws.set_column("B:B", 14)
    ws.set_column("C:C", 62)
    ws.set_column("D:D", 44)
    a, k, fx, ex = audit["issues"], clean["kpi"], clean["fixes"], clean["excluded_by_reason"]
    ex_rows = {key.split(" (")[0]: v["rows"] for key, v in ex.items()}
    ws.set_zoom(90)
    ws.set_row(0, 40)
    ws.merge_range("A1:D1", "รายงานคุณภาพข้อมูล: ก่อน → หลังทำความสะอาด", F(bold=True, font_size=20, font_color="white", bg_color=NAVY, valign="vcenter", indent=1))

    def qa_section(row, text):
        ws.set_row(row, 24)
        ws.merge_range(row, 0, row, 3, text, f_sec)

    def qa(z=False, bold=False, num=False, bg=None, color=None):
        p = {"text_wrap": True, "valign": "top", "bottom": 1, "bottom_color": LINE}
        if bg or z:
            p["bg_color"] = bg or ZEBRA
        if bold:
            p["bold"] = True
        if color:
            p["font_color"] = color
        if num:
            p.update(num_format="#,##0", bold=True, font_color=NAVY)
        return F(**p)
    ws.write("A2", f"ไฟล์ต้นฉบับ {clean['raw_rows']:,} แถว (ชีต {audit['source']['sheet']}) · ไม่มีการลบข้อมูลทิ้ง ทุกแถวถูกจัดเข้า 1 ใน 3 กลุ่ม", f["sub"])

    qa_section(3, "ภาพรวม: ข้อมูลทุกแถวไปอยู่ที่ไหน")
    for c, h in enumerate(["กลุ่ม", "จำนวนแถว", "คำอธิบาย", "อยู่ที่ไหน"]):
        ws.write(4, c, h, head())
    t = clean["tables"]
    overview = [
        ("ไฟล์ต้นฉบับ", clean["raw_rows"], "ข้อมูลดิบทั้งหมดก่อนทำความสะอาด", "online_retail_II.xlsx"),
        ("ยอดขายจริง", t["sales_rows"], "รายการขายสินค้าที่ใช้คำนวณรายงานและแดชบอร์ด", f"ชีต \"{S_DATA}\" / sales_clean.csv"),
        ("คืนสินค้า/ยกเลิก", t["returns_rows"], "ใบยกเลิก (เลขที่บิลขึ้นต้นด้วย C) แยกไว้ ไม่ปนกับยอดขาย", f"ชีต \"{S_RET}\" / data/clean/returns.csv"),
        ("คัดออก (มีเหตุผลกำกับ)", t["excluded_rows"], "แถวซ้ำ, รายการที่ไม่ใช่สินค้า, ปรับสต็อก, ราคา 0, ปรับหนี้, ออร์เดอร์ใหญ่ที่ยกเลิกทั้งรายการ", "data/clean/excluded.csv (คอลัมน์ Reason)"),
    ]
    for i, row in enumerate(overview):
        bg, col = [("#F1F5F9", "#334155"), (TONE["sale"]["tint"], NAVY), (TONE["ret"]["tint"], TONE["ret"]["text"]), ("#F1F5F9", "#64748B")][i]
        ws.write(5 + i, 0, row[0], qa(bold=True, bg=bg, color=col))
        ws.write_number(5 + i, 1, row[1], qa(num=True, bg=bg))
        ws.write(5 + i, 2, row[2], qa(bg=bg))
        ws.write(5 + i, 3, row[3], qa(bg=bg))
    ok = F(bold=True, bg_color=TONE["net"]["tint"], font_color=TONE["net"]["text"], top=2, top_color=TONE["net"]["main"], text_wrap=True, valign="top")
    ws.write(9, 0, "✓ ตรวจยอด: ยอดขาย + คืนสินค้า + คัดออก", ok)
    ws.write_formula(9, 1, "=B7+B8+B9", F(bold=True, bg_color=TONE["net"]["tint"], font_color=TONE["net"]["text"], top=2, top_color=TONE["net"]["main"], num_format="#,##0", valign="top"), t["sales_rows"] + t["returns_rows"] + t["excluded_rows"])
    ws.write(9, 2, "ต้องเท่ากับจำนวนแถวของไฟล์ต้นฉบับ (ไม่มีแถวหาย)", ok)
    ws.write(9, 3, "", ok)

    qa_section(11, "ปัญหาที่พบและวิธีแก้")
    for c, h in enumerate(["ปัญหาที่พบ", "จำนวน (แถว)", "วิธีแก้", "ผลลัพธ์"]):
        ws.write(12, c, h, head("ret") if c < 2 else head("net"))
    issues = [
        ("แถวซ้ำ", ex_rows["แถวซ้ำ"],
         f"รายการเดียวกันในบิลเดียวกันถูกบันทึกซ้ำ (ซ้ำกันทุกคอลัมน์ {a['duplicate_rows_extra']:,} แถว และซ้ำแต่สะกดชื่อสินค้าต่างกันอีก {ex_rows['แถวซ้ำ'] - a['duplicate_rows_extra']:,} แถว) เก็บแถวแรกไว้ ย้ายที่เหลือไปตารางคัดออก",
         "ยอดขายไม่ถูกนับซ้ำ"),
        ("ไม่มีรหัสลูกค้า (Customer ID)", a["customer_id_missing"],
         "เป็นการขายจริง จึงเก็บไว้ในยอดขาย และใส่ค่า \"ไม่ระบุ\" แทนช่องว่าง",
         f"ในยอดขายมี {fx['customer_id_missing_in_sales']:,} แถวที่เป็น \"ไม่ระบุ\" ยังนับยอดขายได้ครบ"),
        ("ออร์เดอร์ยกเลิก/คืนสินค้า (บิลขึ้นต้นด้วย C)", a["cancelled_rows"],
         "แยกรายการสินค้าที่ถูกยกเลิกไปชีต \"คืนสินค้า\" ส่วนค่าธรรมเนียม/ส่วนลดที่อยู่ในใบยกเลิกจัดเป็น \"ไม่ใช่สินค้า\"",
         f"ชีตคืนสินค้า {t['returns_rows']:,} แถว รวม £{-k['returns_gbp']:,.2f}"),
        ("รายการที่ไม่ใช่สินค้า (ค่าส่ง, ค่าธรรมเนียม, ส่วนลด, ปรับยอด, บัตรของขวัญ)", ex_rows["ไม่ใช่สินค้า"],
         "คัดออกจากยอดขายสินค้า โดยดูจากรหัสสินค้า เช่น POST, DOT, M, D, AMAZONFEE, BANK CHARGES",
         "ยอดขายสะท้อนเฉพาะการขายสินค้า"),
        ("จำนวนติดลบที่ไม่ใช่ใบยกเลิก (ปรับสต็อก)", ex_rows["ปรับสต็อก"],
         "เป็นบันทึกภายใน เช่น ของเสียหาย ของหาย ตรวจนับ ไม่ใช่การขาย จึงคัดออก", "ไม่กระทบยอดขาย (มูลค่า £0)"),
        ("ราคาเป็น 0", a["price_zero"],
         f"คัดออก (ส่วนหนึ่งซ้อนกับปัญหาอื่นด้านบน เหลือที่คัดออกด้วยเหตุผลนี้ {ex_rows['ราคาเป็น 0']:,} แถว)", "ไม่มีรายการราคา 0 ในยอดขาย"),
        ("ราคาติดลบ / ปรับหนี้เสีย (บิลขึ้นต้นด้วย A)", ex_rows["ปรับหนี้เสีย"],
         "เป็นรายการทางบัญชี ไม่ใช่การขาย จึงคัดออก", "ไม่มีราคาติดลบในยอดขาย"),
        ("ออร์เดอร์ใหญ่ผิดปกติที่สั่งแล้วยกเลิกทั้งรายการในวันเดียวกัน", ex_rows["สั่งแล้วยกเลิกทั้งรายการในวันเดียวกัน"],
         "ไม่ใช่การขายจริง (เช่น คีย์ออร์เดอร์ผิด) จึงย้ายทั้งรายการขายและใบยกเลิกไปตารางคัดออก", "ยอดขายและอัตราคืนสินค้าไม่สูงเกินจริง สินค้าขายดีไม่ถูกบิดเบือน"),
        ("ชื่อสินค้ามีช่องว่างเกิน", fx["description_whitespace_fixed_rows"],
         "ตัดช่องว่างหน้า/หลัง และช่องว่างซ้อนกลางข้อความ", "ชื่อสินค้าสะอาด ค้นหา/จัดกลุ่มได้ถูกต้อง"),
        ("รหัสสินค้าเดียวกันแต่ชื่อสินค้าหลายแบบ", fx["description_unified_rows"],
         f"พบ {a['stock_codes_with_multiple_descriptions']:,} รหัสที่มีชื่อมากกว่า 1 แบบ ให้ใช้ชื่อที่พบบ่อยที่สุดเป็นชื่อเดียว", "1 รหัสสินค้า = 1 ชื่อ"),
        ("รหัสสินค้าตัวพิมพ์เล็ก/ใหญ่ไม่ตรงกัน", fx["stockcode_uppercased_rows"],
         "แปลงเป็นตัวพิมพ์ใหญ่ทั้งหมด เช่น 85123a → 85123A", "สินค้าเดียวกันไม่ถูกแยกเป็น 2 รายการ"),
        ("ชื่อประเทศไม่เป็นมาตรฐาน", fx["country_renamed_rows"],
         "EIRE → Ireland, RSA → South Africa, USA → United States", "ใช้กับแผนที่ใน Looker Studio ได้"),
        ("ไม่มีชื่อสินค้า (Description ว่าง)", a["description_missing"],
         "ทุกแถวเป็นรายการราคา 0 (ปรับสต็อก) จึงถูกคัดออกไปด้วยเหตุผลข้างต้น", "ไม่มีชื่อสินค้าว่างในยอดขาย"),
    ]
    for i, row in enumerate(issues):
        r = 13 + i
        z = i % 2 == 1
        ws.write(r, 0, row[0], qa(z, bold=True))
        ws.write_number(r, 1, row[1], qa(z, num=True))
        ws.write(r, 2, row[2], qa(z))
        ws.write(r, 3, "✓ " + row[3], qa(z, color=TONE["net"]["text"]))

    r = 13 + len(issues) + 1
    qa_section(r, "ข้อสังเกตสำคัญ")
    exc = pd.read_csv(ROOT / "data" / "clean" / "excluded.csv")
    big = exc[exc["Reason"].str.startswith("สั่งแล้วยกเลิก") & (exc["Amount_GBP"] > 0)]["Amount_GBP"].sort_values(ascending=False).tolist()
    big_txt = " และ ".join(f"£{v:,.2f}" for v in big)
    notes = [
        f"มีออร์เดอร์ขนาดใหญ่ผิดปกติ {len(big)} รายการ ({big_txt}) ที่สั่งแล้วถูกยกเลิกทั้งรายการในวันเดียวกัน จึงย้ายทั้งรายการขายและใบยกเลิกไปตารางคัดออก ไม่นับเป็นยอดขายหรือยอดคืนสินค้า",
        f"ข้อมูลเดือนสุดท้ายถึงวันที่ {last_day.day} {MONTH_TH[last_day.month - 1]} {last_day.year} เท่านั้น ยอดเดือนนั้นจึงต่ำกว่าเดือนเต็ม",
        "อัตราคืนสินค้า = มูลค่าคืนสินค้า ÷ ยอดขายก่อนหักคืน",
    ]
    for i, n in enumerate(notes):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 3, "• " + n, f["wrap"])
        ws.set_row(r + 1 + i, 30)

    r = r + len(notes) + 2
    qa_section(r, "วิธีใช้ไฟล์นี้กับข้อมูลเดือนใหม่")
    howto = [
        f"1) วางข้อมูลใหม่ต่อท้ายชีต \"{S_DATA}\" (คอลัมน์ A–N) โดยให้รายการของบิลเดียวกันอยู่ติดกัน แล้วลากสูตรคอลัมน์ O (OrderFlag) ลงมาให้ครบ",
        f"2) ในชีต \"{S_SUM}\" ตารางยอดขายรายเดือน: แทรกแถวใหม่ก่อนแถว \"รวม\" พิมพ์เดือน (วันที่ 1 ของเดือน) แล้วคัดลอกสูตรจากแถวบน",
        f"3) ถ้ามีสินค้าหรือประเทศใหม่ ให้เพิ่มรหัส/ชื่อในชีต \"{S_PROD}\" หรือ \"{S_CTRY}\" แล้วคัดลอกสูตรจากแถวบน อันดับขายดีจะจัดใหม่เอง",
        f"4) ใบยกเลิก/คืนสินค้า ให้วางต่อท้ายชีต \"{S_RET}\" (ยอดเป็นค่าติดลบ)",
        "5) KPI ด้านบน ตาราง และกราฟ จะอัปเดตเองทั้งหมด เพราะเป็นสูตรที่อ้างทั้งคอลัมน์",
    ]
    for i, n in enumerate(howto):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 3, n, f["wrap"])
        ws.set_row(r + 1 + i, 30)
    r = r + len(howto) + 2
    ws.merge_range(r, 0, r, 3, "ที่มาข้อมูล: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D (CC BY 4.0) · หน่วยเงิน GBP (£)", f["sub"])

    ws_sum.activate()
    wb.close()
    print(f"saved {OUT.name}  ({OUT.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
