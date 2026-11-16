"""สร้างภาพสำหรับลงขาย -> images/

รัน (หลัง clean.py และ build_dashboard.py):  .venv/bin/python build_images.py
ต้องมี Google Chrome ในเครื่อง (ใช้โหมด headless ถ่ายภาพจากไฟล์ HTML)

ผลลัพธ์:
  images/before_after.png        ตารางก่อน/หลังทำความสะอาด (แถวจริงจากข้อมูล)
  images/cover.png               ภาพปก 1280x720
  images/dashboard_desktop.png   แคปแดชบอร์ด HTML จอกว้าง
  images/dashboard_mobile.png    แคปแดชบอร์ด HTML มือถือ
ตัวเลขทุกตัวมาจาก reports/*.json และข้อมูลจริง ไม่มีการแต่งขึ้นเอง
"""
import html
import json
import os
import subprocess
from pathlib import Path

import pandas as pd
from PIL import Image

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "images"
SRC = IMG / "src"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
FONT_LINK = '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap" rel="stylesheet">'
FONT = '"IBM Plex Sans Thai","Noto Sans Thai",Thonburi,"Leelawadee UI",Tahoma,sans-serif'

clean = json.loads((ROOT / "reports" / "clean_summary.json").read_text(encoding="utf-8"))
dash = json.loads((ROOT / "reports" / "dashboard_data.json").read_text(encoding="utf-8"))
COVER_TAG = os.environ.get("COVER_TAG", "โปรเจกต์วิเคราะห์ข้อมูลยอดขาย")  # ข้อความป้ายบนปก ปรับได้ด้วยตัวแปรแวดล้อม
K = clean["kpi"]
T = clean["tables"]


def shot(html_path: Path, out: Path, w: int, h: int, wait_ms: int = 6000) -> None:
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    f"--window-size={w},{h}", f"--virtual-time-budget={wait_ms}", f"--screenshot={out}",
                    html_path.as_uri()], check=True, capture_output=True, timeout=120)


# ------------------------------------------------------------------ ก่อน/หลัง
PICK = [  # (index ในไฟล์ดิบ, ป้ายปัญหา)
    (1443, "Customer ID หาย"),
    (622, "ชื่อว่าง + ราคา 0"),
    (45, "ค่าส่ง ไม่ใช่สินค้า"),
    (141, "ใบยกเลิก/ส่วนลด"),
    (154, "ใบยกเลิก"),
    (87, "เว้นวรรคซ้อน"),
    (5370, "รหัสพิมพ์เล็ก"),
    (494, "รายการ"),
    (517, "แถวซ้ำ"),
]


def money(v: float) -> str:
    return f"-£{abs(v):,.2f}" if v < 0 else f"£{v:,.2f}"


def build_before_after() -> None:
    raw = pd.read_pickle(ROOT / "data" / "interim" / "raw_2010_2011.pkl")
    cs = pd.read_csv(ROOT / "sales_clean.csv", dtype={"Invoice": str, "StockCode": str, "CustomerID": str})
    cr = pd.read_csv(ROOT / "data" / "clean" / "returns.csv", dtype={"Invoice": str, "StockCode": str, "CustomerID": str})
    ce = pd.read_csv(ROOT / "data" / "clean" / "excluded.csv", dtype={"Invoice": str, "StockCode": str, "CustomerID": str})

    def find(tbl, r):
        m = tbl[(tbl.Invoice == str(r.Invoice)) & (tbl.StockCode.str.upper() == str(r.StockCode).upper()) &
                (tbl.Quantity == r.Quantity) & (tbl.InvoiceDate == r.InvoiceDate.strftime("%Y-%m-%d %H:%M:%S"))]
        return m

    left, right = [], []
    seen_dup = set()
    for idx, tag in PICK:
        r = raw.loc[idx]
        key = (str(r.Invoice), str(r.StockCode), r.Quantity, r.InvoiceDate)
        problems = set()
        inv = str(r.Invoice)
        if pd.isna(r["Customer ID"]):
            problems.add("cust")
        if pd.isna(r.Description) or r.Price == 0:
            problems.add("desc")
            problems.add("price")
        if str(r.StockCode) in ("POST", "D", "M"):
            problems.add("code")
        if inv.startswith("C"):
            problems.add("inv")
        if isinstance(r.Description, str) and (r.Description != r.Description.strip() or "  " in r.Description):
            problems.add("desc")
        if str(r.StockCode) != str(r.StockCode).upper():
            problems.add("code")
        is_dup = key in seen_dup
        seen_dup.add(key)
        if is_dup:
            problems.update({"inv", "code", "desc", "qty", "price"})
            tag = "แถวซ้ำ"
        if idx == 494:
            tag = "ต้นฉบับ (มีสำเนา)"
        left.append((r, tag, problems, is_dup))

        # หาว่าไปอยู่ที่ไหน (แถวซ้ำ: แถวแรกอยู่ยอดขาย แถวที่สองถูกคัดออก)
        s, rt, ex = find(cs, r), find(cr, r), find(ce, r)
        if len(s) and not is_dup:
            right.append(("sale", s.iloc[0], None))
        elif len(rt):
            right.append(("ret", rt.iloc[0], None))
        elif len(ex):
            right.append(("ex", ex.iloc[0], ex.iloc[0]["Reason"]))
        else:
            right.append(("ex", None, "คัดออก"))

    def esc(v):
        return html.escape("" if pd.isna(v) else str(v))

    def cell(v, bad=False):
        return f'<td class="{"bad" if bad else ""}">{esc(v) if str(v) != "" else "<i>ว่าง</i>"}</td>'

    rows_l = ""
    for r, tag, pr, dup in left:
        desc = "" if pd.isna(r.Description) else str(r.Description).replace(" ", " ")
        cust = "" if pd.isna(r["Customer ID"]) else str(int(r["Customer ID"]))
        rows_l += ("<tr>" + cell(r.Invoice, "inv" in pr) + cell(r.StockCode, "code" in pr) +
                   f'<td class="{"bad" if "desc" in pr else ""} w">{html.escape(desc) if desc else "<i>ว่าง</i>"}</td>' +
                   cell(r.Quantity, "qty" in pr) + cell(f"{r.Price:g}", "price" in pr) + cell(cust, "cust" in pr) +
                   f'<td class="tag">{esc(tag)}</td></tr>')

    rows_r = ""
    for kind, row, reason in right:
        if kind == "sale":
            cust = "ไม่ระบุ" if row.CustomerID == "ไม่ระบุ" else row.CustomerID
            rows_r += ("<tr>" + f"<td>{esc(row.Invoice)}</td><td>{esc(row.StockCode)}</td><td class='w'>{esc(row.Description)}</td>"
                       f"<td>{row.Quantity}</td><td>{row.Price:g}</td><td class='good'>{money(row.Sales_GBP)}</td>"
                       f"<td>{esc(row.Month)} · {esc(row.Weekday)}</td><td><span class='pill sale'>ยอดขาย</span></td></tr>")
        elif kind == "ret":
            rows_r += ("<tr>" + f"<td>{esc(row.Invoice)}</td><td>{esc(row.StockCode)}</td><td class='w'>{esc(row.Description)}</td>"
                       f"<td>{row.Quantity}</td><td>{row.Price:g}</td><td class='ret'>{money(row.Return_GBP)}</td>"
                       f"<td>{esc(row.Month)} · {esc(row.Weekday)}</td><td><span class='pill ret'>คืนสินค้า</span></td></tr>")
        else:
            short = reason.split(" (")[0] if reason else "คัดออก"
            rows_r += ("<tr class='gray'>" + (f"<td>{esc(row.Invoice)}</td><td>{esc(row.StockCode)}</td>"
                       f"<td class='w'>{esc(row.Description) or '<i>ว่าง</i>'}</td><td>{row.Quantity}</td><td>{row.Price:g}</td>"
                       f"<td>{money(row.Amount_GBP)}</td><td>—</td>" if row is not None else "<td colspan=7></td>") +
                       f"<td><span class='pill ex'>คัดออก: {html.escape(short)}</span></td></tr>")

    page = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">{FONT_LINK}
<style>
*{{box-sizing:border-box}} body{{margin:0;width:1800px;background:#F5F8FC;font-family:{FONT};color:#0F172A;padding:40px 44px 34px}}
h1{{margin:0;font-size:38px;color:#1E3A8A;font-weight:700}} .sub{{color:#475569;font-size:20px;margin:6px 0 22px}}
.grid{{display:grid;grid-template-columns:1fr 1.18fr;gap:26px;align-items:start}}
.panel{{background:#fff;border-radius:16px;border:1px solid #E2E8F0;overflow:hidden}}
.head{{padding:14px 20px;font-size:22px;font-weight:700;color:#fff}} .head small{{font-weight:500;font-size:16px;opacity:.9;margin-left:8px}}
.before .head{{background:#B91C1C}} .after .head{{background:#1D4ED8}}
table{{width:100%;border-collapse:collapse;font-size:16px}} th{{background:#F1F5F9;color:#475569;font-weight:600;text-align:left;padding:9px 10px;font-size:14.5px;white-space:nowrap}}
td{{padding:10px 10px;border-top:1px solid #EEF2F7;white-space:nowrap;font-variant-numeric:tabular-nums}} td.w{{max-width:250px;overflow:hidden;text-overflow:ellipsis}}
td.bad{{background:#FEE2E2;color:#991B1B;font-weight:600}} td i{{color:#B91C1C;font-style:normal}} td.tag{{color:#B91C1C;font-size:15px;font-weight:600}}
.good{{color:#047857;font-weight:600}} .ret{{color:#C2410C;font-weight:600}} tr.gray td{{color:#64748B;background:#FAFAFA}}
.pill{{display:inline-block;padding:3px 11px;border-radius:999px;font-size:14px;font-weight:600}} .pill.sale{{background:#DBEAFE;color:#1E3A8A}} .pill.ret{{background:#FFEDD5;color:#C2410C}} .pill.ex{{background:#E2E8F0;color:#475569}}
.bar{{display:flex;gap:14px;margin-top:22px}} .stat{{flex:1;background:#fff;border:1px solid #E2E8F0;border-radius:14px;padding:14px 20px;border-top:5px solid #1D4ED8}}
.stat b{{display:block;font-size:34px;color:#1E3A8A;line-height:1.2}} .stat span{{color:#475569;font-size:16.5px}}
.stat.o{{border-top-color:#EA580C}} .stat.o b{{color:#C2410C}} .stat.g{{border-top-color:#6B7280}} .stat.g b{{color:#334155}} .stat.r{{border-top-color:#B91C1C}} .stat.r b{{color:#B91C1C}}
.foot{{margin-top:16px;color:#64748B;font-size:14px}}
</style></head><body>
<h1>ก่อน → หลังทำความสะอาด: แถวเดียวกัน ไปอยู่ที่ไหน</h1>
<div class="sub">แถวจริงจากไฟล์ยอดขายร้านค้าออนไลน์ (หน่วยเงิน £) ไม่มีแถวไหนถูกลบทิ้งเงียบ ๆ ทุกแถวถูกจัดเข้ายอดขาย คืนสินค้า หรือคัดออกพร้อมเหตุผล</div>
<div class="grid">
 <div class="panel before"><div class="head">ก่อน: ไฟล์ดิบ<small>ช่องสีแดง = ปัญหา</small></div>
  <table><tr><th>Invoice</th><th>StockCode</th><th>Description</th><th>Qty</th><th>Price</th><th>Customer ID</th><th>ปัญหา</th></tr>{rows_l}</table></div>
 <div class="panel after"><div class="head">หลัง: ข้อมูลพร้อมใช้<small>เพิ่มยอดขาย เดือน วัน และแยกประเภทให้แล้ว</small></div>
  <table><tr><th>Invoice</th><th>StockCode</th><th>Description</th><th>Qty</th><th>Price</th><th>ยอด (£)</th><th>เดือน · วัน</th><th>ไปอยู่ที่</th></tr>{rows_r}</table></div>
</div>
<div class="bar">
 <div class="stat r"><b>{clean['raw_rows']:,}</b><span>แถวในไฟล์ดิบ</span></div>
 <div class="stat"><b>{T['sales_rows']:,}</b><span>แถวยอดขายพร้อมใช้</span></div>
 <div class="stat o"><b>{T['returns_rows']:,}</b><span>แถวคืนสินค้า/ยกเลิก (แยกไว้)</span></div>
 <div class="stat g"><b>{T['excluded_rows']:,}</b><span>แถวคัดออก พร้อมเหตุผลกำกับ</span></div>
</div>
<div class="foot">ที่มาข้อมูล: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D (CC BY 4.0)</div>
</body></html>"""
    p = SRC / "before_after.html"
    p.write_text(page, encoding="utf-8")
    out = IMG / "before_after.png"
    shot(p, out, 1800, 1100)
    trim(out)


# ------------------------------------------------------------------ ปก
def build_cover() -> None:
    months = dash["monthly"]
    mx = max(m["sales"] for m in months)
    bars = ""
    for i, m in enumerate(months):
        h = round(m["sales"] / mx * 100)
        col = "#1E3A8A" if m["sales"] == mx else ("#93C5FD" if m["partial"] else "#2563EB")
        bars += f'<i style="height:{h}%;background:{col}"></i>'
    k = dash["kpi"]
    page = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">{FONT_LINK}
<style>
*{{box-sizing:border-box}} html,body{{margin:0;width:1280px;height:720px;overflow:hidden}}
body{{font-family:{FONT};background:linear-gradient(135deg,#FFFFFF 0%,#EAF2FF 60%,#DBEAFE 100%);position:relative;color:#0F172A}}
.blob{{position:absolute;right:-140px;top:-160px;width:620px;height:620px;border-radius:50%;background:#BFDBFE;opacity:.55}}
.blob2{{position:absolute;left:-120px;bottom:-200px;width:420px;height:420px;border-radius:50%;background:#DBEAFE}}
.left{{position:absolute;left:70px;top:92px;width:640px}}
.tag{{display:inline-block;background:#1E3A8A;color:#fff;border-radius:999px;padding:8px 22px;font-size:25px;font-weight:600}}
h1{{margin:26px 0 0;font-size:68px;line-height:1.28;font-weight:700;color:#64748B}}
h1 .old{{display:block;text-decoration:line-through;text-decoration-color:#EF4444;text-decoration-thickness:5px}}
h1 .arrow{{display:block;font-size:54px;color:#2563EB;line-height:1}}
h1 .new{{display:block;color:#1D4ED8;font-size:84px}}
.sub{{margin-top:22px;font-size:29px;color:#334155;line-height:1.5}}
.chips{{margin-top:26px;display:flex;gap:12px;flex-wrap:wrap}} .chips span{{background:#fff;border:2px solid #BFDBFE;color:#1E3A8A;border-radius:12px;padding:8px 18px;font-size:23px;font-weight:600}}
.mess{{position:absolute;left:735px;top:520px;width:300px;background:#fff;border:2px solid #FCA5A5;border-radius:16px;padding:14px;transform:rotate(-6deg);box-shadow:0 14px 34px rgba(185,28,28,.18);font-size:15px}}
.mess b{{display:block;color:#B91C1C;font-size:17px;margin-bottom:6px}} .mess div{{display:flex;justify-content:space-between;padding:5px 8px;border-radius:6px;margin:3px 0;background:#FEE2E2;color:#7F1D1D}}
.mess div:nth-child(odd){{background:#FFF1F1}}
.card{{position:absolute;right:62px;top:92px;width:520px;background:#fff;border-radius:22px;padding:26px 28px;box-shadow:0 26px 60px rgba(30,58,138,.28);transform:rotate(2.5deg);border:1px solid #DBEAFE}}
.card .t{{font-size:21px;color:#1E3A8A;font-weight:700}} .kp{{display:flex;gap:12px;margin-top:14px}}
.kp div{{flex:1;background:#EFF6FF;border-top:4px solid #2563EB;border-radius:10px;padding:10px 12px}} .kp div.o{{background:#FFF7ED;border-top-color:#EA580C}}
.kp small{{display:block;font-size:15px;color:#64748B}} .kp b{{font-size:29px;color:#1E3A8A}} .kp .o b{{color:#C2410C}}
.chart{{margin-top:18px;height:175px;display:flex;align-items:flex-end;gap:7px;border-bottom:2px solid #E2E8F0;padding:0 4px}}
.chart i{{flex:1;border-radius:5px 5px 0 0;display:block}} .cap{{font-size:15px;color:#64748B;margin-top:8px}}
.src{{position:absolute;left:70px;bottom:26px;font-size:16.5px;color:#64748B}}
</style></head><body><div class="blob"></div><div class="blob2"></div>
<div class="left"><span class="tag">{COVER_TAG}</span>
<h1><span class="old">ไฟล์ยอดขายรก</span><span class="arrow">↓</span><span class="new">รายงานพร้อมใช้</span></h1>
<div class="sub">ทำความสะอาดข้อมูล สรุปยอดด้วย Excel<br>และแดชบอร์ดที่เจ้าของร้านอ่านเข้าใจทันที</div>
<div class="chips"><span>ตรวจสุขภาพข้อมูล</span><span>Excel สูตรจริง</span><span>แดชบอร์ดดูง่าย</span></div></div>
<div class="mess"><b>ก่อน: ไฟล์รก</b><div><span>C536379 · Discount</span><span>-1</span></div><div><span>ชื่อสินค้าว่าง · ราคา 0</span><span>56</span></div><div><span>Customer ID หาย</span><span>NaN</span></div></div>
<div class="card"><div class="t">หลัง: แดชบอร์ดยอดขาย (£)</div>
<div class="kp"><div><small>ยอดขายรวม</small><b>£{k['sales']/1e6:.2f}M</b></div><div><small>ออร์เดอร์</small><b>{k['orders']:,}</b></div><div class="o"><small>อัตราคืนสินค้า</small><b>{k['return_rate']:.2f}%</b></div></div>
<div class="chart">{bars}</div><div class="cap">ยอดขายรายเดือน ธ.ค. 2010 – ธ.ค. 2011</div></div>
<div class="src">ตัวอย่างผลงานจากข้อมูลจริง {clean['raw_rows']:,} แถว · Online Retail II, UCI (CC BY 4.0)</div>
</body></html>"""
    p = SRC / "cover.html"
    p.write_text(page, encoding="utf-8")
    shot(p, IMG / "cover.png", 1280, 720)


# ------------------------------------------------------------------ แคปแดชบอร์ด
def trim(path: Path, bottom_pad: int = 0) -> None:
    """ตัดพื้นที่ว่างด้านล่างที่เกินเนื้อหา"""
    im = Image.open(path).convert("RGB")
    bg = im.getpixel((5, im.height - 5))
    px = im.load()
    last = 0
    for y in range(im.height - 1, -1, -1):
        row_diff = any(sum(abs(a - b) for a, b in zip(px[x, y], bg)) > 18 for x in range(0, im.width, 7))
        if row_diff:
            last = y
            break
    im.crop((0, 0, im.width, min(im.height, last + 30 + bottom_pad))).save(path)


def build_dashboard_shots() -> None:
    d = ROOT / "dashboard.html"
    shot(d, IMG / "dashboard_desktop.png", 1280, 2600, 9000)
    trim(IMG / "dashboard_desktop.png")
    # มือถือ: Chrome headless เปิดหน้าต่างแคบกว่า ~500px ไม่ได้ จึงแสดงใน iframe กว้าง 390px จริง (media query ตอบสนองตาม iframe)
    frame = SRC / "mobile_frame.html"
    frame.write_text('<!doctype html><meta charset="utf-8"><body style="margin:0;background:#F5F8FC">'
                     f'<iframe src="{d.as_uri()}" width="390" height="3600" style="border:0;display:block"></iframe>', encoding="utf-8")
    out = IMG / "dashboard_mobile.png"
    shot(frame, out, 700, 3700, 12000)
    Image.open(out).convert("RGB").crop((0, 0, 390, 3700)).save(out)
    trim(out)


if __name__ == "__main__":
    IMG.mkdir(exist_ok=True)
    SRC.mkdir(exist_ok=True)
    build_before_after()
    build_cover()
    build_dashboard_shots()
    for f in ["before_after", "cover", "dashboard_desktop", "dashboard_mobile"]:
        im = Image.open(IMG / f"{f}.png")
        print(f, im.size)
