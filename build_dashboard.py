"""สร้างแดชบอร์ดตัวอย่างแบบ HTML ไฟล์เดียว -> dashboard.html

รัน (หลัง clean.py):  .venv/bin/python build_dashboard.py

ตัวเลขทั้งหมดคำนวณจาก sales_clean.csv และ data/clean/returns.csv แล้วฝังลงในไฟล์ HTML
ไม่ใช้ไลบรารีภายนอก เปิดดูแบบออฟไลน์ได้ หน่วยเงิน GBP (£)
"""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "dashboard.html"
MONTH_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
WEEKDAY_TH = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]


def build_data() -> dict:
    s = pd.read_csv(ROOT / "sales_clean.csv", dtype={"Invoice": str, "StockCode": str, "CustomerID": str})
    r = pd.read_csv(ROOT / "data" / "clean" / "returns.csv", dtype={"Invoice": str, "StockCode": str})
    total = float(s["Sales_GBP"].sum())
    orders = int(s["Invoice"].nunique())
    returns = float(-r["Return_GBP"].sum())
    last = pd.Timestamp(s["Date"].max())
    first = pd.Timestamp(s["Date"].min())

    monthly = []
    g = s.groupby("Month").agg(sales=("Sales_GBP", "sum"), orders=("Invoice", "nunique"))
    for m, row in g.iterrows():
        y, mo = int(m[:4]), int(m[5:])
        partial = (y == last.year and mo == last.month and last.day < last.days_in_month)
        monthly.append({"m": m, "label": f"{MONTH_TH[mo - 1]} {str(y)[2:]}", "full": f"{MONTH_TH[mo - 1]} {y}",
                        "sales": round(float(row.sales), 2), "orders": int(row.orders), "partial": bool(partial)})

    p = (s.groupby("StockCode").agg(name=("Description", "first"), sales=("Sales_GBP", "sum"), qty=("Quantity", "sum"))
         .sort_values("sales", ascending=False).head(10).reset_index())
    products = [{"code": x.StockCode, "name": x.name, "sales": round(float(x.sales), 2), "qty": int(x.qty)}
                for x in p.itertuples(index=False)]

    c = (s.groupby("Country").agg(sales=("Sales_GBP", "sum"), orders=("Invoice", "nunique"))
         .sort_values("sales", ascending=False).head(11).reset_index())
    countries = [{"name": x.Country, "sales": round(float(x.sales), 2), "orders": int(x.orders)}
                 for x in c.itertuples(index=False)]

    h = s.groupby(["WeekdayNum", "Hour"]).agg(sales=("Sales_GBP", "sum"), orders=("Invoice", "nunique")).reset_index()
    heat = [{"d": int(x.WeekdayNum), "h": int(x.Hour), "sales": round(float(x.sales), 2), "orders": int(x.orders)}
            for x in h.itertuples(index=False)]

    return {
        "kpi": {"sales": round(total, 2), "orders": orders, "aov": round(total / orders, 2),
                "returns": round(returns, 2), "return_rate": round(returns / total * 100, 2),
                "net": round(total - returns, 2), "units": int(s["Quantity"].sum()),
                "customers": int(s.loc[s["CustomerID"] != "ไม่ระบุ", "CustomerID"].nunique()),
                "countries": int(s["Country"].nunique()), "products": int(s["StockCode"].nunique())},
        "period": {"from": f"{first.day} {MONTH_TH[first.month - 1]} {first.year}",
                   "to": f"{last.day} {MONTH_TH[last.month - 1]} {last.year}"},
        "monthly": monthly, "products": products, "countries": countries, "heat": heat,
        "weekdays": WEEKDAY_TH,
        "hours": [int(h["Hour"].min()), int(h["Hour"].max())],
    }


TEMPLATE = r"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>แดชบอร์ดยอดขาย (ตัวอย่าง)</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #F5F8FC; --card: #FFFFFF; --ink: #0F172A; --ink2: #475569; --muted: #94A3B8;
    --line: #E2E8F0; --blue: #2563EB; --blue-dark: #1E3A8A; --blue-soft: #DBEAFE; --blue-pale: #EFF6FF;
  }
  * { box-sizing: border-box; }
  html { -webkit-text-size-adjust: 100%; }
  body { margin: 0; background: var(--bg); color: var(--ink);
    font-family: "IBM Plex Sans Thai", "Noto Sans Thai", "Leelawadee UI", Thonburi, Tahoma, sans-serif;
    font-size: 15px; line-height: 1.55; }
  .wrap { max-width: 1120px; margin: 0 auto; padding: 28px 20px 40px; }
  header { display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between; gap: 8px 24px; margin-bottom: 20px; }
  h1 { font-size: 26px; line-height: 1.25; margin: 0; color: var(--blue-dark); font-weight: 700; }
  .sub { color: var(--ink2); margin: 4px 0 0; }
  .pill { display: inline-block; background: var(--blue-soft); color: var(--blue-dark); border-radius: 999px; padding: 4px 12px; font-size: 13px; font-weight: 500; white-space: nowrap; }
  .kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 14px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 18px 20px; min-width: 0; }
  .kpi { border-top: 3px solid var(--blue); }
  .kpi .label { color: var(--ink2); font-size: 14px; }
  .kpi .value { font-size: 28px; font-weight: 700; color: var(--blue-dark); line-height: 1.25; margin: 2px 0; font-variant-numeric: tabular-nums; }
  .kpi .note { color: var(--muted); font-size: 13px; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 14px; }
  h2 { font-size: 17px; margin: 0; font-weight: 600; }
  .desc { color: var(--ink2); font-size: 13.5px; margin: 2px 0 14px; }
  .head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
  svg { display: block; width: 100%; height: auto; overflow: visible; }
  svg text { font-family: inherit; fill: var(--ink2); font-size: 12px; }
  .bar { fill: var(--blue); } .bar.partial { fill: #93C5FD; }
  .bar:hover, .hbar:hover .fill, .cell:hover { filter: brightness(0.9); }
  .hbar { display: grid; grid-template-columns: minmax(0, 46%) 1fr auto; align-items: center; gap: 10px; padding: 5px 0; cursor: default; }
  .hbar .name { font-size: 13.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .hbar .rank { color: var(--muted); display: inline-block; width: 1.6em; font-variant-numeric: tabular-nums; }
  .hbar .track { height: 14px; background: var(--blue-pale); border-radius: 4px; }
  .hbar .fill { height: 100%; background: var(--blue); border-radius: 0 4px 4px 0; min-width: 2px; }
  .hbar .val { font-size: 13.5px; font-weight: 600; font-variant-numeric: tabular-nums; min-width: 5.2em; text-align: right; }
  .seg { display: inline-flex; border: 1px solid var(--line); border-radius: 8px; overflow: hidden; flex: none; }
  .seg button { font: inherit; font-size: 13px; padding: 5px 10px; border: 0; background: #fff; color: var(--ink2); cursor: pointer; }
  .seg button[aria-pressed="true"] { background: var(--blue); color: #fff; }
  .heat { display: grid; gap: 2px; align-items: center; }
  .heat .hl { font-size: 12px; color: var(--ink2); text-align: center; }
  .heat .dl { font-size: 13px; color: var(--ink2); padding-right: 8px; white-space: nowrap; }
  .cell { aspect-ratio: 1.5 / 1; border-radius: 3px; background: #F1F5F9; }
  .legend { display: flex; align-items: center; gap: 8px; color: var(--ink2); font-size: 12.5px; margin-top: 12px; flex-wrap: wrap; }
  .ramp { display: inline-flex; gap: 2px; } .ramp i { width: 22px; height: 12px; border-radius: 2px; display: block; }
  .sw { width: 12px; height: 12px; border-radius: 3px; display: inline-block; vertical-align: -1px; margin-right: 4px; }
  .insight { background: var(--blue-pale); border-radius: 10px; padding: 10px 14px; margin-top: 14px; font-size: 14px; color: var(--blue-dark); }
  footer { color: var(--muted); font-size: 12.5px; margin-top: 22px; line-height: 1.7; }
  footer a { color: inherit; }
  #tip { position: fixed; z-index: 10; pointer-events: none; background: var(--ink); color: #fff; padding: 8px 12px; border-radius: 8px; font-size: 13px; line-height: 1.5; max-width: 260px; opacity: 0; transition: opacity .1s; box-shadow: 0 6px 20px rgba(15,23,42,.18); }
  #tip b { display: block; font-weight: 600; }
  @media (max-width: 980px) {
    .grid2 { grid-template-columns: 1fr; }
  }
  @media (max-width: 760px) {
    .kpis { grid-template-columns: repeat(2, 1fr); }
  }
  @media (max-width: 520px) {
    .wrap { padding: 18px 12px 28px; } h1 { font-size: 21px; } .card { padding: 14px; border-radius: 12px; }
    .kpis { gap: 10px; } .kpi .value { font-size: 21px; } .kpi .label { font-size: 13px; } .kpi .note { font-size: 12px; }
    .hbar { grid-template-columns: 1fr auto; gap: 2px 10px; padding: 6px 0; }
    .hbar .name { grid-column: 1 / 2; } .hbar .val { grid-column: 2 / 3; grid-row: 1; } .hbar .track { grid-column: 1 / 3; height: 10px; }
    .heat .hl { font-size: 10px; } .heat .dl { font-size: 11.5px; padding-right: 4px; } .cell { aspect-ratio: 1 / 1.25; border-radius: 2px; }
    .head { flex-direction: column; gap: 6px; } .desc { margin-top: 6px; }
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>แดชบอร์ดยอดขายร้านค้าออนไลน์</h1>
      <p class="sub">ข้อมูล <span id="period"></span> · หลังทำความสะอาดข้อมูลแล้ว</p>
    </div>
    <span class="pill">หน่วยเงิน: ปอนด์ (GBP, £)</span>
  </header>

  <section class="kpis" id="kpis"></section>

  <section class="card">
    <h2>ยอดขายรายเดือน</h2>
    <p class="desc">หน่วย: ล้านปอนด์ (£) · แตะหรือชี้ที่แท่งเพื่อดูตัวเลข</p>
    <div id="monthly"></div>
    <div class="legend"><span><i class="sw" style="background:#2563EB"></i>เดือนเต็ม</span><span><i class="sw" style="background:#93C5FD"></i><span id="partialNote"></span></span></div>
    <div class="insight" id="insightMonth"></div>
  </section>

  <div class="grid2">
    <section class="card">
      <h2>สินค้าขายดี 10 อันดับ</h2>
      <p class="desc">เรียงตามยอดขาย (£)</p>
      <div id="products"></div>
    </section>
    <section class="card">
      <div class="head">
        <div><h2>ประเทศที่ขายได้มากสุด</h2></div>
        <div class="seg" role="group" aria-label="เลือกมุมมองประเทศ">
          <button type="button" id="btnAll" aria-pressed="true">ทุกประเทศ</button>
          <button type="button" id="btnEx" aria-pressed="false">ไม่รวมอันดับ 1</button>
        </div>
      </div>
      <p class="desc" id="countryDesc"></p>
      <div id="countries"></div>
    </section>
  </div>

  <section class="card" style="margin-top:14px">
    <h2>ช่วงเวลาที่ขายดี</h2>
    <p class="desc">ยอดขาย (£) แยกตามวันในสัปดาห์และชั่วโมง · สีเข้ม = ขายดี</p>
    <div class="heat" id="heat"></div>
    <div class="legend"><span>น้อย</span><span class="ramp" id="ramp"></span><span>มาก</span><span><i class="sw" style="background:#F1F5F9;border:1px solid #E2E8F0"></i>ไม่มียอดขาย</span></div>
    <div class="insight" id="insightHeat"></div>
  </section>

  <footer>
    ผลงานตัวอย่าง: จัดระเบียบไฟล์ยอดขายและทำแดชบอร์ด · ตัวเลขทั้งหมดคำนวณจากข้อมูลจริงหลังทำความสะอาด (ไม่รวมรายการยกเลิก ค่าส่ง ค่าธรรมเนียม และแถวซ้ำ)<br>
    ที่มาข้อมูล: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository.
    <a href="https://doi.org/10.24432/C5CG6D">https://doi.org/10.24432/C5CG6D</a> (CC BY 4.0)
  </footer>
</div>
<div id="tip" role="tooltip"></div>

<script>
const DATA = __DATA__;
const $ = (id) => document.getElementById(id);
const gbp = (v, d = 2) => "£" + v.toLocaleString("en-GB", { minimumFractionDigits: d, maximumFractionDigits: d });
const num = (v) => v.toLocaleString("en-GB");
const mil = (v) => "£" + (v / 1e6).toFixed(2) + " ล้าน";
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ---------- KPI ----------
const k = DATA.kpi;
$("period").textContent = DATA.period.from + " – " + DATA.period.to;
$("kpis").innerHTML = [
  ["ยอดขายรวม", mil(k.sales), gbp(k.sales) + " (ก่อนหักคืนสินค้า)"],
  ["จำนวนออร์เดอร์", num(k.orders), "สินค้า " + num(k.units) + " ชิ้น"],
  ["ยอดเฉลี่ยต่อออร์เดอร์", gbp(k.aov), "ยอดขาย ÷ จำนวนออร์เดอร์"],
  ["อัตราคืนสินค้า", k.return_rate.toFixed(2) + "%", "ยอดคืน " + gbp(k.returns)],
].map(([l, v, n]) => `<div class="card kpi"><div class="label">${l}</div><div class="value">${v}</div><div class="note">${n}</div></div>`).join("");

// ---------- ยอดขายรายเดือน ----------
function drawMonthly() {
  const el = $("monthly"), W = Math.max(300, el.clientWidth), small = W < 520;
  const H = small ? 220 : 280, m = { l: 34, r: 6, t: 22, b: 26 };
  const data = DATA.monthly, max = Math.max(...data.map((d) => d.sales));
  const step = 0.4e6, top = Math.ceil(max / step) * step;
  const iw = W - m.l - m.r, ih = H - m.t - m.b, bw = iw / data.length, w = Math.min(44, bw - (small ? 5 : 12));
  const y = (v) => m.t + ih - (v / top) * ih;
  const best = data.reduce((a, b) => (b.sales > a.sales ? b : a));
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="กราฟแท่งยอดขายรายเดือน">`;
  for (let v = 0; v <= top + 1; v += step) {
    s += `<line x1="${m.l}" x2="${W - m.r}" y1="${y(v)}" y2="${y(v)}" stroke="${v ? "#E2E8F0" : "#CBD5E1"}"/>`;
    s += `<text x="${m.l - 6}" y="${y(v) + 4}" text-anchor="end">${(v / 1e6).toFixed(1)}</text>`;
  }
  data.forEach((d, i) => {
    const x = m.l + i * bw + (bw - w) / 2, yy = y(d.sales), h = m.t + ih - yy, r = Math.min(4, w / 2, h);
    const tip = `<b>${d.full}${d.partial ? " (ไม่ครบเดือน)" : ""}</b>ยอดขาย ${gbp(d.sales)}<br>${num(d.orders)} ออร์เดอร์`;
    s += `<path class="bar${d.partial ? " partial" : ""}" data-tip="${esc(tip)}" d="M${x},${yy + h} V${yy + r} Q${x},${yy} ${x + r},${yy} H${x + w - r} Q${x + w},${yy} ${x + w},${yy + r} V${yy + h} Z"/>`;
    if (d === best) s += `<text x="${x + w / 2}" y="${yy - 6}" text-anchor="middle" style="font-weight:600;fill:#1E3A8A">${(d.sales / 1e6).toFixed(2)}</text>`;
    if (!small || i % 3 === 0) s += `<text x="${x + w / 2}" y="${H - 8}" text-anchor="middle">${d.label}</text>`;
  });
  el.innerHTML = s + "</svg>";
}
(() => {
  const data = DATA.monthly, best = data.reduce((a, b) => (b.sales > a.sales ? b : a));
  const full = data.filter((d) => !d.partial), low = full.reduce((a, b) => (b.sales < a.sales ? b : a));
  const p = data.find((d) => d.partial);
  $("partialNote").textContent = p ? `${p.full} ข้อมูลถึงวันที่ ${DATA.period.to.split(" ")[0]} (ไม่ครบเดือน)` : "เดือนที่ข้อมูลไม่ครบ";
  $("insightMonth").innerHTML = `เดือนที่ขายดีที่สุดคือ <b>${best.full}</b> ยอดขาย ${gbp(best.sales)} สูงกว่าเดือนที่ขายน้อยที่สุด (${low.full}) ${(best.sales / low.sales).toFixed(1)} เท่า`;
})();

// ---------- แท่งแนวนอน ----------
function hbars(el, rows, startRank) {
  const max = Math.max(...rows.map((r) => r.sales));
  el.innerHTML = rows.map((r, i) =>
    `<div class="hbar" data-tip="${esc(r.tip)}"><div class="name" title="${esc(r.name)}"><span class="rank">${i + startRank}</span>${esc(r.name)}</div>` +
    `<div class="track"><div class="fill" style="width:${(r.sales / max * 100).toFixed(2)}%"></div></div><div class="val">${gbp(r.sales, 0)}</div></div>`).join("");
}
hbars($("products"), DATA.products.map((p) => ({ ...p,
  tip: `<b>${p.name}</b>รหัส ${p.code}<br>ยอดขาย ${gbp(p.sales)}<br>ขายได้ ${num(p.qty)} ชิ้น · ${(p.sales / k.sales * 100).toFixed(2)}% ของยอดขายรวม` })), 1);

function drawCountries(excludeTop) {
  const all = DATA.countries, top = all[0];
  const rows = (excludeTop ? all.slice(1, 11) : all.slice(0, 10)).map((c) => ({ ...c,
    tip: `<b>${c.name}</b>ยอดขาย ${gbp(c.sales)}<br>${num(c.orders)} ออร์เดอร์ · ${(c.sales / k.sales * 100).toFixed(2)}% ของยอดขายรวม` }));
  hbars($("countries"), rows, excludeTop ? 2 : 1);
  $("countryDesc").textContent = excludeTop
    ? `อันดับ 2–11 (ไม่รวม ${top.name}) · ยอดขาย (£)`
    : `${top.name} คิดเป็น ${(top.sales / k.sales * 100).toFixed(1)}% ของยอดขายรวม · ยอดขาย (£)`;
  $("btnAll").setAttribute("aria-pressed", String(!excludeTop));
  $("btnEx").setAttribute("aria-pressed", String(excludeTop));
}
$("btnAll").onclick = () => drawCountries(false);
$("btnEx").onclick = () => drawCountries(true);
drawCountries(false);

// ---------- Heatmap วัน × ชั่วโมง ----------
(() => {
  const RAMP = ["#DBEAFE", "#BFDBFE", "#93C5FD", "#60A5FA", "#3B82F6", "#2563EB", "#1D4ED8", "#1E3A8A"];
  const [h0, h1] = DATA.hours, hours = [];
  for (let h = h0; h <= h1; h++) hours.push(h);
  const map = new Map(DATA.heat.map((c) => [c.d + "-" + c.h, c]));
  const max = Math.max(...DATA.heat.map((c) => c.sales));
  const el = $("heat");
  el.style.gridTemplateColumns = `auto repeat(${hours.length}, minmax(0, 1fr))`;
  let s = "<div></div>" + hours.map((h) => `<div class="hl">${h}</div>`).join("");
  DATA.weekdays.forEach((name, i) => {
    s += `<div class="dl">${name}</div>`;
    hours.forEach((h) => {
      const c = map.get(i + 1 + "-" + h);
      if (!c) { s += `<div class="cell" data-tip="${esc(`<b>วัน${name} ${h}:00–${h}:59</b>ไม่มียอดขาย`)}"></div>`; return; }
      const idx = Math.min(RAMP.length - 1, Math.floor(c.sales / max * RAMP.length));
      s += `<div class="cell" style="background:${RAMP[idx]}" data-tip="${esc(`<b>วัน${name} ${h}:00–${h}:59</b>ยอดขาย ${gbp(c.sales)}<br>${num(c.orders)} ออร์เดอร์`)}"></div>`;
    });
  });
  el.innerHTML = s;
  $("ramp").innerHTML = RAMP.map((c) => `<i style="background:${c}"></i>`).join("");
  const best = DATA.heat.reduce((a, b) => (b.sales > a.sales ? b : a));
  const byDay = DATA.weekdays.map((n, i) => ({ n, v: DATA.heat.filter((c) => c.d === i + 1).reduce((t, c) => t + c.sales, 0) }));
  const bd = byDay.reduce((a, b) => (b.v > a.v ? b : a)), none = byDay.filter((d) => d.v === 0).map((d) => "วัน" + d.n);
  $("insightHeat").innerHTML = `ช่วงที่ขายดีที่สุดคือ <b>วัน${DATA.weekdays[best.d - 1]} ${best.h}:00–${best.h}:59</b> (${gbp(best.sales)}) · วันที่ยอดขายรวมสูงสุดคือ <b>วัน${bd.n}</b>` +
    (none.length ? ` · ไม่มียอดขายใน${none.join(", ")}` : "");
})();

// ---------- Tooltip ----------
(() => {
  const tip = $("tip");
  const show = (e) => {
    const t = e.target.closest("[data-tip]");
    if (!t) { tip.style.opacity = 0; return; }
    tip.innerHTML = t.dataset.tip;
    const pad = 12, w = tip.offsetWidth, h = tip.offsetHeight;
    let x = e.clientX + pad, y = e.clientY + pad;
    if (x + w > innerWidth - 8) x = Math.max(8, e.clientX - w - pad);
    if (y + h > innerHeight - 8) y = Math.max(8, e.clientY - h - pad);
    tip.style.left = x + "px"; tip.style.top = y + "px"; tip.style.opacity = 1;
  };
  document.addEventListener("pointermove", show);
  document.addEventListener("pointerdown", show);
  document.addEventListener("scroll", () => (tip.style.opacity = 0), { passive: true });
})();

drawMonthly();
let rt; addEventListener("resize", () => { clearTimeout(rt); rt = setTimeout(drawMonthly, 100); });
if (document.fonts && document.fonts.ready) document.fonts.ready.then(drawMonthly);
</script>
</body>
</html>
"""


def main() -> None:
    data = build_data()
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    OUT.write_text(TEMPLATE.replace("__DATA__", payload), encoding="utf-8")
    (ROOT / "reports" / "dashboard_data.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"saved {OUT.name} ({OUT.stat().st_size / 1e3:.0f} KB)")
    print(json.dumps(data["kpi"], ensure_ascii=False))


if __name__ == "__main__":
    main()
