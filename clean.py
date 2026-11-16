"""ทำความสะอาดข้อมูลยอดขาย — Online Retail II ชีตปี 2010-2011

รัน:  .venv/bin/python clean.py

หลักการ: ไม่ลบข้อมูลทิ้งเงียบ ๆ ทุกแถวของไฟล์ต้นฉบับจะไปอยู่ใน 1 ใน 3 ตารางเสมอ
  1) sales_clean.csv            ยอดขายจริง (ใช้ทำรายงาน/แดชบอร์ด)
  2) data/clean/returns.csv     ใบยกเลิก/คืนสินค้า (Invoice ขึ้นต้นด้วย C)
  3) data/clean/excluded.csv    แถวที่คัดออก พร้อมคอลัมน์ Reason บอกเหตุผล
และสรุปตัวเลขไว้ที่ reports/clean_summary.json

หน่วยเงินทั้งหมดเป็น GBP (£)
"""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW_XLSX = ROOT / "data" / "raw" / "online_retail_II.xlsx"
CACHE = ROOT / "data" / "interim" / "raw_2010_2011.pkl"
SHEET = "Year 2010-2011"

OUT_SALES = ROOT / "sales_clean.csv"
OUT_RETURNS = ROOT / "data" / "clean" / "returns.csv"
OUT_EXCLUDED = ROOT / "data" / "clean" / "excluded.csv"
OUT_SUMMARY = ROOT / "reports" / "clean_summary.json"

# รหัสสินค้าจริง = ตัวเลข 5 หลัก (+ตัวอักษรต่อท้าย) หรือรหัส DCGS ของร้าน
# ที่เหลือ (POST, DOT, M, D, AMAZONFEE, BANK CHARGES, บัตรของขวัญ ฯลฯ) ถือว่า "ไม่ใช่สินค้า"
PRODUCT_CODE = r"^(?:\d{5}[A-Z]*|DCGS\w+)$"
# รหัสที่หน้าตาเหมือนสินค้า แต่ดูจากชื่อแล้วเป็นค่าบริการ/บัตรของขวัญ (เจอจากการตรวจข้อมูล)
NON_PRODUCT_EXTRA = {"23444", "23574", "22016"}  # Next Day Carriage, PACKING CHARGE, Gift Voucher £100

# เหตุผลที่คัดออก (ตรวจตามลำดับ แถวหนึ่งได้เหตุผลแรกที่เข้าเงื่อนไข)
R_DUP = "แถวซ้ำ (รายการเดียวกันในบิลเดียวกัน เก็บแถวแรกไว้)"
R_ADJ = "ปรับหนี้เสีย (Invoice ขึ้นต้น A ไม่ใช่การขาย)"
R_NONPROD = "ไม่ใช่สินค้า (ค่าส่ง/ค่าธรรมเนียม/ส่วนลด/ปรับยอด/บัตรของขวัญ)"
R_STOCK = "ปรับสต็อก (จำนวนติดลบ แต่ไม่ใช่ใบยกเลิก)"
R_ZERO = "ราคาเป็น 0 (ของแถม/รายการภายใน)"
R_VOID = "สั่งแล้วยกเลิกทั้งรายการในวันเดียวกัน (ออร์เดอร์ใหญ่ผิดปกติ)"
# ออร์เดอร์ที่ถูกยกเลิกทั้งรายการในวันเดียวกัน และมูลค่าตั้งแต่เท่านี้ขึ้นไป ถือว่าไม่ใช่การขายจริง
# (ในข้อมูลนี้มี 2 รายการ: £168,469.60 และ £77,183.60 รายการถัดไปมูลค่าเพียง £1,416)
VOID_MIN_GBP = 10_000

COUNTRY_FIX = {"EIRE": "Ireland", "RSA": "South Africa", "USA": "United States"}
WEEKDAY_TH = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]


def load_raw() -> pd.DataFrame:
    """อ่านชีต 2010-2011 (ครั้งแรกช้า ~1-2 นาที จึงเก็บ cache ไว้)"""
    if CACHE.exists():
        return pd.read_pickle(CACHE)
    df = pd.read_excel(RAW_XLSX, sheet_name=SHEET, dtype={"Invoice": str, "StockCode": str})
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    df.to_pickle(CACHE)
    return df


def standardize(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """จัดรูปแบบข้อความให้เป็นมาตรฐาน (ไม่เปลี่ยนจำนวนแถว)"""
    out = df.copy()
    out["Invoice"] = out["Invoice"].astype(str).str.strip().str.upper()

    code = out["StockCode"].astype(str).str.strip()
    is_product_raw = code.str.upper().str.match(PRODUCT_CODE)
    # รหัสสินค้าทำเป็นตัวพิมพ์ใหญ่ (85123a -> 85123A) ส่วนรหัสที่ไม่ใช่สินค้าคงเดิม
    out["StockCode"] = code.where(~is_product_raw, code.str.upper())

    desc = out["Description"].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
    stats = {
        "stockcode_uppercased_rows": int((is_product_raw & (code != code.str.upper())).sum()),
        "description_whitespace_fixed_rows": int((desc != out["Description"].astype("string")).fillna(False).sum()),
    }
    out["Description"] = desc
    out["Country"] = out["Country"].astype(str).str.strip().replace(COUNTRY_FIX)
    stats["country_renamed_rows"] = int(df["Country"].isin(COUNTRY_FIX.keys()).sum())

    cid = out["Customer ID"]
    out["CustomerID"] = cid.astype("Int64").astype("string").fillna("ไม่ระบุ")
    out = out.drop(columns=["Customer ID"])
    return out, stats


def classify(df: pd.DataFrame) -> pd.Series:
    """คืนค่า 'sale' / 'return' / เหตุผลที่คัดออก ของแต่ละแถว"""
    inv, qty, price = df["Invoice"], df["Quantity"], df["Price"]
    # แถวซ้ำ: เทียบหลังจัดรูปแบบแล้ว และไม่เทียบชื่อสินค้า
    # (รหัสเดียวกันในบิลเดียวกัน เวลา/จำนวน/ราคาเท่ากัน แต่สะกดชื่อต่างกัน ก็นับว่าซ้ำ)
    key = ["Invoice", "StockCode", "Quantity", "InvoiceDate", "Price", "CustomerID", "Country"]
    is_dup = df.duplicated(subset=key, keep="first")
    is_product = df["StockCode"].str.match(PRODUCT_CODE) & ~df["StockCode"].isin(NON_PRODUCT_EXTRA)
    is_cancel = inv.str.startswith("C")

    status = pd.Series("sale", index=df.index, dtype="object")
    # ใส่จากลำดับท้ายไปหน้า เพื่อให้เหตุผลลำดับแรกทับเป็นตัวสุดท้าย
    status[price == 0] = R_ZERO
    status[(qty < 0) & ~is_cancel] = R_STOCK
    status[is_cancel] = "return"
    status[~is_product] = R_NONPROD
    status[inv.str.match(r"^[A-Z]") & ~is_cancel] = R_ADJ
    status[is_dup] = R_DUP
    return status


def mark_voided(df: pd.DataFrame, status: pd.Series) -> pd.Series:
    """หาออร์เดอร์ใหญ่ที่สั่งแล้วยกเลิกทั้งรายการในวันเดียวกัน แล้วย้ายทั้งรายการขายและใบยกเลิกไปตารางคัดออก

    จับคู่เมื่อ ลูกค้า + รหัสสินค้า + ราคา + วันที่ ตรงกัน และจำนวนที่ยกเลิกเท่ากับจำนวนที่สั่ง
    """
    d = df.assign(day=df["InvoiceDate"].dt.date, absq=df["Quantity"].abs(), idx=df.index)
    d = d[d["CustomerID"] != "ไม่ระบุ"]
    key = ["CustomerID", "StockCode", "Price", "day", "absq"]
    s = d[status[d.index] == "sale"]
    r = d[status[d.index] == "return"]
    s = s[s["absq"] * s["Price"] >= VOID_MIN_GBP]
    s = s.assign(n=s.groupby(key).cumcount())
    r = r.assign(n=r.groupby(key).cumcount())
    m = s.merge(r, on=key + ["n"], suffixes=("_s", "_r"))
    m = m[m["InvoiceDate_s"] <= m["InvoiceDate_r"]]
    out = status.copy()
    out[m["idx_s"]] = R_VOID
    out[m["idx_r"]] = R_VOID
    return out


def unify_descriptions(df: pd.DataFrame, status: pd.Series) -> tuple[pd.DataFrame, int]:
    """รหัสสินค้าเดียวกันให้ใช้ชื่อเดียว = ชื่อที่พบบ่อยที่สุดในยอดขายจริง"""
    sales = df[(status == "sale") & df["Description"].notna()]
    canonical = (
        sales.groupby(["StockCode", "Description"]).size().reset_index(name="n")
        .sort_values(["StockCode", "n", "Description"], ascending=[True, False, True])
        .drop_duplicates("StockCode").set_index("StockCode")["Description"]
    )
    new_desc = df["StockCode"].map(canonical).astype("string")
    target = status.isin(["sale", "return"])  # ไม่แตะบันทึกในแถวที่คัดออก (เช่น "damaged")
    final = df["Description"].where(~target | new_desc.isna(), new_desc)
    changed = int((target & (final != df["Description"]).fillna(True)).sum())
    out = df.copy()
    out["Description"] = final
    return out, changed


def add_columns(df: pd.DataFrame) -> pd.DataFrame:
    """เพิ่มคอลัมน์ที่เจ้าของร้านใช้จริง"""
    out = df.copy()
    dt = out["InvoiceDate"]
    out["Sales_GBP"] = (out["Quantity"] * out["Price"]).round(2)
    out["Date"] = dt.dt.strftime("%Y-%m-%d")
    out["Month"] = dt.dt.strftime("%Y-%m")
    out["WeekdayNum"] = dt.dt.day_of_week + 1  # 1=จันทร์ ... 7=อาทิตย์
    out["Weekday"] = dt.dt.day_of_week.map(lambda i: WEEKDAY_TH[i])
    out["Hour"] = dt.dt.hour
    out["InvoiceDate"] = dt.dt.strftime("%Y-%m-%d %H:%M:%S")
    cols = ["Invoice", "InvoiceDate", "Date", "Month", "WeekdayNum", "Weekday", "Hour",
            "StockCode", "Description", "Quantity", "Price", "Sales_GBP", "CustomerID", "Country"]
    return out[cols]


def main() -> None:
    raw = load_raw()
    df, std_stats = standardize(raw)
    status = classify(df)
    status = mark_voided(df, status)
    df, desc_changed = unify_descriptions(df, status)
    full = add_columns(df)

    sales = full[status == "sale"]
    returns = full[status == "return"].rename(columns={"Sales_GBP": "Return_GBP"})
    excluded = full[~status.isin(["sale", "return"])].rename(columns={"Sales_GBP": "Amount_GBP"})
    excluded.insert(0, "Reason", status[excluded.index])

    assert len(sales) + len(returns) + len(excluded) == len(raw), "จำนวนแถวไม่ครบ"

    OUT_RETURNS.parent.mkdir(parents=True, exist_ok=True)
    OUT_SUMMARY.parent.mkdir(parents=True, exist_ok=True)
    # utf-8-sig เพื่อให้ Excel แสดงภาษาไทยถูกต้อง
    sales.to_csv(OUT_SALES, index=False, encoding="utf-8-sig")
    returns.to_csv(OUT_RETURNS, index=False, encoding="utf-8-sig")
    excluded.to_csv(OUT_EXCLUDED, index=False, encoding="utf-8-sig")

    total_sales = float(sales["Sales_GBP"].sum())
    total_returns = float(returns["Return_GBP"].sum())  # ค่าติดลบ
    n_orders = int(sales["Invoice"].nunique())
    excl = excluded.groupby("Reason").agg(rows=("Reason", "size"), amount_gbp=("Amount_GBP", "sum"))

    summary = {
        "currency": "GBP",
        "raw_rows": int(len(raw)),
        "tables": {"sales_rows": int(len(sales)), "returns_rows": int(len(returns)),
                   "excluded_rows": int(len(excluded))},
        "excluded_by_reason": {k: {"rows": int(v.rows), "amount_gbp": round(float(v.amount_gbp), 2)}
                               for k, v in excl.sort_values("rows", ascending=False).iterrows()},
        "fixes": {**std_stats, "description_unified_rows": desc_changed,
                  "customer_id_missing_in_sales": int((sales["CustomerID"] == "ไม่ระบุ").sum())},
        "kpi": {
            "total_sales_gbp": round(total_sales, 2),
            "orders": n_orders,
            "avg_order_value_gbp": round(total_sales / n_orders, 2),
            "units_sold": int(sales["Quantity"].sum()),
            "returns_gbp": round(total_returns, 2),
            "return_invoices": int(returns["Invoice"].nunique()),
            "return_rate_pct": round(-total_returns / total_sales * 100, 2),
            "net_sales_gbp": round(total_sales + total_returns, 2),
            "customers_identified": int(sales.loc[sales["CustomerID"] != "ไม่ระบุ", "CustomerID"].nunique()),
            "products": int(sales["StockCode"].nunique()),
            "countries": int(sales["Country"].nunique()),
            "date_min": sales["Date"].min(), "date_max": sales["Date"].max(),
        },
        "monthly_sales_gbp": {k: round(float(v), 2) for k, v in sales.groupby("Month")["Sales_GBP"].sum().items()},
    }
    OUT_SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
