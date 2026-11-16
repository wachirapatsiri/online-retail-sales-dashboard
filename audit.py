"""ตรวจสุขภาพข้อมูล (Data Audit) — Online Retail II ชีตปี 2010-2011

รัน:  .venv/bin/python audit.py
ผลลัพธ์: reports/audit_summary.json  (ตัวเลขทุกตัวนับจากข้อมูลจริง)

สคริปต์นี้ "นับปัญหา" อย่างเดียว ไม่แก้ไขข้อมูล
"""
import json
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW_XLSX = ROOT / "data" / "raw" / "online_retail_II.xlsx"
CACHE = ROOT / "data" / "interim" / "raw_2010_2011.pkl"
OUT = ROOT / "reports" / "audit_summary.json"
SHEET = "Year 2010-2011"

# รหัสสินค้าปกติ = ตัวเลข 5 หลัก ตามด้วยตัวอักษรได้ (เช่น 85123A)
PRODUCT_CODE = re.compile(r"^\d{5}[A-Za-z]*$")


def load_raw() -> pd.DataFrame:
    """อ่านชีต 2010-2011 (ครั้งแรกช้า ~1-2 นาที จึงเก็บ cache ไว้)"""
    if CACHE.exists():
        return pd.read_pickle(CACHE)
    df = pd.read_excel(RAW_XLSX, sheet_name=SHEET, dtype={"Invoice": str, "StockCode": str})
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    df.to_pickle(CACHE)
    return df


def main() -> None:
    df = load_raw()
    n = len(df)
    inv = df["Invoice"].astype(str).str.strip()
    code = df["StockCode"].astype(str).str.strip()
    desc = df["Description"]
    qty, price = df["Quantity"], df["Price"]
    amount = qty * price

    is_cancel = inv.str.upper().str.startswith("C")
    is_other_prefix = inv.str.match(r"^[A-Za-z]") & ~is_cancel
    is_non_product = ~code.str.match(PRODUCT_CODE)
    desc_str = desc.astype("string")

    # รหัสสินค้าเดียวกันแต่มี Description หลายแบบ
    desc_per_code = (
        df.loc[desc.notna() & ~is_non_product]
        .assign(code=code.str.upper(), d=desc_str.str.strip())
        .groupby("code")["d"].nunique()
    )

    non_product_table = (
        df.loc[is_non_product]
        .assign(code=code, amount=amount)
        .groupby("code")
        .agg(rows=("code", "size"), example=("Description", "first"), amount_gbp=("amount", "sum"))
        .sort_values("rows", ascending=False)
    )

    sale_like = ~is_cancel & ~is_other_prefix & (qty > 0) & (price > 0) & ~is_non_product

    summary = {
        "source": {"file": RAW_XLSX.name, "sheet": SHEET, "currency": "GBP"},
        "shape": {"rows": n, "columns": list(df.columns)},
        "date_range": {"min": str(df["InvoiceDate"].min()), "max": str(df["InvoiceDate"].max())},
        "unique": {
            "invoices": int(inv.nunique()),
            "stock_codes": int(code.nunique()),
            "customers": int(df["Customer ID"].nunique()),
            "countries": int(df["Country"].nunique()),
        },
        "issues": {
            "duplicate_rows_extra": int(df.duplicated().sum()),
            "duplicate_rows_all_involved": int(df.duplicated(keep=False).sum()),
            "customer_id_missing": int(df["Customer ID"].isna().sum()),
            "description_missing": int(desc.isna().sum()),
            "description_leading_trailing_space": int((desc_str != desc_str.str.strip()).fillna(False).sum()),
            "description_not_uppercase": int((desc_str != desc_str.str.upper()).fillna(False).sum()),
            "stock_codes_with_multiple_descriptions": int((desc_per_code > 1).sum()),
            "stock_code_lowercase_variant_rows": int((code != code.str.upper()).sum()),
            "quantity_negative": int((qty < 0).sum()),
            "quantity_zero": int((qty == 0).sum()),
            "cancelled_rows": int(is_cancel.sum()),
            "cancelled_invoices": int(inv[is_cancel].nunique()),
            "other_letter_prefix_invoice_rows": int(is_other_prefix.sum()),
            "negative_qty_not_cancelled": int(((qty < 0) & ~is_cancel).sum()),
            "cancelled_but_positive_qty": int((is_cancel & (qty > 0)).sum()),
            "price_zero": int((price == 0).sum()),
            "price_negative": int((price < 0).sum()),
            "non_product_rows": int(is_non_product.sum()),
            "non_product_codes": int(code[is_non_product].nunique()),
            "country_unspecified_rows": int((df["Country"] == "Unspecified").sum()),
        },
        "non_product_codes": [
            {"code": c, "rows": int(r.rows), "example": None if pd.isna(r.example) else str(r.example),
             "amount_gbp": round(float(r.amount_gbp), 2)}
            for c, r in non_product_table.iterrows()
        ],
        "other_prefix_invoices": df.loc[is_other_prefix, ["Invoice", "StockCode", "Description", "Quantity", "Price"]]
        .astype(str).to_dict("records"),
        "extremes": {
            "quantity_max": int(qty.max()), "quantity_min": int(qty.min()),
            "price_max": float(price.max()), "price_min": float(price.min()),
        },
        "amount_gbp": {
            "all_rows_sum": round(float(amount.sum()), 2),
            "cancelled_rows_sum": round(float(amount[is_cancel].sum()), 2),
            "non_product_rows_sum": round(float(amount[is_non_product].sum()), 2),
            "preliminary_clean_sales_sum": round(float(amount[sale_like].sum()), 2),
        },
        "preliminary_clean_sales": {
            "rows": int(sale_like.sum()),
            "invoices": int(inv[sale_like].nunique()),
            "rows_without_customer_id": int((sale_like & df["Customer ID"].isna()).sum()),
            "duplicate_rows_extra": int(df[sale_like].duplicated().sum()),
        },
        "rows_per_month": {str(k): int(v) for k, v in df.groupby(df["InvoiceDate"].dt.to_period("M")).size().items()},
        "top_countries_rows": {k: int(v) for k, v in df["Country"].value_counts().head(8).items()},
        "sample_lowercase_descriptions": {
            k: int(v) for k, v in desc_str[(desc_str != desc_str.str.upper()).fillna(False)].value_counts().head(25).items()
        },
    }

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
