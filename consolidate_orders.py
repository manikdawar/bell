"""
Consolidates Amazon / Flipkart / Meesho order-report xlsx files in this folder
into Consolidated_Orders.xlsx (sheets: Combined, Amazon, Flipkart, Meesho).

Incremental: if Consolidated_Orders.xlsx already exists, only report files not
already recorded in it (tracked by filename) are parsed and appended. If it
doesn't exist yet, it is built from scratch from every matching file here.

Usage: drop new monthly report files into this folder (any filename containing
"amazon", "flipkart", or "meesho", case-insensitive), then run:
    python3 consolidate_orders.py

Requires: pandas, openpyxl
"""
import glob
import os

import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(BASE, "Consolidated_Orders.xlsx")


def find_files(keyword):
    files = glob.glob(os.path.join(BASE, "*.xlsx"))
    return sorted(
        f for f in files
        if keyword in os.path.basename(f).lower()
        and os.path.basename(f) != os.path.basename(OUT_PATH)
        and not os.path.basename(f).startswith("~$")
    )


def clean_str(x):
    if isinstance(x, str):
        return x.strip().strip('"')
    return x


def month_label(dates):
    dates = pd.to_datetime(dates, errors="coerce", utc=True).dt.tz_localize(None)
    dates = dates.dropna()
    if dates.empty:
        return "Unknown"
    return dates.dt.strftime("%b %Y").mode()[0]


def load_existing_sheet(sheet_name):
    if not os.path.exists(OUT_PATH):
        return pd.DataFrame()
    try:
        return pd.read_excel(OUT_PATH, sheet_name=sheet_name)
    except ValueError:
        return pd.DataFrame()


def new_files_only(files, existing_df):
    if existing_df.empty or "source_file" not in existing_df.columns:
        return files
    already = set(existing_df["source_file"].dropna())
    return [f for f in files if os.path.basename(f) not in already]


def parse_amazon(fname):
    df = pd.read_excel(fname, sheet_name=0)
    df.columns = [c.strip() for c in df.columns]
    df.insert(0, "report_month", month_label(df["purchase-date"]))
    df.insert(1, "source_file", os.path.basename(fname))
    return df


def parse_flipkart(fname):
    df = pd.read_excel(fname, sheet_name="Orders")
    df.columns = [c.strip() for c in df.columns]
    for col in ("sku", "fsn", "product_title"):
        if col in df.columns:
            df[col] = df[col].apply(clean_str).str.replace(r"^SKU:", "", regex=True)
    df.insert(0, "report_month", month_label(df["order_date"]))
    df.insert(1, "source_file", os.path.basename(fname))
    return df


def parse_meesho(fname):
    df = pd.read_excel(fname, sheet_name=0)
    df.columns = [c.strip() for c in df.columns]
    df.insert(0, "report_month", month_label(df["order_date"]))
    df.insert(1, "source_file", os.path.basename(fname))
    return df


def build_platform(keyword, parse_fn, sheet_name):
    existing = load_existing_sheet(sheet_name)
    new_files = new_files_only(find_files(keyword), existing)
    new_frames = [parse_fn(f) for f in new_files]
    if new_files:
        print(f"{sheet_name}: adding {len(new_files)} new file(s) -> {[os.path.basename(f) for f in new_files]}")
    else:
        print(f"{sheet_name}: no new files")
    frames = [existing] + new_frames if not existing.empty else new_frames
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


amazon_all = build_platform("amazon", parse_amazon, "Amazon")
flipkart_all = build_platform("flipkart", parse_flipkart, "Flipkart")
meesho_all = build_platform("meesho", parse_meesho, "Meesho")


# ---------- COMBINED NORMALIZED VIEW ----------
def norm_amazon(df):
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame({
        "platform": "Amazon",
        "report_month": df["report_month"],
        "order_id": df["amazon-order-id"],
        "order_date": pd.to_datetime(df["purchase-date"], errors="coerce", utc=True).dt.tz_localize(None),
        "product_name": df["product-name"],
        "sku": df["sku"],
        "quantity": df["quantity"],
        "amount": df["item-price"],
        "tax": df["item-tax"],
        "status": df["order-status"],
        "city": df["ship-city"],
        "state": df["ship-state"],
        "source_file": df["source_file"],
    })


def norm_flipkart(df):
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame({
        "platform": "Flipkart",
        "report_month": df["report_month"],
        "order_id": df["order_id"],
        "order_date": pd.to_datetime(df["order_date"], errors="coerce"),
        "product_name": df["product_title"],
        "sku": df["sku"],
        "quantity": df["quantity"],
        "amount": pd.NA,
        "tax": pd.NA,
        "status": df["order_item_status"],
        "city": pd.NA,
        "state": pd.NA,
        "source_file": df["source_file"],
    })


def norm_meesho(df):
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame({
        "platform": "Meesho",
        "report_month": df["report_month"],
        "order_id": df["sub_order_num"],
        "order_date": pd.to_datetime(df["order_date"], errors="coerce"),
        "product_name": pd.NA,
        "sku": pd.NA,
        "quantity": df["quantity"],
        "amount": df["total_invoice_value"],
        "tax": df["tax_amount"],
        "status": df["transaction_type"],
        "city": pd.NA,
        "state": df["end_customer_state_new"],
        "source_file": df["source_file"],
    })


combined = pd.concat(
    [norm_amazon(amazon_all), norm_flipkart(flipkart_all), norm_meesho(meesho_all)],
    ignore_index=True,
)

# ---------- WRITE OUTPUT ----------
with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
    combined.to_excel(writer, sheet_name="Combined", index=False)
    amazon_all.to_excel(writer, sheet_name="Amazon", index=False)
    flipkart_all.to_excel(writer, sheet_name="Flipkart", index=False)
    meesho_all.to_excel(writer, sheet_name="Meesho", index=False)

print(f"Combined: {len(combined)} rows | Amazon: {len(amazon_all)} | Flipkart: {len(flipkart_all)} | Meesho: {len(meesho_all)}")
print(f"Saved to {OUT_PATH}")
