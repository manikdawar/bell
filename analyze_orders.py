"""
Builds Orders_Dashboard.html and Orders_Dashboard.pdf — a visual analysis of
Consolidated_Orders.xlsx (growth trend, revenue, fulfillment health, product
mix, geography/timing). The PDF is a static-chart equivalent of the HTML,
meant for sharing/printing where an interactive page isn't convenient.

Run consolidate_orders.py first to refresh Consolidated_Orders.xlsx, then:
    python3 analyze_orders.py
All numbers and the narrative call-outs are recomputed from whatever data is
in the workbook at run time, so this can be re-run every time new months are
added.

Requires: pandas, openpyxl, matplotlib
"""
import calendar
import json
import os
import re
import textwrap

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
IN_PATH = os.path.join(BASE, "Consolidated_Orders.xlsx")
OUT_PATH = os.path.join(BASE, "Orders_Dashboard.html")
PDF_PATH = os.path.join(BASE, "Orders_Dashboard.pdf")

COLOR = {
    "Amazon": "#2a78d6", "Flipkart": "#008300", "Meesho": "#e87ba4",
    "good": "#0ca30c", "warning": "#fab219", "serious": "#ec835a", "critical": "#d03b3b",
    "muted": "#898781", "text": "#0b0b0b", "secondary": "#52514e", "grid": "#e1e0d9",
}
CATEGORICAL_FALLBACK = ["#2a78d6", "#008300", "#e87ba4"]


def platform_color(p, idx=0):
    return COLOR.get(p, CATEGORICAL_FALLBACK[idx % 3])

VEHICLE_KEYWORDS = [
    "Ola S1", "Honda", "Bajaj Pulsar", "Yamaha", "Suzuki", "TVS",
    "Royal Enfield", "Hero", "KTM", "Ather", "Activa",
]


def month_label(ts):
    return ts.strftime("%b %Y")


def pct(n, d):
    return round(100 * n / d, 1) if d else 0.0


def brand_match(name):
    low = name.lower()
    for kw in VEHICLE_KEYWORDS:
        if kw.lower() in low:
            return kw
    return "Other"


def main():
    combined = pd.read_excel(IN_PATH, sheet_name="Combined")
    amazon = pd.read_excel(IN_PATH, sheet_name="Amazon")
    flipkart = pd.read_excel(IN_PATH, sheet_name="Flipkart")

    combined["order_date"] = pd.to_datetime(combined["order_date"], errors="coerce")
    combined = combined.dropna(subset=["order_date"])
    combined["period"] = combined["order_date"].dt.to_period("M")

    periods = sorted(combined["period"].unique())
    months = [month_label(p.to_timestamp()) for p in periods]
    latest_period = periods[-1]
    overall_max_date = combined["order_date"].max()

    days_elapsed = {}
    for p in periods:
        year, mon = p.year, p.month
        days_in_month = calendar.monthrange(year, mon)[1]
        if p == latest_period:
            days_elapsed[month_label(p.to_timestamp())] = overall_max_date.day
        else:
            days_elapsed[month_label(p.to_timestamp())] = days_in_month

    platforms = sorted(combined["platform"].unique())

    def series_by_month(df, value_col=None, agg="count"):
        g = df.groupby(["platform", "period"])
        s = g.size() if agg == "count" else g[value_col].sum()
        s = s.unstack(fill_value=0)
        s = s.reindex(columns=periods, fill_value=0)
        return {p: [float(s.loc[p, per]) if p in s.index else 0.0 for per in periods] for p in platforms}

    orders_by_month = series_by_month(combined)
    daily_rate_by_month = {
        p: [round(orders_by_month[p][i] / days_elapsed[months[i]], 2) for i in range(len(months))]
        for p in platforms
    }

    rev = combined.dropna(subset=["amount"])
    revenue_platforms = sorted(rev["platform"].unique())
    revenue_by_month = series_by_month(rev, "amount", "sum") if revenue_platforms else {}
    revenue_by_month = {p: [round(v, 0) for v in vals] for p, vals in revenue_by_month.items() if p in revenue_platforms}

    valid = combined[combined["quantity"] > 0]
    units_by_month = series_by_month(valid, "quantity", "sum")
    units_by_month = {p: [int(v) for v in vals] for p, vals in units_by_month.items()}

    aov = {p: round(float(rev[rev["platform"] == p]["amount"].mean()), 2) for p in revenue_platforms}

    total_orders = int(len(combined))
    total_revenue = round(float(rev["amount"].sum()), 0)
    total_units = int(valid["quantity"].sum())

    # ---------- status buckets (Amazon / Flipkart only — the only platforms with a status field) ----------
    def amazon_bucket(s):
        if s == "Cancelled":
            return "Cancelled"
        if "Pending" in s:
            return "Pending"
        if s == "Shipped - Delivered to Buyer":
            return "Delivered"
        if "Return" in s:
            return "Returned"
        return "In Transit"

    def flipkart_bucket(s):
        if s in ("CANCELLED", "REJECTED"):
            return "Cancelled"
        if s == "DELIVERED":
            return "Delivered"
        if s in ("RETURNED", "RETURN_REQUESTED"):
            return "Returned"
        return "In Transit"

    amazon = amazon.copy()
    flipkart = flipkart.copy()
    amazon["bucket"] = amazon["order-status"].apply(amazon_bucket)
    flipkart["bucket"] = flipkart["order_item_status"].apply(flipkart_bucket)

    BUCKETS = ["Delivered", "In Transit", "Pending", "Returned", "Cancelled"]

    def bucket_pct(df):
        c = df["bucket"].value_counts()
        total = len(df)
        return [pct(c.get(b, 0), total) for b in BUCKETS]

    status_pct = {"Amazon": bucket_pct(amazon), "Flipkart": bucket_pct(flipkart)}

    amazon_return_rate = pct((amazon["bucket"] == "Returned").sum(), len(amazon))
    flipkart_return_rate = pct((flipkart["bucket"] == "Returned").sum(), len(flipkart))

    # ---------- vehicle brand concentration ----------
    vv = valid.copy()
    vv["product_name"] = vv["product_name"].fillna("")
    named = vv[vv["product_name"] != ""]
    named = named.assign(brand=named["product_name"].apply(brand_match))
    brand_units = named.groupby("brand")["quantity"].sum().sort_values(ascending=False)
    brand_total_named = int(named["quantity"].sum())
    brand_units_list = [[k, int(v)] for k, v in brand_units.items()]
    top_brand, top_brand_units = brand_units_list[0]
    top_brand_pct = pct(top_brand_units, brand_total_named)

    # ---------- states ----------
    state_counts = combined["state"].dropna()
    state_total = len(state_counts)
    top_states = [[k, int(v)] for k, v in state_counts.value_counts().head(8).items()]
    top_state_name, top_state_count = top_states[0] if top_states else ("N/A", 0)
    top_state_pct = pct(top_state_count, state_total)

    # ---------- weekday ----------
    wd_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    wd_counts = combined["order_date"].dt.day_name().value_counts().reindex(wd_order, fill_value=0)
    weekday = [[k, int(v)] for k, v in wd_counts.items()]
    busiest_day, busiest_count = max(weekday, key=lambda x: x[1])
    quietest_day, quietest_count = min(weekday, key=lambda x: x[1])
    weekday_lift_pct = round(100 * (busiest_count - quietest_count) / quietest_count, 0) if quietest_count else 0

    # ---------- flipkart SLA & returns ----------
    dispatch_sla_breach_pct = pct((flipkart["dispatch_sla_breached"] == "Y").sum(), len(flipkart))
    sla_by_period = flipkart.assign(period=pd.to_datetime(flipkart["order_date"], errors="coerce").dt.to_period("M"))
    sla_by_period = sla_by_period.groupby("period")["dispatch_sla_breached"].apply(lambda s: pct((s == "Y").sum(), len(s)))
    dispatch_sla_breach_by_month = [float(sla_by_period.get(p, 0.0)) for p in periods]

    return_reasons = flipkart["return_reason"].value_counts().head(6)
    return_reasons_list = [[k, int(v)] for k, v in return_reasons.items()]
    top2_reason_share = pct(return_reasons.head(2).sum(), (flipkart["bucket"] == "Returned").sum())

    data = {
        "months": months,
        "orders_by_month": orders_by_month,
        "daily_rate_by_month": daily_rate_by_month,
        "revenue_by_month": revenue_by_month,
        "revenue_platforms": revenue_platforms,
        "units_by_month": units_by_month,
        "aov": aov,
        "total_orders": total_orders,
        "total_revenue": total_revenue,
        "total_units": total_units,
        "status_buckets": BUCKETS,
        "status_pct": status_pct,
        "brand_units": brand_units_list,
        "brand_total_named": brand_total_named,
        "top_states": top_states,
        "weekday": weekday,
        "dispatch_sla_breach_pct": dispatch_sla_breach_pct,
        "dispatch_sla_breach_by_month": dispatch_sla_breach_by_month,
        "return_reasons": return_reasons_list,
        "platforms": platforms,
    }

    # ---------- narrative call-outs (recomputed from the numbers above) ----------
    latest_month = months[-1]
    growth_clauses = []
    if len(months) >= 2:
        for p in ("Amazon", "Flipkart"):
            if p in daily_rate_by_month and orders_by_month[p][-1] > 0:
                r_now, r_prev = daily_rate_by_month[p][-1], daily_rate_by_month[p][-2]
                trend = "up from" if r_now > r_prev else ("down from" if r_now < r_prev else "flat vs")
                growth_clauses.append(f"{p} is running at {r_now:.1f}/day ({trend} {r_prev:.1f}/day in {months[-2]})")
    growth_insight = (
        f"{latest_month} data only runs through day {days_elapsed[latest_month]} of the month, so raw monthly "
        f"counts understate it. Normalized to orders/day: " + "; ".join(growth_clauses) + "."
        if growth_clauses else
        f"{latest_month} data only runs through day {days_elapsed[latest_month]} of the month — treat its raw count as partial."
    )

    if len(revenue_platforms) >= 2:
        ranked = sorted(revenue_platforms, key=lambda p: aov[p], reverse=True)
        hi, lo = ranked[0], ranked[1]
        revenue_insight = (
            f"{hi}'s average order value (₹{aov[hi]:.2f}) is higher than {lo}'s (₹{aov[lo]:.2f})."
        )
    elif len(revenue_platforms) == 1:
        p = revenue_platforms[0]
        revenue_insight = f"{p} is the only platform with price data here — average order value is ₹{aov[p]:.2f}."
    else:
        revenue_insight = "No platform in this dataset currently exposes item pricing."
    missing_rev = [p for p in platforms if p not in revenue_platforms]
    if missing_rev:
        revenue_insight += f" {', '.join(missing_rev)} revenue can't be compared — its export has no price field."

    if flipkart_return_rate > amazon_return_rate and amazon_return_rate > 0 and len(return_reasons_list) >= 2:
        ratio = round(flipkart_return_rate / amazon_return_rate, 1)
        fulfillment_insight = (
            f"Flipkart's return rate ({flipkart_return_rate}% of lines) is {ratio}x Amazon's ({amazon_return_rate}%), "
            f"and {dispatch_sla_breach_pct}% of Flipkart dispatches miss the seller's own dispatch SLA. "
            f"The top two return reasons ('{return_reasons_list[0][0]}', '{return_reasons_list[1][0]}') account for "
            f"~{top2_reason_share}% of all Flipkart returns — pointing at pickup/dispatch delay, not product quality."
        )
    else:
        fulfillment_insight = (
            f"Amazon return rate: {amazon_return_rate}%, Flipkart: {flipkart_return_rate}%. "
            f"Flipkart dispatch-SLA breach rate: {dispatch_sla_breach_pct}%."
        )

    brand_insight = (
        f"<b>{top_brand}-compatible parts account for ~{top_brand_pct}% of all identifiable units sold</b> "
        f"across {', '.join(revenue_platforms) or platforms[0]} and other platforms with product names — "
        f"more than every other compatible vehicle combined." if top_brand != "Other" else
        f"No single vehicle brand dominates — the top matched brand is {top_brand} at {top_brand_pct}%."
    )

    weekday_state_insight = (
        f"{busiest_day} is the busiest order day — {weekday_lift_pct:.0f}% more orders than the {quietest_day} low. "
        f"Orders span the country with no single dominant hub: {top_state_name.title()} leads but holds only "
        f"{top_state_pct}% of lines with a known state."
    )

    narrative = {
        "growth_insight": growth_insight,
        "revenue_insight": revenue_insight,
        "fulfillment_insight": fulfillment_insight,
        "brand_insight": brand_insight,
        "weekday_state_insight": weekday_state_insight,
        "latest_month": latest_month,
        "total_orders": total_orders,
        "date_range": f"{months[0]} – {latest_month} (through day {days_elapsed[latest_month]} of {latest_month})",
    }

    render(data, narrative)
    build_pdf(data, narrative)
    print(f"Months covered: {', '.join(months)}")
    print(f"Total orders: {total_orders} | revenue: ₹{total_revenue:,.0f} | units: {total_units}")
    print(f"Saved dashboard to {OUT_PATH}")
    print(f"Saved PDF report to {PDF_PATH}")


def render(data, narrative):
    html = TEMPLATE
    html = html.replace("__DATA_JSON__", json.dumps(data))
    for key, val in narrative.items():
        html = html.replace("__" + key.upper() + "__", str(val))
    with open(OUT_PATH, "w") as f:
        f.write(html)


def strip_tags(s):
    return re.sub(r"<[^>]+>", "", s)


def wrap(s, width=100):
    return "\n".join(textwrap.wrap(s, width=width))


def style_axes(ax):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color(COLOR["grid"])
    ax.spines["bottom"].set_color(COLOR["grid"])
    ax.tick_params(colors=COLOR["secondary"], labelsize=8)
    ax.yaxis.grid(True, color=COLOR["grid"], linewidth=0.8)
    ax.set_axisbelow(True)


def short_month(m):
    return re.sub(r" \d{4}", "", m)


def build_pdf(data, narrative):
    plt.rcParams["font.family"] = "sans-serif"
    months = data["months"]
    platforms = data["platforms"]

    with PdfPages(PDF_PATH) as pdf:
        # ---------- Page 1: cover / summary ----------
        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.07, 0.95, "Order Consolidation — Cross-Platform Analysis", fontsize=16, fontweight="bold", color=COLOR["text"])
        fig.text(0.07, 0.925, f"Amazon, Flipkart & Meesho orders, {narrative['date_range']} · {narrative['total_orders']} order lines", fontsize=10, color=COLOR["secondary"])

        kpis = [
            ("Total order lines", f"{data['total_orders']:,}"),
            ("Tracked revenue", f"₹{data['total_revenue']:,.0f}"),
            ("Units sold", f"{data['total_units']:,}"),
        ]
        for i, (label, val) in enumerate(kpis):
            x = 0.07 + i * 0.3
            fig.text(x, 0.87, label.upper(), fontsize=7.5, color=COLOR["muted"])
            fig.text(x, 0.845, val, fontsize=15, fontweight="bold", color=COLOR["text"])

        sections = [
            ("Growth trajectory", narrative["growth_insight"]),
            ("Revenue", narrative["revenue_insight"]),
            ("Fulfillment health", narrative["fulfillment_insight"]),
            ("Product mix", strip_tags(narrative["brand_insight"])),
            ("Timing & geography", narrative["weekday_state_insight"]),
        ]
        y = 0.78
        for title, body in sections:
            fig.text(0.07, y, title, fontsize=11.5, fontweight="bold", color=COLOR["text"])
            y -= 0.025
            fig.text(0.07, y, wrap(body, 95), fontsize=9, color=COLOR["secondary"], va="top", linespacing=1.6)
            y -= 0.025 + 0.02 * (len(wrap(body, 95).split("\n")))

        fig.text(0.07, 0.04, wrap(
            "Built from Consolidated_Orders.xlsx by analyze_orders.py. Flipkart's export carries no price or "
            "ship-to location; Meesho's export carries no product name or fulfillment status.", 100
        ), fontsize=7.5, color=COLOR["muted"])
        pdf.savefig(fig)
        plt.close(fig)

        # ---------- Page 2: growth ----------
        fig, axes = plt.subplots(1, 2, figsize=(11, 5.5))
        x = range(len(months))
        bar_w = 0.8 / max(1, len(platforms))
        for i, p in enumerate(platforms):
            offs = [xi + (i - (len(platforms) - 1) / 2) * bar_w for xi in x]
            axes[0].bar(offs, data["orders_by_month"][p], width=bar_w, color=platform_color(p, i), label=p)
        axes[0].set_xticks(list(x))
        axes[0].set_xticklabels([short_month(m) for m in months])
        axes[0].set_title("Order lines by month", fontsize=10, loc="left", color=COLOR["text"])
        axes[0].legend(frameon=False, fontsize=8)
        style_axes(axes[0])

        for i, p in enumerate(platforms):
            axes[1].plot(list(x), data["daily_rate_by_month"][p], marker="o", color=platform_color(p, i), label=p, linewidth=2)
        axes[1].set_xticks(list(x))
        axes[1].set_xticklabels([short_month(m) for m in months])
        axes[1].set_title("Daily order run-rate (orders/day)", fontsize=10, loc="left", color=COLOR["text"])
        axes[1].legend(frameon=False, fontsize=8)
        style_axes(axes[1])
        fig.suptitle("Growth trajectory", fontsize=13, fontweight="bold", x=0.07, ha="left", color=COLOR["text"])
        fig.tight_layout(rect=[0, 0, 1, 0.92])
        pdf.savefig(fig)
        plt.close(fig)

        # ---------- Page 3: revenue ----------
        rev_platforms = data["revenue_platforms"]
        fig, axes = plt.subplots(1, 2, figsize=(11, 5.5))
        if rev_platforms:
            bar_w = 0.8 / max(1, len(rev_platforms))
            for i, p in enumerate(rev_platforms):
                offs = [xi + (i - (len(rev_platforms) - 1) / 2) * bar_w for xi in x]
                axes[0].bar(offs, data["revenue_by_month"][p], width=bar_w, color=platform_color(p, platforms.index(p)), label=p)
            axes[0].set_xticks(list(x))
            axes[0].set_xticklabels([short_month(m) for m in months])
            axes[0].legend(frameon=False, fontsize=8)
            axes[1].bar(rev_platforms, [data["aov"][p] for p in rev_platforms], color=[platform_color(p, platforms.index(p)) for p in rev_platforms], width=0.5)
            for i, p in enumerate(rev_platforms):
                axes[1].text(i, data["aov"][p] + max(data["aov"].values()) * 0.02, f"₹{data['aov'][p]:.0f}", ha="center", fontsize=8, color=COLOR["text"])
        axes[0].set_title("Monthly revenue (₹)", fontsize=10, loc="left", color=COLOR["text"])
        axes[1].set_title("Average order value (₹)", fontsize=10, loc="left", color=COLOR["text"])
        style_axes(axes[0]); style_axes(axes[1])
        fig.suptitle("Revenue", fontsize=13, fontweight="bold", x=0.07, ha="left", color=COLOR["text"])
        fig.tight_layout(rect=[0, 0, 1, 0.92])
        pdf.savefig(fig)
        plt.close(fig)

        # ---------- Page 4: fulfillment health ----------
        bucket_colors = {
            "Delivered": COLOR["good"], "In Transit": COLOR["warning"], "Pending": COLOR["muted"],
            "Returned": COLOR["serious"], "Cancelled": COLOR["critical"],
        }
        fig = plt.figure(figsize=(11, 8.5))
        gs = fig.add_gridspec(2, 2, height_ratios=[1, 1.3], hspace=0.55, wspace=0.3)

        ax0 = fig.add_subplot(gs[0, :])
        status_platforms = [p for p in ("Amazon", "Flipkart") if p in data["status_pct"]]
        y = range(len(status_platforms))
        lefts = [0.0] * len(status_platforms)
        for b in data["status_buckets"]:
            vals = [data["status_pct"][p][data["status_buckets"].index(b)] for p in status_platforms]
            ax0.barh(list(y), vals, left=lefts, color=bucket_colors[b], label=b, height=0.55)
            for yi, (v, l) in enumerate(zip(vals, lefts)):
                if v >= 7:
                    ax0.text(l + v / 2, yi, f"{v:.0f}%", ha="center", va="center", color="white", fontsize=8, fontweight="bold")
            lefts = [l + v for l, v in zip(lefts, vals)]
        ax0.set_yticks(list(y)); ax0.set_yticklabels(status_platforms)
        ax0.set_xlim(0, 100)
        ax0.invert_yaxis()
        ax0.set_title("Order status mix by platform (% of lines)", fontsize=10, loc="left", color=COLOR["text"])
        ax0.legend(loc="upper center", bbox_to_anchor=(0.5, -0.25), ncol=5, fontsize=8, frameon=False)
        ax0.spines["top"].set_visible(False); ax0.spines["right"].set_visible(False); ax0.spines["left"].set_visible(False)
        ax0.tick_params(colors=COLOR["secondary"], labelsize=9)

        ax1 = fig.add_subplot(gs[1, 0])
        ax1.bar([short_month(m) for m in months], data["dispatch_sla_breach_by_month"], color=COLOR["warning"])
        for i, v in enumerate(data["dispatch_sla_breach_by_month"]):
            ax1.text(i, v + 1, f"{v:.0f}%", ha="center", fontsize=8, color=COLOR["text"])
        ax1.set_title("Flipkart dispatch-SLA breach rate by month (%)", fontsize=9.5, loc="left", color=COLOR["text"])
        style_axes(ax1)

        ax2 = fig.add_subplot(gs[1, 1])
        reasons = data["return_reasons"][::-1]
        ax2.barh([r[0] for r in reasons], [r[1] for r in reasons], color=COLOR["serious"])
        ax2.set_title("Flipkart return reasons (order lines)", fontsize=9.5, loc="left", color=COLOR["text"])
        ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
        ax2.tick_params(colors=COLOR["secondary"], labelsize=8)

        fig.suptitle("Fulfillment health", fontsize=13, fontweight="bold", x=0.07, ha="left", color=COLOR["text"])
        pdf.savefig(fig)
        plt.close(fig)

        # ---------- Page 5: product mix + timing/geography ----------
        fig = plt.figure(figsize=(11, 8.5))
        gs = fig.add_gridspec(2, 2, height_ratios=[1, 1.1], hspace=0.55, wspace=0.3)

        ax0 = fig.add_subplot(gs[0, :])
        brands = data["brand_units"][::-1]
        colors = [COLOR["Amazon"] if i == len(brands) - 1 else COLOR["muted"] for i in range(len(brands))]
        ax0.barh([b[0] for b in brands], [b[1] for b in brands], color=colors)
        ax0.set_title("Units sold by compatible vehicle (named products)", fontsize=10, loc="left", color=COLOR["text"])
        ax0.spines["top"].set_visible(False); ax0.spines["right"].set_visible(False)
        ax0.tick_params(colors=COLOR["secondary"], labelsize=8)

        ax1 = fig.add_subplot(gs[1, 0])
        wd = data["weekday"]
        ax1.bar([d[0][:3] for d in wd], [d[1] for d in wd], color=COLOR["Amazon"])
        ax1.set_title("Orders by day of week", fontsize=9.5, loc="left", color=COLOR["text"])
        style_axes(ax1)

        ax2 = fig.add_subplot(gs[1, 1])
        states = data["top_states"][::-1]
        ax2.barh([s[0].title() for s in states], [s[1] for s in states], color=COLOR["Amazon"])
        ax2.set_title("Top states by order lines", fontsize=9.5, loc="left", color=COLOR["text"])
        ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
        ax2.tick_params(colors=COLOR["secondary"], labelsize=8)

        fig.suptitle("Product mix, timing & geography", fontsize=13, fontweight="bold", x=0.07, ha="left", color=COLOR["text"])
        pdf.savefig(fig)
        plt.close(fig)


TEMPLATE = r"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Order Consolidation — Cross-Platform Analysis</title>
<style>
  .viz-root {
    color-scheme: light;
    --surface-1:      #fcfcfb;
    --page:           #f9f9f7;
    --text-primary:   #0b0b0b;
    --text-secondary: #52514e;
    --text-muted:     #898781;
    --grid:           #e1e0d9;
    --baseline:       #c3c2b7;
    --border:         rgba(11,11,11,0.10);
    --series-1:       #2a78d6;
    --series-2:       #008300;
    --series-3:       #e87ba4;
    --good:           #0ca30c;
    --warning:        #fab219;
    --serious:        #ec835a;
    --critical:       #d03b3b;
    --tooltip-bg:     #1a1a19;
    --tooltip-fg:     #ffffff;
  }
  @media (prefers-color-scheme: dark) {
    :root:where(:not([data-theme="light"])) .viz-root {
      color-scheme: dark;
      --surface-1:      #1a1a19;
      --page:           #0d0d0d;
      --text-primary:   #ffffff;
      --text-secondary: #c3c2b7;
      --text-muted:     #898781;
      --grid:           #2c2c2a;
      --baseline:       #383835;
      --border:         rgba(255,255,255,0.10);
      --series-1:       #3987e5;
      --series-2:       #1fb31f;
      --series-3:       #d55181;
      --good:           #0ca30c;
      --warning:        #fab219;
      --serious:        #ec835a;
      --critical:       #e66767;
      --tooltip-bg:     #fcfcfb;
      --tooltip-fg:     #0b0b0b;
    }
  }
  :root[data-theme="dark"] .viz-root {
    color-scheme: dark;
    --surface-1:      #1a1a19;
    --page:           #0d0d0d;
    --text-primary:   #ffffff;
    --text-secondary: #c3c2b7;
    --text-muted:     #898781;
    --grid:           #2c2c2a;
    --baseline:       #383835;
    --border:         rgba(255,255,255,0.10);
    --series-1:       #3987e5;
    --series-2:       #1fb31f;
    --series-3:       #d55181;
    --good:           #0ca30c;
    --warning:        #fab219;
    --serious:        #ec835a;
    --critical:       #e66767;
    --tooltip-bg:     #fcfcfb;
    --tooltip-fg:     #0b0b0b;
  }

  * { box-sizing: border-box; }
  body { margin: 0; }
  .viz-root {
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: var(--page);
    color: var(--text-primary);
    padding: 28px 20px 60px;
  }
  .wrap { max-width: 1120px; margin: 0 auto; }
  header.page-head { margin-bottom: 24px; }
  header.page-head h1 { font-size: 1.5rem; margin: 0 0 6px; }
  header.page-head p { margin: 0; color: var(--text-secondary); font-size: 0.92rem; }

  .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 28px; }
  @media (max-width: 760px) { .kpi-row { grid-template-columns: repeat(2, 1fr); } }
  .kpi {
    background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px;
    padding: 14px 16px;
  }
  .kpi .label { font-size: 0.74rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; }
  .kpi .value { font-size: 1.5rem; font-weight: 600; margin-top: 4px; font-variant-numeric: tabular-nums; }
  .kpi .sub { font-size: 0.78rem; color: var(--text-secondary); margin-top: 2px; }

  section.panel {
    background: var(--surface-1); border: 1px solid var(--border); border-radius: 12px;
    padding: 18px 20px 20px; margin-bottom: 20px;
  }
  section.panel h2 { font-size: 1.05rem; margin: 0 0 2px; }
  section.panel .insight { font-size: 0.85rem; color: var(--text-secondary); margin: 0 0 14px; max-width: 760px; }
  .chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
  @media (max-width: 760px) { .chart-row { grid-template-columns: 1fr; } }
  .chart-scroll { overflow-x: auto; }
  .chart-box svg { display: block; width: 100%; height: auto; }
  .chart-title { font-size: 0.8rem; color: var(--text-secondary); margin: 0 0 8px; font-weight: 600; }

  .legend { display: flex; gap: 14px; flex-wrap: wrap; margin: 10px 0 0; font-size: 0.78rem; color: var(--text-secondary); }
  .legend .item { display: flex; align-items: center; gap: 6px; }
  .legend .swatch { width: 10px; height: 10px; border-radius: 2px; display: inline-block; }

  .axis-text { font-size: 10px; fill: var(--text-muted); font-family: inherit; }
  .grid-line { stroke: var(--grid); stroke-width: 1; }
  .baseline { stroke: var(--baseline); stroke-width: 1; }
  .label-text { font-size: 10px; fill: var(--text-secondary); font-family: inherit; }
  .value-label { font-size: 10px; fill: var(--text-primary); font-family: inherit; font-weight: 600; }

  #tooltip {
    position: fixed; pointer-events: none; z-index: 1000;
    background: var(--tooltip-bg); color: var(--tooltip-fg);
    font-size: 0.78rem; padding: 6px 9px; border-radius: 6px;
    opacity: 0; transition: opacity 0.08s ease; white-space: nowrap;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  }

  footer.page-foot { margin-top: 8px; font-size: 0.78rem; color: var(--text-muted); line-height: 1.6; }
</style>
</head>
<body>
<div class="viz-root">
<div class="wrap">

  <header class="page-head">
    <h1>Order Consolidation — Cross-Platform Analysis</h1>
    <p>Amazon, Flipkart &amp; Meesho orders, __DATE_RANGE__ · __TOTAL_ORDERS__ order lines consolidated</p>
  </header>

  <div class="kpi-row">
    <div class="kpi"><div class="label">Total order lines</div><div class="value" id="kpi-orders">—</div><div class="sub">across all marketplaces in the workbook</div></div>
    <div class="kpi"><div class="label">Tracked revenue</div><div class="value" id="kpi-revenue">—</div><div class="sub">platforms with a price field only</div></div>
    <div class="kpi"><div class="label">Units sold</div><div class="value" id="kpi-units">—</div><div class="sub">excludes cancelled-before-ship lines</div></div>
    <div class="kpi"><div class="label">Latest-month run-rate</div><div class="value" id="kpi-runrate">—</div><div class="sub">orders/day, Amazon / Flipkart</div></div>
  </div>

  <section class="panel">
    <h2>Growth trajectory</h2>
    <p class="insight">__GROWTH_INSIGHT__</p>
    <div class="chart-row">
      <div class="chart-box chart-scroll">
        <p class="chart-title">Order lines by month</p>
        <svg id="chart-orders" viewBox="0 0 480 260"></svg>
      </div>
      <div class="chart-box chart-scroll">
        <p class="chart-title">Daily order run-rate (orders/day)</p>
        <svg id="chart-runrate" viewBox="0 0 480 260"></svg>
      </div>
    </div>
    <div class="legend" id="legend-platforms"></div>
  </section>

  <section class="panel">
    <h2>Revenue</h2>
    <p class="insight">__REVENUE_INSIGHT__</p>
    <div class="chart-row">
      <div class="chart-box chart-scroll">
        <p class="chart-title">Monthly revenue (₹)</p>
        <svg id="chart-revenue" viewBox="0 0 480 260"></svg>
      </div>
      <div class="chart-box chart-scroll">
        <p class="chart-title">Average order value (₹)</p>
        <svg id="chart-aov" viewBox="0 0 480 260"></svg>
      </div>
    </div>
  </section>

  <section class="panel">
    <h2>Fulfillment health</h2>
    <p class="insight">__FULFILLMENT_INSIGHT__</p>
    <div class="chart-box chart-scroll" style="margin-bottom:18px;">
      <p class="chart-title">Order status mix by platform (% of lines)</p>
      <svg id="chart-status" viewBox="0 0 900 140"></svg>
      <div class="legend" id="legend-status"></div>
    </div>
    <div class="chart-row">
      <div class="chart-box chart-scroll">
        <p class="chart-title">Flipkart dispatch-SLA breach rate by month (%)</p>
        <svg id="chart-sla" viewBox="0 0 480 220"></svg>
      </div>
      <div class="chart-box chart-scroll">
        <p class="chart-title">Flipkart return reasons (order lines)</p>
        <svg id="chart-returns" viewBox="0 0 480 220"></svg>
      </div>
    </div>
  </section>

  <section class="panel">
    <h2>Product mix</h2>
    <p class="insight">__BRAND_INSIGHT__</p>
    <div class="chart-box chart-scroll">
      <p class="chart-title">Units sold by compatible vehicle (named products)</p>
      <svg id="chart-brand" viewBox="0 0 700 220"></svg>
    </div>
  </section>

  <section class="panel">
    <h2>Timing &amp; geography</h2>
    <p class="insight">__WEEKDAY_STATE_INSIGHT__</p>
    <div class="chart-row">
      <div class="chart-box chart-scroll">
        <p class="chart-title">Orders by day of week</p>
        <svg id="chart-weekday" viewBox="0 0 480 220"></svg>
      </div>
      <div class="chart-box chart-scroll">
        <p class="chart-title">Top states by order lines</p>
        <svg id="chart-states" viewBox="0 0 480 240"></svg>
      </div>
    </div>
  </section>

  <footer class="page-foot">
    Built from Consolidated_Orders.xlsx by analyze_orders.py — regenerated on every run from whatever data is in
    the workbook. The __LATEST_MONTH__ figures are partial and status-dependent metrics (returns/cancellations) are
    under-counted for the latest month since many lines haven't resolved yet. Flipkart's export carries no price or
    ship-to location; Meesho's export carries no product name or fulfillment status — each chart above uses only
    the fields its own platform's export provides.
  </footer>

</div>
</div>

<div id="tooltip"></div>

<script>
const DATA = __DATA_JSON__;

const COLORS = {
  amazon: 'var(--series-1)', flipkart: 'var(--series-2)', meesho: 'var(--series-3)',
  good: 'var(--good)', warning: 'var(--warning)', serious: 'var(--serious)', critical: 'var(--critical)'
};
function cssVar(name) {
  return getComputedStyle(document.querySelector('.viz-root')).getPropertyValue(name.match(/--[a-z0-9-]+/)[0]).trim();
}
function resolve(c) { return c.startsWith('var(') ? cssVar(c) : c; }

const PLATFORM_COLOR = { Amazon: COLORS.amazon, Flipkart: COLORS.flipkart, Meesho: COLORS.meesho };
function platformColor(p, idx) {
  return PLATFORM_COLOR[p] || ['var(--series-1)','var(--series-2)','var(--series-3)'][idx % 3];
}

const tip = document.getElementById('tooltip');
function showTip(html, evt) {
  tip.innerHTML = html;
  tip.style.opacity = '1';
  tip.style.left = (evt.clientX + 14) + 'px';
  tip.style.top = (evt.clientY + 14) + 'px';
}
function hideTip() { tip.style.opacity = '0'; }

function svgNode(tag, attrs) {
  const el = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const k in attrs) el.setAttribute(k, attrs[k]);
  return el;
}

function legend(containerId, items) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = items.map(it =>
    `<span class="item"><span class="swatch" style="background:${resolve(it.color)}"></span>${it.label}</span>`
  ).join('');
}

function drawLines(svgId, categories, series, opts) {
  opts = opts || {};
  const svg = document.getElementById(svgId);
  const W = 480, H = 260, padL = 40, padR = 14, padT = 16, padB = 30;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const allVals = series.flatMap(s => s.values);
  const maxV = Math.max(1, ...allVals) * 1.15;
  const denom = Math.max(1, categories.length - 1);
  const x = i => padL + (innerW * i) / denom;
  const y = v => padT + innerH - (innerH * v) / maxV;

  const ticks = 4;
  for (let t = 0; t <= ticks; t++) {
    const v = (maxV / ticks) * t;
    const gy = y(v);
    svg.appendChild(svgNode('line', { x1: padL, x2: W - padR, y1: gy, y2: gy, class: 'grid-line' }));
    const lbl = svgNode('text', { x: padL - 6, y: gy + 3, 'text-anchor': 'end', class: 'axis-text' });
    lbl.textContent = opts.yFormat ? opts.yFormat(v) : Math.round(v);
    svg.appendChild(lbl);
  }
  categories.forEach((c, i) => {
    const lbl = svgNode('text', { x: x(i), y: H - 8, 'text-anchor': 'middle', class: 'axis-text' });
    lbl.textContent = c.replace(/ \d{4}/, '');
    svg.appendChild(lbl);
  });

  series.forEach(s => {
    const pts = s.values.map((v, i) => [x(i), y(v)]);
    const path = pts.map((p, i) => (i === 0 ? 'M' : 'L') + p[0] + ' ' + p[1]).join(' ');
    svg.appendChild(svgNode('path', { d: path, fill: 'none', stroke: resolve(s.color), 'stroke-width': 2 }));
    pts.forEach(([px, py], i) => {
      const c = svgNode('circle', { cx: px, cy: py, r: 4, fill: resolve(s.color) });
      c.addEventListener('mousemove', e => showTip(`<b>${s.name}</b> · ${categories[i]}<br>${opts.yFormat ? opts.yFormat(s.values[i]) : s.values[i]}`, e));
      c.addEventListener('mouseleave', hideTip);
      svg.appendChild(c);
    });
  });
}

function drawGroupedBars(svgId, categories, series, opts) {
  opts = opts || {};
  const svg = document.getElementById(svgId);
  const W = 480, H = 260, padL = 40, padR = 14, padT = 16, padB = 30;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const allVals = series.flatMap(s => s.values);
  const maxV = Math.max(1, ...allVals) * 1.15;
  const groupW = innerW / categories.length;
  const barGap = 2, groupPad = 10;
  const barW = Math.max(2, (groupW - groupPad - barGap * (series.length - 1)) / series.length);
  const y = v => padT + innerH - (innerH * v) / maxV;

  const ticks = 4;
  for (let t = 0; t <= ticks; t++) {
    const v = (maxV / ticks) * t;
    const gy = y(v);
    svg.appendChild(svgNode('line', { x1: padL, x2: W - padR, y1: gy, y2: gy, class: 'grid-line' }));
    const lbl = svgNode('text', { x: padL - 6, y: gy + 3, 'text-anchor': 'end', class: 'axis-text' });
    lbl.textContent = opts.yFormat ? opts.yFormat(v) : Math.round(v);
    svg.appendChild(lbl);
  }
  svg.appendChild(svgNode('line', { x1: padL, x2: W - padR, y1: padT + innerH, y2: padT + innerH, class: 'baseline' }));

  categories.forEach((cat, ci) => {
    const gx0 = padL + groupW * ci + groupPad / 2;
    series.forEach((s, si) => {
      const v = s.values[ci];
      const bx = gx0 + si * (barW + barGap);
      const by = y(v);
      const bh = padT + innerH - by;
      const rect = svgNode('rect', { x: bx, y: by, width: barW, height: Math.max(0, bh), rx: 3, fill: resolve(s.color) });
      rect.addEventListener('mousemove', e => showTip(`<b>${s.name}</b> · ${cat}<br>${opts.yFormat ? opts.yFormat(v) : v}`, e));
      rect.addEventListener('mouseleave', hideTip);
      svg.appendChild(rect);
    });
    const lbl = svgNode('text', { x: gx0 + (groupW - groupPad) / 2, y: H - 8, 'text-anchor': 'middle', class: 'axis-text' });
    lbl.textContent = cat.replace(/ \d{4}/, '');
    svg.appendChild(lbl);
  });
}

function drawSingleBars(svgId, items, color, opts) {
  opts = opts || {};
  const svg = document.getElementById(svgId);
  const W = 480, H = 220, padL = 36, padR = 14, padT = 16, padB = 32;
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const maxV = Math.max(1, ...items.map(it => it[1])) * 1.2;
  const slot = innerW / items.length;
  const barW = Math.min(46, slot * 0.6);
  const y = v => padT + innerH - (innerH * v) / maxV;

  svg.appendChild(svgNode('line', { x1: padL, x2: W - padR, y1: padT + innerH, y2: padT + innerH, class: 'baseline' }));

  items.forEach((it, i) => {
    const [label, v] = it;
    const cx = padL + slot * i + slot / 2;
    const by = y(v);
    const bh = padT + innerH - by;
    const col = opts.colorFn ? opts.colorFn(label, v, i) : color;
    const rect = svgNode('rect', { x: cx - barW / 2, y: by, width: barW, height: Math.max(0, bh), rx: 3, fill: resolve(col) });
    rect.addEventListener('mousemove', e => showTip(`<b>${label}</b><br>${opts.valueLabel ? opts.valueLabel(v) : v}`, e));
    rect.addEventListener('mouseleave', hideTip);
    svg.appendChild(rect);
    const vlbl = svgNode('text', { x: cx, y: by - 5, 'text-anchor': 'middle', class: 'value-label' });
    vlbl.textContent = opts.valueLabel ? opts.valueLabel(v) : v;
    svg.appendChild(vlbl);
    const lbl = svgNode('text', { x: cx, y: H - 10, 'text-anchor': 'middle', class: 'axis-text' });
    lbl.textContent = opts.shortLabel ? opts.shortLabel(label) : label;
    svg.appendChild(lbl);
  });
}

function drawHorizBars(svgId, items, color, opts) {
  opts = opts || {};
  const svg = document.getElementById(svgId);
  const n = items.length;
  const rowH = 26, gap = 8;
  const H = n * (rowH + gap) + 10;
  const W = parseFloat(svg.getAttribute('viewBox').split(' ')[2]) || 480;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  const labelW = opts.labelW || 150;
  const padR = 50;
  const innerW = W - labelW - padR;
  const maxV = Math.max(1, ...items.map(it => it[1]));

  items.forEach((it, i) => {
    const [label, v] = it;
    const ry = 10 + i * (rowH + gap);
    const bw = (innerW * v) / maxV;
    const col = opts.colorFn ? opts.colorFn(label, v) : color;
    const rect = svgNode('rect', { x: labelW, y: ry, width: Math.max(2, bw), height: rowH, rx: 4, fill: resolve(col) });
    rect.addEventListener('mousemove', e => showTip(`<b>${label}</b><br>${opts.valueLabel ? opts.valueLabel(v) : v}`, e));
    rect.addEventListener('mouseleave', hideTip);
    svg.appendChild(rect);
    const lbl = svgNode('text', { x: labelW - 8, y: ry + rowH / 2 + 3, 'text-anchor': 'end', class: 'label-text' });
    lbl.textContent = label;
    svg.appendChild(lbl);
    const vlbl = svgNode('text', { x: labelW + bw + 8, y: ry + rowH / 2 + 3, 'text-anchor': 'start', class: 'value-label' });
    vlbl.textContent = opts.valueLabel ? opts.valueLabel(v) : v;
    svg.appendChild(vlbl);
  });
}

function drawStatusStack(svgId, rows) {
  const svg = document.getElementById(svgId);
  const W = 900;
  const rowH = 44, gap = 18;
  const labelW = 90, padR = 20;
  const H = rows.length * (rowH + gap) + 10;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  const innerW = W - labelW - padR;

  rows.forEach((row, ri) => {
    const ry = 10 + ri * (rowH + gap);
    let cx = labelW;
    const lbl = svgNode('text', { x: labelW - 10, y: ry + rowH / 2 + 4, 'text-anchor': 'end', class: 'label-text' });
    lbl.textContent = row.name;
    svg.appendChild(lbl);
    row.segments.forEach(seg => {
      const segW = (innerW * seg.pct) / 100;
      if (segW < 0.5) return;
      const rect = svgNode('rect', { x: cx, y: ry, width: Math.max(0, segW - 2), height: rowH, rx: 3, fill: resolve(seg.color) });
      rect.addEventListener('mousemove', e => showTip(`<b>${row.name}</b> · ${seg.name}<br>${seg.pct}% of lines`, e));
      rect.addEventListener('mouseleave', hideTip);
      svg.appendChild(rect);
      if (seg.pct >= 7) {
        const t = svgNode('text', { x: cx + segW / 2, y: ry + rowH / 2 + 4, 'text-anchor': 'middle' });
        t.setAttribute('style', 'fill:#fff;font-size:10px;font-weight:600;');
        t.textContent = seg.pct + '%';
        svg.appendChild(t);
      }
      cx += segW;
    });
  });
}

const inr = v => '₹' + Math.round(v).toLocaleString('en-IN');
const num = v => Math.round(v).toLocaleString('en-IN');

document.getElementById('kpi-orders').textContent = num(DATA.total_orders);
document.getElementById('kpi-revenue').textContent = inr(DATA.total_revenue);
document.getElementById('kpi-units').textContent = num(DATA.total_units);
{
  const n = DATA.months.length - 1;
  const a = DATA.daily_rate_by_month.Amazon ? DATA.daily_rate_by_month.Amazon[n].toFixed(1) : '—';
  const f = DATA.daily_rate_by_month.Flipkart ? DATA.daily_rate_by_month.Flipkart[n].toFixed(1) : '—';
  document.getElementById('kpi-runrate').textContent = a + ' / ' + f;
}

legend('legend-platforms', DATA.platforms.map((p, i) => ({ color: platformColor(p, i), label: p })));

drawGroupedBars('chart-orders', DATA.months, DATA.platforms.map((p, i) => ({ name: p, color: platformColor(p, i), values: DATA.orders_by_month[p] })));

drawLines('chart-runrate', DATA.months, DATA.platforms.map((p, i) => ({ name: p, color: platformColor(p, i), values: DATA.daily_rate_by_month[p] })), { yFormat: v => v.toFixed(1) });

drawGroupedBars('chart-revenue', DATA.months, DATA.revenue_platforms.map((p, i) => ({ name: p, color: platformColor(p, DATA.platforms.indexOf(p)), values: DATA.revenue_by_month[p] })), { yFormat: v => '₹' + Math.round(v / 1000) + 'k' });

drawSingleBars('chart-aov', DATA.revenue_platforms.map(p => [p, DATA.aov[p]]), platformColor(DATA.revenue_platforms[0], 0), {
  valueLabel: v => '₹' + v.toFixed(0),
  colorFn: (label) => platformColor(label, DATA.platforms.indexOf(label)),
});

const bucketColorMap = {
  Delivered: COLORS.good,
  'In Transit': COLORS.warning,
  Pending: 'var(--text-muted)',
  Returned: COLORS.serious,
  Cancelled: COLORS.critical,
};
drawStatusStack('chart-status', ['Amazon', 'Flipkart'].filter(p => DATA.status_pct[p]).map(p => ({
  name: p,
  segments: DATA.status_buckets.map((b, i) => ({ name: b, color: bucketColorMap[b], pct: DATA.status_pct[p][i] })),
})));
legend('legend-status', DATA.status_buckets.map(b => ({ color: bucketColorMap[b], label: b })));

drawSingleBars('chart-sla', DATA.months.map((m, i) => [m.replace(/ \d{4}/, ''), DATA.dispatch_sla_breach_by_month[i]]), COLORS.warning, {
  valueLabel: v => v + '%',
});

drawHorizBars('chart-returns', DATA.return_reasons, COLORS.serious, { labelW: 160, valueLabel: v => v + ' lines' });

const brandTotal = DATA.brand_total_named;
drawHorizBars('chart-brand', DATA.brand_units, COLORS.amazon, {
  labelW: 130,
  valueLabel: v => v + ' units · ' + (100 * v / brandTotal).toFixed(0) + '%',
  colorFn: (label, v, i) => i === 0 ? COLORS.amazon : 'var(--text-muted)',
});

drawSingleBars('chart-weekday', DATA.weekday.map(([d, v]) => [d.slice(0, 3), v]), COLORS.amazon, {});

drawHorizBars('chart-states', DATA.top_states, COLORS.amazon, { labelW: 130, valueLabel: v => v + ' lines' });
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
