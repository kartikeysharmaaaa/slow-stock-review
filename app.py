"""
Slow-Stock Review - Company X
Upload the month's exports, check the data, classify items and build customer call lists.

Run locally:   streamlit run app.py
"""

import io
import re
import urllib.parse
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Business rules - change here if the rules change
# ---------------------------------------------------------------------------
MONTHS_TO_CLEAR_LIMIT = 3        # R5: months to clear below this -> No Action
TARGET_MIN_PURCHASE = 100_000    # R6: only customers above this purchase value are targeted
TOP_N = 10                       # how many items / customers to show first
BUCKET_NAMES = {1: "Bought before", 2: "Similar business", 3: "Same category"}
CLASS_ORDER = {"Push": 0, "No Action": 1, "Red List": 2}
SHEET_NAMES = {"items": "Item_Master", "stock": "Stock_Aging",
               "sales": "Sales_Lines", "customers": "Customers"}
SAMPLE_FILE = Path(__file__).parent / "case_data.xlsx"
IST = ZoneInfo("Asia/Kolkata")

st.set_page_config(page_title="Slow-Stock Review", layout="wide")
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1150px;}
    @media (max-width: 640px) {
        .block-container {padding-left: 0.75rem; padding-right: 0.75rem;}
        h1 {font-size: 1.6rem !important;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def clean_sku(x):
    """'sku1025 ' -> 'SKU1025'"""
    if pd.isna(x):
        return None
    s = re.sub(r"\s+", "", str(x)).upper()
    return s or None


def clean_text(x):
    """Trim and collapse repeated spaces. Blank/NaN -> ''"""
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x)).strip()


def parse_month(x):
    """Reads 'Apr-2026', '2026-06', real Excel dates etc. Returns a monthly Period or None."""
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, datetime)):
        return pd.Period(x, freq="M")
    s = str(x).strip()
    for fmt in ("%b-%Y", "%B-%Y", "%Y-%m", "%m-%Y", "%b %Y", "%B %Y",
                "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return pd.Period(datetime.strptime(s, fmt), freq="M")
        except ValueError:
            continue
    return None


def month_label(p):
    return p.strftime("%b-%Y")


def inr(x):
    """163841 -> '₹1,63,841' (Indian digit grouping)."""
    try:
        n = int(round(float(x)))
    except (TypeError, ValueError):
        return "-"
    sign, s = ("-" if n < 0 else ""), str(abs(n))
    if len(s) > 3:
        head = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", s[:-3])
        s = f"{head},{s[-3:]}"
    return f"{sign}₹{s}"


def find_col(columns, key):
    key = key.replace(" ", "").lower()
    for c in columns:
        if key in str(c).replace(" ", "").lower():
            return c
    return None


def col_cfg():
    money = st.column_config.NumberColumn(format="₹%d")
    return {
        "Unit Cost": money, "Aged Value": money, "Item Aged Value": money,
        "6M Purchase Value": money,
        "Avg / Month": st.column_config.NumberColumn(format="%.1f"),
        "Months to Clear": st.column_config.NumberColumn(format="%.2f"),
    }


class IssueLog:
    """Every data problem found, and what we did about it."""
    COLS = ["Sheet", "Problem", "Count", "How it was handled", "Status", "Details"]

    def __init__(self):
        self.rows = []

    def add(self, sheet, problem, count, action, status="Fixed", detail=""):
        count = int(count)
        if count > 0:
            self.rows.append(dict(zip(self.COLS, [sheet, problem, count, action, status, detail])))

    def frame(self):
        return pd.DataFrame(self.rows, columns=self.COLS)


# ---------------------------------------------------------------------------
# Loading - sheets are recognised by their columns, not their tab names,
# because the separate files all arrive as "Sheet1".
# ---------------------------------------------------------------------------
def identify(df):
    cols = {str(c).strip().lower() for c in df.columns}
    if "master pack qty" in cols:
        return "items"
    if any("180+" in c for c in cols):
        return "stock"
    if {"month", "qty"} <= cols:
        return "sales"
    if {"customerid", "area"} <= cols:
        return "customers"
    return None


def load_files(blobs):
    found, notes = {}, []
    for name, data in blobs:
        try:
            if name.lower().endswith(".csv"):
                sheets = {"csv": pd.read_csv(io.BytesIO(data))}
            else:
                xl = pd.ExcelFile(io.BytesIO(data))
                sheets = {sh: xl.parse(sh) for sh in xl.sheet_names}
        except Exception as e:
            notes.append(f"Could not read {name} ({e.__class__.__name__})")
            continue
        for df in sheets.values():
            df.columns = [str(c).strip() for c in df.columns]
            kind = identify(df)
            if kind:
                found[kind] = df
    return found, notes


# ---------------------------------------------------------------------------
# Rule R5 helpers
# ---------------------------------------------------------------------------
def classify(r):
    if r["Sold 6M"] < r["Master Pack"]:
        return "Red List"
    if r["Aged Qty"] <= 0 or r["Months to Clear"] < MONTHS_TO_CLEAR_LIMIT:
        return "No Action"
    return "Push"


def reason(r):
    if r["Class"] == "Red List":
        return f"Sold {r['Sold 6M']:g} units, less than one pack of {r['Master Pack']:g}"
    if r["Class"] == "No Action":
        if r["Aged Qty"] <= 0:
            return "No stock older than 90 days"
        return f"Clears in {r['Months to Clear']:.1f} months (under {MONTHS_TO_CLEAR_LIMIT})"
    return f"{r['Aged Qty']:g} units older than 90 days and {r['Months to Clear']:.1f} months to clear"


# ---------------------------------------------------------------------------
# The review
# ---------------------------------------------------------------------------
def run_review(raw, read_notes):
    log = IssueLog()
    for note in read_notes:
        log.add("Upload", note, 1, "File skipped", "Flagged")

    # ======================= 1. ITEM MASTER =======================
    im = raw["items"].copy()
    for c in ["Alternate Item (name)", "Alternate Flag", "Status"]:
        if c not in im.columns:
            im[c] = ""
    im = im[im["SKU"].notna()].copy()
    im["SKU"] = im["SKU"].map(clean_sku)

    dup = im["SKU"].duplicated()
    log.add("Item_Master", "Same SKU listed more than once", dup.sum(),
            "Kept the first row", "Fixed", ", ".join(im.loc[dup, "SKU"]))
    im = im[~dup].copy()

    im["Item Name"] = im["Item Name"].map(clean_text)
    im["Category"] = im["Category"].map(clean_text)
    im["Master Pack"] = pd.to_numeric(im["Master Pack Qty"], errors="coerce")
    im["Unit Cost"] = pd.to_numeric(im["Unit Cost (Rs)"], errors="coerce")
    bad = (im["Master Pack"].isna() | (im["Master Pack"] <= 0)
           | im["Unit Cost"].isna() | (im["Unit Cost"] < 0))
    log.add("Item_Master", "Master pack or unit cost missing / invalid", bad.sum(),
            "Pack set to 1 and cost to 0 so the run can finish - fix in the master",
            "Flagged", ", ".join(im.loc[bad, "SKU"]))
    im.loc[im["Master Pack"].isna() | (im["Master Pack"] <= 0), "Master Pack"] = 1
    im["Unit Cost"] = im["Unit Cost"].fillna(0).clip(lower=0)

    im["Status"] = im["Status"].map(clean_text).str.title()
    odd = ~im["Status"].isin(["Active", "Inactive"])
    log.add("Item_Master", "Status is neither Active nor Inactive", odd.sum(),
            "Treated as Inactive (left out)", "Flagged", ", ".join(im.loc[odd, "SKU"]))
    im.loc[odd, "Status"] = "Inactive"
    status = dict(zip(im["SKU"], im["Status"]))
    cost = dict(zip(im["SKU"], im["Unit Cost"]))
    name_of = dict(zip(im["SKU"], im["Item Name"]))

    # Alternate item is stored by NAME, so look the SKU up from the name
    alt_raw = im["Alternate Item (name)"].fillna("").astype(str)
    im["Alt Name"] = alt_raw.map(clean_text)
    im["Alt Flag"] = im["Alternate Flag"].map(clean_text).str.upper()
    spaced = (im["Alt Name"] != "") & (alt_raw != im["Alt Name"])
    log.add("Item_Master", "Alternate item name has extra spaces", spaced.sum(),
            "Spaces removed before matching the name to a SKU", "Fixed",
            ", ".join(f"{s}: '{a}'" for s, a in zip(im.loc[spaced, "SKU"], alt_raw[spaced])))

    name_to_sku = {n.lower(): s for n, s in zip(im["Item Name"], im["SKU"])}
    im["Alt SKU"] = im["Alt Name"].str.lower().map(name_to_sku)
    has_alt = im["Alt Name"] != ""
    unmatched = has_alt & im["Alt SKU"].isna()
    log.add("Item_Master", "Alternate item name not found in the item list", unmatched.sum(),
            "Switch not applied - item analysed on its own", "Flagged",
            ", ".join(f"{s} → '{a}'" for s, a in zip(im.loc[unmatched, "SKU"], im.loc[unmatched, "Alt Name"])))
    no_flag = has_alt & ~unmatched & (im["Alt Flag"] == "")
    log.add("Item_Master", "Alternate item given but Alternate Flag is blank", no_flag.sum(),
            "Not merged (R1 only covers SWITCH) - confirm with Purchase", "Flagged",
            ", ".join(f"{s} → {a}" for s, a in zip(im.loc[no_flag, "SKU"], im.loc[no_flag, "Alt SKU"])))
    unknown_flag = (im["Alt Flag"] != "") & (im["Alt Flag"] != "SWITCH")
    log.add("Item_Master", "Alternate Flag value not covered by the rules", unknown_flag.sum(),
            "Not merged - treated as a normal item", "Flagged",
            ", ".join(f"{s} '{f}' → {a}" for s, f, a in zip(im.loc[unknown_flag, "SKU"],
                                                          im.loc[unknown_flag, "Alt Flag"],
                                                          im.loc[unknown_flag, "Alt SKU"].fillna("?"))))

    # R1 - pack-size switches: old code -> new code
    sw = im[(im["Alt Flag"] == "SWITCH") & im["Alt SKU"].notna() & (im["Alt SKU"] != im["SKU"])]
    switch = {o: n for o, n in zip(sw["SKU"], sw["Alt SKU"])
              if status.get(o) == "Active" and status.get(n) == "Active"}

    def resolve(sku):
        seen = set()
        while sku in switch and sku not in seen:   # handles chains A -> B -> C
            seen.add(sku)
            sku = switch[sku]
        return sku

    log.add("Item_Master", "Pack-size switch (Alternate Flag = SWITCH)", len(switch),
            "Old and new code merged; reported under the new SKU, pack and cost (R1)", "Info",
            "; ".join(f"{o} → {n}" for o, n in switch.items()))
    cost_diff = [f"{o} ₹{cost[o]:g} vs {resolve(o)} ₹{cost[resolve(o)]:g}"
                 for o in switch if cost[o] != cost[resolve(o)]]
    log.add("Item_Master", "Unit cost changes between old and new code of a switch", len(cost_diff),
            "New code's cost used for aged value (R1) - worth confirming with Finance", "Flagged",
            "; ".join(cost_diff))

    # ======================= 2. STOCK AGING =======================
    stk = raw["stock"].copy()
    stk = stk[stk["SKU"].notna()].copy()
    stk["SKU"] = stk["SKU"].map(clean_sku)
    bcols = {}
    for key in ["0-30", "31-60", "61-90", "91-180", "180+"]:
        col = find_col(stk.columns, key)
        if col is None:
            raise ValueError(f"Stock_Aging: could not find the '{key} days' column")
        bcols[key] = col
    B = list(bcols.values())
    for c in B:
        stk[c] = pd.to_numeric(stk[c], errors="coerce")
    log.add("Stock_Aging", "Blank or non-numeric stock quantity", stk[B].isna().sum().sum(),
            "Treated as 0", "Fixed")
    stk[B] = stk[B].fillna(0)

    neg = [f"{r['SKU']} '{c}' = {r[c]:g}" for _, r in stk.iterrows() for c in B if r[c] < 0]
    log.add("Stock_Aging", "Negative stock quantity", len(neg),
            "Set to 0 - physical stock cannot be negative", "Fixed", "; ".join(neg))
    stk[B] = stk[B].clip(lower=0)

    d = stk["SKU"].duplicated()
    log.add("Stock_Aging", "SKU appears on more than one row", d.sum(), "Rows added together",
            "Fixed", ", ".join(sorted(set(stk.loc[d, "SKU"]))))
    stk = stk.groupby("SKU", as_index=False)[B].sum()
    stk["Total Stock"] = stk[B].sum(axis=1)
    stk["Aged Qty"] = stk[bcols["91-180"]] + stk[bcols["180+"]]          # R3

    unk = ~stk["SKU"].isin(im["SKU"])
    log.add("Stock_Aging", "Stock for a SKU that is not in Item_Master", unk.sum(), "Left out",
            "Excluded", ", ".join(stk.loc[unk, "SKU"]))
    stk = stk[~unk].copy()
    missing_stock = im[(im["Status"] == "Active") & ~im["SKU"].isin(stk["SKU"])]
    log.add("Stock_Aging", "Active item with no stock row", len(missing_stock), "Stock taken as 0",
            "Info", ", ".join(missing_stock["SKU"]))

    # ======================= 3. SALES LINES =======================
    sl = raw["sales"].copy().dropna(how="all")
    vcol = next((c for c in sl.columns if "value" in c.lower()), None)
    if vcol is None:
        raise ValueError("Sales_Lines: no Value column found")
    sl = sl.rename(columns={vcol: "Value"})
    for c in ["CustomerName", "Industry"]:
        if c not in sl.columns:
            sl[c] = ""
    sl["Excel Row"] = sl.index + 2

    raw_sku = sl["SKU"].astype(str)
    sl["SKU"] = sl["SKU"].map(clean_sku)
    fixed = sl["SKU"].notna() & (raw_sku != sl["SKU"])
    examples = ", ".join(f"'{x}'" for x in raw_sku[fixed].unique()[:5])
    log.add("Sales_Lines", "SKU code in lower case and/or with extra spaces", fixed.sum(),
            "Trimmed and changed to upper case so it matches the master", "Fixed", f"e.g. {examples}")
    blank = sl["SKU"].isna()
    log.add("Sales_Lines", "SKU missing", blank.sum(), "Row left out", "Excluded")
    sl = sl[~blank].copy()

    month_txt = sl["Month"].astype(str).str.strip()
    sl["Period"] = sl["Month"].map(parse_month)
    bad_m = sl["Period"].isna()
    log.add("Sales_Lines", "Month could not be read", bad_m.sum(), "Row left out", "Excluded",
            "Excel rows " + ", ".join(map(str, sl.loc[bad_m, "Excel Row"].head(20))))
    odd_m = ~month_txt.str.fullmatch(r"[A-Za-z]{3}-\d{4}") & ~bad_m
    log.add("Sales_Lines", "Month written in a different format", odd_m.sum(),
            "Read as the same calendar month as the standard format", "Fixed",
            "Formats seen: " + ", ".join(sorted(month_txt[odd_m].unique())[:5]))
    sl = sl[~bad_m].copy()
    sl["Period"] = pd.PeriodIndex(sl["Period"].tolist(), freq="M")

    sl["Qty"] = pd.to_numeric(sl["Qty"], errors="coerce")
    sl["Value"] = pd.to_numeric(sl["Value"], errors="coerce")
    badq = sl["Qty"].isna() | (sl["Qty"] <= 0) | sl["Value"].isna() | (sl["Value"] < 0)
    log.add("Sales_Lines", "Quantity or value missing, zero or negative", badq.sum(),
            "Row left out", "Excluded",
            "Excel rows " + ", ".join(map(str, sl.loc[badq, "Excel Row"].head(20))))
    sl = sl[~badq].copy()
    sl["CustomerID"] = sl["CustomerID"].map(clean_text).str.upper()

    # Exact duplicates
    key = ["Period", "CustomerID", "SKU", "Qty", "Value"]
    dmask = sl.duplicated(subset=key, keep="first")
    det = "; ".join(f"{month_label(r.Period)} {r.CustomerID} {r.SKU} × {r.Qty:g}"
                    for r in sl[dmask].itertuples())
    log.add("Sales_Lines", "Exact duplicate sales line", dmask.sum(),
            f"Second copy removed ({sl.loc[dmask, 'Qty'].sum():g} units, {inr(sl.loc[dmask, 'Value'].sum())})",
            "Fixed", det)
    sl = sl[~dmask].copy()

    # Same customer + month + qty + value on two different SKUs -> possibly wrong SKU keyed
    n_sku = sl.groupby(["Period", "CustomerID", "Qty", "Value"])["SKU"].transform("nunique")
    sus = sl[n_sku > 1]
    groups = sus.groupby(["Period", "CustomerID", "Qty"])
    det = "; ".join(f"{month_label(p)} {c}: {' & '.join(sorted(g['SKU']))} × {q:g}"
                    for (p, c, q), g in groups)
    log.add("Sales_Lines", "Same customer, month, qty and value on two different SKUs (possible wrong SKU)",
            groups.ngroups, "Kept as is - please check against the invoice", "Flagged", det)

    unk = ~sl["SKU"].isin(im["SKU"])
    log.add("Sales_Lines", "SKU not found in Item_Master", unk.sum(), "Row left out", "Excluded",
            ", ".join(sorted(sl.loc[unk, "SKU"].unique())))
    sl = sl[~unk].copy()

    # ======================= 4. CUSTOMERS =======================
    cu = raw["customers"].copy()
    for c in ["CustomerName", "Industry", "Area"]:
        if c not in cu.columns:
            cu[c] = ""
    cu = cu[cu["CustomerID"].notna()].copy()
    cu["CustomerID"] = cu["CustomerID"].map(clean_text).str.upper()
    for c in ["CustomerName", "Industry", "Area"]:
        cu[c] = cu[c].map(clean_text)
    dc = cu["CustomerID"].duplicated()
    log.add("Customers", "Same CustomerID listed more than once", dc.sum(), "Kept the first row",
            "Fixed", ", ".join(cu.loc[dc, "CustomerID"]))
    cu = cu[~dc].copy()
    no_ind = cu["Industry"] == ""
    log.add("Customers", "Industry missing", no_ind.sum(),
            "Not used for the 'Similar business' bucket (R6)", "Flagged",
            ", ".join(cu.loc[no_ind, "CustomerID"]))

    not_in_master = sorted(set(sl["CustomerID"]) - set(cu["CustomerID"]))
    log.add("Sales_Lines", "Customer not in the Customers sheet", len(not_in_master),
            "Kept; name and industry taken from the sales line", "Flagged", ", ".join(not_in_master))
    if not_in_master:
        extra = (sl[sl["CustomerID"].isin(not_in_master)]
                 .groupby("CustomerID", as_index=False)
                 .agg(CustomerName=("CustomerName", "first"), Industry=("Industry", "first")))
        extra["CustomerName"] = extra["CustomerName"].map(clean_text)
        extra["Industry"] = extra["Industry"].map(clean_text)
        extra["Area"] = ""
        cu = pd.concat([cu, extra], ignore_index=True)

    chk = sl[["CustomerID", "Industry"]].copy()
    chk["Industry"] = chk["Industry"].map(clean_text)
    chk = chk.merge(cu[["CustomerID", "Industry"]].rename(columns={"Industry": "Master"}),
                    on="CustomerID", how="left")
    chk["Master"] = chk["Master"].fillna("")
    mm = (chk["Master"] != "") & (chk["Industry"] != "") & \
         (chk["Industry"].str.lower() != chk["Master"].str.lower())
    log.add("Sales_Lines", "Industry on sales line differs from Customers sheet", mm.sum(),
            "Customers sheet used", "Fixed", ", ".join(sorted(chk.loc[mm, "CustomerID"].unique())))
    no_sales = sorted(set(cu["CustomerID"]) - set(sl["CustomerID"]))
    log.add("Customers", "Customer with no sales in the period", len(no_sales), "No action needed",
            "Info", ", ".join(no_sales))

    # ======================= 5. INACTIVE ITEMS (R2) =======================
    inact_sales = sl[sl["SKU"].map(status) != "Active"]
    log.add("Sales_Lines", "Sale of an inactive item", len(inact_sales),
            "Left out of item analysis; still counted in the customer's purchase value", "Flagged",
            "; ".join(f"{month_label(r.Period)} {r.CustomerID} {r.SKU} × {r.Qty:g}"
                      for r in inact_sales.itertuples()))
    inact_stock = stk[(stk["SKU"].map(status) == "Inactive") & (stk["Total Stock"] > 0)]
    warnings = []
    for _, r in inact_stock.iterrows():
        val = r["Total Stock"] * cost[r["SKU"]]
        warnings.append(
            f"{r['SKU']} ({name_of[r['SKU']]}) is Inactive but still holds {r['Total Stock']:g} units "
            f"({inr(val)} at cost, {r['Aged Qty']:g} units older than 90 days). It is left out of the review.")
    log.add("Stock_Aging", "Inactive item still holds stock", len(inact_stock),
            "Left out of the analysis; warning shown on the Summary", "Flagged", " ".join(warnings))

    # ======================= 6. PERIOD =======================
    periods = sorted(sl["Period"].unique())
    if not periods:
        raise ValueError("no usable sales lines")
    n_months = len(periods)
    period_txt = f"{month_label(periods[0])} to {month_label(periods[-1])}"
    log.add("Sales_Lines", "Months of sales in the file", n_months,
            f"{period_txt}; average monthly sales = total ÷ {n_months}", "Info")

    # ======================= 7. CUSTOMER VALUE (R6) =======================
    cust_val = sl.groupby("CustomerID")["Value"].sum()
    cu["6M Purchase Value"] = cu["CustomerID"].map(cust_val).fillna(0)
    cu["Target Customer"] = cu["6M Purchase Value"] > TARGET_MIN_PURCHASE

    # ======================= 8. ITEM CLASSIFICATION (R1, R3, R4, R5) =======================
    act = im[im["Status"] == "Active"].copy()
    act_set = set(act["SKU"])
    report = act[~act["SKU"].isin(switch.keys())].copy()     # old switch codes fold into new ones
    report["Merged From"] = report["SKU"].map(
        lambda s: ", ".join(sorted(o for o in switch if resolve(o) == s)))

    st_act = stk[stk["SKU"].isin(act_set)].copy()
    st_act["Report SKU"] = st_act["SKU"].map(resolve)
    sg = st_act.groupby("Report SKU")[["Total Stock", "Aged Qty"]].sum()

    sa = sl[sl["SKU"].isin(act_set)].copy()
    sa["Report SKU"] = sa["SKU"].map(resolve)
    sold = sa.groupby("Report SKU")["Qty"].sum()

    report["Total Stock"] = report["SKU"].map(sg["Total Stock"]).fillna(0)
    report["Aged Qty"] = report["SKU"].map(sg["Aged Qty"]).fillna(0)
    report["Aged Value"] = report["Aged Qty"] * report["Unit Cost"]
    report["Sold 6M"] = report["SKU"].map(sold).fillna(0)
    report["Avg / Month"] = report["Sold 6M"] / n_months
    report["Months to Clear"] = np.where(report["Avg / Month"] > 0,
                                         report["Total Stock"] / report["Avg / Month"].replace(0, np.nan),
                                         np.inf)
    report["Class"] = report.apply(classify, axis=1)
    report["Why"] = report.apply(reason, axis=1)

    items = report[["SKU", "Item Name", "Category", "Class", "Why", "Merged From", "Master Pack",
                    "Unit Cost", "Total Stock", "Aged Qty", "Aged Value", "Sold 6M",
                    "Avg / Month", "Months to Clear"]].copy()
    items["Months to Clear"] = items["Months to Clear"].replace(np.inf, np.nan)
    items = (items.assign(_o=items["Class"].map(CLASS_ORDER))
                  .sort_values(["_o", "Aged Value"], ascending=[True, False])
                  .drop(columns="_o").reset_index(drop=True))

    # ======================= 9. CUSTOMER TARGETS (R6, R7) =======================
    cat_of = dict(zip(items["SKU"], items["Category"]))
    buyers = sa.groupby("Report SKU")["CustomerID"].agg(lambda s: set(s)).to_dict()
    bought = sa.groupby("CustomerID")["Report SKU"].agg(lambda s: set(s)).to_dict()
    ind_of = {c: i.lower() for c, i in zip(cu["CustomerID"], cu["Industry"])}
    lp = (sa.groupby(["Report SKU", "CustomerID", "Period"], as_index=False)["Qty"].sum()
            .sort_values("Period"))
    lastbuy = lp.groupby(["Report SKU", "CustomerID"]).last()

    tc = cu[cu["Target Customer"]]
    push = items[items["Class"] == "Push"]
    rows = []
    for _, it in push.iterrows():
        sku = it["SKU"]
        b = buyers.get(sku, set())
        b_ind = {ind_of.get(c, "") for c in b} - {""}
        for _, c in tc.iterrows():
            cid, cind = c["CustomerID"], c["Industry"].lower()
            if cid in b:
                bucket = 1
            elif cind and cind in b_ind:
                bucket = 2
            elif any(cat_of.get(x) == it["Category"] for x in bought.get(cid, set()) if x != sku):
                bucket = 3
            else:
                continue
            last = ""
            if bucket == 1 and (sku, cid) in lastbuy.index:
                lb = lastbuy.loc[(sku, cid)]
                last = f"{month_label(lb['Period'])} ({lb['Qty']:g} units)"
            rows.append({
                "SKU": sku, "Item Name": it["Item Name"], "Category": it["Category"],
                "Item Aged Value": it["Aged Value"], "Bucket": bucket,
                "Bucket Name": BUCKET_NAMES[bucket], "CustomerID": cid,
                "CustomerName": c["CustomerName"], "Industry": c["Industry"], "Area": c["Area"],
                "6M Purchase Value": c["6M Purchase Value"], "Last Bought": last,
            })
    tg_cols = ["SKU", "Item Name", "Category", "Item Aged Value", "Bucket", "Bucket Name",
               "CustomerID", "CustomerName", "Industry", "Area", "6M Purchase Value", "Last Bought"]
    tg = pd.DataFrame(rows, columns=tg_cols)
    if not tg.empty:
        tg = tg.sort_values(["Item Aged Value", "SKU", "Bucket", "6M Purchase Value"],
                            ascending=[False, True, True, False]).reset_index(drop=True)
        tg.insert(0, "Call Order", tg.groupby("SKU").cumcount() + 1)
    else:
        tg.insert(0, "Call Order", [])
    no_tg = sorted(set(push["SKU"]) - set(tg["SKU"]))
    log.add("Result", "Push item with no matching target customer", len(no_tg),
            "Shown with an empty call list", "Info", ", ".join(no_tg))

    # ======================= 10. OUTPUTS =======================
    counts = items["Class"].value_counts()
    summary = {
        "Push": int(counts.get("Push", 0)),
        "No Action": int(counts.get("No Action", 0)),
        "Red List": int(counts.get("Red List", 0)),
        "aged_push": float(items.loc[items["Class"] == "Push", "Aged Value"].sum()),
        "aged_red": float(items.loc[items["Class"] == "Red List", "Aged Value"].sum()),
        "items": len(items), "targets": int(tc.shape[0]),
        "period": period_txt, "n_months": n_months, "inactive": warnings,
    }
    cust_out = (cu.sort_values("6M Purchase Value", ascending=False)
                  [["CustomerID", "CustomerName", "Industry", "Area", "6M Purchase Value", "Target Customer"]])
    clean = sl.copy()
    clean["Month"] = clean["Period"].map(month_label)
    clean["Report SKU"] = clean["SKU"].map(lambda s: resolve(s) if status.get(s) == "Active" else "(inactive)")
    clean = clean[["Month", "CustomerID", "CustomerName", "Industry", "SKU", "Report SKU", "Qty", "Value"]]

    return {"items": items, "targets": tg, "customers": cust_out,
            "issues": log.frame(), "clean_sales": clean, "summary": summary}


@st.cache_data(show_spinner=False)
def process(blobs):
    raw, notes = load_files(blobs)
    missing = [SHEET_NAMES[k] for k in SHEET_NAMES if k not in raw]
    if missing:
        return {"error": "Could not find: " + ", ".join(missing)
                         + ". Upload the full workbook, or all four sheets."}
    try:
        return run_review(raw, notes)
    except (KeyError, ValueError) as e:
        return {"error": f"The files don't look like the usual export ({e}). Check the column headings."}


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def build_excel(res):
    s, items = res["summary"], res["items"]
    summary = pd.DataFrame(
        [("Sales period", s["period"]),
         ("Active items analysed", s["items"]),
         ("Push items", s["Push"]),
         ("No Action items", s["No Action"]),
         ("Red List items", s["Red List"]),
         ("Aged value in Push items (Rs)", round(s["aged_push"])),
         ("Aged value in Red List items (Rs)", round(s["aged_red"])),
         ("Target customers (above Rs 1,00,000)", s["targets"]),
         ("Exported on", datetime.now(IST).strftime("%d %b %Y %H:%M"))]
        + [("Warning", w) for w in s["inactive"]],
        columns=["Measure", "Value"])
    sheets = {
        "Summary": summary,
        "Top 10 Push": items[items["Class"] == "Push"].head(TOP_N),
        "Item List": items,
        "Call Lists": res["targets"],
        "Red List": items[items["Class"] == "Red List"],
        "Data Checks": res["issues"],
        "Customers": res["customers"],
        "Clean Sales": res["clean_sales"],
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Screens
# ---------------------------------------------------------------------------
def wa_link(r, name):
    item = re.sub(r"\s*\(.*?\)", "", r["Item Name"]).strip()
    me = f"this is {name} from Company X" if name else "greetings from Company X"
    if r["Bucket"] == 1:
        hook = f"You have ordered {item} from us before, and we have ready stock for immediate dispatch."
    elif r["Bucket"] == 2:
        hook = (f"{item} is a regular buy for {r['Industry'].lower()} businesses we supply, "
                f"and we have ready stock for immediate dispatch.")
    else:
        hook = (f"Since you already buy {r['Category']} items from us, {item} could be a useful add-on. "
                f"We have ready stock for immediate dispatch.")
    msg = f"Hello {r['CustomerName']}, {me}. {hook} Shall I add some to your next order?"
    return "https://wa.me/?text=" + urllib.parse.quote(msg)


def customer_card(r, name):
    last = f" · last took it {r['Last Bought']}" if r["Last Bought"] else ""
    with st.container(border=True):
        st.markdown(
            f"**{r['Call Order']}. {r['CustomerName']}** ({r['CustomerID']})  \n"
            f"Bucket {r['Bucket']} – {r['Bucket Name']}  \n"
            f"{r['Industry'] or 'Industry not recorded'} · {r['Area']} · "
            f"bought {inr(r['6M Purchase Value'])} in the period{last}")
        st.link_button("WhatsApp draft", wa_link(r, name), use_container_width=True)


def show_summary(res):
    s, items = res["summary"], res["items"]
    c = st.columns(4)
    c[0].metric("Push", s["Push"])
    c[1].metric("No Action", s["No Action"])
    c[2].metric("Red List", s["Red List"])
    c[3].metric("Aged value in Push", inr(s["aged_push"]))
    st.caption(f"{s['items']} active items analysed (pack-size switches merged) · "
               f"{s['targets']} customers bought more than ₹1,00,000 · "
               f"Red List holds {inr(s['aged_red'])} of aged stock")
    for w in s["inactive"]:
        st.warning(w)
    st.subheader(f"Top {TOP_N} Push items by aged value")
    top = items[items["Class"] == "Push"].head(TOP_N)
    st.dataframe(top[["SKU", "Item Name", "Category", "Aged Qty", "Aged Value", "Months to Clear"]],
                 hide_index=True, use_container_width=True, column_config=col_cfg())
    if not top.empty:
        st.bar_chart(top.set_index("SKU")["Aged Value"], height=260)


def show_call_list(res):
    tg, items = res["targets"], res["items"]
    push = items[items["Class"] == "Push"]
    if push.empty or tg.empty:
        st.info("No Push items with target customers this month.")
        return
    st.text_input("Your name (goes into the WhatsApp message)", key="exec_name", placeholder="e.g. Ravi")
    name = st.session_state.get("exec_name", "").strip()
    mode = st.radio("View", ["By item", "By customer"], horizontal=True, label_visibility="collapsed")

    if mode == "By item":
        cats = ["All categories"] + sorted(push["Category"].unique())
        cat = st.selectbox("Category", cats)
        if cat != cats[0]:
            push = push[push["Category"] == cat]
        labels = {r["SKU"]: f"{r['SKU']} · {r['Item Name']} · {inr(r['Aged Value'])} aged"
                  for _, r in push.iterrows()}
        sku = st.selectbox("Item to push", list(labels), format_func=labels.get)
        it = push.set_index("SKU").loc[sku]
        st.markdown(f"**{it['Item Name']}** ({it['Category']})  \n"
                    f"Stock {it['Total Stock']:,.0f} · older than 90 days {it['Aged Qty']:,.0f} "
                    f"({inr(it['Aged Value'])}) · clears in {it['Months to Clear']:.1f} months at current pace")
        if it["Merged From"]:
            st.caption(f"Includes old code {it['Merged From']} (pack-size switch). Offer {sku}.")
        rows = tg[tg["SKU"] == sku]
        if rows.empty:
            st.info("No customer above ₹1,00,000 fits this item.")
            return
        st.caption(f"{len(rows)} customers, in calling order")
        for _, r in rows.head(TOP_N).iterrows():
            customer_card(r, name)
        if len(rows) > TOP_N:
            with st.expander(f"Show {len(rows) - TOP_N} more"):
                for _, r in rows.iloc[TOP_N:].iterrows():
                    customer_card(r, name)
    else:
        cust = tg.drop_duplicates("CustomerID").sort_values("6M Purchase Value", ascending=False)
        labels = {r["CustomerID"]: f"{r['CustomerName']} · {r['Industry'] or 'no industry'} · {r['Area']}"
                  for _, r in cust.iterrows()}
        cid = st.selectbox("Customer", list(labels), format_func=labels.get)
        rows = tg[tg["CustomerID"] == cid].sort_values(["Bucket", "Item Aged Value"], ascending=[True, False])
        st.caption(f"{len(rows)} items to offer · bought {inr(rows['6M Purchase Value'].iloc[0])} in the period")
        for _, r in rows.iterrows():
            with st.container(border=True):
                st.markdown(f"**{r['Item Name']}** ({r['SKU']})  \n"
                            f"Bucket {r['Bucket']} – {r['Bucket Name']} · {r['Category']} · "
                            f"{inr(r['Item Aged Value'])} aged")
                st.link_button("WhatsApp draft", wa_link(r, name), use_container_width=True)


def show_items(res):
    items = res["items"]
    c1, c2 = st.columns(2)
    classes = c1.multiselect("Class", list(CLASS_ORDER), default=list(CLASS_ORDER))
    cats_all = sorted(items["Category"].unique())
    cats = c2.multiselect("Category", cats_all, default=cats_all)
    q = st.text_input("Search SKU or item name").strip()
    view = items[items["Class"].isin(classes) & items["Category"].isin(cats)]
    if q:
        view = view[view["SKU"].str.contains(q, case=False, regex=False)
                    | view["Item Name"].str.contains(q, case=False, regex=False)]
    st.caption(f"{len(view)} of {len(items)} items. Red List items are handled separately and get no customer list.")
    st.dataframe(view, hide_index=True, use_container_width=True, column_config=col_cfg())


def show_checks(res):
    iss = res["issues"]
    n = int((iss["Status"] != "Info").sum())
    st.markdown(f"**{n} problems found.** *Fixed* = corrected automatically. *Excluded* = left out. "
                f"*Flagged* = kept as is, needs a check at source. *Info* = for awareness.")
    st.dataframe(iss, hide_index=True, use_container_width=True)


def show_export(res):
    st.write("One Excel file with the summary, item list, call lists for every Push item, "
             "Red List, data checks, customers and the cleaned sales lines.")
    stamp = datetime.now(IST).strftime("%Y-%m-%d")
    st.download_button("Download Excel", data=build_excel(res),
                       file_name=f"slow_stock_review_{stamp}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)


def main():
    st.title("Slow-Stock Review")
    st.caption("Upload the month's exports, check the data, see what to push and who to call.")

    res = st.session_state.get("res")
    with st.expander("Upload this month's files", expanded=res is None):
        files = st.file_uploader("The full workbook, or the four sheets as separate files",
                                 type=["xlsx", "csv"], accept_multiple_files=True)
        c1, c2 = st.columns(2)
        run_clicked = c1.button("Run review", type="primary", disabled=not files, use_container_width=True)
        sample_clicked = c2.button("Try with sample data", disabled=not SAMPLE_FILE.exists(),
                                   use_container_width=True)
        if run_clicked or sample_clicked:
            blobs = (tuple((f.name, f.getvalue()) for f in files) if run_clicked
                     else ((SAMPLE_FILE.name, SAMPLE_FILE.read_bytes()),))
            with st.spinner("Checking the data and applying the rules..."):
                out = process(blobs)
            if "error" in out:
                st.error(out["error"])
            else:
                st.session_state["res"] = out
                st.session_state["run_at"] = datetime.now(IST).strftime("%d %b %Y, %I:%M %p")
                st.rerun()

    if res is None:
        st.info("Upload the files above and press **Run review**. It takes a few seconds.")
        return

    st.caption(f"Last run {st.session_state.get('run_at', '')} · sales {res['summary']['period']}")
    t1, t2, t3, t4, t5 = st.tabs(["Summary", "Call list", "Items", "Data checks", "Export"])
    with t1:
        show_summary(res)
    with t2:
        show_call_list(res)
    with t3:
        show_items(res)
    with t4:
        show_checks(res)
    with t5:
        show_export(res)


main()