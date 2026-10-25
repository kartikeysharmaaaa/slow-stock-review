# Slow-Stock Review – Company X

Turns the four monthly system exports (Item_Master, Stock_Aging, Sales_Lines, Customers)
into a checked, rule-based slow-stock action list with customer call lists for sales.

**Live app:** <paste link after deploying>

## What it does

1. **Upload** the workbook or the four sheets (sheets are detected by their columns, so separate files named "Sheet1" are fine).
2. **Data checks:** every problem found and how it was handled. Nothing is changed silently.
3. **Rules R1–R7:** pack-size switches merged, inactive items excluded, aged stock, months to clear,
   Red List / No Action / Push, customer buckets and call order.
4. **Mobile call list** for sales with WhatsApp drafts (no API key needed), by item or by customer.
5. **Excel export** of the summary, item list, call lists, Red List, data checks, customers and cleaned sales.

## Project layout

```
.
├── app.py                  # the whole Streamlit app
├── requirements.txt        # pinned dependencies
├── .streamlit/config.toml  # theme
└── case_data.xlsx          # sample workbook used by "Try with sample data"
```

`case_data.xlsx` is the case data: the four source exports combined into one workbook, one sheet each,
with the values unchanged. Without it the app still works, but the "Try with sample data" button is disabled.

## Run locally

Requires Python 3.12.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Community Cloud)

1. Push this repo to GitHub.
2. On share.streamlit.io, sign in with GitHub → **Create app** → pick the repo, branch `main`, main file `app.py`.
3. Under **Advanced settings**, set Python to 3.12, then **Deploy**.

Free apps sleep when unused; one click on the "wake up" screen brings it back in about 30 seconds.

## Change a rule

Thresholds are at the top of `app.py`:

| Constant | Default | Meaning |
|---|---|---|
| `MONTHS_TO_CLEAR_LIMIT` | `3` | R5: items clearing faster than this are No Action |
| `TARGET_MIN_PURCHASE` | `100_000` | R6: only customers above this purchase value are targeted |
| `TOP_N` | `10` | How many items / customers to show first |

## Expected result on the sample data

Push 18 · No Action 13 · Red List 7 · Aged value in Push ₹1,63,841.
