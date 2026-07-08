# Adding a new fund

The app pulls fund prices from 4 Google Sheets (Cash, Gold, Equity, Debt) and
uses `fund_mapping.csv` to know which column in which sheet belongs to which
fund. To add a new fund, update both places.

## `fund_mapping.csv` columns

| Column | What to put there |
|---|---|
| `isin` | The fund's ISIN, exactly as it appears in the trade master / custodian holdings CSV. |
| `fund_name` | A readable name, just for your own reference. Not read by the app. |
| `category` | One of: `Indian Equity`, `Foreign Equity`, `Gold`, `Debt`, `Cash Fund`. Drives allocation grouping everywhere in the app. |
| `source` | Which Google Sheet the fund's price history lives in: `cash`, `gold`, `equity`, or `debt`. |
| `sheet_fund_name` | The **exact** column header text for that fund in the sheet named above. Must match character-for-character, including spacing/capitalization. |

## Steps to add a fund

1. **Open the right Google Sheet** (Cash / Gold / Equity / Debt) and add a new
   column with the fund's daily price history. The column header is whatever
   name you choose — you'll reuse it in step 3.
2. **Open `fund_mapping.csv`** in Excel, Google Sheets, or a text editor.
3. **Add one new row** with:
   - `isin` — the fund's ISIN
   - `fund_name` — any readable label
   - `category` — pick from the 5 allowed values above (spelling must match exactly)
   - `source` — which sheet you just edited (`cash` / `gold` / `equity` / `debt`)
   - `sheet_fund_name` — the exact column header you used in step 1
4. Save. Next time the app loads a dashboard, it picks up the new fund
   automatically — no code changes, no redeploy.

## If a fund isn't in a Google Sheet yet

Leave `source` and `sheet_fund_name` blank, but still fill in `category`. The
app will fall back to the custodian snapshot's Unit Price, or the fund's most
recent transaction NAV, while still classifying it correctly by category.
Rows like this already exist in the file (e.g. the two Franklin funds and
ICICI Prudential Regular Gold Savings Fund FOF) — nobody's added those
columns to the Equity/Gold sheets yet.

## Common mistakes

- **Typo in `sheet_fund_name`**: even one extra space or a different
  capitalization means the app can't find the column, and that fund silently
  falls back to a stale/estimated NAV. Copy-paste the header text directly
  from the sheet rather than retyping it.
- **Wrong `source`**: if you put a fund's data in the Equity sheet but write
  `source,debt` in the mapping row, the app looks in the wrong sheet and
  won't find it.
- **Category spelling**: must be exactly `Indian Equity`, `Foreign Equity`,
  `Gold`, `Debt`, or `Cash Fund` — anything else won't group correctly on
  the dashboards.
