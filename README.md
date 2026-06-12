# Stock Summary Dashboard

A Power BI–style monthly closing-stock dashboard for a textile mill, generated
from the monthly `*Stock statement.xlsx` files (Yarn → Sizing → Loom → Grey →
Dyeing → Stitching/Blanket → Finished Goods).

## Updating the data — two ways

**A. In the browser (quick, personal):** open the dashboard, click
**⚙ Manage Data**, and drop the new month's
`XXX-YY Stock statement.xlsx` straight in. It is parsed client-side
(SheetJS), saved in your browser's local storage, and every chart, alert
and table updates instantly. Works on the hosted Vercel site too.
Uploads are visible only in that browser.

**B. Permanent (publishes for everyone):**

1. Drop the new month's xlsx into this folder (same sheet layout).
2. Double-click **`Update Dashboard.bat`** — it rebuilds and opens the
   dashboard. (Or run `python build_dashboard.py`.)
3. `git add -A && git commit && git push` — Vercel redeploys automatically.

The trend chart, month-over-month comparisons, and automated alerts extend to
every month on file automatically.

## What's inside

| File | Purpose |
|---|---|
| `extract_stock_data.py` | Parses the SUMMARY and Store-RM sheets of every monthly xlsx into `stock_data.json` |
| `dashboard_template.html` | Dashboard design (HTML/CSS/JS, Chart.js) with a `__STOCK_DATA__` placeholder |
| `build_dashboard.py` | Runs the extractor and injects the data into the template → `Stock_Dashboard.html` |
| `Stock_Dashboard.html` | The generated, self-contained dashboard — just double-click |

## Dashboard features

- **Hero KPIs** — total stock, finished goods, WIP, raw material, customer
  concentration, no-movement stock — all with MoM deltas
- **Executive summary** — auto-written plain-language commentary of the
  month (total change, biggest driver, risks) with copy-to-clipboard
- **MoM bridge (waterfall)** — decomposes the change in total stock into
  Raw Material / Chemicals / WIP / Finished Goods / Numbal contributions
- **Top Movers** — the eight largest value swings across all line items
- **Trend & composition** — multi-month value trend, stock mix doughnut
- **Production pipeline strip** — WIP value/weight per stage with MoM arrows
- **Finished goods by customer** — bars colored by movement
  (green = grew, red = drew down, yellow = frozen)
- **Action Center** — automated alerts: concentration risk, stale stock,
  no-movement customers and yarn lots, WIP swings, zero-stock customers,
  data-quality failures
- **Registers** — sortable, searchable customer and raw-material tables with
  rate/kg sanity checks, per-row MoM deltas, status badges and CSV export
- **Controls** — month selector, free choice of comparison month,
  Value ₹ / Weight kg view toggle, Print/PDF with a dedicated print layout
- **Data-quality gate** — the extractor verifies components reproduce each
  sheet's GRAND TOTAL and raises a red alert on mismatch

Requires Python 3 with `openpyxl` (`pip install openpyxl`) to rebuild.
The raw xlsx statements are intentionally **not** committed (see `.gitignore`).

## Stock Register ERP (`erp.html`)

A data-entry-first companion module reachable from the dashboard header
(📝 Stock Register ERP). Three role-gated workspaces:

- **Registrar** — keyboard-first movement entry (received / issued /
  damaged per item & batch). Opening and closing stock auto-calculate;
  over-issuing beyond available stock and negative quantities are
  blocked; entries in accountant-locked months are rejected.
- **Accountant** — live stock valuation under **FIFO / LIFO / Weighted
  Average** (lot-based engine), per-entry approve / flag-with-note
  workflow (flagged entries are excluded from stock until corrected),
  and month-end **lock** that freezes all entries in the month.
- **Executive view** — total inventory value, pending review count,
  low-stock alerts (reorder levels), dead/slow-moving stock (no issues
  in 30 days), shrinkage rate, daily movement chart and value-by-category
  doughnut, all updating live as data is entered.

The item master **auto-seeds on first load** with ~169 items (store
yarn, loom weft, dyed-yarn lots, chemicals, fuel) from the latest
embedded statement, complete with opening stock, rates and suggested
reorder levels. Also included: searchable type-ahead item picker,
per-item ledger with running balance, printable **month-end stock
report in the mill's R/MNT/10 paper format**, physical **stock-take
mode** (variances auto-post as pending adjustments), top-consumption
ranking, bulk approval with status filters, an accountant PIN guard,
CSV exports throughout, a full audit trail, and JSON Backup/Restore
(with a 7-day backup reminder). Data persists in the browser
(localStorage) — swap in a shared backend later if multi-user
concurrency is needed.

## Hosting on Vercel

The build emits `index.html` (identical to `Stock_Dashboard.html`), so the
repo deploys as a plain static site:

1. Go to [vercel.com/new](https://vercel.com/new) and import this GitHub repo.
2. Framework preset: **Other**. Leave build command and output directory
   **empty** (no build step — `index.html` is committed).
3. Deploy. Every `git push` to `main` redeploys automatically.

To publish a new month: run `Update Dashboard.bat`, then
`git add -A && git commit -m "JUN-26" && git push`.
