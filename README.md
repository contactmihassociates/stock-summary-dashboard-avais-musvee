# Stock Summary Dashboard

A Power BI–style monthly closing-stock dashboard for a textile mill, generated
from the monthly `*Stock statement.xlsx` files (Yarn → Sizing → Loom → Grey →
Dyeing → Stitching/Blanket → Finished Goods).

## Monthly workflow

1. Drop the new month's `XXX-YY Stock statement.xlsx` into this folder
   (same sheet layout as previous months).
2. Run:

   ```
   python build_dashboard.py
   ```

3. Open `Stock_Dashboard.html` in any browser.

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
- **Trend & composition** — multi-month value trend, stock mix doughnut
- **Production pipeline strip** — WIP value/weight per stage with MoM arrows
- **Finished goods by customer** — bars colored by movement
  (green = grew, red = drew down, yellow = frozen)
- **Action Center** — automated alerts: concentration risk, stale stock,
  no-movement customers, WIP swings, zero-stock customers
- **Registers** — sortable customer stock and raw-material yarn tables with
  rate/kg sanity checks

Requires Python 3 with `openpyxl` (`pip install openpyxl`) to rebuild.
The raw xlsx statements are intentionally **not** committed (see `.gitignore`).
