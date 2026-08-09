# FNP Sales Analysis

Sales analysis of **Ferns N Petals**, an Indian gifting retailer — 1,000 orders across 2023, delivered two ways: an interactive Excel dashboard and a self-contained web dashboard.

**The finding: four festival months carry 68% of the year's revenue.**


---

## The headline

Valentine's (Feb), Holi (Mar), Raksha Bandhan (Aug) and Diwali (Nov) compress the year into four narrow windows. A festival month averages **₹6,00,722** against **₹1,39,762** in a quiet one — roughly **4.3× the run rate**.

That reframes the business. Inventory, staffing and delivery capacity are a *calendar* problem before they are a demand problem.

| Metric | Value |
|---|---|
| Total revenue | ₹35,20,984 |
| Orders | 1,000 |
| Units shipped | 3,045 |
| Average order value | ₹3,521 |
| Revenue per customer | ₹35,210 |
| Average delivery time | 5.5 days |
| Delivered within 7 days | 64% |
| Delivery cities served | 301 |

---

## What the data says

**Bundle size is the strongest lever on revenue.** Five-unit baskets are 21.9% of orders but **35.9%** of revenue. Single-unit orders are almost the same share of orders (19.6%) and only 6.4% of revenue.

**Volume and value point at different occasions.** Anniversary brings the most orders (205) but Raksha Bandhan earns the highest average order value (₹4,785) — 75% above the weakest occasion.

**Premium stock does the earning.** Items above ₹1,000 are 60% of orders and 80% of revenue.

**Gifts don't ship to the buyer.** `Location` is the *delivery* city and matches the buyer's own city in **0.1%** of orders — 301 delivery cities against 86 buyer cities. "Top cities" means something completely different depending on which column you use.

**The delivery tail is the real risk.** The 5.5-day average hides a spread: 46 orders land past ten days. For a dated occasion, a gift that misses the date has failed regardless of the mean.

**Buyer gender is not a useful segmentation axis here** — revenue splits 51 / 49 between male and female buyers.



## Two dashboards

### Excel


Two dropdowns (Occasion, Category) drive 11 charts, 6 KPI tiles, the ranking tables and the matrix. 16,979 formulas, zero hardcoded results — edit a price in the `Products` sheet and the whole dashboard moves.



| Sheet | Purpose |
|---|---|
| `Dashboard` | KPI tiles, 11 charts, top-customer table, category × occasion matrix |
| `Data Dictionary` | Every field with type, key and known traps |
| `Calcs` | Aggregation tables — `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` |
| `Model` | 1,000-row fact table; every derived column a live `INDEX`/`MATCH` |
| `Orders`, `Products`, `Customers` | Cleaned source data |
| `Lists` | Lookups for months, weekdays, price bands, delivery buckets |

**How the filters work.** The dropdowns write into `Calcs!B3` and `Calcs!B4`, which resolve `(All)` into the wildcard `"*"`:

```excel
=IF(Dashboard!$I$7="(All)", "*", Dashboard!$I$7)
```

Every aggregation then reads those two cells as criteria, so one formula pattern serves both the filtered and unfiltered case:

```excel
=SUMIFS(Model!$N$2:$N$1001, Model!$J$2:$J$1001, $B$3, Model!$L$2:$L$1001, $B$4, Model!$R$2:$R$1001, $B7)
```

This replaces PivotTable slicers with plain formulas, so the workbook needs no PivotCache and stays portable.

### Web

one file, no build step, no dependencies. Open it in any browser.

17 views, live filtering, hover detail on every chart, and a light/dark toggle. Charts are hand-written SVG; the dataset is embedded as compact JSON (~35 KB) and every aggregation runs client-side.

> Serve it on GitHub Pages by enabling Pages on the `/docs` folder — `docs/index.html` is a copy.



---

## Data quality notes

These cost me real debugging time. If you work with this dataset, read them first.

**Product names are not unique.** Two names repeat with *different* categories and prices:

| ID | Name | Category | Price |
|---|---|---|---|
| 21 | Quia Gift | Colors | ₹1,561 |
| 56 | Quia Gift | Raksha Bandhan | ₹1,272 |
| 65 | Error Gift | Raksha Bandhan | ₹1,895 |
| 70 | Error Gift | Sweets | ₹866 |

Grouping products by name inflates "Quia Gift" to ₹1,14,476 and pushes two genuine top-10 products off the chart. **Always aggregate on `Product_ID`.**

**Dates are `DD-MM-YYYY`.** Left to Excel's or pandas' default parser, `07-11-2023` becomes 11 July and the entire monthly trend shifts. Parse the format explicitly.

**`Contact_Number` is stored as an integer**, so any leading zero is already gone. Treat it as text on reload.

**Two different `Occasion` columns exist.** `orders.Occasion` is what the order was placed *for*; `products.Occasion` is a merchandising attribute. They disagree often. These dashboards use the order-level column.

**Six deliveries slip into January 2024** — late-December orders. Filtering the delivery date to 2023 silently drops them.

**`Address` is multi-line** in the raw file, with embedded newline escapes that break naive CSV reads.

There are no missing values and no duplicate `Order_ID`s.


---

## Data model

Two one-to-many relationships. **Revenue is derived — there is no revenue column in any source file.**

```
customers.csv                orders.csv                 products.csv
100 rows, 7 fields           1,000 rows, 10 fields      70 rows, 6 fields
                                                     
  Customer_ID (PK) ◄──1:many── Customer_ID (FK)
                               Order_ID (PK)
                               Product_ID (FK) ──many:1──► Product_ID (PK)

  Revenue = orders.Quantity × products.[Price (INR)]
  Lead_Days = (Delivery_Date + Delivery_Time) − (Order_Date + Order_Time)
```

Lead time is measured to the hour, not rounded to whole days.

---

## Reproduce the figures

Every number quoted above is recomputed from the raw CSVs:

```bash
pip install -r scripts/requirements.txt
python scripts/analysis.py
```

Output reconciles exactly with the Excel workbook — total revenue, order counts, all breakdowns, and the data-quality checks.

---

## Repository layout

```
.
├── data/                       # source CSVs, unmodified
│   ├── customers.csv
│   ├── orders.csv
│   └── products.csv
├── dashboard/
│   ├── FNP_Sales_Dashboard.xlsx
│   └── FNP_Sales_Dashboard.html
├── docs/
│   ├── FNP_Sales_Analysis_Summary.pdf
│   ├── summary-source.html     # the PDF's HTML source
│   ├── index.html              # web dashboard, for GitHub Pages
│   └── DATA_DICTIONARY.md
├── images/
├── scripts/
│   ├── analysis.py             # reproducible figures
│   └── requirements.txt
├── LICENSE
└── README.md
```

---

## Built with

Excel (formulas, charts, data validation, conditional formatting) · Python (pandas) for verification · hand-written SVG and vanilla JS for the web dashboard.

---

## A note on the data

This is a **sample dataset** used for analysis practice. Product names and descriptions are placeholder Latin text, and customer records are generated — so treat the product-level findings as an exercise in method rather than a read on the real Ferns N Petals catalogue. The structural findings (festival seasonality, basket-size economics, the delivery-city gap) are properties of the dataset as given.

Not affiliated with or endorsed by Ferns N Petals.

## License

[MIT](LICENSE) — code and analysis. The sample dataset is included for demonstration only.
