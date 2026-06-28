# D2C Sales Performance & Competitive Pricing Dashboard

An end-to-end ETL and sales/CRM analytics project: extracting data from two separate sources, validating and cleaning it, transforming it into pipeline and revenue KPIs, and presenting it through an executive dashboard — the same workflow used in Revenue Operations, Sales Analytics, and Product & Price Management roles.

📊 **[Download the dashboard](D2C_Sales_Performance_Competitive_Pricing_Dashboard.xlsx)**

## What this demonstrates

- **ETL across multiple sources** — a synthetic CRM export (850 deals) and a competitor pricing feed (42 price points across 6 brands), each with realistic data-quality issues
- **Data validation & cleaning** — documented rules for fixing inconsistent region naming, missing sales rep values, and outlier checks (see the `Data_Cleaning_Log` sheet)
- **Live Excel formulas** — SUMIFS, COUNTIFS, AVERAGEIFS, and helper-column-driven MIN/MAX logic; nothing is hardcoded, so the whole workbook recalculates if the raw data changes
- **Sales pipeline analysis** — funnel by stage, win rate and sales cycle by region, performance by product line, lead-source conversion, and a rep leaderboard
- **Competitive pricing benchmark** — own price vs. 6 competitors per product line, with automatic over/underpricing flags
- **Executive dashboard** — KPI cards and charts built for a management/sales audience

## Repo structure

```
├── D2C_Sales_Performance_Competitive_Pricing_Dashboard.xlsx   # the deliverable
├── data/
│   ├── crm_export_raw.csv          # synthetic CRM export (source 1)
│   └── competitor_pricing_raw.csv  # synthetic competitor pricing feed (source 2)
├── generate_data.py                # generates the synthetic datasets
├── build_workbook.py                # builds the full Excel workbook (sheets, formulas, charts) from the CSVs
└── README.md
```

## How it was built

1. `generate_data.py` creates two realistic, intentionally messy CSV exports.
2. `build_workbook.py` loads them into an Excel workbook using `openpyxl`, builds a cleaned dataset with live formulas (not hardcoded values), aggregates pipeline and pricing KPIs, and adds charts to an executive dashboard sheet.
3. Formulas were recalculated and verified error-free using LibreOffice headless recalculation.

## Workbook sheet guide

| Sheet | Purpose |
|---|---|
| `README` | In-workbook project overview |
| `CRM_Raw` | Unprocessed CRM export (source 1) |
| `Pricing_Raw` | Unprocessed competitor pricing feed (source 2) |
| `Data_Cleaning_Log` | Validation rules applied and records affected |
| `CRM_Clean` | Cleaned dataset, formula-linked to `CRM_Raw` |
| `Pipeline_Analysis` | Funnel, win-rate, sales-cycle, and rep performance metrics |
| `Competitor_Benchmark` | Price positioning vs. competitors |
| `Executive_Dashboard` | KPI summary and charts |

## Tools & skills

Excel (formulas, conditional formatting, pivot-style aggregation, charts) · Python (pandas, openpyxl) · ETL design · Data validation & cleaning · Sales pipeline / funnel analysis · Competitive & pricing benchmarking · Dashboard design

## Note on data

All deal and pricing data is synthetic, generated for demonstration purposes. The structure mirrors a real D2C sales/CRM environment (e.g. Pipedrive, HubSpot, or Salesforce exports).
