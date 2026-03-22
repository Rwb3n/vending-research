# The Snack Choice — London Expansion Research

Business intelligence workbook for evaluating and launching a vending machine operation in London.

## Quick Start

1. Download `TheSnackChoice_London_Research.xlsx`
2. Open in Google Sheets or Excel
3. Start with the **Dashboard** tab for an overview
4. Edit blue-highlighted cells to model your scenarios

## Workbook Structure

| Tab | Description |
|-----|-------------|
| Dashboard | Summary metrics pulling from all sheets |
| Product Catalog | 45 products with wholesale costs, vend prices, margins |
| Machine Specs | 14 machines from UK suppliers with pricing |
| Supplier Directory | 18 suppliers (wholesalers, machines, payments, insurance) |
| Competitor Analysis | 14 UK vending operators mapped |
| Startup Costs | Scenario calculator (1-3 machines) with breakeven |
| Location Tracker | Site scouting template with pipeline tracking |
| Weekly P&L | Per-machine weekly revenue and cost tracker |

## Regenerating

```bash
pip install openpyxl
python3 generate.py
```

Data lives in `data/*.json`. Edit those files and regenerate.

## Data Confidence

- **Machine prices**: HIGH — scraped from Vendtrade (Mar 2026)
- **Nayax pricing**: HIGH — confirmed £450+VAT, £10/month, 2.95%
- **Wholesale product costs**: MEDIUM — industry benchmarks (trade prices behind logins)
- **Competitor data**: HIGH — Companies House, public websites
- **Supplier terms**: HIGH — scraped from supplier websites

## Conventions

- Blue text (#0000FF) + yellow background = manual input cells
- Black text = formula cells (do not edit)
- All prices in GBP (£)
