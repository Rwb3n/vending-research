# The Snack Choice — London Expansion Research

Business intelligence workbook for evaluating and launching a vending machine operation targeting M25 corridor industrial estates.

## Quick Start

1. Download `TheSnackChoice_London_Research.xlsx`
2. Open in Google Sheets or Excel
3. Read **Strategy Notes** first — key decisions and rationale
4. Check the **Dashboard** for headline numbers
5. Edit blue-highlighted cells to model your scenarios

## Workbook Structure (13 sheets)

| Tab | Type | Description |
|-----|------|-------------|
| Dashboard | Calculated | Summary metrics pulling from all sheets |
| Strategy Notes | Reference | Key decisions, rationale, gotchas from planning |
| Product Catalog | Data | 45 products, 6 categories, wholesale costs, margins |
| Machine Specs | Data | 14 machines from UK suppliers with confirmed pricing |
| Supplier Directory | Data | 18 suppliers (wholesalers, machines, payments, insurance) |
| Competitor Analysis | Data | 14 UK vending operators mapped |
| Startup Costs | Calculator | Scenario model (1-3 machines) with breakeven |
| Micro-Market Costs | Calculator | Honesty shop alternative with side-by-side comparison |
| Territory Planner | Data | 14 M25 corridor estates researched and tiered |
| Location Tracker | Scaffold | Site scouting template with pipeline tracking |
| Weekly P&L | Scaffold | Per-machine weekly revenue and cost tracker |
| Operations Log | Scaffold | Restock log, maintenance log, operational beliefs & methods |
| Performance Review | Scaffold | Site + product performance tracking, monthly review gate |

## Key Strategic Decisions

- **Target M25 corridor, not inner London** — corner shops every 200m kill the captive audience assumption in zones 1-3. Industrial estates are genuine food deserts.
- **Cluster machines on single estates** — 5 clustered machines at £44/hr effective rate vs 1 scattered machine at £15-25/hr.
- **Free placement over site rent** — breakeven drops from 16 vends/day to 4 vends/day.
- **Hybrid format** — vending machines for public/untrusted sites, micro-markets for closed office/warehouse populations.

See the **Strategy Notes** sheet for full rationale.

## Regenerating

```bash
pip install openpyxl
python3 generate.py
```

Data lives in `data/*.json`. Edit those files and regenerate.

## Data Confidence

| Data | Confidence | Source |
|------|-----------|--------|
| Machine prices | HIGH | Scraped from Vendtrade (Mar 2026) |
| Nayax pricing | HIGH | Confirmed £450+VAT, £10/month, 2.95% |
| Wholesale products | MEDIUM | Industry benchmarks (trade prices behind logins) |
| Competitor data | HIGH | Companies House, public websites |
| Supplier terms | HIGH | Scraped from supplier websites |
| Territory data | HIGH | GLA studies, BID publications, property agents |

## Conventions

- Blue text + yellow background = manual input (editable)
- Black text = formula (do not edit)
- All prices in GBP (£)
- Territory Planner: green rows = Tier 1 targets, yellow = Tier 3
- Performance Review: red = revenue/hr below £15, green = above £30
