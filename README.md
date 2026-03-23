# The Snack Choice — London Expansion Research

Business intelligence workbook and territory analysis for evaluating and launching a vending machine operation targeting M25 corridor industrial estates.

## Quick Start

**Visual overview:** Open `TheSnackChoice_Report.html` in any browser — interactive map + strategy summary.

**Detailed modelling:** Open `TheSnackChoice_London_Research.xlsx` in Google Sheets or Excel:
1. Read **Strategy Notes** first — key decisions and rationale
2. Check the **Dashboard** for headline numbers
3. Edit blue-highlighted cells to model your scenarios

## Deliverables

| File | What it is |
|------|-----------|
| `TheSnackChoice_Report.html` | Interactive territory map + strategy pitch (open in browser) |
| `TheSnackChoice_London_Research.xlsx` | 13-sheet research workbook with calculators |
| `data/*.json` | Raw research data (products, machines, suppliers, competitors, territories) |

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
| Territory Planner | Data | 20 estates researched, scout-verified, and tiered |
| Location Tracker | Scaffold | Site scouting template with pipeline tracking |
| Weekly P&L | Scaffold | Per-machine weekly revenue and cost tracker |
| Operations Log | Scaffold | Restock log, maintenance log, operational beliefs & methods |
| Performance Review | Scaffold | Site + product performance tracking, monthly review gate |

## Territory Analysis (20 estates)

Food desert scoring via Google Places API — searched for food retail within 800m of each estate.

**Tier 1 — Prime Targets:**
- **Park Royal** (NW10) — 52,000 workers, 12 food outlets (1:4,333 ratio). Confirmed food desert.
- **London Gateway** (SS17) — 1,500+ workers, 3 outlets on 1,500 acres. Confirmed food desert.
- **Manor Royal** (RH10) — 30,000 workers. Partial desert — western half has nothing.
- **Brimsdown** (EN3) — 7,500 workers. Partial desert — western interior empty.
- **Crossways** (DA2) — Costa is the only quick food option on the park.

**Downgraded:** Slough Trading Estate — 3 Greggs + Lidl = workers aren't captive.

## Key Strategic Decisions

- **Target M25 corridor, not inner London** — corner shops every 200m kill the captive audience assumption in zones 1-3.
- **Cluster machines on single estates** — 5 clustered = £44/hr effective rate vs 1 scattered at £15-25/hr.
- **Free placement over site rent** — breakeven drops from 16 vends/day to 4 vends/day.
- **Hybrid format** — vending for untrusted sites, micro-markets for closed office/warehouse populations.
- **Revenue per operator-hour is the north star** — not revenue per machine.

See the **Strategy Notes** sheet for full rationale.

## Regenerating

```bash
pip install openpyxl
python3 generate.py          # workbook
python3 generate_report.py   # HTML report
```

Data lives in `data/*.json`. Edit those files and regenerate both outputs.

## Data Confidence

| Data | Confidence | Source |
|------|-----------|--------|
| Machine prices | HIGH | Scraped from Vendtrade (Mar 2026) |
| Nayax pricing | HIGH | Confirmed £450+VAT, £10/month, 2.95% |
| Wholesale products | MEDIUM | Industry benchmarks (trade prices behind logins) |
| Competitor data | HIGH | Companies House, public websites |
| Supplier terms | HIGH | Scraped from supplier websites |
| Territory data | HIGH | Google Places API density scoring, GLA studies, BID publications |

## Conventions

- Blue text + yellow background = manual input (editable)
- Black text = formula (do not edit)
- All prices in GBP (£)
- Territory Planner: green rows = Tier 1, yellow = Tier 3
- Performance Review: red = revenue/hr below £15, green = above £30
- HTML report: green markers = Tier 1, yellow = Tier 2, grey = Tier 3. Size = vending score.
