---
id: spike.003
title: Stockout-rate target
type: spike
status: open
assumption: A9
unblocks: [config/kpis.yaml#stockout_rate_pct]
created: 2026-06-15
grep: "id: spike.003"
---

# Spike 003 — What stockout rate is acceptable?

**Question.** We assume a target of **<5% of machine-days** with a stockout. Zero is
uneconomic (implies over-restocking → kills revenue/operator-hour). What rate balances
lost sales against operator-hours?

**Cheapest experiment.** Once 6+ weeks of `scorecard.jsonl` exist, regress weekly
gross_revenue against stockout_rate per machine. Find the knee where marginal revenue
recovered from an extra restock run < operator cost of that run (use A2 £/hr bands).

**Decision it unblocks.** The green/red thresholds on the `stockout_rate_pct` KPI.

**Time box.** 6 weeks of scorecard data; 1 hour analysis.

## Result
_(open — needs ≥6 weeks of scorecard data)_
