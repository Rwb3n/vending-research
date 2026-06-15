---
id: spike.006
title: Target vends/day and machine-buy trigger
type: spike
status: open
assumption: [A12, A13]
unblocks: [config/kpis.yaml#vends_per_machine_per_day, config/policies.yaml#machine-buy]
created: 2026-06-15
grep: "id: spike.006"
---

# Spike 006 — What is the target vends/day, and when do we buy machine N+1?

**Question.** Breakeven is ~4 vends/day (A4, evidenced). We assume an *operating
target* of **15 vends/day/machine** and a machine-buy trigger at **>15 sustained for
≥6 weeks**. Both numbers are guesses bolted onto the one evidenced anchor.

**Cheapest experiment.** Let the first 4 machines run 6–8 weeks. Read actual
vends/day/machine from `scorecard.jsonl`. Set the operating target at the 60th
percentile of observed live-machine performance (achievable-but-stretch). Set the
buy trigger where an estate's machines are demand-constrained (stockouts present AND
vends/day above target) — i.e. unmet demand a new machine would absorb.

**Decision it unblocks.** `vends_per_machine_per_day` KPI target/thresholds and the
`machine-buy` capital-allocation policy.

**Time box.** 6–8 weeks of live operation; 1 hour analysis.

## Result
_(open — needs 6–8 weeks of live data)_
