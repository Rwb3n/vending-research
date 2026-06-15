---
id: spike.007
title: Planogram par levels per facing
type: spike
status: open
assumption: A14
unblocks: [config/planograms.yaml, tools/reorder.py]
created: 2026-06-15
grep: "id: spike.007"
---

# Spike 007 — What is the right par (units per facing) for each SKU?

**Question.** The planogram *structure* (which SKU, how many facings) is evidenced by
the catalog and the Gusto-8 layout. The `par_each` numbers (snack 10, crisps 8, food
6, bottles 6) are **guesses** — placeholders so `reorder.py` can compute a real order.
Too high = capital tied in stock + stale lines; too low = stockouts between runs.

**Cheapest experiment.** Once per-SKU depletion exists (spike 005 telemetry, or manual
slot counts on a run), set each SKU's par = expected sales between restock visits
(spike 002 cadence) + a safety buffer. Top-facing sellers carry deeper par; slow lines
shrink toward 1 facing then get cut (product-cut policy).

**Decision it unblocks.** The `par_each` values in `planograms.yaml`, which drive the
purchase order quantities and the implied stock investment per machine.

**Time box.** Falls out of spikes 002 + 005; no new data collection of its own.

## Result
_(open — depends on spikes 002 and 005)_
