---
id: spike.001
title: Restock fill-level trigger threshold
type: spike
status: open
assumption: A7
unblocks: [config/policies.yaml#restock-trigger, tools/route.py]
created: 2026-06-15
grep: "id: spike.001"
---

# Spike 001 — At what fill % should we restock?

**Question.** We currently assume restock triggers at **30% fill**. Too high = wasted
operator-hours on near-full machines (kills revenue/operator-hour). Too low = stockouts
(lost sales + site goodwill). What is the right trigger?

**Cheapest experiment.** For the first 4 live machines, log `current_fill` daily for 3
weeks (already captured in `state/machines_state.json` + restock events). Plot
depletion curve per site. Find the fill % at which the *next* day's demand would cause
a stockout of any top-5 SKU. Set trigger one safety-day above that.

**Decision it unblocks.** The `restock-trigger` value in `policies.yaml` and the
threshold `route.py` uses to build a run.

**Time box.** 3 weeks of passive logging; 1 hour of analysis.

## Result
_(open — not yet run)_
