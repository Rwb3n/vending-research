---
id: spike.005
title: Product add/cut rule
type: spike
status: open
assumption: A11
unblocks: [config/policies.yaml#product-cut]
created: 2026-06-15
grep: "id: spike.005"
---

# Spike 005 — When do we cut a product line?

**Question.** We assume: cut a line selling **<3 units/week for ≥4 weeks**. A slot
holding a dead SKU is opportunity cost — it could hold a top-5 seller. But cutting too
fast churns the planogram and confuses regulars. What rule maximises slot revenue?

**Cheapest experiment.** Needs per-SKU sales (Nayax telemetry exports, not yet wired).
Spike step 1: confirm Nayax exports per-selection sales. Step 2: once 6 weeks of
per-SKU data exist, rank by units/slot/week; model swapping bottom-quartile SKUs for a
duplicate of a top seller; compare slot revenue.

**Decision it unblocks.** The `product-cut` policy thresholds and `planogram-update`
playbook trigger.

**Time box.** Blocked on per-SKU telemetry ingestion (separate task). 1 day once data
exists.

## Result
_(open — blocked on per-SKU telemetry ingestion)_
