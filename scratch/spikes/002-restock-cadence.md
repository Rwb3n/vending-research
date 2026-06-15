---
id: spike.002
title: Default restock cadence per machine
type: spike
status: open
assumption: A8
unblocks: [config/cadence.yaml, config/workflows.yaml#weekly-route]
created: 2026-06-15
grep: "id: spike.002"
---

# Spike 002 — How often should each machine be restocked?

**Question.** We assume **weekly**. The real driver is depletion rate (spike 001) vs
drive time. A machine that empties in 4 days needs twice-weekly; one that lasts 12 days
is being over-served.

**Cheapest experiment.** Derive days-to-30%-fill per machine from the spike-001
depletion curves. Cadence per machine = days-to-trigger, rounded down to a route day.
Cluster machines so one drive covers several due on the same day (honours A3).

**Decision it unblocks.** The cadence entries and the `weekly-route` workflow trigger.

**Time box.** Falls out of spike 001 — no extra data collection.

## Result
_(open — depends on spike 001)_
