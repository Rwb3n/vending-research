---
id: spike.004
title: Cash variance tolerance
type: spike
status: open
assumption: A10
unblocks: [config/kpis.yaml#cash_variance_pct, config/policies.yaml#cash-investigate]
created: 2026-06-15
grep: "id: spike.004"
---

# Spike 004 — What cash variance should trigger investigation?

**Question.** We assume **±2%** (collected vs telemetry-expected) is normal noise.
Beyond that = miscount, theft, jammed validator, or telemetry drift. What is the true
noise floor for Nayax cashless + coin machines?

**Cheapest experiment.** For 4 weeks, reconcile every collection against Nayax
telemetry (`cash_ledger.jsonl` vs scorecard `cash_expected`). Compute the standard
deviation of variance%. Set the investigate threshold at ~2σ.

**Decision it unblocks.** `cash_variance_pct` KPI thresholds and the `cash-investigate`
policy.

**Time box.** 4 weeks of reconciliation; built into the weekly cadence anyway.

## Result
_(open — not yet run)_
