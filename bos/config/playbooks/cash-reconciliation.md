---
id: playbook.cash-reconciliation
title: Cash Reconciliation
type: playbook
task: reconcile-cash
trigger: "weekly (after restock runs)"
inputs: [cash_ledger.jsonl, telemetry]
outputs: [incidents.jsonl if variance breach]
feeds_policy: cash-investigate
updated: 2026-06-15
grep: "id: playbook.cash-reconciliation"
---

# Cash Reconciliation

**Goal:** catch shrinkage and validator drift early (risk R3).

1. For each collection this week, ensure a `cash_ledger` entry exists with both
   `collected_gbp` (banked) and `expected_gbp` (Nayax telemetry).
2. Compute variance per machine:
   `python bos/tools/policy.py evaluate --decision cash-investigate --input '{"variance_pct": <v>}' --log`
   where `v = (collected - expected) / expected * 100`.
3. **Act on verdict:**
   - `OK` → done.
   - `WATCH` → recount at next collection; note on the ledger entry.
   - `INVESTIGATE` → open an `incident` (type `cash_variance`); check coin
     validator, note acceptor, and telemetry mapping; consider going cashless-only on
     that machine.
4. Roll the week's totals into the scorecard fields `cash_collected` / `cash_expected`
   so the `cash_variance_pct` KPI tracks the trend.
