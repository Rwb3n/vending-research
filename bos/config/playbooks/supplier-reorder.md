---
id: playbook.supplier-reorder
title: Supplier Reorder
type: playbook
task: build-reorder
trigger: "weekly, before restock runs"
inputs: [reorder.py purchase order, ../data/suppliers.json]
outputs: [staged stock]
protects_kpi: gross_margin_pct
updated: 2026-06-15
grep: "id: playbook.supplier-reorder"
---

# Supplier Reorder

**Goal:** stock the week at the best landed cost without dead capital sitting in a
garage (protects margin, risk R7).

1. `python bos/tools/reorder.py` → purchase order: units needed per product to refill
   the fleet to planogram, with estimated wholesale cost.
2. Split the order across wholesalers by best price / availability
   (`../data/suppliers.json` — Booker, Bestway, etc.). Hit minimum-order thresholds
   without over-buying perishables.
3. Buy to the *next* run's need, not a month ahead — capital tied in stock is capital
   not buying machine N+1.
4. Record actual landed cost; this is the COGS that flows into the weekly scorecard.
5. If a line's wholesale cost has risen enough to push blended margin <50%, run the
   `price-change` policy.
