---
id: playbook.restock-run
title: Restock Run
type: playbook
task: restock-run
trigger: "route.py returns ≥1 machine below restock_trigger_fill"
inputs: [route plan, purchase order]
outputs: [updated machines_state.json, cash_ledger entries]
protects_kpi: route_revenue_per_hour
updated: 2026-06-15
grep: "id: playbook.restock-run"
---

# Restock Run

**Goal:** refill due machines in the fewest operator-hours. Protect route £/hr.

1. **Plan.** `python bos/tools/route.py plan` → ordered, site-clustered stop list. Do
   not drive to a machine that is above trigger just because you're nearby unless it's
   the same site (clustering is free; a detour is not).
2. **Pick.** `python bos/tools/reorder.py` → purchase order. Collect/stage stock.
3. **At each machine:**
   - Note `current_fill` *before* refilling (feeds spike 001 depletion curves).
   - Refill to planogram. Pull anything out of date.
   - Collect cash / confirm cashless settlement; record a `cash_ledger` entry.
   - If a fault is found → log an incident, follow `fault-triage`.
4. **After the run, update state:**
   - Set each restocked machine's `current_fill = 1.0`, `last_restock = today`.
   - `python bos/tools/validate.py bos/state/machines_state.json`
5. **Log time.** Record route hours for the week's scorecard (`route_hours`).

**Stop rule:** if a run is projected below £25 route £/hr (red), batch the stops to a
later day or wait for more machines to hit trigger.
