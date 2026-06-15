---
id: playbook.fault-triage
title: Fault Triage
type: playbook
task: triage-fault
trigger: "telemetry fault flag OR reported fault"
inputs: [incident]
outputs: [updated incident, possibly machines_state status change]
protects_kpi: stockout_rate_pct
updated: 2026-06-15
grep: "id: playbook.fault-triage"
---

# Fault Triage

**Goal:** clear silent revenue loss fast (risk R2). Target: high-severity faults
resolved within 1 operating day.

1. **Log first.** Append an `incident` (type `fault`, set severity). A fault not logged
   is a fault that festers.
2. **Diagnose remotely** via Nayax telemetry before driving:
   - *No comms* → telemetry/SIM issue, not necessarily a vend fault. Severity low if
     still vending.
   - *Sold-out flags* → it's a stockout, not a fault → switch to `restock-run`.
   - *Validator / coin errors* → likely jam or cash path.
3. **On site:** clear jam, test each channel, reset, confirm a live vend.
4. **If unfixable on site:** set machine `status: fault` in `machines_state.json`,
   estimate `revenue_lost_gbp` on the incident, escalate to supplier/engineer.
5. **Close:** set incident `status: resolved` with `resolution`; restore machine
   `status: live`.
