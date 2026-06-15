---
id: playbook.planogram-update
title: Planogram Update
type: playbook
task: update-planogram
trigger: "product-cut verdict CUT, or monthly gate"
inputs: [decision CUT/KEEP, per-SKU sales]
outputs: [updated machines_state planogram]
protects_kpi: vends_per_machine_per_day
updated: 2026-06-15
grep: "id: playbook.planogram-update"
---

# Planogram Update

**Goal:** every slot earns its place. Dead slots are opportunity cost.

1. Rank SKUs by units/slot/week (per-SKU telemetry — see spike 005, currently blocked).
2. Run `product-cut` policy on the bottom lines:
   `python bos/tools/policy.py evaluate --decision product-cut --input '{"units_per_week":<u>,"weeks_below":<w>}' --log`
3. For each `CUT`: replace the slot with a **duplicate of a top-5 seller** (more facings
   of proven demand) rather than an untested new SKU.
4. Change at most ~2 slots/visit — regulars notice churn; keep changes measurable.
5. Update the machine's `planogram` id in `machines_state.json` if the mix materially
   changes; note the swap so its effect on `vends_per_machine_per_day` is attributable.
