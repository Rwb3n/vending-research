---
id: spike.008
title: Three-view OS — shell / ghost / pilot (acts_via)
type: spike
status: built
unblocks: [config/kpis.yaml#acts_via, tools/brief.py, schema/decision.schema.json, demo/week_in_the_life.py]
created: 2026-06-15
grep: "id: spike.008"
---

# Spike 008 — One truth, three readers (machine · ghost · pilot)

> **status: BUILT** — `acts_via` signed off (mode: runnable/proposes/manual), wired into the demo.
> probe: `bos/tools/brief.py` · artifacts (gitignored, regenerated): `bos/demo/out/{state.json,
> brief.agent.json, brief.operator.md}`.

## Thesis

The system is the **shell**, the agent is the **ghost**, the operator is the **pilot**. From one
truth (`kpi.py compute`), the machine deterministically emits **three views**, each shaped for one
reader. No LLM in the render path — the intelligence is a **join over config**, not generation. The
loop is **bi-directional**: render → ghost/pilot acts → state appended → re-render reflects the act.

## Built (this session)

1. `acts_via` on all 7 KPIs in `kpis.yaml` (`mode: runnable | proposes | manual`). **Zero unmapped
   KPIs** — the boundary moved from "gap" to "declared trust level."
2. `decision.schema.json` gained optional `status: proposed | committed` (default committed, no
   migration). The proposal lifecycle now has a contract.
3. `brief.py` resolves actions from `acts_via` and emits runnable commands. Determinism: **PASS**
   (run twice vs same `BOS_STATE_DIR` ⇒ byte-identical; sorted keys, no wall-clock).
4. `week_in_the_life.py` Step 3 is now brief-driven — the **hardwired `decide()` calls are deleted**.
   The ghost runs only `runnable` commands; `proposes`/`manual` are narrated, not faked. UTF-8
   self-fix added (runs cold on Windows + Linux, no env vars).
5. **Weekly logbook (north-star heartbeat).** On a GREEN week the north star still logs a KEEP —
   `decisions.jsonl` is a reproducible weekly history, not just an exception report. Required
   write-path work, now done: `policy.py --log` is **week-anchored** (date = Monday of `--week`,
   not wall-clock) and **idempotent** per `(week, policy, subject)`. `brief.py` emits a `heartbeat`
   block (green only, separate from `flags`) and adds `--week` to *all* policy commands, so the
   whole logbook is reproducible. Guarded by `bos/demo/test_logbook.py` (the write-path test the
   brief-diff can't cover) and the `week` field on `decision.schema.json`.

**Honest outcome vs the old demo:** the old Step 3 logged 3 decisions — 2 of which fired on GREEN
KPIs (hardwired theater). The brief-driven Step 3 logs **1** (cash → WATCH, real arithmetic -2.67%),
proposes 1 (machine-buy), and routes 2 to manual playbooks. Fewer logged decisions, but every one is
computed, not typed. That is the thesis.

## The three findings (the spike was built to confess, not flatter)

1. **The deterministic shell has a clean, self-declared boundary.** Pre-`acts_via`, only 4 of 7 KPIs
   had a playbook binding; the north star and `machines_live` had nothing. That gap *was* the finding:
   it's where the ghost stops joining and starts deciding.
2. **Playbook bindings can't emit a runnable command** (they're human SOPs → `command: null`). To make
   the ghost actionable, KPIs must bind to a tool/policy path — hence `acts_via`.
3. **`tasks.yaml` already names the seam:** `evaluate-flags` is `kind: agent`. The system admits one
   step is judgment. The demo's hardwired `decide()` froze that judgment as tribal knowledge — deleted.

## `acts_via` contract (SIGNED OFF)

```yaml
# in kpis.yaml, per KPI:
acts_via:
  kind: policy                     # policy | playbook
  target: cash-investigate         # policy id or playbook id
  input_map: {variance_pct: value} # KPI fields -> policy inputs (handles name mismatches)
  mode: runnable                   # runnable | proposes | manual  (3-state, not bool)
```

`mode` is the trust three-state — the judgment boundary made machine-readable:

- **`runnable`** — ghost evaluates the policy and logs the `DEC-…` **unattended**. Every input is
  derivable from state. (`rev_per_operator_hour`, `gross_margin_pct`, `cash_variance_pct`)
- **`proposes`** — ghost drafts `status: proposed` (auditable, **not committed**); pilot commits.
  Used where inputs are partial / the call is capital. (`machines_live → machine-buy`, `product-cut`)
- **`manual`** — no policy; follow a human playbook. (`stockout`, `vends`, `route_rev`)

### Bindings (grounded in what the data supports)

| KPI | `acts_via` | `mode` | why |
|---|---|---|---|
| `rev_per_operator_hour` ★ | policy `site-keep-fix-cut` | runnable | KPI value *is* the policy input (1:1) |
| `gross_margin_pct` | policy `price-change` | runnable | 1:1 input match |
| `cash_variance_pct` | policy `cash-investigate` | runnable | `input_map: {variance_pct: value}` (rename) |
| `stockout_rate_pct` | playbook `fault-triage` | manual | human SOP |
| `vends_per_machine_per_day` | playbook `planogram-update` | manual | human SOP |
| `route_revenue_per_hour` | playbook `restock-run` | manual | human SOP |
| `machines_live` | policy `machine-buy` | proposes | needs `weeks_sustained`, `stockout_present` — multi-week context one scorecard can't supply. An information boundary, not an oversight. |

### Proposal lifecycle (fits the existing shell — no new primitive)

`decision.schema.json` **already** had `decided_by: [agent, operator]`. We added `status: proposed |
committed`. `proposes`-mode: ghost appends `{status: proposed, decided_by: agent}`; approval: pilot
appends `{status: committed, decided_by: operator}`. Append-only, auditable, same `decisions.jsonl`.

## Still open (next spec)

- **`--propose` flag on `policy.py`** (writes `status: proposed`) — deferred. The briefs **already
  emit `--propose`** in their templated commands, so this flag must land or those instructions error.
  No caller yet: `machine-buy` (the only proposes-mode KPI) needs multi-week inputs not derivable
  from one week, so building the flag now = infra ahead of a caller.
- **Bi-directional write-back proven by re-render:** log a decision → re-brief → see it reflected.
  (Today: a decision appends + validates; the flag persists because the KPI value is unchanged — by
  design. The proof we still owe is a state mutation that *moves* a KPI and shows the flag clear.)

## Resolved this session

- **North-star-always-confirm:** decided **YES** — green weeks log a KEEP (weekly logbook). Built;
  see item 5 above.
