---
id: bos.demo.scenario
title: Demo Scenario — A Week in the Life
type: demo
status: active
owner: operator
runner: bos/demo/week_in_the_life.py
updated: 2026-06-15
grep: "id: bos.demo.scenario"
---

# Demo — A Week in the Life

A reproducible walkthrough of the BOS weekly loop, driven entirely by the real
deterministic tools. It proves the thesis: the business runs as plain files + tools
that an agent or a new operator can drive with no tribal knowledge.

**Non-destructive.** The runner copies `bos/state/` into a temp sandbox and points the
tools at it via `BOS_STATE_DIR`, so committed state is never mutated. Run it as often
as you like — every run starts from the same seed and ends the same way.

## The scenario (Monday, week 2026-W25)

It's Monday morning. Last week (W25) just closed and the numbers are *almost* good:

1. **Close the week** — append the W25 scorecard (raw inputs).
2. **Read the one number** — compute KPIs. North-star revenue/operator-hour holds green,
   but four KPIs flag amber (vends/machine/day, stockout, cash variance, live machines).
3. **Brief & decide** — one truth → three views (`brief.py`). The green north star still logs
   a weekly **heartbeat** (KEEP) so `decisions.jsonl` is a reproducible logbook, not just an
   exception report. Each flag carries an `acts_via` binding, so the brief partitions the work
   by trust: **runnable** (the ghost evaluates the policy and logs the verdict unattended —
   here, cash variance → WATCH), **proposes** (needs the pilot's multi-week judgment — live
   machines → machine-buy), and **manual** (follow a human playbook — stockout, vends). No
   hand-wired decisions; every action is computed from `kpis.yaml` and week-anchored.
4. **Replenish** — plan the clustered restock route and build the per-SKU purchase order.
5. **Close the loop** — regenerate the operating manual reflecting the new week. The pilot's
   own next-steps are in the regenerated **operator brief** (`bos/demo/out/live/brief.operator.md`).

## Run it

```bash
python bos/demo/week_in_the_life.py
```

Add `--quiet` for just the tool outputs (no narration), or `--no-build` to skip the
manual regeneration step.
