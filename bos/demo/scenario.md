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
   but **vends/machine/day slips to amber** and **cash variance is amber**.
3. **Decide** — the amber/red KPIs and a flagged collection drive three policy
   evaluations (logged, auditable): keep/fix/cut the estate, investigate the cash
   variance, and a margin check.
4. **Replenish** — plan the clustered restock route and build the per-SKU purchase order.
5. **Close the loop** — regenerate the operating manual reflecting the new week.

## Run it

```bash
python bos/demo/week_in_the_life.py
```

Add `--quiet` for just the tool outputs (no narration), or `--no-build` to skip the
manual regeneration step.
