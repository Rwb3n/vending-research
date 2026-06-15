---
id: scratch.readme
title: Scratchpad — Spikes & Throwaway Experiments
type: readme
status: active
owner: operator
updated: 2026-06-15
grep: "id: scratch.readme"
---

# scratch/

Throwaway space. Nothing here is part of the operating loop.

`scratch/spikes/` holds **spikes**: cheap, time-boxed experiments that exist to turn
an unevidenced operating assumption (see `../bos/ASSUMPTIONS.md`) into a fact. Each
spike states the question, the cheapest experiment that answers it, the decision it
unblocks, and — once run — the result.

When a spike is discharged: copy the validated number into the relevant `bos/config/`
file, flip the assumption row to `EVIDENCED`, and the spike becomes history.

This directory is intentionally outside `bos/` so agents reading the operating system
never confuse a guess-under-test with a shipped fact.
