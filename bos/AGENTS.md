---
id: bos.agents
title: The Snack Choice OS — Agent & Operator Entry Point
type: guide
status: active
owner: operator
audience: [agent, operator]
inputs: [bos/config/, bos/state/, bos/tools/manifest.json]
outputs: [bos/state/*.jsonl]
updated: 2026-06-15
grep: "id: bos.agents"
---

# The Snack Choice — Operating System

**Read this first.** This is how the business runs. It is built so an agent can operate
the digital layer cheaply: read this file, read the one state file you need, call one
deterministic tool. Do not read generator source to do a job — read the manifest.

## The one number

Everything optimises **revenue per operator-hour**. Every routine, policy, and tool
exists to push that up. Viability bands: red `<£15`, amber, green `>£30` (see
`config/kpis.yaml`).

## How the system is laid out

| You want to…                        | Read / run                                  |
|-------------------------------------|---------------------------------------------|
| Know what tools exist               | `tools/manifest.json` (the API)             |
| Compute this week's numbers         | `python tools/kpi.py compute --latest`      |
| Decide a site / product / purchase  | `python tools/policy.py evaluate ...`       |
| Plan a restock run                  | `python tools/route.py plan`                |
| Build a purchase order              | `python tools/reorder.py`                   |
| Check a file is valid               | `python tools/validate.py <file>`           |
| Rebuild the workbook + manual       | `python build.py`                           |
| Understand a routine                | `config/cadence.yaml`, `config/playbooks/`  |
| Understand a decision rule          | `config/policies.yaml`                      |
| See what's a guess vs a fact        | `ASSUMPTIONS.md`, `scratch/spikes/`         |

## Work model: tasks → pipelines → workflows

There are **no roles** here, only work. A human or an agent can execute any of it.

- **Task** (`config/tasks.yaml`) — one unit of work. Either a deterministic `tool` call
  or a human `playbook`. Has explicit inputs/outputs.
- **Pipeline** (`config/pipelines.yaml`) — an ordered list of tasks. The output of one
  feeds the next (pipes).
- **Workflow** (`config/workflows.yaml`) — a pipeline bound to a **cadence** trigger
  (daily / weekly / monthly / quarterly) from `config/cadence.yaml`.

Example — the **weekly review** workflow runs the pipeline:
`log-scorecard → compute-kpis → evaluate-flags → record-decisions → regenerate-manual`.

## Contracts & state

- `state/*.jsonl` are append logs. The **first line is a `_meta` header — skip it**;
  every other line is one record validated against `schema/<name>.schema.json`.
- `state/machines_state.json` is the current fleet snapshot (mutable).
- Never hand-edit derived numbers. Append raw inputs; let `tools/kpi.py` derive metrics.

## Rules for an agent operating here

1. **Tools do math, you do judgment.** Never compute a KPI in your head — call `kpi.py`.
2. **Validate before you write.** Run `validate.py` on anything you append.
3. **Cite the policy.** When you log a decision, record which policy fired and why.
4. **Evidence or spike.** If you act on a number not in `ASSUMPTIONS.md` as EVIDENCED,
   say so and open a spike in `scratch/spikes/` — do not silently harden a guess.
5. **Stay token-economic.** Read the smallest file that answers the question.

## Reference data (read-only)

The OS reads, but never writes, the market-research datasets:
`../data/products.json` (catalog + costs), `../data/machines.json` (specs). Those are
owned by the market-research track (`../generate.py`).
