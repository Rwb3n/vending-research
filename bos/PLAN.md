---
id: bos.plan
title: The Snack Choice — Business Operating System (BOS) Build Plan
type: plan
status: active
owner: operator
optimised_for: agentic-automation
inputs: []
outputs: [bos/, scratch/, generate_bos.py, generate_bos_manual.py]
updated: 2026-06-15
grep: "id: bos.plan"
---

# The Snack Choice — Business Operating System

## What this is

A **second research track** layered on top of the existing vending market-research
repo. The market-research workbook answers *"is this opportunity worth pursuing, and
where?"*. This Business Operating System (BOS) answers the next question:

> **"How do we run it — every day, every week — so it is profitable, delegable, and
> automatable?"**

It is a *research experiment* in the sense that the business is treated as an
instrumented system tuned against one number: **revenue per operator-hour**.

## Design principles (the constraint that shapes everything)

The whole digital layer is **optimised for agentic automation**:

1. **Everything is a file.** State, config, contracts, and docs are all plain files.
2. **Deterministic tools, not LLM math.** Computation lives in small Python CLIs
   (`bos/tools/*.py`). Agents spend tokens on *judgment*, never on arithmetic a tool
   can do exactly and repeatably.
3. **Pipes.** Tools read JSON from files/stdin and write JSON to stdout. They compose.
4. **An API an agent can discover.** `bos/tools/manifest.json` is the registry; an
   agent reads one small file to learn every capability — never the generator source.
5. **Token-economic.** Small, single-concern files. An agent reads `bos/AGENTS.md`,
   then the one state file it needs, then calls one tool. No blob-scanning.
6. **Grep-native headers.** Every file opens with a YAML-style header (`id:`,
   `purpose:`, `inputs:`, `outputs:`). `grep -r "id: bos."` maps the whole system.
7. **Contracts.** JSON Schemas in `bos/schema/` validate every state record.
8. **Tasks, pipelines, workflows — not roles.** The unit of work is a *task* (one
   tool or playbook), composed into *pipelines*, triggered by *workflows* on a
   cadence. A task can be run by a human or an agent; the OS does not care which.
9. **Evidence or spike.** Any operating assumption without evidence is flagged and
   gets a *spike* (a cheap experiment) in `scratch/spikes/`. See `ASSUMPTIONS.md`.

## Module map (your 7 modules → this structure)

| Module            | Lives in                                  |
|-------------------|-------------------------------------------|
| Vision & Targets  | `config/kpis.yaml` (north star + targets) + `AGENTS.md` |
| Scorecard         | `tools/kpi.py` + `state/scorecard.jsonl`  |
| Operating Cadence | `config/cadence.yaml`                     |
| Playbooks (SOPs)  | `config/playbooks/*.md`                   |
| Decision Policies | `config/policies.yaml` + `tools/policy.py`|
| ~~Roles~~ → Work  | `config/tasks.yaml`, `pipelines.yaml`, `workflows.yaml` |
| Risk Register     | `config/risks.yaml`                       |

## Directory layout

```
bos/
  AGENTS.md            entry point (read first)
  PLAN.md              this file
  ASSUMPTIONS.md       operating assumptions + evidence status
  schema/              JSON Schema contracts for state records
  config/              slow-changing DEFINITIONS (YAML, human/agent-authored)
    kpis · cadence · tasks · pipelines · workflows · policies · risks .yaml
    playbooks/*.md
  state/               fast-changing RECORDS (JSONL append logs + machine state)
  tools/               deterministic CLIs + manifest.json (the API)
  build.py             orchestrator: validate → compute → generate
generate_bos.py        → TheSnackChoice_OperatingSystem.xlsx
generate_bos_manual.py → TheSnackChoice_OS_Manual.html
scratch/               spikes (experiments) for unevidenced assumptions
```

## Layering / blast radius

Self-contained. The BOS **reads** `data/products.json` and `data/machines.json` as
reference data but never writes them. The existing `generate.py`, `generate_report.py`,
and all market-research artefacts are untouched.

## Build order

1. `PLAN.md`, `ASSUMPTIONS.md`, `scratch/` spikes
2. `AGENTS.md` (agent entry point)
3. `schema/*.schema.json` (contracts)
4. `config/*.yaml` + `config/playbooks/*.md`
5. `state/*` (seeded with realistic Snack Choice data)
6. `tools/*.py` + `manifest.json`
7. `build.py`
8. `generate_bos.py` (workbook) + `generate_bos_manual.py` (HTML)
9. `requirements.txt`, README update
10. Build, verify, commit, push

## File-type header convention

- `.md`  → YAML front-matter (`--- ... ---`)
- `.yaml`→ top-of-file `# ---` comment block + `_meta:` mapping
- `.py`  → module docstring opening with a `--- ... ---` YAML block
- `.json`→ first key `"_meta": {...}`
- `.jsonl`→ first line is a `{"_meta": ...}` header record; tools skip it
