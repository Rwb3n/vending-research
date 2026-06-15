---
id: bos.assumptions
title: Operating Assumptions & Evidence Status
type: register
status: active
owner: operator
inputs: [data/products.json, README.md, TheSnackChoice_London_Research.xlsx]
outputs: [scratch/spikes/]
updated: 2026-06-15
grep: "id: bos.assumptions"
---

# Operating Assumptions

Every belief the OS encodes is listed here with its evidence status. The rule:

> **Evidence or spike.** If an assumption drives a threshold/policy and has no
> evidence, it is marked `SPIKE` and gets a cheap experiment in `scratch/spikes/`.

`EVIDENCED` = traceable to a source in this repo or the world.
`SPIKE`     = currently a guess; the linked spike says how to validate it cheaply.

| # | Assumption | Status | Source / Spike |
|---|-----------|--------|----------------|
| A1 | The north-star metric is **revenue per operator-hour** (not per machine). | EVIDENCED | README "Key Strategic Decisions"; Performance Review sheet. |
| A2 | Revenue/operator-hour viability bands: red `<£15`, green `>£30`. | EVIDENCED | Performance Review sheet conventions (README). |
| A3 | Clustering machines on one estate lifts effective rate to ~£44/hr vs £15–25/hr scattered. | EVIDENCED | Strategy Notes (README). Drives route policy (cluster-first). |
| A4 | Breakeven is ~**4 vends/day/machine** with free placement (16/day if paying rent). | EVIDENCED | Startup Costs sheet (README). |
| A5 | Blended gross margin target is **≥50%**. | EVIDENCED | `data/products.json` (e.g. Mars £0.55 cost → £1.40 vend ≈ 61%). |
| A6 | Site go/no-go: vending_score `≥8` GO, `5–7` WATCH, `<5` NO. | EVIDENCED | Territory scoring 1–10 + tier cutoffs (README). |
| A7 | Restock should trigger when a machine drops below **30% fill**. | **SPIKE** | `scratch/spikes/001-restock-threshold.md` |
| A8 | Default restock cadence is **weekly** per live machine. | **SPIKE** | `scratch/spikes/002-restock-cadence.md` |
| A9 | Stockout-rate target is **<5%** of machine-days. | **SPIKE** | `scratch/spikes/003-stockout-target.md` |
| A10 | Acceptable cash variance (collected vs expected) is **±2%**. | **SPIKE** | `scratch/spikes/004-cash-variance-tolerance.md` |
| A11 | Product-cut rule: cut a line selling `<3 units/week` for `≥4 weeks`. | **SPIKE** | `scratch/spikes/005-product-cut-rule.md` |
| A12 | Target operating run-rate is **15 vends/day/machine** (≈3.75× breakeven). | **SPIKE** | `scratch/spikes/006-vends-day-target.md` |
| A13 | Machine-buy trigger: avg `>15 vends/day` sustained `≥6 weeks` across estate. | **SPIKE** | depends on A12; `scratch/spikes/006-vends-day-target.md` |

## How to discharge a spike

1. Run the experiment described in the spike file (cheap, time-boxed).
2. Record the observed number + n + date in the spike's `## Result` section.
3. Update the assumption row here to `EVIDENCED` with the source.
4. Update the corresponding value in `config/` (e.g. `policies.yaml`, `kpis.yaml`).

The spikes are *not* code we ship into the operating loop until discharged — they
live in `scratch/` precisely so the OS does not pretend a guess is a fact.
