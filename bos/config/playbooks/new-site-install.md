---
id: playbook.new-site-install
title: New Site Install
type: playbook
task: install-site
trigger: "site-go-nogo verdict GO logged in decisions.jsonl"
inputs: [decision GO, machine from inventory]
outputs: [new machines_state.json record]
updated: 2026-06-15
grep: "id: playbook.new-site-install"
---

# New Site Install

**Pre-req:** a logged `site-go-nogo` GO decision for the site.

1. **Terms in writing.** Confirm free placement (or terms) in writing with the site
   contact. No paper, no machine (risk R1).
2. **Siting.** Place in the highest-footfall captive spot — canteen, near clock-in,
   warehouse break area. Power + (ideally) 4G signal for Nayax telemetry.
3. **Commission.** Install machine, load planogram, enable Nayax, test one vend on each
   payment channel (coin / note / cashless).
4. **Register in the OS** — append a `machine_state` record:
   ```json
   {"machine_id":"Mxx","model":"<from data/machines.json>","site":"<name>",
    "postcode":"<pc>","status":"live","installed":"<today>","capacity":<n>,
    "current_fill":1.0,"last_restock":"<today>","planogram":"standard-industrial",
    "telemetry":"nayax"}
   ```
   Then `python bos/tools/validate.py bos/state/machines_state.json`.
5. **Baseline.** First restock visit ~ within the machine's expected days-to-trigger;
   note depletion to seed spikes 001/002 for this site.
