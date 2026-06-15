"""
---
id: bos.tools.kpi
purpose: Compute KPIs deterministically from scorecard inputs + KPI formulas. This is
         where ALL operating arithmetic lives — agents call this, never compute by hand.
inputs: bos/config/kpis.yaml, bos/state/scorecard.jsonl, bos/state/machines_state.json
outputs: KPI report (stdout JSON) with value + green/amber/red status per KPI
usage: "python bos/tools/kpi.py compute [--latest | --week 2026-W24 | --all]"
updated: 2026-06-15
---
"""
import os
import sys

import common


def load_inputs():
    kpis = common.load_yaml(os.path.join(common.CONFIG_DIR, "kpis.yaml"))["kpis"]
    weeks = common.load_jsonl(os.path.join(common.STATE_DIR, "scorecard.jsonl"))
    machines = common.load_json(os.path.join(common.STATE_DIR, "machines_state.json"))["machines"]
    return kpis, weeks, machines


def compute_week(kpis, record, machines):
    out = []
    for kpi in kpis:
        if kpi.get("source") == "machines_state":
            # only machines_live today; computed directly, not via the record
            value = sum(1 for m in machines if m.get("status") == "live")
        else:
            value = common.safe_eval(kpi["formula"], record)
            if value is not None:
                value = round(value, 2)
        out.append({
            "id": kpi["id"], "label": kpi["label"], "unit": kpi["unit"],
            "value": value, "target": kpi.get("target"),
            "status": common.status_for(value, kpi["direction"], kpi["thresholds"]),
            "north_star": kpi.get("north_star", False),
        })
    return out


def main():
    if len(sys.argv) < 2 or sys.argv[1] != "compute":
        common.emit({"error": "usage: kpi.py compute [--latest|--week W|--all]"})
        sys.exit(2)
    kpis, weeks, machines = load_inputs()
    if not weeks:
        common.emit({"error": "no scorecard records"})
        sys.exit(1)

    flag = sys.argv[2] if len(sys.argv) > 2 else "--latest"
    if flag == "--all":
        common.emit({"weeks": [
            {"week": w["week"], "kpis": compute_week(kpis, w, machines)} for w in weeks
        ]})
        return
    if flag == "--week":
        target = sys.argv[3]
        record = next((w for w in weeks if w["week"] == target), None)
        if record is None:
            common.emit({"error": f"week {target} not found"})
            sys.exit(1)
    else:  # --latest
        record = weeks[-1]

    common.emit({"week": record["week"], "kpis": compute_week(kpis, record, machines)})


if __name__ == "__main__":
    main()
