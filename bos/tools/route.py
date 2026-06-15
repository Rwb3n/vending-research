"""
---
id: bos.tools.route
purpose: Plan a restock route. Selects live machines below the restock trigger and
         clusters stops by site (A3 — clustering protects route revenue/hour). Pure
         function of machine state + the operational trigger in policies.yaml.
inputs: bos/state/machines_state.json, bos/config/policies.yaml
outputs: route plan (stdout JSON) — ordered, site-clustered stops + units to refill
usage: "python bos/tools/route.py plan"
updated: 2026-06-15
---
"""
import os
import sys

import common


def main():
    if len(sys.argv) < 2 or sys.argv[1] != "plan":
        common.emit({"error": "usage: route.py plan"})
        sys.exit(2)

    ops = common.load_yaml(os.path.join(common.CONFIG_DIR, "policies.yaml"))["operational"]
    trigger = ops["restock_trigger_fill"]
    cluster = ops.get("cluster_routes", True)
    machines = common.load_json(os.path.join(common.STATE_DIR, "machines_state.json"))["machines"]

    due = [m for m in machines if m.get("status") == "live" and m.get("current_fill", 1) < trigger]

    # group by site (cluster); order sites by their most-urgent (lowest-fill) machine
    sites = {}
    for m in due:
        sites.setdefault((m["site"], m.get("postcode", "")), []).append(m)

    stops = []
    for (site, postcode), ms in sites.items():
        ms.sort(key=lambda x: x["current_fill"])
        stops.append({
            "site": site, "postcode": postcode,
            "machines": [{
                "machine_id": m["machine_id"],
                "current_fill": m["current_fill"],
                "units_to_refill": round(m["capacity"] * (1 - m["current_fill"])),
                "planogram": m.get("planogram"),
            } for m in ms],
            "urgency": min(m["current_fill"] for m in ms),
        })
    stops.sort(key=lambda s: s["urgency"])  # most urgent site first

    common.emit({
        "restock_trigger_fill": trigger,
        "clustered": cluster,
        "stops": len(stops),
        "machines_due": len(due),
        "total_units": sum(mm["units_to_refill"] for s in stops for mm in s["machines"]),
        "route": stops,
    })


if __name__ == "__main__":
    main()
