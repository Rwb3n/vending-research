"""
---
id: bos.tools.reorder
purpose: Build a per-SKU supplier purchase order from current fleet fill levels and each
         machine's planogram. For every due machine, refills each SKU toward par scaled
         by how empty the machine is, costs the order against the catalog, and reports
         the true blended margin. Deterministic.
inputs: bos/state/machines_state.json, bos/config/policies.yaml,
        bos/config/planograms.yaml, data/products.json
outputs: purchase order (stdout JSON) — per-SKU lines + category subtotals + true margin
usage: "python bos/tools/reorder.py [--all]   (--all = every live machine, not just due)"
note: "Planogram structure is evidenced; par levels are SPIKE 007 (tune par_each)."
updated: 2026-06-15
---
"""
import os
import sys
from collections import defaultdict

import common


def main():
    ops = common.load_yaml(os.path.join(common.CONFIG_DIR, "policies.yaml"))["operational"]
    trigger = ops["restock_trigger_fill"]
    machines = common.load_json(os.path.join(common.STATE_DIR, "machines_state.json"))["machines"]
    planograms = common.load_yaml(os.path.join(common.CONFIG_DIR, "planograms.yaml"))["planograms"]
    catalog = {p["name"]: p for p in common.load_json(os.path.join(common.DATA_DIR, "products.json"))}

    include_all = "--all" in sys.argv
    selected = [m for m in machines if m.get("status") == "live"
                and (include_all or m.get("current_fill", 1) < trigger)]

    # aggregate units needed per SKU across selected machines: refill toward par,
    # scaled by how empty each machine is (current_fill is the depletion proxy).
    sku_units = defaultdict(int)
    errors = []
    for m in selected:
        pg = planograms.get(m.get("planogram"))
        if pg is None:
            errors.append(f"{m['machine_id']}: unknown planogram '{m.get('planogram')}'")
            continue
        deficit = 1 - m.get("current_fill", 1)
        for line in pg["lines"]:
            if line["sku"] not in catalog:
                errors.append(f"planogram sku not in catalog: {line['sku']}")
                continue
            sku_units[line["sku"]] += round(line["facings"] * line["par_each"] * deficit)

    if errors:
        common.emit({"ok": False, "errors": errors})
        sys.exit(1)

    lines = []
    cat_totals = defaultdict(lambda: {"units": 0, "cost": 0.0, "retail": 0.0})
    total_units = total_cost = total_retail = 0
    for sku in sorted(sku_units, key=lambda s: (catalog[s]["category"], s)):
        units = sku_units[sku]
        if units <= 0:
            continue
        p = catalog[sku]
        cost = round(units * p["wholesale_cost"], 2)
        retail = round(units * p["vend_price"], 2)
        lines.append({"sku": sku, "category": p["category"], "units": units,
                      "wholesale_cost": p["wholesale_cost"], "est_cost": cost,
                      "vend_price": p["vend_price"], "est_retail": retail})
        ct = cat_totals[p["category"]]
        ct["units"] += units
        ct["cost"] += cost
        ct["retail"] += retail
        total_units += units
        total_cost += cost
        total_retail += retail

    by_category = [{"category": c, "units": v["units"], "est_cost": round(v["cost"], 2),
                    "est_retail": round(v["retail"], 2)} for c, v in sorted(cat_totals.items())]
    margin = round((total_retail - total_cost) / total_retail * 100, 1) if total_retail else None

    common.emit({
        "machines_selected": [m["machine_id"] for m in selected],
        "scope": "all live" if include_all else f"below {trigger:.0%} fill",
        "total_units": total_units,
        "est_cost_gbp": round(total_cost, 2),
        "est_retail_gbp": round(total_retail, 2),
        "implied_margin_pct": margin,
        "lines": lines,
        "by_category": by_category,
        "note": "Per-SKU refill toward planogram par, scaled by machine current_fill (proxy). Par levels: spike 007.",
    })


if __name__ == "__main__":
    main()
