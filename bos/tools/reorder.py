"""
---
id: bos.tools.reorder
purpose: Build a supplier purchase order from current fleet fill levels. Estimates units
         needed to refill due machines to capacity, costs them against the product
         catalog, and breaks the order down by category. Deterministic estimate.
inputs: bos/state/machines_state.json, bos/config/policies.yaml, data/products.json
outputs: purchase order (stdout JSON) — units + estimated wholesale cost, by category
usage: "python bos/tools/reorder.py [--all]   (--all = every live machine, not just due)"
note: "Catalog has no per-planogram slot map yet (see spike 005); allocation is by
       catalog category share. Costs use catalog wholesale_cost / vend_price."
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
    products = common.load_json(os.path.join(common.DATA_DIR, "products.json"))

    include_all = "--all" in sys.argv
    selected = [m for m in machines if m.get("status") == "live"
                and (include_all or m.get("current_fill", 1) < trigger)]
    total_units = round(sum(m["capacity"] * (1 - m["current_fill"]) for m in selected))

    # category shares from the catalog (proxy for planogram mix; see spike 005)
    cat_counts = defaultdict(int)
    cat_wholesale = defaultdict(list)
    cat_vend = defaultdict(list)
    for p in products:
        c = p["category"]
        cat_counts[c] += 1
        cat_wholesale[c].append(p["wholesale_cost"])
        cat_vend[c].append(p["vend_price"])
    n_products = len(products)

    by_category = []
    total_cost = total_retail = 0.0
    for c in sorted(cat_counts):
        share = cat_counts[c] / n_products
        units = round(total_units * share)
        avg_w = sum(cat_wholesale[c]) / len(cat_wholesale[c])
        avg_v = sum(cat_vend[c]) / len(cat_vend[c])
        cost = round(units * avg_w, 2)
        retail = round(units * avg_v, 2)
        total_cost += cost
        total_retail += retail
        by_category.append({"category": c, "units": units, "avg_wholesale": round(avg_w, 2),
                            "est_cost": cost, "est_retail": retail})

    margin = round((total_retail - total_cost) / total_retail * 100, 1) if total_retail else None
    common.emit({
        "machines_selected": [m["machine_id"] for m in selected],
        "scope": "all live" if include_all else f"below {trigger:.0%} fill",
        "total_units": total_units,
        "est_cost_gbp": round(total_cost, 2),
        "est_retail_gbp": round(total_retail, 2),
        "implied_margin_pct": margin,
        "by_category": by_category,
        "note": "Estimate — allocation by catalog category share until planogram slot map exists (spike 005).",
    })


if __name__ == "__main__":
    main()
