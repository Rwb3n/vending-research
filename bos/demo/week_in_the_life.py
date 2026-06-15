#!/usr/bin/env python3
"""
---
id: bos.demo.runner
purpose: Reproducible "week in the life" demo. Drives the REAL deterministic tools
         through the weekly loop (close -> read -> decide -> replenish -> regenerate),
         against a throwaway sandbox copy of state so committed data is never touched.
inputs: bos/tools/*, bos/config/*, bos/state/* (copied to a temp sandbox)
outputs: narrated transcript (stdout); demo manual at bos/demo/out/
usage: "python bos/demo/week_in_the_life.py [--quiet] [--no-build]"
updated: 2026-06-15
---
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
BOS = os.path.dirname(HERE)
REPO = os.path.dirname(BOS)
PY = sys.executable
QUIET = "--quiet" in sys.argv
NO_BUILD = "--no-build" in sys.argv

# the not-quite-good W25 numbers that make the demo interesting
W25 = {"week": "2026-W25", "operator_hours": 16, "route_hours": 9.5,
       "gross_revenue": 561.0, "cogs": 230.0, "vends": 404, "machine_days": 28,
       "stockout_days": 2, "cash_collected": 546.0, "cash_expected": 561.0,
       "note": "demo: vends slip to amber; cash variance amber."}


def say(*lines):
    if not QUIET:
        for ln in lines:
            print(ln)


def banner(n, title):
    say("", "\033[1;36m" + "=" * 64 + "\033[0m",
        f"\033[1;36m  STEP {n} — {title}\033[0m",
        "\033[1;36m" + "=" * 64 + "\033[0m")


def run(args, env, show=True):
    """Run a tool, echo a compact view, return parsed JSON (or None)."""
    cmd = [PY] + args
    if not QUIET:
        print(f"\n\033[2m$ {' '.join(args)}\033[0m")
    res = subprocess.run(cmd, cwd=REPO, env=env, capture_output=True, text=True)
    out = res.stdout.strip()
    if res.returncode != 0 and not out:
        print(res.stderr.strip())
        return None
    if QUIET:  # quiet mode = raw tool output for every step, no narration
        print(out)
    try:
        data = json.loads(out)
    except json.JSONDecodeError:
        if show and not QUIET:
            print(out)
        return None
    return data


def kpi_table(report):
    rows = []
    for k in report["kpis"]:
        dot = {"green": "🟢", "amber": "🟡", "red": "🔴", "na": "⚪"}[k["status"]]
        star = "★" if k["north_star"] else " "
        val = "—" if k["value"] is None else k["value"]
        rows.append(f"   {dot} {star} {k['label']:<34} {str(val):>8} {k['unit']:<10} (target {k['target']})")
    return "\n".join(rows)


def main():
    env = dict(os.environ)
    sandbox = tempfile.mkdtemp(prefix="bos-demo-state-")
    out_dir = os.path.join(HERE, "out")
    os.makedirs(out_dir, exist_ok=True)
    try:
        # sandbox state so committed data is never mutated
        shutil.copytree(os.path.join(BOS, "state"), sandbox, dirs_exist_ok=True)
        env["BOS_STATE_DIR"] = sandbox

        say("\n\033[1mThe Snack Choice — a week in the life\033[0m",
            "Monday, 2026-W25. Driving the real tools against a sandbox copy of state.",
            f"(sandbox: {sandbox})")

        # 1 — close the week
        banner(1, "Close the week (log raw inputs)")
        say("Last week's totals go in as raw inputs. KPIs are derived later — never typed.")
        r = run(["bos/tools/scorecard.py", "append", "--input", json.dumps(W25)], env)
        say(f"   → appended {r.get('appended') if r else '?'}")

        # 2 — read the one number
        banner(2, "Read the one number (compute KPIs)")
        say("North star is revenue per operator-hour. Tools do the math; we read status.")
        report = run(["bos/tools/kpi.py", "compute", "--latest"], env, show=False)
        say(kpi_table(report))
        ambers = [k for k in report["kpis"] if k["status"] in ("amber", "red")]
        say("", f"   {len(ambers)} KPI(s) need a look: " + ", ".join(k["label"] for k in ambers))

        # 3 — decide (policies, logged + auditable)
        banner(3, "Decide (evaluate policies, log verdicts)")
        say("Each soft KPI maps to a policy. Verdicts are logged with reasons — auditable.")

        def decide(decision, inputs, subject, note):
            say(f"\n• {note}")
            d = run(["bos/tools/policy.py", "evaluate", "--decision", decision,
                     "--input", json.dumps(inputs), "--subject", subject, "--log"], env, show=False)
            if d:
                say(f"   → \033[1m{d['verdict']}\033[0m — {d['reasons'][0]}"
                    + (f"  [logged {d['logged']}]" if d.get("logged") else ""))
            return d

        ns = next(k for k in report["kpis"] if k["north_star"])
        decide("site-keep-fix-cut", {"rev_per_operator_hour": ns["value"]},
               "Beddington Lane estate",
               "Estate keep/fix/cut — against north-star revenue/operator-hour:")
        decide("cash-investigate", {"variance_pct": -8.3}, "M03 collection 2026-06-12",
               "Cash variance — a collection came in 8.3% light at M03:")
        margin = next(k for k in report["kpis"] if k["id"] == "gross_margin_pct")
        decide("price-change", {"gross_margin_pct": margin["value"]}, "fleet blended margin",
               "Margin check — confirm prices/terms still healthy:")

        # 4 — replenish
        banner(4, "Replenish (clustered route + per-SKU order)")
        say("Restock by trigger, not habit. Route clusters same-site stops to protect £/hr.")
        route = run(["bos/tools/route.py", "plan"], env, show=False)
        for s in route["route"]:
            ids = ", ".join(f"{m['machine_id']} ({m['current_fill']:.0%})" for m in s["machines"])
            say(f"   stop: {s['site']} [{s['postcode']}] → {ids}")
        say(f"   {route['machines_due']} machines due, {route['total_units']} units, {route['stops']} clustered stop(s)")

        say("\nPurchase order to refill those machines to planogram par:")
        po = run(["bos/tools/reorder.py"], env, show=False)
        say(f"   {len(po['lines'])} SKU lines · {po['total_units']} units · "
            f"cost £{po['est_cost_gbp']} → retail £{po['est_retail_gbp']} · "
            f"margin {po['implied_margin_pct']}%")
        for line in po["lines"][:5]:
            say(f"     - {line['sku']:<34} {line['units']:>3}u  £{line['est_cost']}")
        say(f"     … and {len(po['lines']) - 5} more lines")

        # 5 — close the loop
        if not NO_BUILD:
            banner(5, "Close the loop (regenerate the manual)")
            say("The manual + scorecard regenerate from the new state — always current.")
            manual_out = os.path.join(out_dir, "TheSnackChoice_OS_Manual.html")
            benv = dict(env)
            benv["BOS_MANUAL_OUT"] = manual_out
            run(["generate_bos_manual.py"], benv)
            say(f"   → {os.path.relpath(manual_out, REPO)} (open in a browser)")

        # wrap
        say("", "\033[1;32m" + "─" * 64 + "\033[0m",
            "\033[1;32m  Loop closed.\033[0m  3 decisions logged · route planned · PO ready · manual refreshed.",
            "  Committed state was never touched (sandbox).",
            "\033[1;32m" + "─" * 64 + "\033[0m")
    finally:
        shutil.rmtree(sandbox, ignore_errors=True)


if __name__ == "__main__":
    main()
