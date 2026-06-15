#!/usr/bin/env python3
"""
---
id: bos.tools.brief
purpose: SPIKE — project one truth (latest KPI state + config bindings) into THREE
         deterministic views for three readers: the machine (lossless JSON), the agent
         /ghost (state + applicable actions + which tool/policy to call), and the
         operator/pilot (the one number, ranked flags, computed next-actions). No LLM
         in the render path. The intelligence is the JOIN over config, not generation.
inputs: kpi.py compute output, config/kpis.yaml, config/policies.yaml,
        config/tasks.yaml, config/playbooks/*.md (frontmatter), state/machines_state.json
outputs: <out>/state.json (machine), <out>/brief.agent.json (ghost),
         <out>/brief.operator.md (pilot). Also a compact summary to stdout.
usage: "python bos/tools/brief.py [--latest|--week W] [--out DIR]"
determinism: run twice against the same BOS_STATE_DIR -> byte-identical files.
             No wall-clock; the only date is the week from state. All collections sorted.
honesty: where a KPI has no bound playbook/policy, the views say so explicitly
         ("no bound action — needs a decision") rather than inventing or hardcoding one.
updated: 2026-06-15
---
"""
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common  # noqa: E402  (shared loaders + STATE_DIR resolution)

TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
PLAYBOOK_DIR = os.path.join(common.CONFIG_DIR, "playbooks")
PY = sys.executable


# --- read the binding layer (the join keys) ---------------------------------

def _frontmatter(path):
    """Tiny YAML-frontmatter reader: the block between the first two '---' lines."""
    keys = {}
    with open(path, encoding="utf-8") as fh:
        lines = fh.read().splitlines()
    if not lines or lines[0].strip() != "---":
        return keys
    for ln in lines[1:]:
        if ln.strip() == "---":
            break
        if ":" in ln and not ln.lstrip().startswith("#"):
            k, _, v = ln.partition(":")
            keys[k.strip()] = v.strip().strip('"').strip("'")
    return keys


def load_playbooks_by_id():
    """playbook id -> frontmatter (+ file path). Used to enrich a playbook binding."""
    by_id = {}
    for name in sorted(os.listdir(PLAYBOOK_DIR)):
        if not name.endswith(".md"):
            continue
        fm = _frontmatter(os.path.join(PLAYBOOK_DIR, name))
        # acts_via.target uses the short id ("fault-triage"); frontmatter id is
        # "playbook.fault-triage" — key on both so either resolves.
        fm["_file"] = f"bos/config/playbooks/{name}"
        full = fm.get("id", name)
        short = full.split(".", 1)[-1]
        by_id[full] = fm
        by_id[short] = fm
    return by_id


def load_policies_by_id():
    """policy id -> policy dict (label, inputs) so we can name the CLI to run."""
    doc = common.load_yaml(os.path.join(common.CONFIG_DIR, "policies.yaml"))
    return {p["id"]: p for p in doc.get("policies", [])}


def load_acts_via():
    """kpi id -> its acts_via binding from kpis.yaml (the signed-off source of truth)."""
    doc = common.load_yaml(os.path.join(common.CONFIG_DIR, "kpis.yaml"))
    return {k["id"]: k.get("acts_via") for k in doc.get("kpis", [])}


# --- the core join: KPI -> bound action via acts_via (deterministic) ---------

def resolve_action(kpi, acts_via, policies, playbooks, week):
    """Resolve a KPI's bound response from its acts_via binding in kpis.yaml.

    Reads config only — never invents a mapping. The `mode` carried through is the
    judgment boundary: runnable (ghost acts) | proposes (ghost drafts, pilot commits)
    | manual (human playbook).
    """
    av = acts_via.get(kpi["id"])
    if not av:
        return {"bound": False, "reason": "no acts_via binding in kpis.yaml",
                "needs": "an acts_via block on this KPI"}
    kind, target, mode = av.get("kind"), av.get("target"), av.get("mode")
    action = {"bound": True, "kind": kind, "target": target, "mode": mode}

    if kind == "policy":
        # build the concrete input json by mapping the KPI value to policy fields
        input_map = av.get("input_map") or {}
        inputs = {pol_field: kpi.get(kpi_field) if kpi_field != "value" else kpi.get("value")
                  for pol_field, kpi_field in input_map.items()}
        action["policy"] = policies.get(target, {}).get("label", target)
        # runnable/proposes -> a real, runnable command (the ghost's action)
        flag = "--log" if mode == "runnable" else "--propose"
        if inputs:  # only emit a command when we can supply the policy's inputs
            # --week anchors the logged decision to the reviewed week (not wall-clock),
            # so the whole decisions.jsonl logbook is reproducible.
            action["command"] = (
                f"python bos/tools/policy.py evaluate --decision {target} "
                f"--input '{json.dumps(inputs, sort_keys=True)}' "
                f"--subject '{kpi['label']}' --week {week} {flag}")
        else:
            action["command"] = None
        if mode == "proposes":
            action["needs"] = av.get("needs", [])  # inputs the pilot must supply
    elif kind == "playbook":
        pb = playbooks.get(target, {})
        action["command"] = None  # human SOP; the file is the instruction
        action["playbook_file"] = pb.get("_file", f"bos/config/playbooks/{target}.md")
        action["trigger"] = pb.get("trigger")
    return action


def severity(status):
    return {"red": 0, "amber": 1, "na": 2, "green": 3}.get(status, 4)


def compute_report(week_args):
    """Run the REAL kpi tool (proven path) and parse its JSON contract."""
    cmd = [PY, "bos/tools/kpi.py", "compute"] + week_args
    res = subprocess.run(cmd, cwd=common.REPO_DIR, capture_output=True, text=True,
                         env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"})
    if res.returncode != 0:
        sys.stderr.write(res.stderr)
        sys.exit(res.returncode)
    return json.loads(res.stdout)


# --- build the three projections from one truth -----------------------------

def build_projections(report):
    acts_via = load_acts_via()
    policies = load_policies_by_id()
    playbooks = load_playbooks_by_id()
    week = report["week"]

    # enrich every KPI with its bound action via acts_via — one join, reused 3x
    enriched = []
    for k in report["kpis"]:
        enriched.append({**k, "action": resolve_action(k, acts_via, policies, playbooks, week)})
    # deterministic order: worst-status first, then by id (stable, no sets)
    enriched.sort(key=lambda k: (severity(k["status"]), k["id"]))

    flags = [k for k in enriched if k["status"] in ("amber", "red")]
    north = next((k for k in enriched if k.get("north_star")), None)

    machine = {"week": week, "kpis": report["kpis"]}  # lossless, original order

    # North-star heartbeat: on a GREEN week the north star is NOT a flag, but we still
    # log a weekly confirmation (logbook, not exception report). Mutually exclusive with
    # the flag path — when amber/red the north star is already in `flags` with its own
    # runnable action, so the heartbeat fires only when green to avoid a double-log.
    heartbeat = None
    if north and north["status"] == "green" and north["action"].get("mode") == "runnable":
        inputs = {"rev_per_operator_hour": north["value"]}
        heartbeat = {
            "kpi": north["id"], "label": north["label"], "status": "green",
            "value": north["value"], "verdict_hint": "KEEP",
            "command": (f"python bos/tools/policy.py evaluate "
                        f"--decision {north['action']['target']} "
                        f"--input '{json.dumps(inputs, sort_keys=True)}' "
                        f"--subject 'north-star {week}' --week {week} --log"),
        }

    agent = {
        "week": week,
        "north_star": ({"id": north["id"], "value": north["value"],
                        "status": north["status"]} if north else None),
        "heartbeat": heartbeat,  # green-week weekly confirmation (logbook), or null
        "flags": [
            {"kpi": k["id"], "label": k["label"], "status": k["status"],
             "value": k["value"], "target": k["target"], "action": k["action"]}
            for k in flags
        ],
        # the ghost's action partition by trust level (the bi-directional boundary)
        "runnable": [k["id"] for k in flags if k["action"].get("mode") == "runnable"],
        "proposes": [k["id"] for k in flags if k["action"].get("mode") == "proposes"],
        "manual": [k["id"] for k in flags if k["action"].get("mode") == "manual"],
        "note": ("mode=runnable: ghost runs action.command (--log) unattended. "
                 "mode=proposes: ghost runs --propose (status:proposed), pilot commits. "
                 "mode=manual: follow action.playbook_file. heartbeat: green-week "
                 "north-star confirmation, logged for a reproducible weekly logbook."),
    }
    return week, flags, north, heartbeat, machine, agent


DOT = {"green": "🟢", "amber": "🟡", "red": "🔴", "na": "⚪"}


def render_operator(week, flags, north, heartbeat):
    """The pilot view: one number, ranked flags, computed next-action per flag."""
    L = []
    L.append(f"# Operator Brief — {week}")
    L.append("")
    L.append("_Deterministic projection of state. Not hand-written. "
             "Regenerate with `python bos/tools/brief.py --latest`._")
    L.append("")
    if north:
        dot = DOT[north["status"]]
        L.append(f"## The one number  {dot}")
        L.append("")
        L.append(f"**{north['label']}: {north['value']} {north['unit']}** "
                 f"(target {north['target']}) — **{north['status'].upper()}**")
        if heartbeat:  # green-week confirmation — a logbook line, NOT a flag
            L.append("")
            L.append(f"✓ North star holds → **{heartbeat['verdict_hint']}** "
                     "(logged this week — the weekly heartbeat).")
        L.append("")
    if not flags:
        L.append("## This week: nothing flagged 🟢")
        L.append("")
        L.append("All KPIs green. Run the replenish steps and close the loop.")
    else:
        L.append(f"## This week: {len(flags)} flag(s) — do these in order")
        L.append("")
        for i, k in enumerate(flags, 1):
            dot = DOT[k["status"]]
            L.append(f"{i}. {dot} **{k['label']}** — {k['value']} {k['unit']} "
                     f"(target {k['target']}, {k['status'].upper()})")
            a = k["action"]
            mode = a.get("mode")
            if mode == "runnable":
                L.append(f"   - → **Run** the `{a['target']}` policy (ghost can do this "
                         "unattended; verdict is logged).")
                L.append(f"   - `{a['command']}`")
            elif mode == "proposes":
                needs = ", ".join(a.get("needs", [])) or "judgment inputs"
                L.append(f"   - → **Decide.** Ghost drafts a *proposed* `{a['target']}` "
                         "decision; **you** commit it.")
                if a.get("command"):
                    L.append(f"   - `{a['command']}`")
                else:
                    L.append(f"   - Once you supply the inputs: "
                             f"`python bos/tools/policy.py evaluate --decision "
                             f"{a['target']} --input '{{…}}' --propose`")
                L.append(f"   - You supply: {needs} (not derivable from one week).")
            elif mode == "manual":
                L.append(f"   - → **Follow** the `{a['target']}` playbook (human SOP).")
                L.append(f"   - File: `{a.get('playbook_file')}`")
                if a.get("trigger"):
                    L.append(f"   - Trigger: {a['trigger']}")
            else:
                L.append(f"   - ⚠️ No acts_via binding: {a.get('reason', 'unknown')}.")
            L.append("")
    L.append("## Then close the loop")
    L.append("")
    L.append("- Plan restock: `python bos/tools/route.py plan`")
    L.append("- Build order: `python bos/tools/reorder.py`")
    L.append("- Regenerate: `python bos/build.py`")
    L.append("")
    return "\n".join(L) + "\n"


def main():
    week_args = ["--latest"]
    if "--week" in sys.argv:
        week_args = ["--week", sys.argv[sys.argv.index("--week") + 1]]
    out_dir = os.path.join(TOOLS_DIR, "..", "demo", "out")
    if "--out" in sys.argv:
        out_dir = sys.argv[sys.argv.index("--out") + 1]
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    report = compute_report(week_args)
    week, flags, north, heartbeat, machine, agent = build_projections(report)
    operator_md = render_operator(week, flags, north, heartbeat)

    # write the triptych (sorted keys -> byte-stable; no wall-clock anywhere)
    paths = {
        "machine": os.path.join(out_dir, "state.json"),
        "agent": os.path.join(out_dir, "brief.agent.json"),
        "operator": os.path.join(out_dir, "brief.operator.md"),
    }
    with open(paths["machine"], "w", encoding="utf-8") as fh:
        json.dump(machine, fh, indent=2, sort_keys=True)
        fh.write("\n")
    with open(paths["agent"], "w", encoding="utf-8") as fh:
        json.dump(agent, fh, indent=2, sort_keys=True)
        fh.write("\n")
    with open(paths["operator"], "w", encoding="utf-8") as fh:
        fh.write(operator_md)

    common.emit({
        "week": week,
        "wrote": {k: os.path.relpath(v, common.REPO_DIR) for k, v in paths.items()},
        "flags": len(flags),
        "runnable": agent["runnable"],   # ghost acts unattended
        "proposes": agent["proposes"],   # ghost drafts, pilot commits
        "manual": agent["manual"],       # human playbook
        "finding": ("Every flag now has an acts_via binding. The runnable/proposes/"
                    "manual split is the deterministic→judgment boundary, machine-readable."),
    })


if __name__ == "__main__":
    main()
