#!/usr/bin/env python3
"""
---
id: bos.generate_manual
purpose: Build the self-contained HTML operating manual — a readable handbook plus a
         live scorecard view (computed KPIs, red/amber/green). Renders bos/config + state.
inputs: bos/config/*, bos/state/*, bos/tools/kpi.py
outputs: TheSnackChoice_OS_Manual.html
usage: "python generate_bos_manual.py"
updated: 2026-06-15
---
"""
import datetime
import glob
import html
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "bos", "tools"))
import common  # noqa: E402
import kpi as kpi_tool  # noqa: E402

OUTPUT = os.environ.get("BOS_MANUAL_OUT") or os.path.join(HERE, "TheSnackChoice_OS_Manual.html")
E = lambda s: html.escape(str(s))

CSS = """
:root{--navy:#1F3864;--ink:#222;--mut:#666;--line:#e3e6ee;--bg:#f6f7fb}
*{box-sizing:border-box}body{margin:0;font:15px/1.55 -apple-system,Segoe UI,Roboto,Helvetica,Arial;color:var(--ink);background:var(--bg)}
header{background:var(--navy);color:#fff;padding:30px 40px}
header h1{margin:0 0 4px;font-size:26px}header p{margin:0;opacity:.85}
.wrap{max-width:1080px;margin:0 auto;padding:24px 40px 80px}
nav{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);padding:10px 40px;font-size:13px;z-index:5}
nav a{color:var(--navy);text-decoration:none;margin-right:16px;font-weight:600}
section{background:#fff;border:1px solid var(--line);border-radius:10px;padding:22px 26px;margin:20px 0}
h2{color:var(--navy);margin:0 0 4px;font-size:20px}.sub{color:var(--mut);margin:0 0 16px;font-size:13px}
.kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px}
.kpi{border:1px solid var(--line);border-radius:8px;padding:14px;border-left:5px solid #ccc}
.kpi.green{border-left-color:#2e9e5b}.kpi.amber{border-left-color:#d9a400}.kpi.red{border-left-color:#d6455b}.kpi.na{border-left-color:#bbb}
.kpi .v{font-size:26px;font-weight:700}.kpi .l{font-size:13px;color:var(--mut)}.kpi .t{font-size:12px;color:var(--mut)}
.star{color:#d9a400}
table{width:100%;border-collapse:collapse;font-size:13.5px}th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
th{background:#eef1f8;color:var(--navy)}
code{background:#eef1f8;padding:1px 5px;border-radius:4px;font-size:12.5px}
.pill{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px}
.pill.green{background:#d7f0e0;color:#1c6b3e}.pill.amber{background:#fbeec2;color:#8a6a00}.pill.red{background:#f7d4da;color:#9b2233}
.tag{font-size:11px;background:#eef1f8;color:var(--navy);padding:2px 7px;border-radius:5px;margin-right:4px}
.spike{color:#9b2233;font-weight:700}.evid{color:#1c6b3e;font-weight:700}
ul.r{margin:4px 0 0;padding-left:18px}ul.r li{margin:2px 0}
.foot{color:var(--mut);font-size:12px;text-align:center;margin-top:30px}
"""


def front_matter_and_body(path):
    import yaml
    text = open(path).read()
    fm, body = {}, text
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            fm = yaml.safe_load(text[3:end]) or {}
            body = text[end + 3:].strip()
    return fm, body


def render_kpis():
    kpis, weeks, machines = kpi_tool.load_inputs()
    latest = weeks[-1]
    cards = []
    for k in kpi_tool.compute_week(kpis, latest, machines):
        star = '<span class="star">★</span> ' if k["north_star"] else ""
        val = "—" if k["value"] is None else k["value"]
        cards.append(f"""<div class="kpi {k['status']}">
          <div class="v">{star}{E(val)} <span class="l">{E(k['unit'])}</span></div>
          <div class="l">{E(k['label'])}</div>
          <div class="t">target {E(k['target'])} · <span class="pill {k['status']}">{k['status'].upper()}</span></div>
        </div>""")
    return latest["week"], "\n".join(cards)


def section(sid, h, sub, body):
    return f'<section id="{sid}"><h2>{E(h)}</h2><p class="sub">{E(sub)}</p>{body}</section>'


def table(cols, rows):
    head = "".join(f"<th>{E(c)}</th>" for c in cols)
    body = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f"<table><tr>{head}</tr>{body}</table>"


def main():
    cfg = lambda name: common.load_yaml(os.path.join(common.CONFIG_DIR, name))

    week, kpi_cards = render_kpis()

    # cadence
    cad_rows = [[f"<b>{E(c['id'])}</b>", E(c["trigger"]), E(c["purpose"]),
                 "<ul class='r'>" + "".join(f"<li>{E(x)}</li>" for x in c["routines"]) + "</ul>"]
                for c in cfg("cadence.yaml")["cadence"]]

    # workflows + pipelines
    pl = {p["id"]: p for p in cfg("pipelines.yaml")["pipelines"]}
    wf_rows = []
    for w in cfg("workflows.yaml")["workflows"]:
        steps = []
        for pid in w.get("runs", []):
            steps.extend(pl.get(pid, {}).get("steps", []))
        chain = " → ".join(f"<code>{E(s)}</code>" for s in steps) or ", ".join(f"<code>{E(t)}</code>" for t in w.get("tasks", [])) or "(event-driven)"
        wf_rows.append([f"<b>{E(w['id'])}</b>", f"<span class='tag'>{E(w['cadence'])}</span>", chain, E(w["done_when"])])

    # policies
    pol_rows = []
    for p in cfg("policies.yaml")["policies"]:
        rules = "<ul class='r'>" + "".join(f"<li>if <code>{E(ru['when'])}</code> → <b>{E(ru['verdict'])}</b></li>" for ru in p.get("rules", []))
        rules += f"<li>else → <b>{E(p.get('default', {}).get('verdict', '?'))}</b></li></ul>"
        ev = p.get("evidence", "")
        ev_html = f'<span class="{"spike" if "SPIKE" in ev else "evid"}">{E(ev)}</span>'
        pol_rows.append([f"<code>{E(p['id'])}</code>", E(p["label"]), rules, ev_html])

    # playbooks
    pb_rows = []
    for path in sorted(glob.glob(os.path.join(common.CONFIG_DIR, "playbooks", "*.md"))):
        fm, _ = front_matter_and_body(path)
        pb_rows.append([f"<b>{E(fm.get('title',''))}</b>", E(str(fm.get("trigger", ""))),
                        E(fm.get("protects_kpi", "—")), f"<code>config/playbooks/{E(os.path.basename(path))}</code>"])

    # risks
    rk_rows = [[E(x["id"]), E(x["risk"]), f"{E(x['likelihood'])}/{E(x['impact'])}", E(x["signal"]), E(x["mitigation"])]
               for x in cfg("risks.yaml")["risks"]]

    # spikes
    sp_rows = []
    for path in sorted(glob.glob(os.path.join(HERE, "scratch", "spikes", "*.md"))):
        fm, _ = front_matter_and_body(path)
        unb = fm.get("unblocks", [])
        unb = ", ".join(unb) if isinstance(unb, list) else str(unb)
        sp_rows.append([f"<code>{E(fm.get('id',''))}</code>", E(fm.get("title", "")),
                        f"<span class='pill red'>{E(fm.get('status',''))}</span>", f"<code>{E(unb)}</code>"])

    # machines
    mc = common.load_json(os.path.join(common.STATE_DIR, "machines_state.json"))["machines"]
    mc_rows = [[E(m["machine_id"]), E(m["site"]), E(m["status"]),
                f"{round(m['current_fill']*100)}%", E(m.get("planogram", "")), E(m.get("last_restock", ""))] for m in mc]

    # planograms (summary: per planogram, facings + items + category mix)
    pg_rows = []
    for pid, pg in cfg("planograms.yaml")["planograms"].items():
        facings = sum(ln["facings"] for ln in pg["lines"])
        items = sum(ln["facings"] * ln["par_each"] for ln in pg["lines"])
        cats = {}
        for ln in pg["lines"]:
            cats[ln["category"]] = cats.get(ln["category"], 0) + ln["facings"]
        mix = ", ".join(f"{c} {n}" for c, n in sorted(cats.items(), key=lambda x: -x[1]))
        pg_rows.append([f"<code>{E(pid)}</code>", E(pg.get("label", "")), f"{facings} sel / {items} items", E(mix)])

    doc = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>The Snack Choice — Operating Manual</title><style>{CSS}</style></head><body>
<header><h1>The Snack Choice — Operating Manual</h1>
<p>The system that runs the business. North star: <b>revenue per operator-hour</b>. Optimised for agentic automation — everything is a file, tools are deterministic.</p></header>
<nav><a href="#scorecard">Scorecard</a><a href="#cadence">Cadence</a><a href="#work">Work</a>
<a href="#playbooks">Playbooks</a><a href="#policies">Policies</a><a href="#machines">Fleet</a><a href="#planograms">Planograms</a>
<a href="#risks">Risks</a><a href="#spikes">Assumptions</a><a href="#agents">For agents</a></nav>
<div class="wrap">
{section("scorecard", f"Scorecard — {week}", "Computed by bos/tools/kpi.py. ★ = north star. Never hand-edited.", f'<div class="kpis">{kpi_cards}</div>')}
{section("cadence", "Operating Cadence", "The rhythm. Workflows bind to these triggers.", table(["Cadence","Trigger","Purpose","Routines"], cad_rows))}
{section("work", "Tasks → Pipelines → Workflows", "No roles, only work. Any task runs by agent or human.", table(["Workflow","Cadence","Pipeline (tasks)","Done when"], wf_rows))}
{section("playbooks", "Playbooks (SOPs)", "Human-runnable procedures. Full text in bos/config/playbooks/.", table(["Playbook","Trigger","Protects KPI","File"], pb_rows))}
{section("policies", "Decision Policies", "Deterministic rules via bos/tools/policy.py. Verdicts logged to decisions.jsonl.", table(["Policy","Question","Rules (first match wins)","Evidence"], pol_rows))}
{section("machines", "Fleet", "Current state. Machines under 30% fill are due a restock.", table(["ID","Site","Status","Fill","Planogram","Last restock"], mc_rows))}
{section("planograms", "Planograms", "Slot→SKU→par plan per machine; reorder.py refills to par. Full lines in bos/config/planograms.yaml. Par levels are spike 007.", table(["ID","Label","Size","Facing mix"], pg_rows))}
{section("risks", "Risk Register", "Failure modes, early-warning signals, mitigations.", table(["ID","Risk","L/I","Signal","Mitigation"], rk_rows))}
{section("spikes", "Assumptions & Spikes", "Evidence or spike. Open spikes are guesses under test in scratch/spikes/.", table(["Spike","Question","Status","Unblocks"], sp_rows))}
{section("agents", "For agents", "Operate the digital layer cheaply.", '''
<p>Entry point: <code>bos/AGENTS.md</code>. Tool registry (the API): <code>bos/tools/manifest.json</code>.</p>
<p>A weekly review is: <code>scorecard.py append</code> → <code>kpi.py compute --latest</code> →
evaluate red KPIs with <code>policy.py evaluate --log</code> → <code>build.py</code>.
Tools do math; you do judgment. Validate before you write. Cite the policy that fired.</p>''')}
<p class="foot">Generated by generate_bos_manual.py on {datetime.date.today().isoformat()} · data: bos/config + bos/state</p>
</div></body></html>"""

    with open(OUTPUT, "w", encoding="utf-8") as fh:
        fh.write(doc)
    print(f"Wrote {OUTPUT} ({len(doc)} bytes)")


if __name__ == "__main__":
    main()
