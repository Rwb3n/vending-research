"""
---
id: bos.tools.policy
purpose: Evaluate a deterministic decision policy against an input and (optionally) log
         the verdict to state/decisions.jsonl. Rules are evaluated top-down; first match
         wins; else the policy default. Every verdict is auditable (policy + reasons).
inputs: bos/config/policies.yaml, input JSON
outputs: verdict (stdout JSON); optional appended decision record
usage: |
  python bos/tools/policy.py list
  python bos/tools/policy.py evaluate --decision site-go-nogo --input '{"vending_score": 8}' [--subject "Name"] [--log]
updated: 2026-06-15
---
"""
import datetime
import json
import os
import sys

import common

POLICIES_PATH = os.path.join(common.CONFIG_DIR, "policies.yaml")
DECISIONS_PATH = os.path.join(common.STATE_DIR, "decisions.jsonl")


def get_arg(name, default=None):
    return sys.argv[sys.argv.index(name) + 1] if name in sys.argv else default


def parse_input(raw):
    if raw is None:
        return {}
    if raw.startswith("@"):
        return common.load_json(raw[1:])
    return json.loads(raw)


def evaluate(policy, inputs):
    for rule in policy.get("rules", []):
        if common.safe_eval(rule["when"], inputs) is True:
            return rule["verdict"], [rule.get("reason", "")]
    default = policy.get("default", {})
    return default.get("verdict", "UNDECIDED"), [default.get("reason", "no rule matched")]


def next_decision_id():
    year = datetime.date.today().year
    existing = common.load_jsonl(DECISIONS_PATH) if os.path.exists(DECISIONS_PATH) else []
    nums = [int(d["id"].split("-")[-1]) for d in existing if d.get("id", "").startswith(f"DEC-{year}-")]
    return f"DEC-{year}-{max(nums, default=0) + 1:03d}"


def main():
    policies = {p["id"]: p for p in common.load_yaml(POLICIES_PATH)["policies"]}
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""

    if cmd == "list":
        common.emit([{"id": p["id"], "label": p["label"], "inputs": p["inputs"]} for p in policies.values()])
        return

    if cmd != "evaluate":
        common.emit({"error": "usage: policy.py evaluate --decision <id> --input <json> [--subject S] [--log]"})
        sys.exit(2)

    pid = get_arg("--decision")
    if pid not in policies:
        common.emit({"error": f"unknown policy {pid}", "known": list(policies)})
        sys.exit(1)
    policy = policies[pid]
    inputs = parse_input(get_arg("--input"))
    verdict, reasons = evaluate(policy, inputs)

    result = {"policy": pid, "subject": get_arg("--subject"), "verdict": verdict,
              "reasons": reasons, "inputs": inputs}

    if "--log" in sys.argv:
        record = {"id": next_decision_id(), "date": datetime.date.today().isoformat(),
                  "policy": pid, "subject": get_arg("--subject", pid), "verdict": verdict,
                  "reasons": reasons, "inputs": inputs, "decided_by": "agent"}
        common.append_jsonl(DECISIONS_PATH, record)
        result["logged"] = record["id"]

    common.emit(result)


if __name__ == "__main__":
    main()
