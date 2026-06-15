"""
---
id: bos.tools.scorecard
purpose: Append a validated weekly scorecard entry. Validates against the contract
         BEFORE writing (never append a bad record). Raw inputs only — KPIs are derived
         later by kpi.py.
inputs: scorecard_entry JSON, bos/schema/scorecard_entry.schema.json
outputs: appended record in bos/state/scorecard.jsonl
usage: |
  python bos/tools/scorecard.py append --input '{"week":"2026-W25", ...}'
  python bos/tools/scorecard.py append --file entry.json
updated: 2026-06-15
---
"""
import json
import os
import sys

from jsonschema import Draft202012Validator

import common

PATH = os.path.join(common.STATE_DIR, "scorecard.jsonl")
SCHEMA = os.path.join(common.SCHEMA_DIR, "scorecard_entry.schema.json")


def get_arg(name):
    return sys.argv[sys.argv.index(name) + 1] if name in sys.argv else None


def main():
    if len(sys.argv) < 2 or sys.argv[1] != "append":
        common.emit({"error": "usage: scorecard.py append --input <json> | --file <path>"})
        sys.exit(2)

    raw = get_arg("--input")
    record = json.loads(raw) if raw else common.load_json(get_arg("--file"))

    schema = common.load_json(SCHEMA)
    errors = [{"path": list(e.path), "message": e.message}
              for e in Draft202012Validator(schema).iter_errors(record)]
    if errors:
        common.emit({"ok": False, "errors": errors})
        sys.exit(1)

    if any(w["week"] == record["week"] for w in common.load_jsonl(PATH)):
        common.emit({"ok": False, "error": f"week {record['week']} already logged"})
        sys.exit(1)

    common.append_jsonl(PATH, record)
    common.emit({"ok": True, "appended": record["week"]})


if __name__ == "__main__":
    main()
