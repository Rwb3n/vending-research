"""
---
id: bos.tools.validate
purpose: Validate state files against their JSON Schema contracts. Run before writing
         or in build.py. Deterministic; exits non-zero on any violation.
inputs: bos/state/*, bos/schema/*
outputs: validation report (stdout JSON), exit code
usage: "python bos/tools/validate.py [<file> | --all]"
updated: 2026-06-15
---
"""
import os
import sys

from jsonschema import Draft202012Validator

import common

# file -> (schema name, how to extract the list of records to check)
TARGETS = {
    "machines_state.json": ("machine_state", lambda p: common.load_json(p)["machines"]),
    "scorecard.jsonl": ("scorecard_entry", common.load_jsonl),
    "cash_ledger.jsonl": ("cash_ledger_entry", common.load_jsonl),
    "incidents.jsonl": ("incident", common.load_jsonl),
    "decisions.jsonl": ("decision", common.load_jsonl),
}


def validate_file(fname):
    schema_name, extract = TARGETS[fname]
    schema = common.load_json(os.path.join(common.SCHEMA_DIR, f"{schema_name}.schema.json"))
    validator = Draft202012Validator(schema)
    records = extract(os.path.join(common.STATE_DIR, fname))
    errors = []
    for i, rec in enumerate(records):
        for err in validator.iter_errors(rec):
            errors.append({"file": fname, "record": i, "path": list(err.path), "message": err.message})
    return len(records), errors


def main():
    args = sys.argv[1:]
    if not args or args[0] == "--all":
        files = list(TARGETS)
    else:
        files = [os.path.basename(args[0])]
        if files[0] not in TARGETS:
            common.emit({"ok": False, "error": f"No schema mapping for {files[0]}", "known": list(TARGETS)})
            sys.exit(2)

    report = {"ok": True, "checked": {}, "errors": []}
    for fname in files:
        count, errors = validate_file(fname)
        report["checked"][fname] = count
        report["errors"].extend(errors)
    report["ok"] = not report["errors"]
    common.emit(report)
    sys.exit(0 if report["ok"] else 1)


if __name__ == "__main__":
    main()
