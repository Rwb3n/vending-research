#!/usr/bin/env python3
"""
---
id: bos.demo.test_logbook
purpose: Guard the WRITE path that brief.py's run-twice determinism diff cannot see.
         Asserts policy.py --log is (a) week-anchored (date = Monday of the reviewed
         week, not wall-clock) and (b) idempotent per (week, policy, subject) — so a
         weekly logbook is reproducible and re-running a review never double-logs.
usage: "python bos/demo/test_logbook.py"  (exits non-zero on any failure)
updated: 2026-06-15
---
"""
import datetime
import os
import shutil
import subprocess
import sys
import tempfile

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
BOS = os.path.dirname(HERE)
REPO = os.path.dirname(BOS)
PY = sys.executable

CASES = 0
FAILS = 0


def check(cond, msg):
    global CASES, FAILS
    CASES += 1
    mark = "\033[32mPASS\033[0m" if cond else "\033[31mFAIL\033[0m"
    print(f"  [{mark}] {msg}")
    if not cond:
        FAILS += 1


def policy_log(env, week, decision="site-keep-fix-cut",
               inputs='{"rev_per_operator_hour": 36.9}', subject="north-star test"):
    cmd = [PY, "bos/tools/policy.py", "evaluate", "--decision", decision,
           "--input", inputs, "--subject", subject, "--week", week, "--log"]
    cenv = {**env, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    res = subprocess.run(cmd, cwd=REPO, env=cenv, capture_output=True,
                         text=True, encoding="utf-8")
    return res


def main():
    sandbox = tempfile.mkdtemp(prefix="bos-logbook-test-")
    # Point BOTH this process's reads AND the subprocess writes at the sandbox. Must be
    # set in os.environ BEFORE importing common (it resolves STATE_DIR at import time).
    os.environ["BOS_STATE_DIR"] = sandbox
    env = dict(os.environ)
    try:
        shutil.copytree(os.path.join(BOS, "state"), sandbox, dirs_exist_ok=True)
        # read records via the tool's OWN resolved path (avoids bash/win path skew)
        sys.path.insert(0, os.path.join(BOS, "tools"))
        import common  # noqa: E402
        decisions_path = os.path.join(common.STATE_DIR, "decisions.jsonl")

        def records():
            return common.load_jsonl(decisions_path)

        before = len(records())
        week = "2026-W24"

        print("Logbook write-path guard (sandboxed; committed state untouched)")

        # 1 — first log appends exactly one record
        policy_log(env, week)
        after_one = records()
        check(len(after_one) == before + 1, "first --log appends exactly one record")

        # 2 — week-anchored date == Monday of that ISO week (not wall-clock)
        expected = datetime.date.fromisocalendar(2026, 24, 1).isoformat()
        rec = [r for r in after_one if r.get("week") == week
               and r.get("subject") == "north-star test"][0]
        check(rec["date"] == expected,
              f"date is week-anchored ({rec['date']} == Monday {expected})")
        check(rec.get("week") == week, "record carries the week key")
        check(rec["date"] != datetime.date.today().isoformat()
              or expected == datetime.date.today().isoformat(),
              "date is not wall-clock today (unless today IS that Monday)")

        # 3 — idempotent: re-logging the same (week, policy, subject) is a no-op
        policy_log(env, week)
        policy_log(env, week)
        same = [r for r in records() if r.get("week") == week
                and r.get("subject") == "north-star test"]
        check(len(same) == 1, "re-logging same (week, policy, subject) does NOT duplicate")

        # 4 — a DIFFERENT week is a distinct record
        policy_log(env, "2026-W25")
        w25 = [r for r in records() if r.get("week") == "2026-W25"
               and r.get("subject") == "north-star test"]
        check(len(w25) == 1, "a different week logs a new, distinct record")

        print(f"\n{CASES - FAILS}/{CASES} checks passed.")
        sys.exit(1 if FAILS else 0)
    finally:
        shutil.rmtree(sandbox, ignore_errors=True)


if __name__ == "__main__":
    main()
