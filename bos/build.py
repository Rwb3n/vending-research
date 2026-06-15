"""
---
id: bos.build
purpose: Orchestrator — the one entrypoint that runs the digital layer end to end:
         validate all state -> compute latest KPIs -> regenerate workbook + HTML manual.
         Pipes: each step's failure stops the build (non-zero exit).
inputs: bos/config/*, bos/state/*
outputs: TheSnackChoice_OperatingSystem.xlsx, TheSnackChoice_OS_Manual.html
usage: "python bos/build.py"
updated: 2026-06-15
---
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)


def run(label, cmd):
    print(f"\n=== {label} ===")
    result = subprocess.run([sys.executable, *cmd], cwd=REPO)
    if result.returncode != 0:
        print(f"!! {label} failed (exit {result.returncode}); stopping build.")
        sys.exit(result.returncode)


def main():
    run("validate state", ["bos/tools/validate.py", "--all"])
    run("compute KPIs (latest week)", ["bos/tools/kpi.py", "compute", "--latest"])
    run("generate workbook", ["generate_bos.py"])
    run("generate HTML manual", ["generate_bos_manual.py"])
    print("\nBuild complete: TheSnackChoice_OperatingSystem.xlsx + TheSnackChoice_OS_Manual.html")


if __name__ == "__main__":
    main()
