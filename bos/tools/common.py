"""
---
id: bos.tools.common
purpose: Shared helpers for the BOS deterministic tools — file IO (everything is a
         file), JSONL header-skipping, a safe arithmetic evaluator (no eval() of
         arbitrary code), and KPI status banding. No LLM, fully deterministic.
inputs: bos/config, bos/state, bos/schema, data/
consumed_by: [validate.py, kpi.py, policy.py, route.py, reorder.py, scorecard.py]
updated: 2026-06-15
---
"""
import ast
import json
import operator
import os

import yaml

# --- paths (everything is a file; resolve relative to this tool's location) ---
TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
BOS_DIR = os.path.dirname(TOOLS_DIR)
REPO_DIR = os.path.dirname(BOS_DIR)
CONFIG_DIR = os.path.join(BOS_DIR, "config")
STATE_DIR = os.environ.get("BOS_STATE_DIR") or os.path.join(BOS_DIR, "state")
SCHEMA_DIR = os.path.join(BOS_DIR, "schema")
DATA_DIR = os.path.join(REPO_DIR, "data")


def load_yaml(path):
    with open(path) as fh:
        return yaml.safe_load(fh)


def load_json(path):
    with open(path) as fh:
        return json.load(fh)


def load_jsonl(path):
    """Load a JSONL file, skipping the leading `_meta` header record."""
    records = []
    with open(path) as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            obj = json.loads(line)
            if isinstance(obj, dict) and "_meta" in obj and len(obj) == 1:
                continue  # header line
            records.append(obj)
    return records


def append_jsonl(path, record):
    """Append one record to a JSONL log (the header must already exist).

    Tolerates a file whose last line lacks a trailing newline, so records can never
    be concatenated onto an existing line."""
    needs_nl = False
    if os.path.exists(path) and os.path.getsize(path) > 0:
        with open(path, "rb") as fh:
            fh.seek(-1, os.SEEK_END)
            needs_nl = fh.read(1) != b"\n"
    with open(path, "a") as fh:
        fh.write(("\n" if needs_nl else "") + json.dumps(record) + "\n")


# --- safe arithmetic / comparison evaluator (for kpi & policy expressions) ---
_BIN_OPS = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod, ast.Pow: operator.pow,
}
_CMP_OPS = {
    ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt,
    ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge,
}
_FUNCS = {"abs": abs, "min": min, "max": max, "round": round}


def safe_eval(expr, names):
    """Evaluate a restricted arithmetic/boolean expression against `names`.

    Supports + - * / // % **, comparisons, and/or/not, and abs/min/max/round.
    Raises on anything else. Returns None on division by zero / missing name."""
    try:
        tree = ast.parse(expr, mode="eval")
        return _eval_node(tree.body, names)
    except (ZeroDivisionError, _MissingName):
        return None


class _MissingName(Exception):
    pass


def _eval_node(node, names):
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name):
        if node.id not in names:
            raise _MissingName(node.id)
        return names[node.id]
    if isinstance(node, ast.BinOp) and type(node.op) in _BIN_OPS:
        return _BIN_OPS[type(node.op)](_eval_node(node.left, names), _eval_node(node.right, names))
    if isinstance(node, ast.UnaryOp):
        val = _eval_node(node.operand, names)
        if isinstance(node.op, ast.USub):
            return -val
        if isinstance(node.op, ast.UAdd):
            return +val
        if isinstance(node.op, ast.Not):
            return not val
    if isinstance(node, ast.BoolOp):
        vals = [_eval_node(v, names) for v in node.values]
        return all(vals) if isinstance(node.op, ast.And) else any(vals)
    if isinstance(node, ast.Compare):
        left = _eval_node(node.left, names)
        for op, comp in zip(node.ops, node.comparators):
            right = _eval_node(comp, names)
            if type(op) not in _CMP_OPS or not _CMP_OPS[type(op)](left, right):
                return False
            left = right
        return True
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in _FUNCS:
        args = [_eval_node(a, names) for a in node.args]
        return _FUNCS[node.func.id](*args)
    raise ValueError(f"Disallowed expression element: {ast.dump(node)}")


def status_for(value, direction, thresholds):
    """Band a value green / amber / red per a KPI's direction and thresholds."""
    if value is None:
        return "na"
    green, red = thresholds.get("green"), thresholds.get("red")
    if direction == "higher_better":
        if value >= green:
            return "green"
        return "red" if value < red else "amber"
    if direction == "lower_better":
        if value <= green:
            return "green"
        return "red" if value > red else "amber"
    if direction == "abs_lower_better":
        v = abs(value)
        if v <= green:
            return "green"
        return "red" if v > red else "amber"
    return "na"


def emit(obj):
    """Print JSON to stdout (pipe-friendly)."""
    print(json.dumps(obj, indent=2))
