#!/usr/bin/env python3
"""AI Scope Builder v6 - strict allow-list wrapper.

This version wraps V5 and removes irrelevant suggestions using hard allow-lists
for the object + intent detected in the description.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from pathlib import Path

import ai_scope_builder_v5 as v5
import ai_scope_builder_v4 as v4
import ai_scope_builder_v3 as v3

ALLOW = {
    ("radiator", "replace"): {"RAD-FIT", "ADD-RAD-REMOVE", "RAD-VALVES", "PIPE-ALTER-M", "RAD-BALANCE"},
    ("radiator", "move"): {"RAD-MOVE", "PIPE-ALTER-M", "ADD-RAD-REMOVE", "RAD-FIT", "RAD-BALANCE"},
    ("radiator", "remove"): {"ADD-RAD-REMOVE", "PIPE-CAP", "ADD-PIPE-REMOVE-CAP"},
    ("toilet", "repair"): {"WC-FILL", "WC-FLUSH"},
    ("toilet", "unblock"): {"WC-UNBLOCK"},
    ("toilet", "replace"): {"WC-INSTALL", "WC-REMOVE", "ADD-WC-WASTE-M", "PIPE-ALTER-M"},
    ("toilet", "install"): {"WC-INSTALL", "ADD-WC-WASTE-M", "PIPE-ALTER-M"},
    ("toilet", "remove"): {"WC-REMOVE", "PIPE-CAP", "ADD-PIPE-REMOVE-CAP"},
    ("basin", "remove"): {"BASIN-REMOVE", "PIPE-CAP", "ADD-PIPE-REMOVE-CAP", "BASIN-WASTE"},
    ("basin", "replace"): {"BASIN-REMOVE", "BASIN-VANITY", "PB-BASIN-PED", "BASIN-WASTE", "TAP-BASIN", "PIPE-ALTER-M"},
    ("basin", "install"): {"BASIN-VANITY", "PB-BASIN-PED", "BASIN-WASTE", "TAP-BASIN", "PIPE-ALTER-M"},
    ("outside tap", "install"): {"OUTSIDE-TAP", "ADD-HC-PIPE-M", "PIPE-ALTER-M"},
    ("pipework", "cap"): {"PIPE-CAP", "ADD-PIPE-REMOVE-CAP"},
    ("pipework", "remove"): {"PIPE-CAP", "ADD-PIPE-REMOVE-CAP"},
    ("pipework", "move"): {"PIPE-ALTER-M", "ADD-HC-PIPE-M"},
}


def allowed_codes(description: str) -> tuple[set[str], set[str], str]:
    locked = v4.detect_primary_objects(description)
    intent = v5.infer_trade_intent(description)
    allowed: set[str] = set()
    for obj in locked:
        allowed |= ALLOW.get((obj, intent), set())
    return allowed, locked, intent


def build_scope_v6(description: str, workbook_path: Path) -> v3.V3Result:
    base = v5.build_scope_v5(description, workbook_path)
    allowed, locked, intent = allowed_codes(description)

    if not allowed:
        warnings = list(base.warnings)
        warnings.append("V6 strict allow-list: no specific allow-list found, using V5 result")
        return v3.V3Result(base.description, base.workbook_path, base.confidence, base.suggestions, base.questions, warnings)

    kept = [s for s in base.suggestions if s.code.upper() in allowed]

    # Add expected missing removal tasks such as BASIN-REMOVE/WC-REMOVE.
    workbook_codes = {t.code.upper() for t in v3.read_task_library(workbook_path)}
    existing_codes = {s.code.upper() for s in kept}
    for code in sorted(allowed):
        if code in existing_codes or code in workbook_codes:
            continue
        if code.endswith("-REMOVE"):
            kept.append(v3.V3Suggestion(
                code=code,
                name=code,
                confidence=90,
                source="missing-expected-task",
                exists_in_workbook=False,
                workbook_name=None,
                reason="Expected for this object/intent but not currently in the Task Library.",
            ))

    warnings = list(base.warnings)
    warnings.append("V6 strict allow-list active: " + ", ".join(sorted(allowed)))
    warnings.append("V6 object lock: " + ", ".join(sorted(locked)))
    warnings.append("V6 intent lock: " + intent)

    missing = [s.code for s in kept if not s.exists_in_workbook]
    if missing:
        warnings.append("Suggested task codes not currently in workbook Task Library: " + ", ".join(missing))

    confidence = round(sum(s.confidence for s in kept) / len(kept)) if kept else 0
    return v3.V3Result(base.description, base.workbook_path, confidence, kept, base.questions, warnings)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("description", nargs="*")
    p.add_argument("--workbook")
    p.add_argument("--json", action="store_true")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    description = " ".join(args.description).strip() or input("Works description: ").strip()
    if not description:
        print("No description entered.")
        return 1
    workbook_path = v3.choose_workbook(args.workbook)
    result = build_scope_v6(description, workbook_path)
    if args.json:
        print(json.dumps(asdict(result), indent=2))
    else:
        v3.print_result(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
