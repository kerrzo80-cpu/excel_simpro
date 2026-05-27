#!/usr/bin/env python3
"""
AI Scope Builder v4 - object locked.

V4 builds on V3 but adds strict object locking:
  - If the description is about a radiator, it does not show bath/toilet/tap tasks.
  - If the description is about a toilet, it does not show radiator/bath/tap tasks.
  - Generic pipework tasks are only allowed as supporting tasks.

Read-only by default. It does not edit the workbook.

Examples:
  python3 ai_scope_builder_v4.py "replace radiator with new double panel radiator"
  python3 ai_scope_builder_v4.py "toilet not filling"
  python3 ai_scope_builder_v4.py "remove basin and cap off supplies"
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import asdict
from pathlib import Path
from typing import Sequence

import ai_scope_builder_v3 as v3

OBJECT_GROUPS = {
    "radiator": {"radiator", "rad", "trv", "valve"},
    "toilet": {"toilet", "wc", "cistern", "flush", "fill", "syphon", "siphon", "pan", "seat"},
    "basin": {"basin", "sink", "whb", "vanity", "waste", "trap", "tap"},
    "bath": {"bath", "waste", "tap", "panel"},
    "shower": {"shower", "screen", "tray", "valve", "mixer"},
    "kitchen": {"kitchen", "sink", "tap", "mixer"},
    "outside tap": {"outside", "garden", "external", "tap"},
    "pipework": {"pipework", "pipe", "cap", "blank", "alter", "move", "supplies", "hot", "cold"},
}

PRIMARY_OBJECT_ORDER = ["outside tap", "radiator", "toilet", "basin", "shower", "bath", "kitchen", "pipework"]
SUPPORTING_OBJECTS = {"pipework"}


def has_phrase(text: str, phrase: str) -> bool:
    return re.search(rf"\b{re.escape(v3.norm(phrase))}\b", v3.norm(text)) is not None


def detect_primary_objects(description: str) -> set[str]:
    tokens = set(v3.tokenise(description))
    text = v3.norm(description)
    found: set[str] = set()

    if has_phrase(text, "outside tap") or has_phrase(text, "garden tap") or has_phrase(text, "external tap"):
        found.add("outside tap")

    # Single-word object matching with whole tokens only.
    if {"radiator"} & tokens:
        found.add("radiator")
    if {"toilet", "wc", "cistern", "syphon", "siphon"} & tokens:
        found.add("toilet")
    if {"basin", "sink", "whb", "vanity"} & tokens:
        found.add("basin")
    if {"shower"} & tokens:
        found.add("shower")
    if {"bath"} & tokens:
        found.add("bath")
    if {"kitchen"} & tokens or has_phrase(text, "kitchen tap") or has_phrase(text, "kitchen sink"):
        found.add("kitchen")
    if {"pipework", "pipe", "pipes", "cap", "blank", "supplies"} & tokens:
        found.add("pipework")

    # If a specific object exists, don't let pipework become the main object.
    specific = found - SUPPORTING_OBJECTS
    return specific or found


def task_object_group(task: v3.WorkbookTask) -> set[str]:
    combined = v3.norm(" ".join([task.code, task.category, task.name, task.notes]))
    tokens = set(v3.tokenise(combined))
    groups: set[str] = set()

    code = task.code.upper()
    name = v3.norm(task.name)
    category = v3.norm(task.category)

    if code.startswith("RAD") or "radiator" in tokens:
        groups.add("radiator")
    if code.startswith("WC") or {"toilet", "wc", "cistern", "flush", "fill", "syphon", "siphon"} & tokens:
        groups.add("toilet")
    if code.startswith("BASIN") or {"basin", "vanity"} & tokens or "basin" in name:
        groups.add("basin")
    if code.startswith("BATH") or "bath" in tokens:
        groups.add("bath")
    if code.startswith("SHWR") or "shower" in tokens:
        groups.add("shower")
    if "kitchen" in tokens or code.startswith("TAP-KITCHEN"):
        groups.add("kitchen")
    if code in {"PIPE-CAP", "PIPE-ALTER-M", "ADD-PIPE-REMOVE-CAP", "ADD-HC-PIPE-M"} or {"pipework", "pipe", "cap", "alter"} & tokens:
        groups.add("pipework")
    if code == "OUTSIDE-TAP" or has_phrase(name, "outside tap") or has_phrase(category, "outside tap"):
        groups.add("outside tap")

    return groups


def is_relevant_to_locked_objects(task: v3.WorkbookTask, locked: set[str]) -> bool:
    if not locked:
        return True
    groups = task_object_group(task)
    if not groups:
        return False
    if groups & locked:
        return True
    # Pipework is allowed as a supporting task for specific objects.
    if "pipework" in groups and (locked - {"pipework"}):
        return True
    return False


def supporting_allowed(task: v3.WorkbookTask, locked: set[str]) -> bool:
    groups = task_object_group(task)
    return "pipework" in groups and bool(locked - {"pipework"})


def score_workbook_task_locked(description: str, task: v3.WorkbookTask, locked: set[str]) -> int:
    if not is_relevant_to_locked_objects(task, locked):
        return 0

    score = v3.score_workbook_task(description, task)
    groups = task_object_group(task)

    if locked and groups & locked:
        score += 90
    elif locked and supporting_allowed(task, locked):
        score += 20
        # Supporting pipework should not outrank the main object task.
        score = min(score, 82)

    # Penalise replacement/repair generic matches that do not contain the main object.
    if locked and not (groups & locked) and not supporting_allowed(task, locked):
        score -= 120

    return max(0, score)


def merge_suggestions_locked(description: str, tasks: Sequence[v3.WorkbookTask]):
    locked = detect_primary_objects(description)
    workbook_codes = v3.task_by_code(tasks)
    rule_result = v3.rules.base.build_scope(description)

    suggestions: dict[str, v3.V3Suggestion] = {}
    warnings = list(rule_result.warnings)
    questions = list(rule_result.questions)

    for item in rule_result.suggested_tasks:
        code = item.code.upper()
        wb_task = workbook_codes.get(code)
        exists = wb_task is not None
        if exists and not is_relevant_to_locked_objects(wb_task, locked):
            continue
        suggestions[code] = v3.V3Suggestion(
            code=code,
            name=wb_task.name if wb_task else item.name,
            confidence=item.confidence,
            source=item.source,
            exists_in_workbook=exists,
            workbook_name=wb_task.name if wb_task else None,
            reason=item.reason,
        )

    scored = sorted(((score_workbook_task_locked(description, task, locked), task) for task in tasks), key=lambda x: x[0], reverse=True)
    for score, task in scored[:12]:
        if score < 70:
            continue
        code = task.code.upper()
        confidence = min(94, max(50, score))
        existing = suggestions.get(code)
        if existing and existing.confidence >= confidence:
            continue
        suggestions[code] = v3.V3Suggestion(
            code=code,
            name=task.name,
            confidence=confidence,
            source="workbook-match-object-locked",
            exists_in_workbook=True,
            workbook_name=task.name,
            reason="Matched against the real Task Library and passed object-lock filtering.",
        )

    ordered = sorted(suggestions.values(), key=lambda s: (s.exists_in_workbook, s.confidence), reverse=True)

    # Put missing exact object tasks near the top, but after existing workbook tasks.
    missing = [s.code for s in ordered if not s.exists_in_workbook]
    if missing:
        msg = "Suggested task codes not currently in workbook Task Library: " + ", ".join(missing)
        if msg not in warnings:
            warnings.append(msg)

    if locked:
        warnings.append("Object lock: " + ", ".join(sorted(locked)))

    if not ordered:
        confidence = 0
        warnings.append("No task suggestion found. Price this once, then save it as a learned rule.")
    else:
        confidence = round(sum(s.confidence for s in ordered[:5]) / min(5, len(ordered)))

    return ordered[:8], questions[:6], warnings, confidence


def build_scope_v4(description: str, workbook_path: Path) -> v3.V3Result:
    tasks = v3.read_task_library(workbook_path)
    suggestions, questions, warnings, confidence = merge_suggestions_locked(description, tasks)
    return v3.V3Result(
        description=description,
        workbook_path=str(workbook_path),
        confidence=confidence,
        suggestions=suggestions,
        questions=questions,
        warnings=warnings,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("description", nargs="*", help="Plain-English description of works")
    parser.add_argument("--workbook", help="Path to Works Pricing Tool .xlsm")
    parser.add_argument("--json", action="store_true", help="Output JSON")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    description = " ".join(args.description).strip()
    if not description:
        description = input("Works description: ").strip()
    if not description:
        print("No description entered.")
        return 1

    workbook_path = v3.choose_workbook(args.workbook)
    result = build_scope_v4(description, workbook_path)
    if args.json:
        print(json.dumps(asdict(result), indent=2))
    else:
        v3.print_result(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
