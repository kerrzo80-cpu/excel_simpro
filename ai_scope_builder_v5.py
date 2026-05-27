#!/usr/bin/env python3
"""
AI Scope Builder v5 - object + intent locked.

V5 builds on V4 and adds trade-intent filtering:
  - "toilet not filling" only shows likely repair/fill/flush tasks.
  - "remove basin" favours removal/cap-off/waste tasks, not install/vanity tasks.
  - "move radiator" favours move/alter tasks, not every radiator task.

Read-only by default. It does not edit the workbook.
"""

from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from pathlib import Path
from typing import Sequence

import ai_scope_builder_v4 as v4
import ai_scope_builder_v3 as v3

INTENT_ALLOWED = {
    "remove": {"remove", "cap", "blank", "disconnect", "waste", "trap", "make", "safe"},
    "move": {"move", "relocate", "alter", "pipework", "pipe", "cap", "remove", "fit", "refit"},
    "replace": {"replace", "renew", "swap", "change", "remove", "fit", "install", "valve", "waste", "trap"},
    "install": {"install", "fit", "supply", "run", "pipework", "pipe"},
    "repair": {"repair", "replace", "fill", "flush", "valve", "syphon", "siphon", "leak", "test", "unblock"},
    "unblock": {"unblock", "blocked", "blockage", "clear"},
    "cap": {"cap", "blank", "disconnect", "make", "safe", "remove"},
}

INTENT_FORBIDDEN = {
    "remove": {"install", "fit", "vanity", "pedestal", "supply"},
    "move": {"seat", "flush", "fill", "unblock"},
    "replace": {"unblock"},
    "install": {"unblock", "seat"},
    "repair": {"seat", "install", "fit", "wall", "hung", "close", "coupled", "back", "waste", "pipe", "pipework", "unblock"},
    "unblock": {"install", "fit", "seat", "waste", "pipework"},
    "cap": {"install", "fit", "seat", "vanity"},
}

# Specific job phrases override broad intent detection.
REPAIR_PHRASES = ["not filling", "not flushing", "leaking", "leak", "dripping", "not working", "faulty"]


def infer_trade_intent(description: str) -> str:
    text = v3.norm(description)
    if any(p in text for p in REPAIR_PHRASES):
        return "repair"
    return v3.infer_intent(description)


def task_intent_tokens(task: v3.WorkbookTask) -> set[str]:
    return set(v3.tokenise(" ".join([task.code, task.category, task.name, task.notes, task.driver_type, task.driver_unit])))


def is_intent_relevant(description: str, task: v3.WorkbookTask) -> bool:
    intent = infer_trade_intent(description)
    if intent == "unknown":
        return True

    tokens = task_intent_tokens(task)
    code = task.code.upper()
    name = v3.norm(task.name)

    # Strong exact allowances for known codes.
    if intent == "repair":
        if code in {"WC-FILL", "WC-FLUSH"}:
            return True
        if "fill valve" in name or "flush valve" in name or "siphon" in name or "syphon" in name:
            return True
    if intent == "move":
        if code in {"RAD-MOVE", "PIPE-ALTER-M"}:
            return True
        if "move" in name or "alter pipework" in name:
            return True
    if intent == "remove":
        if code in {"PIPE-CAP", "ADD-PIPE-REMOVE-CAP", "ADD-RAD-REMOVE"}:
            return True
        if "remove" in name or "cap off" in name or "waste" in name or "trap" in name:
            return True
    if intent == "replace":
        if any(word in name for word in ["replace", "fit", "install", "remove"]):
            return True

    allowed = INTENT_ALLOWED.get(intent, set())
    forbidden = INTENT_FORBIDDEN.get(intent, set())

    if forbidden & tokens:
        # Permit if it also has a very strong allowed token for the same intent.
        if not (allowed & tokens):
            return False

    return bool(allowed & tokens)


def score_task_v5(description: str, task: v3.WorkbookTask, locked: set[str]) -> int:
    if not v4.is_relevant_to_locked_objects(task, locked):
        return 0
    if not is_intent_relevant(description, task):
        return 0

    score = v4.score_workbook_task_locked(description, task, locked)
    intent = infer_trade_intent(description)
    tokens = task_intent_tokens(task)
    name = v3.norm(task.name)
    code = task.code.upper()

    allowed = INTENT_ALLOWED.get(intent, set())
    forbidden = INTENT_FORBIDDEN.get(intent, set())

    if allowed & tokens:
        score += 55
    if forbidden & tokens:
        score -= 60

    # Intent-specific boosts/penalties.
    if intent == "repair":
        if code == "WC-FILL" or "fill valve" in name:
            score += 90
        if code == "WC-FLUSH" or "flush valve" in name or "siphon" in name:
            score += 45
        if any(x in name for x in ["seat", "wall hung", "close coupled", "back to wall", "waste pipe", "unblock"]):
            score -= 180

    if intent == "remove":
        if "remove" in name:
            score += 80
        if "cap off" in name:
            score += 60
        if any(x in name for x in ["install", "fit basin", "vanity", "pedestal", "taps"]):
            score -= 170

    if intent == "move":
        if "move" in name:
            score += 95
        if "alter pipework" in name or code == "PIPE-ALTER-M":
            score += 55
        if any(x in name for x in ["replace", "valve", "balance", "remove"]):
            score -= 80

    if intent == "replace":
        if "replace" in name or "fit" in name:
            score += 55
        if "unblock" in name:
            score -= 100

    return max(0, score)


def merge_suggestions_v5(description: str, tasks: Sequence[v3.WorkbookTask]):
    locked = v4.detect_primary_objects(description)
    intent = infer_trade_intent(description)
    workbook_codes = v3.task_by_code(tasks)
    rule_result = v3.rules.base.build_scope(description)

    suggestions: dict[str, v3.V3Suggestion] = {}
    warnings = list(rule_result.warnings)
    questions = list(rule_result.questions)

    # Keep rule suggestions only if they pass object and intent checks when present in workbook.
    for item in rule_result.suggested_tasks:
        code = item.code.upper()
        wb_task = workbook_codes.get(code)
        exists = wb_task is not None
        if exists:
            if not v4.is_relevant_to_locked_objects(wb_task, locked):
                continue
            if not is_intent_relevant(description, wb_task):
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

    scored = sorted(((score_task_v5(description, task, locked), task) for task in tasks), key=lambda x: x[0], reverse=True)
    for score, task in scored[:12]:
        if score < 95:
            continue
        code = task.code.upper()
        confidence = min(96, max(50, score))
        existing = suggestions.get(code)
        if existing and existing.confidence >= confidence:
            continue
        suggestions[code] = v3.V3Suggestion(
            code=code,
            name=task.name,
            confidence=confidence,
            source="workbook-match-object-intent-locked",
            exists_in_workbook=True,
            workbook_name=task.name,
            reason="Matched real Task Library and passed object + intent filtering.",
        )

    ordered = sorted(suggestions.values(), key=lambda s: (s.exists_in_workbook, s.confidence), reverse=True)

    missing = [s.code for s in ordered if not s.exists_in_workbook]
    if missing:
        msg = "Suggested task codes not currently in workbook Task Library: " + ", ".join(missing)
        if msg not in warnings:
            warnings.append(msg)

    if locked:
        warnings.append("Object lock: " + ", ".join(sorted(locked)))
    warnings.append("Intent lock: " + intent)

    if not ordered:
        confidence = 0
        warnings.append("No task suggestion found. Price this once, then save it as a learned rule.")
    else:
        confidence = round(sum(s.confidence for s in ordered[:5]) / min(5, len(ordered)))

    return ordered[:8], questions[:6], warnings, confidence


def build_scope_v5(description: str, workbook_path: Path) -> v3.V3Result:
    tasks = v3.read_task_library(workbook_path)
    suggestions, questions, warnings, confidence = merge_suggestions_v5(description, tasks)
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
    result = build_scope_v5(description, workbook_path)
    if args.json:
        print(json.dumps(asdict(result), indent=2))
    else:
        v3.print_result(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
