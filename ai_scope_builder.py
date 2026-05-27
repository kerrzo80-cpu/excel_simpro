#!/usr/bin/env python3
"""
AI Scope Builder v1 for the EWG Excel -> simPRO quoting workflow.

Purpose:
  Turn a plain-English job description into a suggested scope of works using
  the current pricing workbook task codes where possible.

This first version is rule-based and local. It does not call ChatGPT/OpenAI yet.
That is intentional: it gives us a safe, testable foundation before wiring it
into Excel or simPRO.

Examples:
  python3 ai_scope_builder.py "remove basin and cap off supplies"
  python3 ai_scope_builder.py "replace radiator in same position"
  python3 ai_scope_builder.py --learn "remove basin" BASIN-REMOVE PIPE-CAP BASIN-WASTE
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable, List, Sequence

BASE_DIR = Path(__file__).resolve().parent
RULES_FILE = BASE_DIR / "scope_rules.csv"
LEARNED_FILE = BASE_DIR / "learned_scope_rules.csv"


@dataclass(frozen=True)
class SuggestedTask:
    code: str
    name: str
    reason: str
    confidence: int
    source: str = "rule"


@dataclass(frozen=True)
class ScopeResult:
    description: str
    suggested_tasks: List[SuggestedTask]
    questions: List[str]
    warnings: List[str]
    confidence: int


def clean(value: object) -> str:
    return str(value or "").strip()


def norm(text: str) -> str:
    text = clean(text).lower()
    text = text.replace("/", " ").replace("-", " ")
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def words(text: str) -> set[str]:
    return set(norm(text).split())


def contains_any(text: str, phrases: Iterable[str]) -> bool:
    n = norm(text)
    return any(norm(phrase) in n for phrase in phrases if clean(phrase))


def ensure_rules_file() -> None:
    """Create a starter rule file if one does not exist."""
    if RULES_FILE.exists():
        return

    rows = [
        # intent, object, trigger_phrases, task_code, task_name, reason, confidence, questions
        ["remove", "basin", "remove basin|take out basin|rip out basin|strip out basin|remove sink|take out sink", "BASIN-REMOVE", "Remove basin", "Customer wants basin removed. This task may need adding to Task Library.", "95", "Is the basin being replaced or removed permanently?"],
        ["remove", "basin", "remove basin|take out basin|rip out basin|strip out basin|remove sink|take out sink|cap off basin", "PIPE-CAP", "Cap off pipework", "Hot/cold supplies usually need capping or making safe.", "80", "Are hot and cold supplies being capped or reused?"],
        ["remove", "basin", "remove basin|take out basin|rip out basin|strip out basin|remove sink|take out sink", "BASIN-WASTE", "Remove/alter basin waste or trap", "Waste/trap often needs disconnected or altered.", "65", "Is the waste being capped, altered, or reused?"],

        ["replace", "basin", "replace basin|new basin|change basin|swap basin|replace sink|new sink", "BASIN-REMOVE", "Remove existing basin", "Existing basin usually needs removed before replacement. This task may need adding to Task Library.", "75", "Is there an existing basin to remove?"],
        ["replace", "basin", "replace basin|new basin|change basin|swap basin|replace sink|new sink|vanity", "BASIN-VANITY", "Install basin / vanity unit", "Workbook contains basin/vanity install task.", "90", "Is the basin wall hung, pedestal, or vanity unit?"],
        ["replace", "basin", "replace basin|new basin|change basin|swap basin|replace sink|new sink", "BASIN-WASTE", "Replace basin waste / trap", "Replacement basin commonly needs waste/trap allowed for.", "80", "Is a new waste/trap included?"],
        ["replace", "basin", "replace basin taps|new basin taps|replace basin|new basin", "TAP-BASIN", "Replace basin taps / mono tap", "Taps may be part of basin replacement.", "55", "Are new taps required or reusing existing?"],

        ["remove", "radiator", "remove radiator|take off radiator|remove rad|take off rad", "ADD-RAD-REMOVE", "Remove radiator", "Workbook contains radiator removal task.", "95", "Is the radiator being removed permanently or refitted later?"],
        ["remove", "radiator", "remove radiator|take off radiator|remove rad|cap radiator", "PIPE-CAP", "Cap off pipework", "Pipework may need capping if radiator is removed permanently.", "65", "Are radiator pipes being capped below floor or left visible?"],

        ["move", "radiator", "move radiator|relocate radiator|shift radiator|move rad", "RAD-MOVE", "Move radiator", "Workbook contains radiator move task.", "95", "How far is the radiator moving?"],
        ["replace", "radiator", "replace radiator|new radiator|swap radiator|replace rad", "ADD-RAD-REMOVE", "Remove existing radiator", "Existing radiator usually needs removed first.", "80", "Is the old radiator being disposed of?"],
        ["replace", "radiator", "replace radiator|new radiator|swap radiator|replace rad", "RAD-FIT", "Fit radiator", "Workbook contains radiator fitting task.", "90", "Is the radiator supplied by EWG or client?"],
        ["replace", "radiator", "replace radiator|new radiator|swap radiator|replace rad|trv|valves", "RAD-VALVES", "Fit/replace radiator valves", "Valves/TRVs are commonly required with radiator works.", "65", "Are TRVs required?"],

        ["replace", "toilet", "toilet fill valve|toilet not filling|cistern filling|ball valve", "WC-FILL", "Replace toilet fill valve", "Workbook contains toilet fill valve task.", "95", "Is the cistern accessible and standard type?"],
        ["replace", "toilet", "toilet flush valve|syphon|siphon|toilet not flushing", "WC-FLUSH", "Replace toilet flush valve / syphon", "Workbook contains toilet flush repair task.", "95", "Is it a concealed cistern?"],
        ["install", "toilet", "install toilet|fit toilet|new toilet|replace toilet", "WC-INSTALL", "Install toilet", "Workbook contains toilet install task.", "90", "Close-coupled, back-to-wall, or concealed cistern?"],
        ["remove", "toilet", "remove toilet|take out toilet|rip out toilet|remove wc", "WC-REMOVE", "Remove toilet", "This task may need adding to Task Library.", "90", "Is the toilet being replaced or removed permanently?"],
        ["unblock", "toilet", "blocked toilet|toilet blocked|unblock toilet|wc blocked", "WC-UNBLOCK", "Unblock toilet", "Workbook contains toilet unblock task.", "95", "Is the blockage local to the pan or affecting other drains?"],

        ["install", "outside tap", "outside tap|garden tap|external tap", "OUTSIDE-TAP", "Install outside tap", "Workbook contains outside tap task.", "95", "Is there accessible cold water nearby?"],
        ["cap", "pipework", "cap off|blank off|make safe|disconnect pipework", "PIPE-CAP", "Cap off pipework", "Workbook contains cap-off pipework task.", "90", "What pipework is being capped and where?"],
        ["alter", "pipework", "alter pipework|reroute pipework|move pipework|extend pipework", "PIPE-ALTER-M", "Alter pipework", "Workbook contains pipework alteration task.", "85", "Approximate metres and pipe size?"],
    ]

    with RULES_FILE.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["intent", "object", "trigger_phrases", "task_code", "task_name", "reason", "confidence", "questions"])
        writer.writerows(rows)


def load_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="") as f:
        return list(csv.DictReader(f))


def infer_intent(description: str) -> str:
    d = norm(description)
    if contains_any(d, ["remove", "take out", "rip out", "strip out", "disconnect"]):
        return "remove"
    if contains_any(d, ["move", "relocate", "shift"]):
        return "move"
    if contains_any(d, ["replace", "swap", "change", "renew", "new"]):
        return "replace"
    if contains_any(d, ["install", "fit", "supply and fit"]):
        return "install"
    if contains_any(d, ["cap", "blank", "make safe"]):
        return "cap"
    if contains_any(d, ["unblock", "blocked", "blockage"]):
        return "unblock"
    return "unknown"


def infer_objects(description: str) -> set[str]:
    d_words = words(description)
    d = norm(description)
    objects = set()
    aliases = {
        "basin": ["basin", "sink", "whb", "wash hand basin", "vanity"],
        "radiator": ["radiator", "rad", "rads", "trv"],
        "toilet": ["toilet", "wc", "loo", "pan", "cistern", "syphon", "siphon"],
        "outside tap": ["outside tap", "garden tap", "external tap"],
        "pipework": ["pipe", "pipes", "pipework", "cap", "blank", "supplies", "hot", "cold"],
        "shower": ["shower"],
        "bath": ["bath"],
        "tap": ["tap", "taps", "mixer"],
        "waste": ["waste", "trap"],
    }
    for obj, terms in aliases.items():
        if any(term in d_words or norm(term) in d for term in terms):
            objects.add(obj)
    return objects


def rule_matches(description: str, row: dict[str, str], inferred_intent: str, inferred_objects: set[str]) -> bool:
    trigger_phrases = [p.strip() for p in clean(row.get("trigger_phrases")).split("|") if p.strip()]
    if contains_any(description, trigger_phrases):
        return True

    row_intent = norm(row.get("intent"))
    row_object = norm(row.get("object"))
    if inferred_intent and row_intent == inferred_intent and row_object in inferred_objects:
        return True
    return False


def build_scope(description: str) -> ScopeResult:
    ensure_rules_file()
    rows = load_rows(RULES_FILE) + load_rows(LEARNED_FILE)
    inferred_intent = infer_intent(description)
    inferred_objects = infer_objects(description)

    tasks_by_code: dict[str, SuggestedTask] = {}
    questions: list[str] = []
    warnings: list[str] = []

    for row in rows:
        if not rule_matches(description, row, inferred_intent, inferred_objects):
            continue

        code = clean(row.get("task_code"))
        if not code:
            continue

        try:
            confidence = int(float(clean(row.get("confidence")) or 50))
        except ValueError:
            confidence = 50

        task = SuggestedTask(
            code=code,
            name=clean(row.get("task_name")) or code,
            reason=clean(row.get("reason")),
            confidence=max(1, min(100, confidence)),
            source="learned" if row in load_rows(LEARNED_FILE) else "rule",
        )

        existing = tasks_by_code.get(code)
        if existing is None or task.confidence > existing.confidence:
            tasks_by_code[code] = task

        for q in clean(row.get("questions")).split("|"):
            q = clean(q)
            if q and q not in questions:
                questions.append(q)

    suggested_tasks = sorted(tasks_by_code.values(), key=lambda t: t.confidence, reverse=True)

    missing_codes = [t.code for t in suggested_tasks if t.code.endswith("-REMOVE") and t.code not in {"ADD-RAD-REMOVE"}]
    if missing_codes:
        warnings.append(
            "Some suggested removal tasks may need adding to the Excel Task Library first: " + ", ".join(missing_codes)
        )

    if not suggested_tasks:
        warnings.append("No confident scope found yet. Add a learned rule after pricing this once.")
        confidence = 0
    else:
        confidence = round(sum(t.confidence for t in suggested_tasks) / len(suggested_tasks))
        if inferred_intent == "unknown":
            questions.append("Is this a remove, replace, move, install, repair, or cap-off job?")
            confidence = min(confidence, 60)
        if not inferred_objects:
            questions.append("What item is being worked on? For example basin, radiator, toilet, tap, pipework.")
            confidence = min(confidence, 55)

    return ScopeResult(
        description=description,
        suggested_tasks=suggested_tasks,
        questions=questions[:6],
        warnings=warnings,
        confidence=confidence,
    )


def save_learned_rule(description: str, task_codes: Sequence[str]) -> None:
    ensure_rules_file()
    existing_header = LEARNED_FILE.exists()
    inferred_intent = infer_intent(description)
    inferred_objects = infer_objects(description)
    obj = sorted(inferred_objects)[0] if inferred_objects else "unknown"

    with LEARNED_FILE.open("a", newline="") as f:
        writer = csv.writer(f)
        if not existing_header:
            writer.writerow(["intent", "object", "trigger_phrases", "task_code", "task_name", "reason", "confidence", "questions"])
        for code in task_codes:
            code = clean(code).upper()
            if not code:
                continue
            writer.writerow([
                inferred_intent,
                obj,
                description,
                code,
                code,
                "Learned from estimator correction/approval.",
                "95",
                "",
            ])


def print_human(result: ScopeResult) -> None:
    print("")
    print(f"Description: {result.description}")
    print(f"Overall confidence: {result.confidence}%")
    print("")

    if result.suggested_tasks:
        print("Suggested scope:")
        for idx, task in enumerate(result.suggested_tasks, 1):
            print(f"{idx}. {task.code} - {task.name} ({task.confidence}%)")
            if task.reason:
                print(f"   Reason: {task.reason}")
    else:
        print("Suggested scope: none yet")

    if result.questions:
        print("")
        print("Questions to confirm:")
        for q in result.questions:
            print(f"- {q}")

    if result.warnings:
        print("")
        print("Warnings:")
        for w in result.warnings:
            print(f"- {w}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("description", nargs="*", help="Plain-English description of works")
    parser.add_argument("--json", action="store_true", help="Output JSON")
    parser.add_argument("--learn", action="store_true", help="Save a learned mapping: --learn 'description' TASK1 TASK2")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.learn:
        if len(args.description) < 2:
            print("Usage: python3 ai_scope_builder.py --learn 'remove basin' BASIN-REMOVE PIPE-CAP")
            return 2
        desc = args.description[0]
        codes = args.description[1:]
        save_learned_rule(desc, codes)
        print(f"Learned rule saved for: {desc}")
        print("Task codes:", ", ".join(codes))
        return 0

    description = " ".join(args.description).strip()
    if not description:
        print("Describe the works, for example: remove basin and cap off supplies")
        description = input("Works description: ").strip()
    if not description:
        print("No description entered.")
        return 1

    result = build_scope(description)
    if args.json:
        print(json.dumps(asdict(result), indent=2))
    else:
        print_human(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
