#!/usr/bin/env python3
"""
Safer plain-English task search for the EWG Excel to simPRO pricing workbook.

This version is designed to stop weak fuzzy matches like "remove basin" selecting
unrelated items such as "move radiator".

Usage:
  python3 task_search_fixed.py
  python3 task_search_fixed.py "remove basin"
  python3 task_search_fixed.py --workbook "/path/to/Works Pricing Tool.xlsm" "remove basin"
"""

import argparse
import csv
import re
from difflib import SequenceMatcher
from pathlib import Path

import xlwings as xw

DEFAULT_WORKBOOKS = [
    Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm",
    Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Pricing/Works Pricing Tool.xlsm",
    Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V2.xlsm",
]

ALIASES_FILE = Path.home() / "Downloads" / "task_search_aliases.csv"
TASK_SHEET = "Task Library"
WORKS_SHEET = "Works To Be Carried Out"
FIRST_TASK_ROW = 6
FIRST_WORK_ROW = 12
LAST_WORK_ROW = 200
PLACEHOLDER = "Select Work Task"

# Keep these simple. Do not expand exact removal phrases into install/replace words.
DEFAULT_ALIASES = [
    ("rad", "radiator"),
    ("rads", "radiator"),
    ("move rad", "move radiator"),
    ("replace rad", "replace radiator"),
    ("new radiator", "replace radiator"),
    ("trv", "radiator valve"),
    ("loo", "toilet"),
    ("wc", "toilet"),
    ("pan", "toilet"),
    ("cistern", "toilet cistern"),
    ("syphon", "toilet flush valve"),
    ("siphon", "toilet flush valve"),
    ("sink", "basin"),
    ("wash hand basin", "basin"),
    ("whb", "basin"),
    ("vanity", "basin vanity"),
    ("take out", "remove"),
    ("rip out", "remove"),
    ("strip out", "remove"),
    ("disconnect", "remove cap off"),
    ("blank off", "cap off pipework"),
    ("blank", "cap off pipework"),
    ("cap", "cap off pipework"),
    ("remove basin", "remove basin cap off pipework"),
    ("remove sink", "remove basin cap off pipework"),
    ("remove radiator", "remove radiator cap off pipework"),
    ("toilet not filling", "toilet fill valve"),
    ("ball valve", "toilet fill valve"),
]

BAD_ALIAS_WORDS = {"install", "replace", "vanity", "waste"}
ACTION_WORDS = {"remove", "move", "replace", "install", "fit", "cap", "disconnect"}
OBJECT_WORDS = {
    "basin", "sink", "radiator", "toilet", "wc", "tap", "shower", "bath", "pipe", "pipework",
    "valve", "cistern", "waste", "trap", "door", "window", "wall", "ceiling", "floor",
}

STOPWORDS = {
    "a", "an", "and", "at", "for", "from", "in", "into", "of", "on", "the", "to", "with",
    "customer", "wants", "needs", "need", "looking", "look", "please", "can", "we", "i", "am",
    "old", "new", "job", "work", "works", "would", "like", "quote", "price",
}


def clean(value):
    return str(value or "").strip()


def norm(text):
    text = clean(text).lower()
    text = text.replace("/", " ").replace("-", " ")
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokens(text):
    return [t for t in norm(text).split() if t and t not in STOPWORDS]


def choose_workbook(cli_path=None):
    if cli_path:
        path = Path(cli_path).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"Workbook not found: {path}")
        return str(path)

    for path in DEFAULT_WORKBOOKS:
        if path.exists():
            return str(path)

    raise FileNotFoundError("Could not find a Works Pricing Tool workbook. Use --workbook /path/to/file.xlsm")


def ensure_alias_file():
    if ALIASES_FILE.exists():
        return
    with ALIASES_FILE.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["what_user_might_type", "search_words_to_add"])
        writer.writerows(DEFAULT_ALIASES)


def load_aliases():
    ensure_alias_file()
    aliases = list(DEFAULT_ALIASES)
    try:
        with ALIASES_FILE.open(newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                phrase = clean(row.get("what_user_might_type"))
                extra = clean(row.get("search_words_to_add"))
                if not phrase or not extra:
                    continue

                # The old file may already contain the bad alias:
                # remove basin -> install basin vanity replace basin waste...
                # Ignore that because it causes the wrong work type to win.
                if norm(phrase) in {"remove basin", "remove sink"} and BAD_ALIAS_WORDS & set(tokens(extra)):
                    continue

                aliases.append((phrase, extra))
    except Exception:
        pass
    return aliases


def expand_query(query):
    q = norm(query)
    additions = []
    for phrase, extra in load_aliases():
        if norm(phrase) in q:
            additions.append(extra)
    return norm(q + " " + " ".join(additions))


def query_parts(query):
    q_tokens = set(tokens(expand_query(query)))
    actions = q_tokens & ACTION_WORDS
    objects = q_tokens & OBJECT_WORDS
    return q_tokens, actions, objects


def open_book(path):
    app = xw.App(visible=False)
    book = app.books.open(path)
    return app, book


def read_tasks(book):
    ws = book.sheets[TASK_SHEET]
    tasks = []
    row = FIRST_TASK_ROW
    blank_streak = 0

    while row < 1000:
        code = clean(ws.range(f"A{row}").value)
        category = clean(ws.range(f"B{row}").value)
        name = clean(ws.range(f"C{row}").value)
        driver = clean(ws.range(f"D{row}").value)
        unit = clean(ws.range(f"E{row}").value)
        hours = ws.range(f"F{row}").value
        notes = clean(ws.range(f"I{row}").value)

        if not code and not name:
            blank_streak += 1
            if blank_streak >= 5:
                break
            row += 1
            continue
        blank_streak = 0

        if code and name:
            haystack = norm(" ".join([code, category, name, driver, unit, notes]))
            tasks.append({
                "row": row,
                "code": code,
                "category": category,
                "name": name,
                "driver": driver,
                "unit": unit,
                "hours": hours,
                "notes": notes,
                "haystack": haystack,
                "tokens": set(tokens(haystack)),
            })
        row += 1
    return tasks


def score_task(query, task):
    q_expanded = expand_query(query)
    q_tokens, q_actions, q_objects = query_parts(query)
    h_tokens = task["tokens"]
    h = task["haystack"]
    name_norm = norm(task["name"])
    code_norm = norm(task["code"])

    if not q_tokens:
        return 0

    score = 0

    # Exact and phrase matches are very strong.
    if q_expanded == name_norm:
        score += 160
    if norm(query) in name_norm:
        score += 110
    if q_expanded in h:
        score += 80

    overlap = q_tokens & h_tokens
    score += len(overlap) * 20

    # Object gating: if the customer says basin, radiator tasks should be pushed down hard.
    if q_objects:
        object_overlap = q_objects & h_tokens
        if object_overlap:
            score += 60 + (len(object_overlap) * 20)
        else:
            score -= 120

    # Action matching: remove/move/replace are very different.
    if q_actions:
        action_overlap = q_actions & h_tokens
        if action_overlap:
            score += 55 + (len(action_overlap) * 15)
        else:
            score -= 35

    if "remove" in q_actions or "disconnect" in q_actions:
        if {"remove", "cap", "disconnect"} & h_tokens:
            score += 55
        if {"move", "relocate"} & h_tokens:
            score -= 90
        if {"install", "fit"} & h_tokens and "remove" not in h_tokens:
            score -= 50

    if "move" in q_actions:
        if {"move", "relocate"} & h_tokens:
            score += 60
        if "remove" in h_tokens:
            score -= 70

    if "replace" in q_actions:
        if {"replace", "renew", "swap"} & h_tokens:
            score += 60
        if "remove" in h_tokens and not ({"replace", "renew", "swap"} & h_tokens):
            score -= 45

    # Fuzzy matching is now a small tie-breaker only, not enough to beat wrong objects/actions.
    score += int(SequenceMatcher(None, q_expanded, name_norm).ratio() * 15)
    score += int(SequenceMatcher(None, q_expanded, code_norm).ratio() * 5)
    return score


def best_matches(query, tasks, limit=8):
    ranked = []
    for task in tasks:
        score = score_task(query, task)
        if score > 0:
            ranked.append((score, task))
    ranked.sort(key=lambda item: item[0], reverse=True)
    return ranked[:limit]


def next_blank_work_row(book):
    ws = book.sheets[WORKS_SHEET]
    for row in range(FIRST_WORK_ROW, LAST_WORK_ROW + 1):
        current = clean(ws.range(f"C{row}").value)
        if not current or PLACEHOLDER.lower() in current.lower():
            return row
    raise RuntimeError("No blank work rows found between rows 12 and 200.")


def insert_task(book, task, count=1, measurement=None, room="Plumbing "):
    ws = book.sheets[WORKS_SHEET]
    row = next_blank_work_row(book)
    ws.range(f"B{row}").value = room
    ws.range(f"C{row}").value = task["name"]
    if not clean(ws.range(f"D{row}").value):
        ws.range(f"D{row}").value = task["code"]
    ws.range(f"E{row}").value = count

    if measurement is not None:
        ws.range(f"F{row}").value = measurement
    elif not clean(ws.range(f"F{row}").value):
        driver = task["driver"].lower()
        ws.range(f"F{row}").value = 1 if driver in ["quantity", "per item", "per point", "per radiator"] else 0
    return row


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("query", nargs="*", help="Plain-English work task, e.g. remove basin")
    parser.add_argument("--workbook", help="Path to Works Pricing Tool .xlsm")
    parser.add_argument("--list-only", action="store_true", help="Show matches but do not add to workbook")
    return parser.parse_args()


def main():
    args = parse_args()
    query = " ".join(args.query).strip()
    if not query:
        print("What work task are you looking for? Examples: remove basin, replace rad, toilet fill valve")
        query = input("Search: ").strip()
    if not query:
        print("No search entered.")
        return

    workbook_path = choose_workbook(args.workbook)
    app, book = open_book(workbook_path)
    try:
        tasks = read_tasks(book)
        matches = best_matches(query, tasks)
        if not matches:
            print("No good match found. Add keywords to ~/Downloads/task_search_aliases.csv and try again.")
            return

        print("")
        print(f"Workbook: {workbook_path}")
        print(f"Search: {query}")
        print("Best matches:")
        for idx, (score, task) in enumerate(matches, 1):
            print(f"{idx}. {task['name']}  [{task['code']}]  score={score}")

        if args.list_only:
            return

        top_score, top_task = matches[0]
        second_score = matches[1][0] if len(matches) > 1 else 0
        confident = top_score >= 95 and (top_score - second_score) >= 25

        print("")
        if confident:
            choice = input(f"Add best match '{top_task['name']}' to next blank work row? [Y/n or 1-8]: ").strip().lower()
        else:
            print("Confidence is not high enough to auto-pick safely.")
            choice = input("Choose a number to add, or press Enter to cancel: ").strip().lower()

        if not choice and not confident:
            print("Cancelled.")
            return
        if choice in ["n", "no"]:
            print("Cancelled.")
            return
        if choice.isdigit():
            index = int(choice)
            if not (1 <= index <= len(matches)):
                print("Cancelled: invalid choice.")
                return
            top_task = matches[index - 1][1]

        count_text = input("Count [default 1]: ").strip()
        try:
            count = float(count_text) if count_text else 1
        except ValueError:
            count = 1

        meas = input("Measurement/Qty [press Enter for workbook default]: ").strip()
        measurement = None
        if meas:
            try:
                measurement = float(meas)
            except ValueError:
                measurement = None

        row = insert_task(book, top_task, count=count, measurement=measurement)
        book.save()
        print("")
        print(f"Added to row {row}: {top_task['name']} [{top_task['code']}]")
        print("Workbook saved.")
    finally:
        book.close()
        app.quit()


if __name__ == "__main__":
    main()
