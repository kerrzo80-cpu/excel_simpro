#!/usr/bin/env python3
"""
Plain-English task search for the EWG Excel to simPRO pricing workbook.

Usage:
  python3 ~/Downloads/task_search.py
  python3 ~/Downloads/task_search.py "remove basin"

What it does:
  - Reads existing Task Library from the workbook.
  - Scores tasks using simple keywords + fuzzy matching + editable aliases.
  - Inserts the best selected task name into the next blank line on Works To Be Carried Out.

It does NOT rebuild or restyle the workbook.
"""

import csv
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

import xlwings as xw

WORKBOOK = str(Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm")
ALIASES_FILE = Path.home() / "Downloads" / "task_search_aliases.csv"

TASK_SHEET = "Task Library"
WORKS_SHEET = "Works To Be Carried Out"
FIRST_TASK_ROW = 6
FIRST_WORK_ROW = 12
LAST_WORK_ROW = 200
PLACEHOLDER = "Select Work Task"

DEFAULT_ALIASES = [
    ("rad", "radiator"),
    ("rads", "radiator"),
    ("trv", "radiator valve"),
    ("valves", "valve"),
    ("loo", "toilet"),
    ("wc", "toilet"),
    ("pan", "toilet"),
    ("cistern", "toilet flush valve"),
    ("syphon", "toilet flush valve"),
    ("siphon", "toilet flush valve"),
    ("sink", "basin"),
    ("wash hand basin", "basin"),
    ("whb", "basin"),
    ("vanity", "basin vanity"),
    ("tap", "taps"),
    ("mixer", "mixer tap"),
    ("outside tap", "outside tap"),
    ("shower screen", "shower screen"),
    ("mixer shower", "shower mixer valve"),
    ("shower valve", "shower mixer valve"),
    ("waste", "waste trap"),
    ("trap", "waste trap"),
    ("cap", "cap off pipework"),
    ("blank", "cap off pipework"),
    ("blank off", "cap off pipework"),
    ("disconnect", "remove cap off"),
    ("take out", "remove"),
    ("rip out", "remove"),
    ("strip out", "remove"),
    ("remove basin", "install basin vanity replace basin waste cap off pipework"),
    ("remove sink", "install basin vanity replace basin waste cap off pipework"),
    ("remove radiator", "remove radiator"),
    ("move rad", "move radiator"),
    ("swap radiator", "fit replace radiator"),
    ("replace rad", "fit replace radiator"),
    ("new radiator", "fit replace radiator"),
]

STOPWORDS = {
    "a", "an", "and", "at", "for", "from", "in", "into", "of", "on", "the", "to", "with",
    "customer", "wants", "needs", "need", "looking", "look", "please", "can", "we", "i", "am",
    "old", "new", "job", "work", "works", "fit", "fitting", "replace", "replacement", "install", "installation",
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


def ensure_alias_file():
    if ALIASES_FILE.exists():
        return
    with ALIASES_FILE.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["what_user_might_type", "search_words_to_add"])
        for row in DEFAULT_ALIASES:
            w.writerow(row)


def load_aliases():
    ensure_alias_file()
    aliases = list(DEFAULT_ALIASES)
    try:
        with ALIASES_FILE.open(newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                a = clean(row.get("what_user_might_type"))
                b = clean(row.get("search_words_to_add"))
                if a and b:
                    aliases.append((a, b))
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


def open_book(path):
    app = xw.App(visible=False)
    book = app.books.open(path)
    return app, book


def read_tasks(book):
    ws = book.sheets[TASK_SHEET]
    tasks = []
    row = FIRST_TASK_ROW
    while row < 1000:
        code = clean(ws.range(f"A{row}").value)
        category = clean(ws.range(f"B{row}").value)
        name = clean(ws.range(f"C{row}").value)
        driver = clean(ws.range(f"D{row}").value)
        unit = clean(ws.range(f"E{row}").value)
        hours = ws.range(f"F{row}").value
        notes = clean(ws.range(f"I{row}").value)
        if not code and not name:
            blank_count = 0
            for r in range(row, min(row + 10, 1000)):
                if not clean(ws.range(f"A{r}").value) and not clean(ws.range(f"C{r}").value):
                    blank_count += 1
            if blank_count >= 5:
                break
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
    q_tokens = set(tokens(q_expanded))
    h = task["haystack"]
    if not q_expanded:
        return 0

    score = 0
    name_norm = norm(task["name"])
    code_norm = norm(task["code"])

    if q_expanded == name_norm:
        score += 100
    if q_expanded in name_norm:
        score += 70
    if q_expanded in h:
        score += 45

    overlap = q_tokens & task["tokens"]
    score += len(overlap) * 14

    # Strong action weighting: remove should prefer removal tasks where available.
    if "remove" in q_tokens or "disconnect" in q_tokens:
        if "remove" in task["tokens"] or "cap" in task["tokens"]:
            score += 25
        if any(x in task["tokens"] for x in ["install", "fit"]):
            score -= 8

    if "radiator" in q_tokens and any(x in task["tokens"] for x in ["radiator", "rad"]):
        score += 18
    if "basin" in q_tokens and any(x in task["tokens"] for x in ["basin", "vanity", "waste"]):
        score += 18
    if "toilet" in q_tokens and any(x in task["tokens"] for x in ["toilet", "wc"]):
        score += 18

    score += int(SequenceMatcher(None, q_expanded, name_norm).ratio() * 35)
    score += int(SequenceMatcher(None, q_expanded, code_norm).ratio() * 10)
    return score


def best_matches(query, tasks, limit=8):
    ranked = []
    for task in tasks:
        s = score_task(query, task)
        if s > 0:
            ranked.append((s, task))
    ranked.sort(key=lambda x: x[0], reverse=True)
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
    # Leave formulas/validation to workbook where possible, but set basics for reliability.
    if not clean(ws.range(f"D{row}").value):
        ws.range(f"D{row}").value = task["code"]
    ws.range(f"E{row}").value = count
    if measurement is not None:
        ws.range(f"F{row}").value = measurement
    elif not clean(ws.range(f"F{row}").value):
        ws.range(f"F{row}").value = 1 if task["driver"].lower() in ["quantity", "per item", "per point", "per radiator"] else 0
    return row


def main():
    query = " ".join(sys.argv[1:]).strip()
    if not query:
        print("What work task are you looking for? Examples: remove basin, replace rad, toilet fill valve")
        query = input("Search: ").strip()
    if not query:
        print("No search entered.")
        return

    app, book = open_book(WORKBOOK)
    try:
        tasks = read_tasks(book)
        matches = best_matches(query, tasks)
        if not matches:
            print("No match found. Add keywords to ~/Downloads/task_search_aliases.csv and try again.")
            return

        print("")
        print(f"Search: {query}")
        print("Best matches:")
        for idx, (score, task) in enumerate(matches, 1):
            print(f"{idx}. {task['name']}  [{task['code']}]  score={score}")

        top_score, top_task = matches[0]
        if top_score < 45:
            print("")
            print("Confidence is low. I won’t auto-add this.")
            choice = input("Choose a number to add, or press Enter to cancel: ").strip()
            if not choice:
                return
            try:
                top_task = matches[int(choice) - 1][1]
            except Exception:
                print("Cancelled: invalid choice.")
                return
        else:
            print("")
            choice = input(f"Add best match '{top_task['name']}' to next blank work row? [Y/n or 1-8]: ").strip().lower()
            if choice in ["n", "no"]:
                print("Cancelled.")
                return
            if choice.isdigit():
                i = int(choice)
                if 1 <= i <= len(matches):
                    top_task = matches[i - 1][1]

        count = input("Count [default 1]: ").strip()
        try:
            count = float(count) if count else 1
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
        print("You can now reopen Excel and check the line/totals before sending to simPRO.")
    finally:
        book.close()
        app.quit()


if __name__ == "__main__":
    main()
