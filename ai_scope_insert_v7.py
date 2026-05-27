#!/usr/bin/env python3
"""AI Scope Insert V7.

Uses ai_scope_builder_v6.py to build a scope from a plain-English description.
By default it is read-only. Use --insert to write existing workbook task codes
into the Works To Be Carried Out sheet.

Examples:
  python3 ai_scope_insert_v7.py "replace radiator with new double panel radiator"
  python3 ai_scope_insert_v7.py --insert "replace radiator with new double panel radiator"
  python3 ai_scope_insert_v7.py --insert --yes "toilet not filling"
"""

from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

import ai_scope_builder_v3 as v3
import ai_scope_builder_v6 as v6

WORKS_SHEET = "Works To Be Carried Out"
FIRST_WORK_ROW = 12
LAST_WORK_ROW = 200
PLACEHOLDER = "Select Work Task"
DEFAULT_ROOM = "Plumbing "


def clean(value: object) -> str:
    return str(value or "").strip()


def open_book(path: Path):
    app = xw.App(visible=False)
    book = app.books.open(str(path))
    return app, book


def safe_close(book, app):
    try:
        if book is not None:
            book.close()
    except Exception as exc:
        print(f"Note: workbook close warning ignored: {exc}")
    try:
        if app is not None:
            app.quit()
    except Exception as exc:
        print(f"Note: Excel quit warning ignored: {exc}")


def next_blank_work_row(book) -> int:
    ws = book.sheets[WORKS_SHEET]
    for row in range(FIRST_WORK_ROW, LAST_WORK_ROW + 1):
        current = clean(ws.range(f"C{row}").value)
        if not current or PLACEHOLDER.lower() in current.lower():
            return row
    raise RuntimeError(f"No blank work rows found between rows {FIRST_WORK_ROW} and {LAST_WORK_ROW}.")


def existing_insertable_suggestions(result):
    return [s for s in result.suggestions if s.exists_in_workbook]


def insert_tasks(workbook_path: Path, suggestions, room: str = DEFAULT_ROOM):
    app = None
    book = None
    inserted = []
    try:
        app, book = open_book(workbook_path)
        ws = book.sheets[WORKS_SHEET]
        for item in suggestions:
            row = next_blank_work_row(book)
            ws.range(f"B{row}").value = room
            ws.range(f"C{row}").value = item.name
            if not clean(ws.range(f"D{row}").value):
                ws.range(f"D{row}").value = item.code
            if not clean(ws.range(f"E{row}").value):
                ws.range(f"E{row}").value = 1
            inserted.append((row, item.code, item.name))
        book.save()
        return inserted
    finally:
        safe_close(book, app)


def print_scope(result):
    v3.print_result(result)
    insertable = existing_insertable_suggestions(result)
    skipped = [s for s in result.suggestions if not s.exists_in_workbook]

    print("")
    print("Insertable existing workbook tasks:")
    if not insertable:
        print("- None")
    else:
        for item in insertable:
            print(f"- {item.code} - {item.name}")

    if skipped:
        print("")
        print("Skipped missing tasks:")
        for item in skipped:
            print(f"- {item.code} - {item.name}")


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("description", nargs="*", help="Plain-English work description")
    parser.add_argument("--workbook", help="Path to Works Pricing Tool .xlsm")
    parser.add_argument("--insert", action="store_true", help="Insert existing workbook tasks into Works To Be Carried Out")
    parser.add_argument("--yes", action="store_true", help="Do not ask for confirmation before inserting")
    parser.add_argument("--room", default=DEFAULT_ROOM, help="Room/section value to put in column B")
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
    result = v6.build_scope_v6(description, workbook_path)
    print_scope(result)

    if not args.insert:
        print("")
        print("Read-only mode. Add --insert to write these tasks into the workbook.")
        return 0

    insertable = existing_insertable_suggestions(result)
    if not insertable:
        print("No existing workbook tasks to insert.")
        return 1

    if not args.yes:
        print("")
        answer = input(f"Insert {len(insertable)} task(s) into the workbook? [y/N]: ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled.")
            return 0

    inserted = insert_tasks(workbook_path, insertable, room=args.room)
    print("")
    print("Inserted tasks:")
    for row, code, name in inserted:
        print(f"- Row {row}: {code} - {name}")
    print("Workbook saved.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
