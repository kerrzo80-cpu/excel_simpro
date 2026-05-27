#!/usr/bin/env python3
"""AI Scope Insert V8.

Safe upgrade from V7:
- Uses V6 scope builder.
- Writes a customer-facing Description of Works to Final Quote B6.
- Skips duplicate task codes already present on the quote.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

import ai_scope_builder_v3 as v3
import ai_scope_builder_v6 as v6

WORKS_SHEET = "Works To Be Carried Out"
FINAL_SHEET = "Final Quote"
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


def get_sheet(book, wanted_name: str):
    for sheet in book.sheets:
        if sheet.name.strip() == wanted_name.strip():
            return sheet
    raise RuntimeError(f"Could not find sheet: {wanted_name}")


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


def existing_task_codes(book) -> set[str]:
    ws = get_sheet(book, WORKS_SHEET)
    codes = set()
    for row in range(FIRST_WORK_ROW, LAST_WORK_ROW + 1):
        code = clean(ws.range(f"D{row}").value).upper()
        if code:
            codes.add(code)
    return codes


def next_blank_work_row(book) -> int:
    ws = get_sheet(book, WORKS_SHEET)
    for row in range(FIRST_WORK_ROW, LAST_WORK_ROW + 1):
        current = clean(ws.range(f"C{row}").value)
        if not current or PLACEHOLDER.lower() in current.lower():
            return row
    raise RuntimeError("No blank work rows found.")


def insertable(result):
    return [s for s in result.suggestions if s.exists_in_workbook]


def build_description(original_description: str, suggestions) -> str:
    codes = {s.code.upper() for s in suggestions}
    lines = []

    if original_description.strip():
        text = original_description.strip()
        lines.append(text[:1].upper() + text[1:].rstrip(".") + ".")

    mapping = [
        ("WC-FILL", "Replace toilet fill valve and test cistern operation."),
        ("WC-FLUSH", "Check and replace toilet flush valve / siphon where required."),
        ("ADD-RAD-REMOVE", "Remove existing radiator and make pipework safe as required."),
        ("RAD-FIT", "Fit / replace radiator in the agreed location."),
        ("RAD-VALVES", "Supply and fit radiator valve pair / TRV as required."),
        ("RAD-MOVE", "Move radiator from existing position to agreed new location."),
        ("PIPE-ALTER-M", "Alter associated pipework to suit the new layout."),
        ("PIPE-CAP", "Cap off redundant pipework and leave safe."),
        ("ADD-PIPE-REMOVE-CAP", "Remove redundant pipework and cap off as required."),
        ("BASIN-WASTE", "Disconnect / alter basin waste and trap as required."),
        ("BASIN-VANITY", "Install basin / vanity unit and connect services."),
        ("PB-BASIN-PED", "Fit basin and pedestal / unit and connect services."),
        ("TAP-BASIN", "Fit basin taps / mono tap as required."),
        ("OUTSIDE-TAP", "Install outside tap and connect to suitable cold water supply."),
    ]

    for code, sentence in mapping:
        if code in codes:
            lines.append(sentence)

    lines.append("All works to be tested on completion and working area left tidy.")

    unique = []
    seen = set()
    for line in lines:
        line = clean(line)
        if line and line not in seen:
            unique.append(line)
            seen.add(line)
    return "\n\n".join(unique)


def write_to_workbook(workbook_path: Path, suggestions, description: str, room: str):
    app = None
    book = None
    inserted = []
    duplicates = []
    try:
        app, book = open_book(workbook_path)
        works = get_sheet(book, WORKS_SHEET)
        final = get_sheet(book, FINAL_SHEET)

        existing = existing_task_codes(book)
        final.range("B6").value = build_description(description, suggestions)

        for item in suggestions:
            code = item.code.upper()
            if code in existing:
                duplicates.append((code, item.name))
                continue
            row = next_blank_work_row(book)
            works.range(f"B{row}").value = room
            works.range(f"C{row}").value = item.name
            works.range(f"D{row}").value = item.code
            works.range(f"E{row}").value = 1
            inserted.append((row, item.code, item.name))
            existing.add(code)

        book.save()
        return inserted, duplicates
    finally:
        safe_close(book, app)


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("description", nargs="*")
    p.add_argument("--workbook")
    p.add_argument("--insert", action="store_true")
    p.add_argument("--yes", action="store_true")
    p.add_argument("--room", default=DEFAULT_ROOM)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    description = " ".join(args.description).strip() or input("Works description: ").strip()
    if not description:
        print("No description entered.")
        return 1

    workbook_path = v3.choose_workbook(args.workbook)
    result = v6.build_scope_v6(description, workbook_path)
    v3.print_result(result)

    suggestions = insertable(result)
    print("")
    print("Insertable existing workbook tasks:")
    for s in suggestions:
        print(f"- {s.code} - {s.name}")

    missing = [s for s in result.suggestions if not s.exists_in_workbook]
    if missing:
        print("")
        print("Skipped missing tasks:")
        for s in missing:
            print(f"- {s.code} - {s.name}")

    print("")
    print("Description preview:")
    print(build_description(description, suggestions))

    if not args.insert:
        print("")
        print("Read-only mode. Add --insert to write to workbook.")
        return 0

    if not suggestions:
        print("No existing workbook tasks to insert.")
        return 1

    if not args.yes:
        answer = input(f"Insert {len(suggestions)} task(s) and write description? [y/N]: ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled.")
            return 0

    inserted, duplicates = write_to_workbook(workbook_path, suggestions, description, args.room)
    print("")
    print("Inserted tasks:")
    if inserted:
        for row, code, name in inserted:
            print(f"- Row {row}: {code} - {name}")
    else:
        print("- None; all suggested tasks were already on the quote.")

    if duplicates:
        print("")
        print("Skipped duplicates already on quote:")
        for code, name in duplicates:
            print(f"- {code} - {name}")

    print("")
    print("Description of Works written to Final Quote B6.")
    print("Workbook saved.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
