#!/usr/bin/env python3
"""Clear the active quote area in the Works Pricing Tool.

This is for starting a clean test/live quote so old work rows do not get sent
to simPRO by mistake.

Usage:
  python3 clear_quote.py
  python3 clear_quote.py --yes
"""

from __future__ import annotations

import argparse
from pathlib import Path

import xlwings as xw

WORKBOOK = Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm"
FINAL_SHEET = "Final Quote"
WORKS_SHEET = "Works To Be Carried Out"


def get_sheet(book, name):
    for sheet in book.sheets:
        if sheet.name.strip() == name:
            return sheet
    raise RuntimeError(f"Missing sheet: {name}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--yes", action="store_true")
    args = parser.parse_args()

    if not args.yes:
        print("This will clear customer/site/lead/description and work rows 12:200.")
        answer = input("Continue? [y/N]: ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled.")
            return 0

    app = xw.App(visible=False, add_book=False)
    book = app.books.open(str(WORKBOOK))
    try:
        final = get_sheet(book, FINAL_SHEET)
        works = get_sheet(book, WORKS_SHEET)

        # Final quote header fields
        final.range("B3:B6").clear_contents()
        final.range("A20:D35").clear_contents()

        # Works area: clear room/task/code/count/measurement/notes style columns.
        # We clear broad columns to remove old copied formulas and stale work items.
        works.range("A12:H200").clear_contents()

        book.save()
        print("Cleared quote header and works rows 12:200.")
        print("Workbook saved:", WORKBOOK)
    finally:
        try:
            book.close()
        finally:
            app.quit()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
