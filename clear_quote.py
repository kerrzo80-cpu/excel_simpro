#!/usr/bin/env python3
"""Clear only safe header/search fields in the Works Pricing Tool.

IMPORTANT: This script deliberately does NOT clear the Works To Be Carried Out
entry area, because that area contains dropdown/formula setup. Clear work rows
manually inside Excel until we build a proper reset macro that preserves data
validation and formulas.

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
        print("This will clear only Final Quote B3:B6 and lead search result rows A20:E35.")
        print("It will NOT touch the works/dropdown area.")
        answer = input("Continue? [y/N]: ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled.")
            return 0

    app = xw.App(visible=False, add_book=False)
    book = app.books.open(str(WORKBOOK))
    try:
        final = get_sheet(book, FINAL_SHEET)
        final.range("B3:B6").clear_contents()
        final.range("A20:E35").clear_contents()
        book.save()
        print("Cleared Final Quote B3:B6 and search result rows A20:E35 only.")
        print("Works/dropdown area was not touched.")
        print("Workbook saved:", WORKBOOK)
    finally:
        try:
            book.close()
        finally:
            app.quit()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
