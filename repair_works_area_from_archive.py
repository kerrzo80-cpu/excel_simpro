#!/usr/bin/env python3
"""Repair the Works To Be Carried Out entry area using a donor workbook.

This keeps the CURRENT live workbook and copies only the broken input/template
range from a known-good archived workbook. It does not replace the whole file.

Usage:
  python3 repair_works_area_from_archive.py --donor "/path/to/good/archive.xlsm"

Default live workbook:
  OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm
"""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
from datetime import datetime

import xlwings as xw

LIVE_WORKBOOK = Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm"
SHEET_NAME = "Works To Be Carried Out"
RANGE_TO_REPAIR = "A12:H200"


def get_sheet(book, name):
    for sheet in book.sheets:
        if sheet.name.strip() == name:
            return sheet
    raise RuntimeError(f"Missing sheet: {name}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--donor", required=True, help="Path to an archived workbook where dropdowns still work")
    parser.add_argument("--live", default=str(LIVE_WORKBOOK), help="Path to the live workbook to repair")
    parser.add_argument("--yes", action="store_true")
    args = parser.parse_args()

    donor = Path(args.donor).expanduser()
    live = Path(args.live).expanduser()

    if not donor.exists():
        print("Donor workbook not found:", donor)
        return 1
    if not live.exists():
        print("Live workbook not found:", live)
        return 1

    print("This will repair only:", SHEET_NAME, RANGE_TO_REPAIR)
    print("Live workbook:", live)
    print("Donor workbook:", donor)
    print("It will create a timestamped backup of the live workbook first.")

    if not args.yes:
        answer = input("Continue? [y/N]: ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled.")
            return 0

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup = live.with_name(live.stem + f" BACKUP before works repair {stamp}" + live.suffix)
    shutil.copy2(live, backup)
    print("Backup created:", backup)

    app = xw.App(visible=False, add_book=False)
    try:
        donor_book = app.books.open(str(donor))
        live_book = app.books.open(str(live))
        try:
            donor_sheet = get_sheet(donor_book, SHEET_NAME)
            live_sheet = get_sheet(live_book, SHEET_NAME)

            # Copy all formatting, formulas, values, data validation and dropdowns
            # for the entry area from the donor into the live workbook.
            donor_sheet.range(RANGE_TO_REPAIR).api.Copy(live_sheet.range(RANGE_TO_REPAIR).api)

            live_book.save()
            print("Repair complete. Copied", RANGE_TO_REPAIR, "from donor into live workbook.")
            print("Open Excel and check the dropdowns.")
        finally:
            donor_book.close()
            live_book.close()
    finally:
        app.quit()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
