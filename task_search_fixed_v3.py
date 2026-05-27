#!/usr/bin/env python3
"""
Wrapper around task_search_fixed_v2.py that ignores Excel close/quit cleanup errors.
Use this if xlwings/appscript raises OSERROR -1728 after printing the results.
"""

import argparse

from task_search_fixed_v2 import (
    choose_workbook,
    open_book,
    read_tasks,
    best_matches,
    insert_task,
)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("query", nargs="*", help="Plain-English work task, e.g. remove basin")
    parser.add_argument("--workbook", help="Path to Works Pricing Tool .xlsm")
    parser.add_argument("--list-only", action="store_true", help="Show matches but do not add to workbook")
    return parser.parse_args()


def safe_close(book, app):
    try:
        if book is not None:
            book.close()
    except Exception as exc:
        print(f"Note: Excel workbook close warning ignored: {exc}")
    try:
        if app is not None:
            app.quit()
    except Exception as exc:
        print(f"Note: Excel app quit warning ignored: {exc}")


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
    app = None
    book = None
    try:
        app, book = open_book(workbook_path)
        tasks = read_tasks(book)
        matches = best_matches(query, tasks)

        print("")
        print(f"Workbook: {workbook_path}")
        print(f"Search: {query}")

        if not matches:
            print("No good match found. This likely means the workbook needs a new task for this phrase.")
            return

        print("Best matches:")
        for idx, (score, task) in enumerate(matches, 1):
            print(f"{idx}. {task['name']}  [{task['code']}]  score={score}")

        if args.list_only:
            return

        top_score, top_task = matches[0]
        second_score = matches[1][0] if len(matches) > 1 else 0
        confident = top_score >= 125 and (top_score - second_score) >= 35

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
        safe_close(book, app)


if __name__ == "__main__":
    main()
