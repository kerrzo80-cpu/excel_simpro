#!/usr/bin/env python3
"""Safe wrapper for sending Excel quotes to simPRO.

This prevents the duplicate-lead problem:
- If Final Quote B5 contains a Lead ID, this script refuses to create a new quote.
- That stops us creating a quote while leaving the original simPRO Lead untouched.

Next step is to implement the proper simPRO lead-to-quote conversion once the
correct API endpoint/workflow is confirmed.
"""

from __future__ import annotations

import sys
from pathlib import Path

import excel_to_simpro_quote_with_lines as sender


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python3 excel_to_simpro_quote_safe.py /path/to/workbook.xlsm")
        return 2

    workbook_path = Path(sys.argv[1]).expanduser()
    quote, lines = sender.read_final_quote(str(workbook_path))

    lead_id = sender.clean(quote.get("lead_id"))
    description = sender.clean(quote.get("description"))

    print("Customer:", quote.get("customer_name"))
    print("Site:", quote.get("site_name"))
    print("Lead ID:", lead_id or "None")
    print("Description:", description or "MISSING")
    print("Lines found:", len(lines))

    if not description:
        print("")
        print("STOPPED: Final Quote B6 is empty.")
        print("Add a customer-facing description of works before sending to simPRO.")
        return 1

    if lead_id:
        print("")
        print("STOPPED: This quote came from a simPRO Lead.")
        print("Creating a brand-new quote would leave the Lead open and create duplicates.")
        print("Next fix required: convert/update Lead ID", lead_id, "into a Quote instead of POSTing /quotes/.")
        return 1

    # No lead ID: safe to create a standalone quote using the existing sender.
    return sender.main() or 0


if __name__ == "__main__":
    raise SystemExit(main())
