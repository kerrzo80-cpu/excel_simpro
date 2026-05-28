#!/usr/bin/env python3
"""Safe sender for Excel quotes to simPRO.

If Final Quote B5 contains a simPRO Lead ID, this converts that Lead to a Quote
first, then uploads the Excel quote lines into the converted Quote. This avoids
creating duplicate quotes while leaving the original Lead open.
"""

from __future__ import annotations

import sys
from pathlib import Path

import requests

import excel_to_simpro_quote_with_lines as sender


def convert_lead_to_quote(token: str, company_id: int, lead_id: str) -> dict:
    url = f"{sender.BASE_URL}/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/"
    response = requests.post(url, headers=sender.headers(token), json={"Type": sender.DEFAULT_QUOTE_TYPE})
    print("Lead convert status:", response.status_code)
    print(response.text)
    response.raise_for_status()
    return response.json()


def update_quote_header(token: str, quote_id: int, quote: dict) -> None:
    """Best-effort update of converted quote name/description.

    If simPRO rejects PATCH on this endpoint, do not stop the send. The quote
    lines still upload and the converted lead link remains correct.
    """
    description = sender.clean(quote.get("description"))
    if not description:
        return

    url = f"{sender.BASE_URL}/api/v1.0/companies/{sender.COMPANY_ID}/quotes/{quote_id}/"
    payload = {
        "Name": description,
        "Description": description,
        "Type": sender.DEFAULT_QUOTE_TYPE,
    }
    response = requests.patch(url, headers=sender.headers(token), json=payload)
    print("Quote header update status:", response.status_code)
    print(response.text)
    if response.status_code >= 400:
        print("Warning: quote header update failed; continuing with line upload.")


def upload_lines_to_quote(token: str, quote_id: int, quote: dict, lines: list[dict]) -> None:
    section = sender.create_section(token, quote_id, quote["section_name"])
    section_id = section["ID"]

    created_cost_centres: dict[int, int] = {}

    for line in lines:
        task_name = sender.clean(line.get("task_name"))
        task_code = sender.clean(line.get("task_code"))

        if not task_name or not task_code:
            continue
        if "select work task" in task_name.lower() or "select work task" in task_code.lower():
            print("Skipping placeholder row:", task_name, task_code)
            continue

        cost_centre_id = sender.COST_CENTRE_MAP.get(line["cost_centre_name"], 6)

        if cost_centre_id not in created_cost_centres:
            quote_cost_centre = sender.add_cost_centre(token, quote_id, section_id, cost_centre_id)
            created_cost_centres[cost_centre_id] = quote_cost_centre["ID"]

        quote_cost_centre_id = created_cost_centres[cost_centre_id]
        labor_type_id = sender.get_labor_type_for_line(line)

        print("Adding labour:", line["task_name"], line["labour_hours"], "hrs", "LaborType:", labor_type_id)
        sender.add_labor(token, quote_id, section_id, quote_cost_centre_id, labor_type_id, line["labour_hours"])

        for material in line["materials"]:
            qty = material["fixed_qty"] + (material["qty_per_unit"] * line["measurement"])
            qty = qty * line["count"]

            material_code = material["material_code"]
            price_info = line["material_prices"].get(material_code, {})
            sell_price = material["unit_sell_ex_vat"] or price_info.get("sell_ex_vat", 0) or 0
            material_name = price_info.get("item") or material["material_item"]
            spec = price_info.get("spec", "")
            unit = material["unit"] or price_info.get("unit", "")

            description = f"{material_name} {spec}".strip()
            description = f"{description} | Code: {material_code} | Unit: {unit}"
            if sell_price == 0:
                description += " | PRICE MISSING"

            print("Adding material:", description, "Qty:", qty, "Sell:", sell_price)
            sender.add_material(token, quote_id, section_id, quote_cost_centre_id, description, qty, sell_price)


def run_with_converted_lead(workbook_path: Path, quote: dict, lines: list[dict], lead_id: str) -> int:
    settings = sender.read_settings(str(workbook_path))
    token = sender.get_access_token(settings)

    converted_quote = convert_lead_to_quote(token, settings["company_id"], lead_id)
    quote_id = converted_quote["ID"]
    print("Converted lead", lead_id, "to quote:", quote_id)

    update_quote_header(token, quote_id, quote)
    upload_lines_to_quote(token, quote_id, quote, lines)
    sender.archive_and_reset_workbook(str(workbook_path), quote_id, quote)
    print("Finished converting lead and uploading quote with native labour and itemised materials.")
    return 0


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
        return run_with_converted_lead(workbook_path, quote, lines, lead_id)

    return sender.main() or 0


if __name__ == "__main__":
    raise SystemExit(main())
