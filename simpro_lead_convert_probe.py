#!/usr/bin/env python3
"""Safe simPRO Lead -> Quote conversion endpoint probe.

Read-only by default. It searches the lead list because simPRO may not support
GET /leads/{id}/ even though /leads/ list works.

Use this to inspect whether a converted lead still appears in the leads list,
and to inspect available lead status/stage fields before/after conversion.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import requests

import excel_to_simpro_quote_with_lines as sender

BASE_URL = sender.BASE_URL

LIKELY_ENDPOINTS = [
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/",
]


def clean(value: Any) -> str:
    return str(value or "").strip()


def read_lead_id_from_workbook(workbook_path: Path) -> str:
    quote, _ = sender.read_final_quote(str(workbook_path))
    return clean(quote.get("lead_id"))


def short(text: str, limit: int = 1200) -> str:
    return (text or "")[:limit].replace("\n", " ")


def pretty(value: Any) -> str:
    try:
        return json.dumps(value, indent=2, sort_keys=True)
    except Exception:
        return str(value)


def request_probe(method: str, url: str, token: str, payload: dict | None = None):
    try:
        response = requests.request(method, url, headers=sender.headers(token), json=payload, timeout=20)
        print(f"{method:<7} {url}")
        print("Status:", response.status_code)
        if response.text:
            print("Body:", short(response.text))
        print("")
        return response
    except Exception as exc:
        print(f"{method:<7} {url}")
        print("ERROR:", exc)
        print("")
        return None


def find_lead_in_list(token: str, company_id: int, lead_id: str, quiet: bool = False):
    headers = sender.headers(token)
    for page in range(1, 51):
        url = f"{BASE_URL}/api/v1.0/companies/{company_id}/leads/"
        response = requests.get(url, headers=headers, params={"page": page, "pageSize": 100}, timeout=20)
        if not quiet:
            print(f"Lead list page {page} status:", response.status_code)
        if response.status_code != 200:
            print(short(response.text))
            return None
        data = response.json()
        leads = data.get("Items") if isinstance(data, dict) else data
        if not leads:
            return None
        for lead in leads:
            current_id = clean(lead.get("ID") or lead.get("id"))
            if current_id == lead_id:
                return lead
        if len(leads) < 100:
            return None
    return None


def print_lead_summary(label: str, lead: dict | None) -> None:
    print(label)
    if not lead:
        print("- Lead not found in /leads/ list")
        print("")
        return
    keys_of_interest = [
        "ID", "Name", "LeadName", "Description", "Stage", "Status", "IsClosed", "Closed", "Archived", "ArchiveReason",
        "Converted", "ConvertedFromLead", "ConvertedToQuote", "Quote", "DateCreated", "DateModified", "Customer", "Site",
    ]
    for key in keys_of_interest:
        if key in lead:
            print(f"- {key}: {pretty(lead.get(key))}")
    print("Full lead object:")
    print(pretty(lead))
    print("")


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--workbook", required=False)
    p.add_argument("--lead-id", required=False)
    p.add_argument("--do-post", action="store_true", help="DANGEROUS: attempt POST conversion. Use only on a test lead.")
    p.add_argument("--endpoint", help="Specific endpoint path to probe/POST.")
    p.add_argument("--inspect-only", action="store_true", help="Only inspect whether the lead is still listed; do not probe endpoints.")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    workbook_path = Path(args.workbook).expanduser() if args.workbook else Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm"

    settings = sender.read_settings(str(workbook_path))
    token = sender.get_access_token(settings)
    company_id = settings["company_id"]

    lead_id = clean(args.lead_id) or read_lead_id_from_workbook(workbook_path)
    if not lead_id:
        print("No lead ID supplied and Final Quote B5 is empty.")
        return 1

    print("Workbook:", workbook_path)
    print("Company ID:", company_id)
    print("Lead ID:", lead_id)
    print("")

    lead_before = find_lead_in_list(token, company_id, lead_id)
    print_lead_summary("Lead before probe/conversion:", lead_before)

    if args.inspect_only:
        print("Inspect-only mode finished.")
        return 0

    endpoints = [args.endpoint] if args.endpoint else LIKELY_ENDPOINTS
    converted_quote = None
    for template in endpoints:
        path = template.format(company_id=company_id, lead_id=lead_id)
        url = BASE_URL + path
        request_probe("OPTIONS", url, token)
        request_probe("GET", url, token)
        if args.do_post and path.endswith("/convert/"):
            print("POST conversion attempt enabled. Only use this on a disposable test lead.")
            response = request_probe("POST", url, token, payload={"Type": "Project"})
            if response is not None and response.status_code in (200, 201):
                try:
                    converted_quote = response.json()
                    print("Converted quote object:")
                    print(pretty(converted_quote))
                    print("")
                except Exception:
                    pass

    lead_after = find_lead_in_list(token, company_id, lead_id, quiet=True)
    print_lead_summary("Lead after probe/conversion:", lead_after)
    if converted_quote:
        print("Converted quote ID:", converted_quote.get("ID"))
        print("ConvertedFromLead:", pretty(converted_quote.get("ConvertedFromLead")))
    print("Probe finished.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
