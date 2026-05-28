#!/usr/bin/env python3
"""Safe simPRO Lead -> Quote conversion endpoint probe.

Read-only by default. It searches the lead list first because simPRO may not
support GET /leads/{id}/ even though /leads/ list works.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import requests

import excel_to_simpro_quote_with_lines as sender

BASE_URL = sender.BASE_URL

LIKELY_ENDPOINTS = [
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convertTo/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert-to/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/quotes/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/quotes/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/",
]


def clean(value: Any) -> str:
    return str(value or "").strip()


def read_lead_id_from_workbook(workbook_path: Path) -> str:
    quote, _ = sender.read_final_quote(str(workbook_path))
    return clean(quote.get("lead_id"))


def short(text: str, limit: int = 700) -> str:
    return (text or "")[:limit].replace("\n", " ")


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


def find_lead_in_list(token: str, company_id: int, lead_id: str):
    headers = sender.headers(token)
    for page in range(1, 51):
        url = f"{BASE_URL}/api/v1.0/companies/{company_id}/leads/"
        response = requests.get(url, headers=headers, params={"page": page, "pageSize": 100}, timeout=20)
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
                print("Found lead in list:")
                print(short(str(lead), 1200))
                print("")
                return lead
        if len(leads) < 100:
            return None
    return None


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--workbook", required=False)
    p.add_argument("--lead-id", required=False)
    p.add_argument("--do-post", action="store_true", help="DANGEROUS: attempt POST conversion. Use only on a test lead.")
    p.add_argument("--endpoint", help="Specific endpoint path to probe/POST.")
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

    lead = find_lead_in_list(token, company_id, lead_id)
    if not lead:
        print("Lead was not found in the open leads list.")
        print("This may mean the lead has already been converted/closed, or the ID is not an open lead ID.")
        print("The probe will still test likely endpoint shapes below.")
        print("")

    endpoints = [args.endpoint] if args.endpoint else LIKELY_ENDPOINTS
    for template in endpoints:
        path = template.format(company_id=company_id, lead_id=lead_id)
        url = BASE_URL + path
        request_probe("OPTIONS", url, token)
        request_probe("GET", url, token)
        if args.do_post:
            print("POST conversion attempt enabled. Only use this on a disposable test lead.")
            request_probe("POST", url, token, payload={"Type": "Project"})

    print("Probe finished.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
