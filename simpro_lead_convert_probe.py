#!/usr/bin/env python3
"""Safe simPRO Lead -> Quote conversion endpoint probe.

This script is intentionally read-only by default.
It helps discover which simPRO API route supports the same action as:
Lead > Options > Convert To > Quote.

It will:
  1. Read simPRO settings/token from the workbook.
  2. Read a lead ID from --lead-id or Final Quote B5.
  3. GET the lead to prove the ID is valid.
  4. Probe likely conversion endpoints using OPTIONS/GET where possible.

It will NOT run a POST conversion unless --do-post is supplied.
Only use --do-post on a disposable test lead.
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
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convertTo/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert-to/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/quotes/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/quotes/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/quote/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/convert/",
]


def clean(value: Any) -> str:
    return str(value or "").strip()


def read_lead_id_from_workbook(workbook_path: Path) -> str:
    quote, _ = sender.read_final_quote(str(workbook_path))
    return clean(quote.get("lead_id"))


def short(text: str, limit: int = 700) -> str:
    text = text or ""
    return text[:limit].replace("\n", " ")


def request_probe(method: str, url: str, token: str, payload: dict | None = None):
    try:
        response = requests.request(
            method,
            url,
            headers=sender.headers(token),
            json=payload,
            timeout=20,
        )
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


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--workbook", required=False, help="Workbook path. Defaults to the V3 quoting workbook.")
    p.add_argument("--lead-id", required=False, help="simPRO lead ID. If omitted, reads Final Quote B5.")
    p.add_argument("--do-post", action="store_true", help="DANGEROUS: attempt POST conversion. Use only on a test lead.")
    p.add_argument("--endpoint", help="Specific endpoint path to POST/OPTIONS, e.g. /api/v1.0/companies/{company_id}/leads/{lead_id}/convert/quote/")
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
        print("Use: python3 simpro_lead_convert_probe.py --lead-id 123")
        return 1

    print("Workbook:", workbook_path)
    print("Company ID:", company_id)
    print("Lead ID:", lead_id)
    print("")

    lead_url = f"{BASE_URL}/api/v1.0/companies/{company_id}/leads/{lead_id}/"
    response = request_probe("GET", lead_url, token)
    if response is None or response.status_code >= 400:
        print("Lead GET failed. Stop here and check the lead ID.")
        return 1

    endpoints = [args.endpoint] if args.endpoint else LIKELY_ENDPOINTS
    for template in endpoints:
        path = template.format(company_id=company_id, lead_id=lead_id)
        url = BASE_URL + path
        request_probe("OPTIONS", url, token)
        request_probe("GET", url, token)
        if args.do_post:
            print("POST conversion attempt enabled. Only use this on a disposable test lead.")
            payload = {"Type": "Project"}
            request_probe("POST", url, token, payload=payload)

    print("Probe finished.")
    print("If one endpoint returns 200/201/204 or an Allowed methods response, paste the output back here.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
