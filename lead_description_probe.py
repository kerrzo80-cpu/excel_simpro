#!/usr/bin/env python3
from pathlib import Path
import json
import requests
import excel_to_simpro_quote_with_lines as sender

BASE_URL = sender.BASE_URL
WORKBOOK = Path.home() / "Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm"

ENDPOINTS = [
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/notes/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/activities/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/logs/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/attachments/",
    "/api/v1.0/companies/{company_id}/leads/{lead_id}/customFields/",
]


def short(text, limit=2500):
    return (text or "")[:limit].replace("\n", " ")


def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--lead-id", required=True)
    args = p.parse_args()

    settings = sender.read_settings(str(WORKBOOK))
    token = sender.get_access_token(settings)
    company_id = settings["company_id"]
    headers = sender.headers(token)

    print("Company ID:", company_id)
    print("Lead ID:", args.lead_id)
    print("")

    for endpoint in ENDPOINTS:
        url = BASE_URL + endpoint.format(company_id=company_id, lead_id=args.lead_id)
        print("GET", url)
        try:
            r = requests.get(url, headers=headers, timeout=20)
            print("Status:", r.status_code)
            if r.text:
                print(short(r.text))
            print("")
        except Exception as e:
            print("ERROR:", e)
            print("")

    print("Done.")

if __name__ == "__main__":
    main()
