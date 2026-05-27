import requests
from pathlib import Path
import xlwings as xw

BASE_URL = "https://errolwatson.simprosuite.com"
WORKBOOK = "/Users/ewgcoomercial/Library/CloudStorage/OneDrive-ErrolWatsonGroup/Quoting/Works Pricing Tool V3 BACKUP before visual style 2026-05-27 09-11.xlsm"
TOKEN_FILE = Path.home() / "Downloads" / "simpro_refresh_token.txt"


def clean(value):
    return str(value or "").strip()


def get_sheet(book, wanted_name):
    for sheet in book.sheets:
        if sheet.name.strip() == wanted_name.strip():
            return sheet
    raise RuntimeError(f"Could not find sheet: {wanted_name}")


def open_book(path):
    app = xw.App(visible=False, add_book=False)
    book = app.books.open(str(path))
    return app, book


def read_settings_from_workbook(path):
    app, book = open_book(path)
    try:
        settings = get_sheet(book, "Settings")
        return {
            "company_id": int(settings.range("B3").value),
            "client_id": clean(settings.range("B4").value),
            "client_secret": clean(settings.range("B5").value),
        }
    finally:
        book.close()
        app.quit()


def get_token(settings):
    response = requests.post(
        f"{BASE_URL}/oauth2/token",
        data={
            "grant_type": "refresh_token",
            "client_id": settings["client_id"],
            "client_secret": settings["client_secret"],
            "refresh_token": TOKEN_FILE.read_text().strip(),
        },
    )

    print("Token status:", response.status_code)
    response.raise_for_status()
    data = response.json()

    if data.get("refresh_token"):
        TOKEN_FILE.write_text(data["refresh_token"])
        print("Refresh token updated")

    return data["access_token"]


def search_leads(token, company_id, search):
    search = clean(search).lower()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    matches = []
    seen = set()

    def flatten(value):
        parts = []
        if isinstance(value, dict):
            for v in value.values():
                parts.append(flatten(v))
        elif isinstance(value, list):
            for v in value:
                parts.append(flatten(v))
        elif value is not None:
            parts.append(str(value))
        return " ".join(parts)

    def pick_address(site):
        if not isinstance(site, dict):
            return ""
        addr = site.get("Address") or {}
        if isinstance(addr, dict):
            pieces = [
                addr.get("Address"),
                addr.get("Address2"),
                addr.get("Suburb"),
                addr.get("State"),
                addr.get("Postcode"),
                addr.get("PostalCode"),
            ]
            return clean(", ".join([str(x) for x in pieces if x]))
        return clean(addr)

    # Pull multiple pages. simPRO often defaults to a limited first page.
    for page in range(1, 21):
        response = requests.get(
            f"{BASE_URL}/api/v1.0/companies/{company_id}/leads/",
            headers=headers,
            params={"page": page, "pageSize": 100},
        )

        print(f"Lead page {page} status:", response.status_code)

        if response.status_code != 200:
            print(response.text[:1000])
            break

        data = response.json()
        leads = data.get("Items") if isinstance(data, dict) else data

        if not leads:
            break

        for lead in leads:
            lead_id = lead.get("ID") or lead.get("id")
            if lead_id in seen:
                continue
            seen.add(lead_id)

            customer = lead.get("Customer") or lead.get("Client") or {}
            site = lead.get("Site") or {}

            company = clean(customer.get("CompanyName")) if isinstance(customer, dict) else ""
            given = clean(customer.get("GivenName")) if isinstance(customer, dict) else ""
            family = clean(customer.get("FamilyName")) if isinstance(customer, dict) else ""
            customer_name = clean(company or f"{given} {family}".strip())

            site_name = clean(site.get("Name")) if isinstance(site, dict) else ""
            site_address = pick_address(site)

            lead_name = clean(lead.get("Name") or lead.get("Description") or lead.get("Subject"))
            lead_number = clean(lead.get("LeadNumber") or lead.get("Number") or lead_id)

            full_text = flatten(lead).lower()

            # Match all words typed in B3, so "Brian Kerr" can match across fields.
            words = [w for w in search.split() if w]
            if search in full_text or all(w in full_text for w in words):
                matches.append({
                    "lead_id": lead_id,
                    "customer": customer_name or lead_name or search.title(),
                    "site": site_address or site_name,
                    "lead_name": lead_name,
                    "lead_number": lead_number,
                })

    return matches

def write_results_to_workbook(path, matches):
    app, book = open_book(path)
    try:
        final = get_sheet(book, "Final Quote")

        # Clear old visible result rows only
        final.range("A20:D35").clear_contents()

        results = book.sheets["Lead Search Results"] if "Lead Search Results" in [s.name for s in book.sheets] else book.sheets.add("Lead Search Results")
        results.clear_contents()
        results.range("A1").value = "Lead ID"
        results.range("B1").value = "Customer"
        results.range("C1").value = "Site Address"
        results.range("D1").value = "Selected"

        if not matches:
            results.range("A2").value = "No matching leads found"
            book.save()
            return

        rows = [[m["lead_id"], m["customer"], m["site"], ""] for m in matches[:10]]
        results.range("A2").value = rows

        # Load the first match into the quote header
        best = matches[0]
        final.range("B3").value = best["customer"]
        final.range("B4").value = best["site"]
        final.range("B5").value = best["lead_id"]

        book.save()
    finally:
        book.close()
        app.quit()


def main():
    workbook = Path(WORKBOOK)

    app, book = open_book(workbook)
    try:
        final = get_sheet(book, "Final Quote")
        search = clean(final.range("B3").value).lower()
    finally:
        book.close()
        app.quit()

    if not search:
        print("Type a customer name or address in Final Quote B3 first.")
        return

    settings = read_settings_from_workbook(workbook)
    token = get_token(settings)
    matches = search_leads(token, settings["company_id"], search)

    write_results_to_workbook(workbook, matches)

    if matches:
        print(f"Found {len(matches)} matching lead(s). Loaded first match into B3:B5.")
        for match in matches[:10]:
            print(f"{match['lead_id']} | {match['customer']} | {match['site']}")
    else:
        print("No matching leads found.")


if __name__ == "__main__":
    main()
