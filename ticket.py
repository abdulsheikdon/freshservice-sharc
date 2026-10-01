#!/usr/bin/env python3
"""
Abdullahi Sheikdon
this intro was made by claude

tickets_to_excel.py - export tickets from a ticketing system to an Excel workbook.

Supported providers: freshservice, zendesk (add more by subclassing Provider).

Setup:
    pip install requests openpyxl

Credentials (environment variables):
    Freshservice: FRESHSERVICE_DOMAIN (e.g. "acme" for acme.freshservice.com)
                  FRESHSERVICE_API_KEY
    Zendesk:      ZENDESK_SUBDOMAIN, ZENDESK_EMAIL, ZENDESK_API_TOKEN

Usage:
    python tickets_to_excel.py --provider freshservice
    python tickets_to_excel.py --provider zendesk --since 2026-01-01 -o zendesk.xlsx
"""
import argparse
import os
import sys
import time
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Iterator, Optional

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Every provider is normalized to these columns, so the sheet looks the same
# regardless of the source system.
COLUMNS = [
    "ID", "Subject", "Status", "Priority", "Type", "Requester ID",
    "Assignee ID", "Created", "Updated", "Due", "Tags", "Source", "URL",
]


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def parse_dt(value: Optional[str]) -> Optional[datetime]:
    """ISO-8601 string -> naive datetime (Excel can't store timezones)."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        return None


def get_json(session: requests.Session, url: str, params=None, max_retries: int = 5) -> dict:
    """GET with automatic backoff when the API rate-limits us (HTTP 429)."""
    for attempt in range(max_retries):
        resp = session.get(url, params=params, timeout=30)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 2 ** attempt))
            print(f"  rate limited, waiting {wait}s...", file=sys.stderr)
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    raise RuntimeError("Gave up after repeated rate-limit responses")


def require_env(*names: str) -> list:
    missing = [n for n in names if not os.environ.get(n)]
    if missing:
        sys.exit(f"Missing environment variable(s): {', '.join(missing)}")
    return [os.environ[n] for n in names]


# --------------------------------------------------------------------------- #
# providrs
# --------------------------------------------------------------------------- #
class Provider(ABC):
    name: str

    @abstractmethod
    def fetch(self, since: Optional[datetime]) -> Iterator[dict]:
        """Yield normalized ticket dicts keyed by COLUMNS."""


class Freshservice(Provider):
    name = "Freshservice"
    STATUS = {2: "Open", 3: "Pending", 4: "Resolved", 5: "Closed"}
    PRIORITY = {1: "Low", 2: "Medium", 3: "High", 4: "Urgent"}
    PER_PAGE = 100

    def __init__(self):
        domain, api_key = require_env("FRESHSERVICE_DOMAIN", "FRESHSERVICE_API_KEY")
        self.base = f"https://{domain}.freshservice.com"
        self.session = requests.Session()
        self.session.auth = (api_key, "X") 

    def fetch(self, since):
        updated_since = (since or datetime(2010, 1, 1)).strftime("%Y-%m-%dT%H:%M:%SZ")
        page = 1
        while True:
            data = get_json(
                self.session,
                f"{self.base}/api/v2/tickets",
                params={"per_page": self.PER_PAGE, "page": page,
                        "updated_since": updated_since, "order_type": "asc"},
            )
            batch = data.get("tickets", [])
            for t in batch:
                yield {
                    "ID": t["id"],
                    "Subject": t.get("subject"),
                    "Status": self.STATUS.get(t.get("status"), t.get("status")),
                    "Priority": self.PRIORITY.get(t.get("priority"), t.get("priority")),
                    "Type": t.get("type"),
                    "Requester ID": t.get("requester_id"),
                    "Assignee ID": t.get("responder_id"),
                    "Created": parse_dt(t.get("created_at")),
                    "Updated": parse_dt(t.get("updated_at")),
                    "Due": parse_dt(t.get("due_by")),
                    "Tags": ", ".join(t.get("tags") or []),
                    "Source": self.name,
                    "URL": f"{self.base}/a/tickets/{t['id']}",
                }
            if len(batch) < self.PER_PAGE:
                break
            page += 1


class Zendesk(Provider):
    name = "Zendesk"

    def __init__(self):
        sub, email, token = require_env("ZENDESK_SUBDOMAIN", "ZENDESK_EMAIL", "ZENDESK_API_TOKEN")
        self.base = f"https://{sub}.zendesk.com"
        self.session = requests.Session()
        self.session.auth = (f"{email}/token", token)

    def fetch(self, since):
        url = f"{self.base}/api/v2/tickets.json"
        params = {"page[size]": 100}
        while url:
            data = get_json(self.session, url, params=params)
            params = None 
            for t in data.get("tickets", []):
                updated = parse_dt(t.get("updated_at"))
                if since and updated and updated < since:
                    continue
                yield {
                    "ID": t["id"],
                    "Subject": t.get("subject"),
                    "Status": (t.get("status") or "").title(),
                    "Priority": (t.get("priority") or "").title() or None,
                    "Type": t.get("type"),
                    "Requester ID": t.get("requester_id"),
                    "Assignee ID": t.get("assignee_id"),
                    "Created": parse_dt(t.get("created_at")),
                    "Updated": updated,
                    "Due": parse_dt(t.get("due_at")),
                    "Tags": ", ".join(t.get("tags") or []),
                    "Source": self.name,
                    "URL": f"{self.base}/agent/tickets/{t['id']}",
                }
            url = data.get("links", {}).get("next") if data.get("meta", {}).get("has_more") else None


PROVIDERS = {"freshservice": Freshservice, "zendesk": Zendesk}


# --------------------------------------------------------------------------- #
# excel 
# --------------------------------------------------------------------------- #
def write_workbook(tickets: list, path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tickets"

    base_font = Font(name="Arial", size=10)
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="1F3864")

    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font, cell.fill = header_font, header_fill
        cell.alignment = Alignment(vertical="center")

    for t in tickets:
        ws.append([t.get(c) for c in COLUMNS])

    date_cols = {COLUMNS.index(c) + 1 for c in ("Created", "Updated", "Due")}
    url_col = COLUMNS.index("URL") + 1
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = base_font
            if cell.column in date_cols:
                cell.number_format = "yyyy-mm-dd hh:mm"
            if cell.column == url_col and cell.value:
                cell.hyperlink = cell.value
                cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")

    for i, col in enumerate(COLUMNS, start=1):
        longest = max((len(str(c.value)) for c in ws[get_column_letter(i)] if c.value), default=10)
        ws.column_dimensions[get_column_letter(i)].width = min(max(longest + 2, 10), 60)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Summary sheet: live formulas, so it updates if you edit or filter the data
    summary = wb.create_sheet("Summary")
    summary.append(["Status", "Tickets"])
    for cell in summary[1]:
        cell.font, cell.fill = header_font, header_fill
    status_col = get_column_letter(COLUMNS.index("Status") + 1)
    statuses = sorted({str(t["Status"]) for t in tickets if t.get("Status")})
    for r, status in enumerate(statuses, start=2):
        summary.cell(r, 1, status).font = base_font
        summary.cell(r, 2, f'=COUNTIF(Tickets!${status_col}:${status_col},A{r})').font = base_font
    total_row = len(statuses) + 2
    summary.cell(total_row, 1, "Total").font = Font(name="Arial", size=10, bold=True)
    summary.cell(total_row, 2, f"=SUM(B2:B{total_row - 1})").font = Font(name="Arial", size=10, bold=True)
    summary.column_dimensions["A"].width = 18
    summary.column_dimensions["B"].width = 12

    wb.save(path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> None:
    p = argparse.ArgumentParser(description="Export tickets to Excel")
    p.add_argument("--provider", required=True, choices=PROVIDERS)
    p.add_argument("--since", help="only tickets updated on/after this date (YYYY-MM-DD)")
    p.add_argument("-o", "--output", help="output .xlsx path")
    args = p.parse_args()

    since = datetime.strptime(args.since, "%Y-%m-%d") if args.since else None
    output = args.output or f"{args.provider}_tickets_{datetime.now():%Y%m%d_%H%M}.xlsx"

    provider = PROVIDERS[args.provider]()
    print(f"Fetching tickets from {provider.name}...")
    try:
        tickets = []
        for t in provider.fetch(since):
            tickets.append(t)
            if len(tickets) % 100 == 0:
                print(f"  {len(tickets)} tickets...")
    except requests.HTTPError as e:
        sys.exit(f"API error: {e.response.status_code} {e.response.text[:300]}")

    write_workbook(tickets, output)
    print(f"Wrote {len(tickets)} tickets to {output}")


if __name__ == "__main__":
    main()