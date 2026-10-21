"""Load optional, email-keyed roster fields for declarative scheduling rules."""

from __future__ import annotations

import csv
import re
from pathlib import Path


SUPPORTED_FIELDS = {
    "accessstatus", "classyear", "driverstatus", "driverrank", "primarystatus",
    "rolestatus", "role", "scheduleradmin",
}


def _key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def _rows(path: str | Path):
    with Path(path).open(newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise ValueError("Roster CSV must have a header row")
        headers = {name: _key(name) for name in reader.fieldnames if name}
        email_header = next((name for name, key in headers.items()
                             if key in {"brownemail", "email", "emailaddress"}), None)
        if email_header is None:
            raise ValueError("Roster CSV must include a BrownEmail or Email column")
        fields = [name for name, key in headers.items()
                  if key in SUPPORTED_FIELDS]
        return email_header, fields, list(reader)


def load_attribute_catalog(path: str | Path) -> dict[str, list[str]]:
    """Return supported roster fields and their nonblank values for editor choices."""
    _, fields, rows = _rows(path)
    return {
        name: sorted({str(row.get(name, "")).strip() for row in rows
                      if str(row.get(name, "")).strip()}, key=str.casefold)
        for name in fields
    }


def apply_member_attributes(path: str | Path, people) -> list[str]:
    """Attach roster fields to scheduler members by email; warn on unmatched rows."""
    email_header, fields, rows = _rows(path)
    by_email = {}
    for row in rows:
        email = str(row.get(email_header, "")).strip().casefold()
        if not email:
            continue
        if email in by_email:
            raise ValueError(f"Duplicate email in roster CSV: {email}")
        by_email[email] = {name: str(row.get(name, "")).strip() for name in fields}
    people_by_email = {person.email.casefold(): person for person in people}
    for email, attributes in by_email.items():
        person = people_by_email.get(email)
        if person is not None:
            person.attributes.update(attributes)
    missing = sorted(email for email in people_by_email if email not in by_email)
    return [f"{len(missing)} form respondent(s) have no matching row in the roster CSV." ] if missing else []
