"""
parse_form.py
=============
Reads the Google Form CSV export and returns (volunteers, bert_members).

The form (see FORM_GUIDE.md) has one column per date per question grid:
  - "Day Shifts ... [Mon 4/27]"      cells contain AM and/or PM
  - "Night Shifts [Mon 4/27]"        cells contain NIGHT
  - "Weekend Day [Sat 5/2]"          cells contain DAY
  - BERT availability grid [date]    cells contain A/B/C/D
Cells may also contain "NOT AVAILABLE"; any real shift token in the same cell
still counts as availability.

Duplicate submissions are resolved by keeping the latest timestamp per email.
Ambulance EMTs' campus availability is inferred from their AM/PM availability
(AM -> blocks A,B; PM -> blocks C,D).
"""

from __future__ import annotations

import csv
import re
from datetime import date, datetime, timedelta
from typing import Optional

from models import (
    CAMPUS_BLOCKS,
    BertMember,
    Volunteer,
    is_weekend,
    shift_types_for,
)

# ── Header / role markers ────────────────────────────────────────────────────

COL_ROLE = "Are you an ambulance EMT or BERT member?"
COL_EMAIL = "Email Address"
COL_EMAIL_FALLBACK = "Username"
COL_TIMESTAMP = "Timestamp"
COL_DRIVER = "Driver Status"
COL_DIFFICULTIES = "Do you foresee"   # freeform blackout question (appears once per section)

ROLE_EMT_MARKERS = ("ambulance emt", "emt only", "dual-role")
ROLE_BERT_MARKERS = ("bert member", "bert only")

_BRACKET_DATE = re.compile(r"\[(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s*(\d{1,2})/(\d{1,2})", re.IGNORECASE)


# ── Small parsing helpers ────────────────────────────────────────────────────

def _safe(row: list[str], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    return (row[idx] or "").strip()


def _parse_timestamp(ts_str: str) -> datetime:
    # Strip a trailing timezone code ("EST", "GMT-4") but never an AM/PM marker.
    s = re.sub(r"\s+(?!AM\b|PM\b)[A-Z]{2,4}(?:[+-]\d{1,2})?\s*$", "", (ts_str or "").strip())
    for fmt in (
        "%m/%d/%Y %H:%M:%S",       # Google Sheets CSV export (current forms)
        "%Y/%m/%d %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return datetime.min


def _header_date(month: int, day: int, block_start: date, block_end: date) -> Optional[date]:
    """Resolve a M/D header to a date inside the block (handles year rollover)."""
    for year in (block_start.year, block_start.year + 1):
        try:
            d = date(year, month, day)
        except ValueError:
            return None
        if block_start <= d <= block_end:
            return d
    return None


def _tokens(cell: str, pattern: str) -> list[str]:
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", "no"):
        return []
    seen: list[str] = []
    for token in re.findall(pattern, cell, re.IGNORECASE):
        t = token.upper()
        if t not in seen:
            seen.append(t)
    return seen


def _shift_tokens(cell: str) -> list[str]:
    return _tokens(cell, r"\b(AM|PM|NIGHT|DAY)\b")


def _block_tokens(cell: str) -> list[str]:
    return [t for t in _tokens(cell, r"\b([ABCD])\b") if t in CAMPUS_BLOCKS]


def normalise_driver(raw: str) -> str:
    r = (raw or "").strip().upper()
    if "EVDT" in r:
        return "EVDT"
    if "AUTH" in r:
        return "Auth"
    return "EMT"


def _is_emt_role(role: str) -> bool:
    return any(m in (role or "").lower() for m in ROLE_EMT_MARKERS)


def _is_bert_role(role: str) -> bool:
    return any(m in (role or "").lower() for m in ROLE_BERT_MARKERS)


def _find_nth_header(headers: list[str], exact: str, n: int) -> int:
    """Index of the (n+1)-th header exactly equal to `exact`, or -1."""
    count = 0
    for i, h in enumerate(headers):
        if (h or "").strip() == exact:
            if count == n:
                return i
            count += 1
    return -1


# ── Column map ───────────────────────────────────────────────────────────────

def _build_column_maps(headers: list[str], block_start: date, block_end: date) -> dict:
    emt_day, emt_night, emt_weekend, emt_week, bert = {}, {}, {}, {}, {}

    def find(exact: str) -> int:
        return next((i for i, h in enumerate(headers) if (h or "").strip() == exact), -1)

    def find_contains(sub: str, after: int = -1) -> int:
        return next((i for i, h in enumerate(headers) if i > after and h and sub in h), -1)

    # The Fall 2026 shopping-period form keeps its visible grid titles generic.
    # Its BERT grids follow the second First Name field, while the EMT grids are
    # explicitly marked Weekdays / Weekend Days / Weekend Nights.
    idx_bert_first = _find_nth_header(headers, "First Name", 1)

    for i, h in enumerate(headers):
        hh = (h or "").strip()
        m = _BRACKET_DATE.search(hh)
        if not m:
            continue
        d = _header_date(int(m.group(1)), int(m.group(2)), block_start, block_end)
        if d is None:
            continue
        low = hh.lower()
        # Legacy form: one grid per shift type.
        if low.startswith("day shifts"):
            emt_day[i] = d
        elif low.startswith("night shifts"):
            emt_night[i] = d
        elif low.startswith("weekend day"):
            emt_weekend[i] = d
        # Current form (FORM_GUIDE.md): one grid per week, all shifts as columns.
        elif "ambulance availability" in low:
            emt_week[i] = d
        elif "campus response availability" in low or ("availability" in low and "a/b" in low):
            bert[i] = d
        # Shopping-period form: every EMT grid is explicitly labelled with a
        # weekday/weekend qualifier, including the Sunday weekday-night grids.
        elif any(label in low for label in (
            "(weekdays)", "(weekday nights)", "(weekend nights)", "(weekend days)",
        )):
            emt_week[i] = d
        # Its Campus Response grids have the same generic title but appear in
        # the BERT section and contain A/B/C/D response values.
        elif idx_bert_first >= 0 and i > idx_bert_first:
            bert[i] = d

    # The EMT and BERT sections each have their own First/Last Name and
    # difficulties question; BERT's are the second occurrence of each.
    idx_emt_diff = find_contains(COL_DIFFICULTIES)
    return {
        "emt_day": emt_day,
        "emt_night": emt_night,
        "emt_weekend": emt_weekend,
        "emt_week": emt_week,
        "bert": bert,
        "idx_role": find_contains(COL_ROLE),
        "idx_email": find(COL_EMAIL) if find(COL_EMAIL) >= 0 else find(COL_EMAIL_FALLBACK),
        "idx_ts": max(find_contains(COL_TIMESTAMP), 0),
        "idx_driver": find_contains(COL_DRIVER),
        "idx_emt_first": _find_nth_header(headers, "First Name", 0),
        "idx_emt_last": _find_nth_header(headers, "Last Name", 0),
        "idx_emt_diff": idx_emt_diff,
        "idx_bert_first": idx_bert_first,
        "idx_bert_last": _find_nth_header(headers, "Last Name", 1),
        "idx_bert_diff": find_contains(COL_DIFFICULTIES, after=idx_emt_diff),
    }


# ── Freeform blackout parsing ────────────────────────────────────────────────

def parse_blackouts(raw: str, year: int) -> tuple[set, set]:
    """
    Parse the freeform "difficulties" answer into (slots, whole_days).
    Understands "5/3", "5/3-5/6", and shift/block tokens next to dates,
    e.g. "5/3 NIGHT; 5/6 AM" or "5/4 A B". Entries split on ';' or newlines.
    """
    slots: set = set()
    days: set = set()
    if not raw or raw.strip().upper() in ("N/A", "NA", ""):
        return slots, days

    for entry in re.split(r"[;\n]+", raw):
        entry = entry.strip()
        if not entry:
            continue
        shifts = [t.upper() for t in re.findall(r"\b(AM|PM|NIGHT|DAY|[ABCD])\b", entry, re.IGNORECASE)]
        dates: list[date] = []
        range_m = re.search(r"(\d{1,2})/(\d{1,2})\s*[-–]\s*(\d{1,2})/(\d{1,2})", entry)
        if range_m:
            m1, d1, m2, d2 = map(int, range_m.groups())
            try:
                cur, end = date(year, m1, d1), date(year, m2, d2)
            except ValueError:
                continue
            while cur <= end:
                dates.append(cur)
                cur += timedelta(days=1)
        else:
            for m, d in re.findall(r"(\d{1,2})/(\d{1,2})", entry):
                try:
                    dates.append(date(year, int(m), int(d)))
                except ValueError:
                    continue
        for dt in dates:
            if shifts:
                slots.update((dt, s) for s in shifts)
            else:
                days.add(dt)
    return slots, days


# ── Row expansion ────────────────────────────────────────────────────────────

def _expand_emt_row(row: list[str], maps: dict) -> set:
    available: set = set()
    for idx, d in maps["emt_day"].items():
        available.update((d, s) for s in _shift_tokens(_safe(row, idx)) if s in ("AM", "PM"))
    for idx, d in maps["emt_night"].items():
        if "NIGHT" in _shift_tokens(_safe(row, idx)):
            available.add((d, "NIGHT"))
    for idx, d in maps["emt_weekend"].items():
        if "DAY" in _shift_tokens(_safe(row, idx)):
            available.add((d, "DAY"))
    for idx, d in maps["emt_week"].items():
        valid = ("DAY", "NIGHT") if is_weekend(d) else ("AM", "PM", "NIGHT")
        available.update((d, s) for s in _shift_tokens(_safe(row, idx)) if s in valid)
    return available


def _apply_blackouts(available: set, blackout_slots: set, blackout_dates: set,
                     block_start: date, block_end: date) -> None:
    for bd in blackout_dates:
        if block_start <= bd <= block_end:
            for s in shift_types_for(bd):
                available.discard((bd, s))
    for key in blackout_slots:
        available.discard(key)


def infer_campus_availability(v: Volunteer) -> set:
    """EMTs' campus availability: AM availability -> blocks A,B; PM -> C,D."""
    result = set()
    for (d, s) in v.available:
        if is_weekend(d):
            continue
        if s == "AM":
            result.update({(d, "A"), (d, "B")})
        elif s == "PM":
            result.update({(d, "C"), (d, "D")})
    for bd in v.blackout_dates:
        for b in CAMPUS_BLOCKS:
            result.discard((bd, b))
    for (d, s) in v.blackout_slots:
        if s == "AM" or s == "DAY":
            result -= {(d, "A"), (d, "B")}
        if s == "PM" or s == "DAY":
            result -= {(d, "C"), (d, "D")}
        if s in CAMPUS_BLOCKS:
            result.discard((d, s))
    return result


# ── Entry point ──────────────────────────────────────────────────────────────

def load_all_responses(
    csv_path: str,
    block_start: date,
    block_end: date,
) -> tuple[list[Volunteer], list[BertMember]]:
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise ValueError("CSV file is empty.")

    headers, data_rows = rows[0], rows[1:]
    maps = _build_column_maps(headers, block_start, block_end)
    year = block_start.year

    # Latest submission per email, split by role.
    latest_emt: dict[str, tuple[datetime, list[str]]] = {}
    latest_bert: dict[str, tuple[datetime, list[str]]] = {}
    for row in data_rows:
        email = _safe(row, maps["idx_email"]).lower()
        if not email:
            continue
        ts = _parse_timestamp(_safe(row, maps["idx_ts"]))
        role = _safe(row, maps["idx_role"])
        bucket = latest_bert if _is_bert_role(role) else latest_emt if _is_emt_role(role) else None
        if bucket is not None and (email not in bucket or ts > bucket[email][0]):
            bucket[email] = (ts, row)

    volunteers: list[Volunteer] = []
    for email, (_, row) in latest_emt.items():
        blackout_slots, blackout_dates = parse_blackouts(_safe(row, maps["idx_emt_diff"]), year)
        available = _expand_emt_row(row, maps)
        _apply_blackouts(available, blackout_slots, blackout_dates, block_start, block_end)
        v = Volunteer(
            first_name=_safe(row, maps["idx_emt_first"]),
            last_name=_safe(row, maps["idx_emt_last"]),
            email=email,
            certification=normalise_driver(_safe(row, maps["idx_driver"])),
            available=available,
            blackout_slots=blackout_slots,
            blackout_dates=blackout_dates,
        )
        v.campus_available = infer_campus_availability(v)
        volunteers.append(v)

    bert_members: list[BertMember] = []
    for email, (_, row) in latest_bert.items():
        blackout_slots, blackout_dates = parse_blackouts(_safe(row, maps["idx_bert_diff"]), year)
        campus_available = set()
        for idx, d in maps["bert"].items():
            campus_available.update((d, b) for b in _block_tokens(_safe(row, idx)))
        for bd in blackout_dates:
            for b in CAMPUS_BLOCKS:
                campus_available.discard((bd, b))
        for (d, tok) in blackout_slots:
            if tok in CAMPUS_BLOCKS:
                campus_available.discard((d, tok))
        bert_members.append(BertMember(
            first_name=_safe(row, maps["idx_bert_first"]),
            last_name=_safe(row, maps["idx_bert_last"]),
            email=email,
            campus_available=campus_available,
            blackout_slots=blackout_slots,
            blackout_dates=blackout_dates,
        ))

    print(f"  Loaded {len(volunteers)} Ambulance EMT volunteers from form.")
    print(f"  Loaded {len(bert_members)} BERT members from form.")
    return volunteers, bert_members
