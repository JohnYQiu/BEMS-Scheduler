"""
parse_form.py
=============
Reads the Google Form CSV export and returns Volunteer / BertMember lists.
Handles per-block date columns, duplicate submissions (latest timestamp wins),
driver status → EVDT / Auth / EMT, and campus block availability for BERT.
"""

import csv
import re
from datetime import date, timedelta, datetime
from dataclasses import dataclass, field
from typing import Optional, Tuple, Dict, List, Set, Any


COL_ROLE = "Are you an ambulance EMT or BERT member?"
COL_EMAIL = "Username"
COL_TIMESTAMP = "Timestamp"
COL_DRIVER = "Driver Status"

# Ambulance EMT role substring (Google Form text may change slightly)
ROLE_EMT_MARKERS = ("ambulance emt", "emt only", "dual-role")
ROLE_BERT_MARKERS = ("bert member", "bert only")

CAMPUS_BLOCKS = ("A", "B", "C", "D")
CAMPUS_BLOCK_HOURS = 3

SHIFT_HOURS = {"AM": 6, "PM": 6, "NIGHT": 12, "DAY": 12}


@dataclass
class Volunteer:
    first_name:      str
    last_name:       str
    email:           str
    certification:   str
    available:       set = field(default_factory=set)
    blackout_slots:  set = field(default_factory=set)
    blackout_dates:  set = field(default_factory=set)
    scheduled_hours: int = 0
    scheduled_shifts: list = field(default_factory=list)
    campus_available: set = field(default_factory=set)
    campus_scheduled_hours: int = 0
    campus_scheduled_shifts: list = field(default_factory=list)

    @property
    def full_name(self): return f"{self.first_name} {self.last_name}"
    @property
    def is_evdt(self): return self.certification == "EVDT"
    @property
    def is_auth(self): return self.certification in ("EVDT", "Auth")


@dataclass
class BertMember:
    first_name:      str
    last_name:       str
    email:           str
    certification:   str = "BERT"  # for reports / hour summary (BERT-only, no ambulance driver level)
    campus_available: set = field(default_factory=set)
    blackout_slots:  set = field(default_factory=set)
    blackout_dates:  set = field(default_factory=set)
    campus_scheduled_hours: int = 0
    campus_scheduled_shifts: list = field(default_factory=list)

    @property
    def full_name(self): return f"{self.first_name} {self.last_name}"


def normalise_driver(raw: str) -> str:
    r = raw.strip().upper()
    if "EVDT" in r:
        return "EVDT"
    if "AUTHORIZED" in r or "AUTH" in r:
        return "Auth"
    if "NOT A DRIVER" in r or r == "N/A":
        return "EMT"
    return "EMT"


_DATE_IN_HEADER = re.compile(
    r"Please indicate your availability for the below dates and shifts\.\s*\[(\d{1,2})/(\d{1,2})"
)
_SAT_HEADER = re.compile(
    r"Please indicate your availability for Saturdays during this block\.\s*\[(\d{1,2})/(\d{1,2})\]"
)
_SUN_HEADER = re.compile(
    r"Please indicate your availability for Sundays during this block\.\s*\[(\d{1,2})/(\d{1,2})\]"
)


def _parse_timestamp(ts_str: str) -> datetime:
    s = re.sub(r"\s+[A-Z]{2,4}\s*$", "", (ts_str or "").strip())
    for fmt in ("%Y/%m/%d %I:%M:%S %p", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return datetime.min


def _header_date_to_date(month: int, day: int, block_start: date, block_end: date) -> Optional[date]:
    y = block_start.year
    try:
        d = date(y, month, day)
    except ValueError:
        return None
    if d < block_start - timedelta(days=60) or d > block_end + timedelta(days=60):
        try:
            d = date(y + 1, month, day)
        except ValueError:
            return None
    if block_start <= d <= block_end:
        return d
    return None


def _parse_shifts_from_cell(cell: str) -> list[str]:
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", "no", ""):
        return []
    found: list[str] = []
    for token in re.findall(r"\b(AM|PM|NIGHT|DAY)\b", cell, re.IGNORECASE):
        s = token.upper()
        if s in SHIFT_HOURS and s not in found:
            found.append(s)
    return found


def _parse_blocks_from_cell(cell: str) -> list[str]:
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", "no", ""):
        return []
    found: list[str] = []
    for token in re.findall(r"\b([ABCD])\b", cell, re.IGNORECASE):
        b = token.upper()
        if b in CAMPUS_BLOCKS and b not in found:
            found.append(b)
    return found


def _build_column_maps(headers: List[str], block_start: date, block_end: date) -> Dict[str, Any]:
    """Return dicts: emt_col -> (date, shift_types), sat/sun column lists."""
    emt_weekday: Dict[int, Tuple[date, Tuple[str, ...]]] = {}
    emt_sat: list[tuple[int, date]] = []
    emt_sun: list[tuple[int, date]] = []

    bert_weekday: dict[int, tuple[date, tuple[str, ...]]] = {}

    for i, h in enumerate(headers):
        if (m := _DATE_IN_HEADER.search(h or "")):
            mo, dy = int(m.group(1)), int(m.group(2))
            d = _header_date_to_date(mo, dy, block_start, block_end)
            if d is None:
                continue
            if d.weekday() < 5:
                emt_weekday[i] = (d, ("AM", "PM", "NIGHT"))
            else:
                emt_weekday[i] = (d, ("DAY", "NIGHT"))
        elif (m := _SAT_HEADER.search(h or "")):
            mo, dy = int(m.group(1)), int(m.group(2))
            d = _header_date_to_date(mo, dy, block_start, block_end)
            if d is not None:
                emt_sat.append((i, d))
        elif (m := _SUN_HEADER.search(h or "")):
            mo, dy = int(m.group(1)), int(m.group(2))
            d = _header_date_to_date(mo, dy, block_start, block_end)
            if d is not None:
                emt_sun.append((i, d))

    idx_role = next((i for i, h in enumerate(headers) if h and COL_ROLE in h), -1)
    idx_driver = next((i for i, h in enumerate(headers) if h and COL_DRIVER in h), -1)
    idx_bert_last = _find_nth_header(headers, "Last Name", 1)

    for i, h in enumerate(headers):
        if i <= idx_bert_last or idx_bert_last < 0:
            continue
        if (m := _DATE_IN_HEADER.search(h or "")):
            mo, dy = int(m.group(1)), int(m.group(2))
            d = _header_date_to_date(mo, dy, block_start, block_end)
            if d is None or d.weekday() >= 5:
                continue
            bert_weekday[i] = (d, ("A", "B", "C", "D"))

    return {
        "emt_weekday": emt_weekday,
        "emt_sat": emt_sat,
        "emt_sun": emt_sun,
        "bert_weekday": bert_weekday,
        "idx_role": idx_role,
        "idx_email": next((i for i, h in enumerate(headers) if h and COL_EMAIL in h), 1),
        "idx_ts": next((i for i, h in enumerate(headers) if h and COL_TIMESTAMP in h), 0),
        "idx_driver": idx_driver,
        "idx_emt_first": next((i for i, h in enumerate(headers) if h and h.strip() == "First Name"), -1),
        "idx_emt_last": next((i for i, h in enumerate(headers) if h and h.strip() == "Last Name"), -1),
        "idx_bert_first": _find_nth_header(headers, "First Name", 1),
        "idx_bert_last": idx_bert_last,
    }


def _find_nth_header(headers: list[str], exact: str, n: int) -> int:
    count = 0
    for i, h in enumerate(headers):
        if (h or "").strip() == exact:
            if count == n:
                return i
            count += 1
    return -1


def _is_emt_role(role: str) -> bool:
    r = (role or "").lower()
    return any(x in r for x in ROLE_EMT_MARKERS)


def _is_bert_role(role: str) -> bool:
    r = (role or "").lower()
    return any(x in r for x in ROLE_BERT_MARKERS)


def _expand_emt_row(
    row: list[str],
    maps: dict,
    block_start: date,
    block_end: date,
) -> set:
    available: set = set()
    for col_idx, (d, _allowed) in maps["emt_weekday"].items():
        cell = _safe(row, col_idx)
        for s in _parse_shifts_from_cell(cell):
            if (d, s) not in available:
                available.add((d, s))

    for col_idx, d in maps["emt_sat"]:
        cell = _safe(row, col_idx)
        for s in _parse_shifts_from_cell(cell):
            if d.weekday() == 5 and s in ("DAY", "NIGHT"):
                available.add((d, s))

    for col_idx, d in maps["emt_sun"]:
        cell = _safe(row, col_idx)
        for s in _parse_shifts_from_cell(cell):
            if d.weekday() == 6 and s in ("DAY", "NIGHT"):
                available.add((d, s))

    return available


def _expand_bert_row(row: list[str], maps: dict, block_start: date, block_end: date) -> set:
    result = set()
    for col_idx, (d, _blocks) in maps["bert_weekday"].items():
        cell = _safe(row, col_idx)
        for b in _parse_blocks_from_cell(cell):
            result.add((d, b))
    return result


def _safe(row: list[str], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    return (row[idx] or "").strip()


def parse_blackouts(raw: str, year: int) -> tuple[set, set]:
    """Same freeform parser as before; supports AM/PM/NIGHT/DAY and A–D."""
    slots: set = set()
    days: set = set()

    if not raw or raw.strip().upper() in ("N/A", "NA", ""):
        return slots, days

    shift_map = {"AM": "AM", "PM": "PM", "NIGHT": "NIGHT", "DAY": "DAY",
                 "A": "A", "B": "B", "C": "C", "D": "D"}

    entries = re.split(r"[;\n]+", raw)
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        found_shifts = []
        for token in re.findall(r"\b(AM|PM|NIGHT|DAY|A|B|C|D)\b", entry, re.IGNORECASE):
            s = shift_map.get(token.upper())
            if s:
                found_shifts.append(s)

        date_matches = re.findall(r"(\d{1,2})/(\d{1,2})", entry)
        if not date_matches:
            continue

        range_match = re.search(
            r"(\d{1,2})/(\d{1,2})\s*[-–]\s*(\d{1,2})/(\d{1,2})", entry
        )
        if range_match:
            m1, d1, m2, d2 = map(int, range_match.groups())
            start = date(year, m1, d1)
            end = date(year, m2, d2)
            current = start
            while current <= end:
                if found_shifts:
                    for s in found_shifts:
                        slots.add((current, s))
                else:
                    days.add(current)
                current += timedelta(days=1)
        else:
            for m, d in date_matches:
                try:
                    dt = date(year, int(m), int(d))
                except ValueError:
                    continue
                if found_shifts:
                    for s in found_shifts:
                        slots.add((dt, s))
                else:
                    days.add(dt)

    return slots, days


def expand_campus_availability(
    weekly_blocks: dict,
    block_start: date,
    block_end: date,
    blackout_slots: set,
    blackout_dates: set,
) -> set:
    DOW_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    result = set()
    current = block_start
    while current <= block_end:
        dow_name = DOW_NAMES[current.weekday()]
        if current.weekday() < 5:
            for b in weekly_blocks.get(dow_name, []):
                key = (current, b)
                if current not in blackout_dates and key not in blackout_slots:
                    result.add(key)
        current += timedelta(days=1)
    return result


def infer_campus_availability_for_ambulance(v: Volunteer) -> set:
    result = set()
    for (d, s) in v.available:
        if d.weekday() >= 5:
            continue
        if s == "AM":
            result.add((d, "A"))
            result.add((d, "B"))
        elif s == "PM":
            result.add((d, "C"))
            result.add((d, "D"))

    for bd in v.blackout_dates:
        if bd.weekday() < 5:
            for b in CAMPUS_BLOCKS:
                result.discard((bd, b))

    for (d, s) in v.blackout_slots:
        if d.weekday() >= 5:
            continue
        if s == "AM":
            result.discard((d, "A"))
            result.discard((d, "B"))
        elif s == "PM":
            result.discard((d, "C"))
            result.discard((d, "D"))
        elif s == "DAY":
            for b in CAMPUS_BLOCKS:
                result.discard((d, b))

    return result


def load_all_responses(
    csv_path: str,
    block_start: date,
    block_end: date,
) -> tuple[list[Volunteer], list[BertMember]]:
    year = block_start.year

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        raw_rows = list(reader)

    if not raw_rows:
        raise ValueError("CSV file is empty.")

    headers = raw_rows[0]
    data_rows = raw_rows[1:]
    maps = _build_column_maps(headers, block_start, block_end)

    latest_ambulance: dict[str, tuple] = {}
    latest_bert: dict[str, tuple] = {}

    idx_role = maps["idx_role"]
    idx_email = maps["idx_email"]
    idx_ts = maps["idx_ts"]

    for row in data_rows:
        role = _safe(row, idx_role)
        email = _safe(row, idx_email).lower()
        if not email:
            continue
        ts = _parse_timestamp(_safe(row, idx_ts))

        if _is_bert_role(role):
            if email not in latest_bert or ts > latest_bert[email][0]:
                latest_bert[email] = (ts, row)
        elif _is_emt_role(role):
            if email not in latest_ambulance or ts > latest_ambulance[email][0]:
                latest_ambulance[email] = (ts, row)

    volunteers: list[Volunteer] = []
    idx_first = maps["idx_emt_first"]
    idx_last = maps["idx_emt_last"]
    idx_driver = maps["idx_driver"]
    idx_emt_difficulties = next(
        (i for i, h in enumerate(headers) if h and "Do you foresee" in h),
        -1,
    )

    for email, (_, row) in latest_ambulance.items():
        first = _safe(row, idx_first)
        last = _safe(row, idx_last)
        cert = normalise_driver(_safe(row, idx_driver))
        available = _expand_emt_row(row, maps, block_start, block_end)
        blackout_raw = _safe(row, idx_emt_difficulties)
        blackout_slots, blackout_dates = parse_blackouts(blackout_raw, year)

        for bd in list(blackout_dates):
            if block_start <= bd <= block_end:
                for s in (["DAY", "NIGHT"] if bd.weekday() >= 5 else ["AM", "PM", "NIGHT"]):
                    available.discard((bd, s))

        for (bd, s) in list(blackout_slots):
            available.discard((bd, s))

        v = Volunteer(
            first_name=first,
            last_name=last,
            email=email,
            certification=cert,
            available=available,
            blackout_slots=blackout_slots,
            blackout_dates=blackout_dates,
        )
        v.campus_available = infer_campus_availability_for_ambulance(v)
        volunteers.append(v)

    bert_members: list[BertMember] = []
    idx_bf = maps["idx_bert_first"]
    idx_bl = maps["idx_bert_last"]

    idx_bert_difficulties = next(
        (i for i in range(idx_bl + 1, len(headers))
         if headers[i] and "Do you foresee" in headers[i]),
        len(headers) - 1,
    )

    for email, (_, row) in latest_bert.items():
        first = _safe(row, idx_bf)
        last = _safe(row, idx_bl)
        campus_available = _expand_bert_row(row, maps, block_start, block_end)
        blackout_raw = _safe(row, idx_bert_difficulties)
        blackout_slots, blackout_dates = parse_blackouts(blackout_raw, year)

        for bd in blackout_dates:
            if block_start <= bd <= block_end and bd.weekday() < 5:
                for b in CAMPUS_BLOCKS:
                    campus_available.discard((bd, b))
        for (bd, tok) in blackout_slots:
            if tok in CAMPUS_BLOCKS:
                campus_available.discard((bd, tok))

        b = BertMember(
            first_name=first,
            last_name=last,
            email=email,
            campus_available=campus_available,
            blackout_slots=blackout_slots,
            blackout_dates=blackout_dates,
        )
        bert_members.append(b)

    print(f"  Loaded {len(volunteers)} Ambulance EMT volunteers from form.")
    print(f"  Loaded {len(bert_members)} BERT members from form.")
    return volunteers, bert_members


def load_responses(csv_path: str, block_start: date, block_end: date) -> list[Volunteer]:
    volunteers, _ = load_all_responses(csv_path, block_start, block_end)
    return volunteers
