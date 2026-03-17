"""
parse_form.py
=============
Reads the Google Form CSV export and returns a list of Volunteer objects.
Handles:
  - Parsing Ambulance EMT and BERT submissions
  - Duplicate submissions (keeps most recent by timestamp)
  - Normalising driver status → EVDT / Auth / EMT
  - Expanding day-of-week availability into concrete (date, shift/block) tuples
  - Parsing freeform blackout date strings for both ambulance and campus blocks
"""

import csv
import re
from datetime import date, timedelta, datetime
from dataclasses import dataclass, field


# ── Column indices (0-based) in the real form CSV ─────────────────────────────
# Row 0 is the merged multi-line header; data rows start at index 1 (row 19+)
# We identify columns by partial header matching rather than fixed indices
# so the code stays robust if columns shift slightly.

COL_TIMESTAMP   = "Timestamp"
COL_EMAIL       = "Email Address"
COL_ROLE        = "Are you an ambulance EMT or BERT member?"
COL_LAST_NAME   = "Last Name"       # first occurrence = EMT section
COL_FIRST_NAME  = "First Name"      # first occurrence = EMT section
COL_DRIVER      = "Driver Status"

# Weekday availability columns (EMT section) — matched by substring
WEEKDAY_COLS = {
    "Sunday":    ("NIGHT",),          # Sunday column only has Night for EMTs
    "Monday":    ("AM", "PM", "NIGHT"),
    "Tuesday":   ("AM", "PM", "NIGHT"),
    "Wednesday": ("AM", "PM", "NIGHT"),
    "Thursday":  ("AM", "PM", "NIGHT"),
    "Friday":    ("AM", "PM"),        # Friday AM/PM handled separately from Friday NIGHT
}

# Special weekend columns — identified by partial name
WEEKEND_COL_MAP = {
    "Friday NIGHT":    ("Friday", "NIGHT"),
    "Saturday DAY":    ("Saturday", "DAY"),
    "Saturday NIGHT":  ("Saturday", "NIGHT"),
    "Sunday DAY":      ("Sunday", "DAY"),
}

COL_BLACKOUT = "Enter dates and shifts that you know you cannot work"

ROLE_EMT  = "Ambulance EMT"
ROLE_BERT = "BERT Member Only"

# Campus response blocks (weekdays only)
CAMPUS_BLOCKS = ("A", "B", "C", "D")  # 07–10, 10–13, 13–16, 16–19
CAMPUS_BLOCK_HOURS = 3


# ── Shift hours lookup ────────────────────────────────────────────────────────
SHIFT_HOURS = {"AM": 6, "PM": 6, "NIGHT": 12, "DAY": 12}


# ── Data class ────────────────────────────────────────────────────────────────
@dataclass
class Volunteer:
    first_name:      str
    last_name:       str
    email:           str
    certification:   str                        # EVDT | Auth | EMT
    available:       set = field(default_factory=set)   # {(date, shift), ...}
    blackout_slots:  set = field(default_factory=set)   # {(date, shift), ...}
    blackout_dates:  set = field(default_factory=set)   # whole days off {date, ...}
    scheduled_hours: int = 0
    scheduled_shifts: list = field(default_factory=list)
    campus_available: set = field(default_factory=set)  # {(date, block), ...} block ∈ A/B/C/D
    campus_scheduled_hours: int = 0
    campus_scheduled_shifts: list = field(default_factory=list)  # {(date, block), ...}

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
    campus_available: set = field(default_factory=set)  # {(date, block), ...}
    blackout_slots:  set = field(default_factory=set)  # {(date, block), ...} for campus blocks
    blackout_dates:  set = field(default_factory=set)  # whole days off {date, ...}
    campus_scheduled_hours: int = 0
    campus_scheduled_shifts: list = field(default_factory=list)

    @property
    def full_name(self): return f"{self.first_name} {self.last_name}"


# ── Driver status normaliser ──────────────────────────────────────────────────
def normalise_driver(raw: str) -> str:
    r = raw.strip().upper()
    if "EVDT" in r:
        return "EVDT"
    if "AUTHORIZED" in r or "AUTH" in r:
        return "Auth"
    return "EMT"


# ── Blackout parser ───────────────────────────────────────────────────────────
def parse_blackouts(raw: str, year: int) -> tuple[set, set]:
    """
    Parses the freeform blackout field.
    Returns (blackout_slots, blackout_dates) where:
      blackout_slots = {(date, shift), ...}   shift-specific blocks
      blackout_dates = {date, ...}            whole-day blocks
    Handles formats seen in real data:
      "3/14, AM"  "3/14, AM/PM/NIGHT"  "3/20-3/21 AM/PM/NIGHT"
      "03/12, PM" "3/13 Night"  "N/A"  blank
    """
    slots: set = set()
    days:  set = set()

    if not raw or raw.strip().upper() in ("N/A", "NA", ""):
        return slots, days

    shift_map = {"AM": "AM", "PM": "PM", "NIGHT": "NIGHT", "DAY": "DAY",
                 "A": "A", "B": "B", "C": "C", "D": "D"}

    # Split on semicolons or newlines
    entries = re.split(r"[;\n]+", raw)
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        # Look for shift keywords in this entry
        found_shifts = []
        for token in re.findall(r"\b(AM|PM|NIGHT|DAY|A|B|C|D)\b", entry, re.IGNORECASE):
            s = shift_map.get(token.upper())
            if s:
                found_shifts.append(s)

        # Extract date range or single date — patterns: M/D or MM/DD
        date_matches = re.findall(r"(\d{1,2})/(\d{1,2})", entry)
        if not date_matches:
            continue

        # Check for range: "3/20-3/21"
        range_match = re.search(
            r"(\d{1,2})/(\d{1,2})\s*[-–]\s*(\d{1,2})/(\d{1,2})", entry
        )
        if range_match:
            m1, d1, m2, d2 = map(int, range_match.groups())
            start = date(year, m1, d1)
            end   = date(year, m2, d2)
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


# ── Day-of-week availability expander ─────────────────────────────────────────
def _parse_shifts_from_cell(cell: str) -> list[str]:
    """Extract shift names from a cell like 'AM (0700-1300), PM (1300-1900)'."""
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", ""):
        return []
    found = []
    for token in re.findall(r"\b(AM|PM|NIGHT|DAY)\b", cell, re.IGNORECASE):
        s = token.upper()
        if s in SHIFT_HOURS and s not in found:
            found.append(s)
    return found


def _parse_blocks_from_cell(cell: str) -> list[str]:
    """Extract campus blocks from a cell like 'C (1300-1600), D (1600-1900)'."""
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", "no", ""):
        return []
    found: list[str] = []
    for token in re.findall(r"\b([ABCD])\b", cell, re.IGNORECASE):
        b = token.upper()
        if b in CAMPUS_BLOCKS and b not in found:
            found.append(b)
    return found


def _parse_specific_dates(cell: str, year: int) -> list:
    """
    Parse a weekend column cell containing specific available dates.
    Handles: "3/14", "3/14, 4/4", "Not Available", blank.
    Returns a list of date objects.
    """
    if not cell or cell.strip().lower() in ("not available", "n/a", "na", "no", ""):
        return []
    dates = []
    for m, d in re.findall(r"(\d{1,2})/(\d{1,2})", cell):
        try:
            dates.append(date(year, int(m), int(d)))
        except ValueError:
            pass
    return dates

def expand_availability(
    weekly_avail: dict,        # {"Monday": ["AM","PM"], "Saturday": ["DAY"], ...}
    block_start: date,
    block_end:   date,
    blackout_slots: set,
    blackout_dates: set,
) -> set:
    """
    Convert recurring weekly availability into concrete (date, shift) tuples
    across the block, then subtract blackouts.
    weekly_avail keys: Monday/Tuesday/Wednesday/Thursday/Friday/Saturday/Sunday
    """
    DOW_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    result = set()
    current = block_start
    while current <= block_end:
        dow_name = DOW_NAMES[current.weekday()]
        shifts = weekly_avail.get(dow_name, [])
        for shift in shifts:
            key = (current, shift)
            if current not in blackout_dates and key not in blackout_slots:
                result.add(key)
        current += timedelta(days=1)
    return result


def expand_campus_availability(
    weekly_blocks: dict,        # {"Monday": ["A","B"], ...}
    block_start: date,
    block_end:   date,
    blackout_slots: set,
    blackout_dates: set,
) -> set:
    """Expand recurring weekday campus blocks to concrete (date, block) tuples."""
    DOW_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    result = set()
    current = block_start
    while current <= block_end:
        dow_name = DOW_NAMES[current.weekday()]
        if current.weekday() < 5:
            blocks = weekly_blocks.get(dow_name, [])
            for b in blocks:
                key = (current, b)
                if current not in blackout_dates and key not in blackout_slots:
                    result.add(key)
        current += timedelta(days=1)
    return result


def infer_campus_availability_for_ambulance(v: Volunteer) -> set:
    """
    Infer campus responder availability from ambulance availability:
      - AM => A + B
      - PM => C + D
    Apply ambulance blackouts to the inferred campus blocks.
    """
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

    # Translate ambulance blackout slots/dates into campus blocks
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


# ── Column finder ─────────────────────────────────────────────────────────────
def _find_col(headers: list[str], keyword: str, occurrence: int = 0) -> int:
    """Return index of the `occurrence`-th header containing `keyword`."""
    count = 0
    for i, h in enumerate(headers):
        if keyword.lower() in h.lower():
            if count == occurrence:
                return i
            count += 1
    return -1


# ── Main loader ───────────────────────────────────────────────────────────────
def load_all_responses(
    csv_path: str,
    block_start: date,
    block_end:   date,
) -> tuple[list[Volunteer], list[BertMember]]:
    """Read the Google Form CSV and return (ambulance_volunteers, bert_members)."""
    year = block_start.year

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        raw_rows = list(reader)

    if not raw_rows:
        raise ValueError("CSV file is empty.")

    # Row 0 is the header (multi-line cells collapsed by csv reader)
    headers = raw_rows[0]
    data_rows = raw_rows[1:]

    # Locate key column indices
    idx_ts      = _find_col(headers, "Timestamp")
    idx_email   = _find_col(headers, "Email Address")
    idx_role    = _find_col(headers, "Are you an ambulance EMT")
    idx_last    = _find_col(headers, "Last Name",  occurrence=0)
    idx_first   = _find_col(headers, "First Name", occurrence=0)
    idx_driver  = _find_col(headers, "Driver Status")
    idx_blackout = _find_col(headers, "Enter dates and shifts that you know you cannot work", occurrence=0)
    idx_blackout_bert = _find_col(headers, "Enter dates and shifts that you know you cannot work", occurrence=1)

    # Weekday columns: Sunday–Friday (general AM/PM/NIGHT)
    idx_sunday    = _find_col(headers, "Sunday",    occurrence=0)
    idx_monday    = _find_col(headers, "Monday",    occurrence=0)
    idx_tuesday   = _find_col(headers, "Tuesday",   occurrence=0)
    idx_wednesday = _find_col(headers, "Wednesday", occurrence=0)
    idx_thursday  = _find_col(headers, "Thursday",  occurrence=0)
    idx_friday    = _find_col(headers, "Friday",    occurrence=0)

    # Weekend / special columns
    idx_fri_night  = _find_col(headers, "Friday NIGHT")
    idx_sat_day    = _find_col(headers, "Saturday DAY")
    idx_sat_night  = _find_col(headers, "Saturday NIGHT")
    idx_sun_day    = _find_col(headers, "Sunday DAY")

    # BERT section columns (campus responder availability)
    idx_last_bert  = _find_col(headers, "Last Name",  occurrence=1)
    idx_first_bert = _find_col(headers, "First Name", occurrence=1)
    idx_mon_bert   = _find_col(headers, "Monday",    occurrence=1)
    idx_tue_bert   = _find_col(headers, "Tuesday",   occurrence=1)
    idx_wed_bert   = _find_col(headers, "Wednesday", occurrence=1)
    idx_thu_bert   = _find_col(headers, "Thursday",  occurrence=1)
    idx_fri_bert   = _find_col(headers, "Friday",    occurrence=1)

    def safe(row, idx):
        if idx < 0 or idx >= len(row):
            return ""
        return row[idx].strip()

    # ── Dedup: keep most recent submission per email (separately per role) ─
    latest_ambulance: dict[str, tuple] = {}  # email → (timestamp, row)
    latest_bert: dict[str, tuple] = {}       # email → (timestamp, row)
    for row in data_rows:
        role = safe(row, idx_role)
        email = safe(row, idx_email).lower()
        if not email:
            continue

        ts_str = safe(row, idx_ts)
        try:
            ts = datetime.strptime(ts_str, "%m/%d/%Y %H:%M:%S")
        except ValueError:
            ts = datetime.min

        if ROLE_BERT.lower() in role.lower():
            if email not in latest_bert or ts > latest_bert[email][0]:
                latest_bert[email] = (ts, row)
        elif ROLE_EMT.lower() in role.lower():
            if email not in latest_ambulance or ts > latest_ambulance[email][0]:
                latest_ambulance[email] = (ts, row)

    volunteers: list[Volunteer] = []
    for email, (_, row) in latest_ambulance.items():
        # Basic info
        first = safe(row, idx_first)
        last  = safe(row, idx_last)
        driver_raw = safe(row, idx_driver)
        cert  = normalise_driver(driver_raw)

        # Blackouts
        blackout_raw = safe(row, idx_blackout)
        blackout_slots, blackout_dates = parse_blackouts(blackout_raw, year)

        # Weekly availability — build dict of day → [shifts]
        weekly: dict[str, list] = {}

        # Sunday (EMT section only has Night for Sunday weekday; DAY is in Sunday DAY col)
        sun_cell = safe(row, idx_sunday)
        weekly["Sunday"] = _parse_shifts_from_cell(sun_cell)

        weekly["Monday"]    = _parse_shifts_from_cell(safe(row, idx_monday))
        weekly["Tuesday"]   = _parse_shifts_from_cell(safe(row, idx_tuesday))
        weekly["Wednesday"] = _parse_shifts_from_cell(safe(row, idx_wednesday))
        weekly["Thursday"]  = _parse_shifts_from_cell(safe(row, idx_thursday))
        weekly["Friday"]    = _parse_shifts_from_cell(safe(row, idx_friday))

        # Expand weekday recurring availability to concrete dates
        available = expand_availability(
            weekly, block_start, block_end, blackout_slots, blackout_dates
        )

        # Weekend columns contain specific dates, not recurring weekly flags.
        # Parse each cell as an explicit list of dates and add directly.

        # Friday NIGHT — specific dates volunteer is available
        for d in _parse_specific_dates(safe(row, idx_fri_night), year):
            key = (d, "NIGHT")
            if d not in blackout_dates and key not in blackout_slots:
                available.add(key)

        # Saturday DAY — specific dates
        for d in _parse_specific_dates(safe(row, idx_sat_day), year):
            key = (d, "DAY")
            if d not in blackout_dates and key not in blackout_slots:
                available.add(key)

        # Saturday NIGHT — specific dates
        for d in _parse_specific_dates(safe(row, idx_sat_night), year):
            key = (d, "NIGHT")
            if d not in blackout_dates and key not in blackout_slots:
                available.add(key)

        # Sunday DAY — specific dates
        for d in _parse_specific_dates(safe(row, idx_sun_day), year):
            key = (d, "DAY")
            if d not in blackout_dates and key not in blackout_slots:
                available.add(key)

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
    for email, (_, row) in latest_bert.items():
        first = safe(row, idx_first_bert)
        last  = safe(row, idx_last_bert)

        blackout_raw = safe(row, idx_blackout_bert)
        blackout_slots, blackout_dates = parse_blackouts(blackout_raw, year)

        weekly_blocks: dict[str, list] = {}
        weekly_blocks["Monday"]    = _parse_blocks_from_cell(safe(row, idx_mon_bert))
        weekly_blocks["Tuesday"]   = _parse_blocks_from_cell(safe(row, idx_tue_bert))
        weekly_blocks["Wednesday"] = _parse_blocks_from_cell(safe(row, idx_wed_bert))
        weekly_blocks["Thursday"]  = _parse_blocks_from_cell(safe(row, idx_thu_bert))
        weekly_blocks["Friday"]    = _parse_blocks_from_cell(safe(row, idx_fri_bert))

        campus_available = expand_campus_availability(
            weekly_blocks, block_start, block_end, blackout_slots, blackout_dates
        )

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
    """Backwards-compatible: return only ambulance volunteers."""
    volunteers, _ = load_all_responses(csv_path, block_start, block_end)
    return volunteers