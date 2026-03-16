"""
parse_form.py
=============
Reads the Google Form CSV export and returns a list of Volunteer objects.
Handles:
  - Filtering out BERT members (only Ambulance EMTs are scheduled)
  - Duplicate submissions (keeps most recent by timestamp)
  - Normalising driver status → EVDT / Auth / EMT
  - Expanding day-of-week availability into concrete (date, shift) tuples
  - Parsing freeform blackout date strings
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

    @property
    def full_name(self): return f"{self.first_name} {self.last_name}"
    @property
    def is_evdt(self): return self.certification == "EVDT"
    @property
    def is_auth(self): return self.certification in ("EVDT", "Auth")


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

    shift_map = {"AM": "AM", "PM": "PM", "NIGHT": "NIGHT", "DAY": "DAY"}

    # Split on semicolons or newlines
    entries = re.split(r"[;\n]+", raw)
    for entry in entries:
        entry = entry.strip()
        if not entry:
            continue

        # Look for shift keywords in this entry
        found_shifts = []
        for token in re.findall(r"\b(AM|PM|NIGHT|DAY)\b", entry, re.IGNORECASE):
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
def load_responses(
    csv_path: str,
    block_start: date,
    block_end:   date,
) -> list[Volunteer]:
    """
    Read the Google Form CSV and return a list of Volunteer objects
    (Ambulance EMTs only, deduplicated by email keeping most recent).
    """
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

    def safe(row, idx):
        if idx < 0 or idx >= len(row):
            return ""
        return row[idx].strip()

    # ── Dedup: keep most recent submission per email ───────────────────────
    latest: dict[str, tuple] = {}  # email → (timestamp, row)
    for row in data_rows:
        role = safe(row, idx_role)
        if ROLE_BERT.lower() in role.lower():
            continue  # skip BERT entirely
        if ROLE_EMT.lower() not in role.lower():
            continue  # skip unknown roles

        email = safe(row, idx_email).lower()
        if not email:
            continue

        ts_str = safe(row, idx_ts)
        try:
            ts = datetime.strptime(ts_str, "%m/%d/%Y %H:%M:%S")
        except ValueError:
            ts = datetime.min

        if email not in latest or ts > latest[email][0]:
            latest[email] = (ts, row)

    volunteers = []
    for email, (_, row) in latest.items():
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

        # Friday NIGHT is a separate column — append to Friday if present
        fri_night_cell = safe(row, idx_fri_night)
        if fri_night_cell and fri_night_cell.lower() not in ("not available", "n/a", "na", ""):
            # Cell value is a specific date like "4/3" meaning they marked that weekend date
            # OR it could be a date indicating they are available on ALL Friday nights
            # In the real form, the value is a specific date they're available (e.g. "4/3")
            # We treat any non-empty, non-NA value as available for ALL Friday NIGHTs in block
            if "NIGHT" not in weekly["Friday"]:
                weekly["Friday"].append("NIGHT")

        # Saturday DAY / NIGHT
        sat_day_cell   = safe(row, idx_sat_day)
        sat_night_cell = safe(row, idx_sat_night)
        sat_shifts = []
        if sat_day_cell and sat_day_cell.lower() not in ("not available", "n/a", "na", ""):
            sat_shifts.append("DAY")
        if sat_night_cell and sat_night_cell.lower() not in ("not available", "n/a", "na", ""):
            sat_shifts.append("NIGHT")
        weekly["Saturday"] = sat_shifts

        # Sunday DAY (the Sunday DAY column is separate from the general Sunday column)
        sun_day_cell = safe(row, idx_sun_day)
        sun_day_avail = (
            sun_day_cell and
            sun_day_cell.lower() not in ("not available", "n/a", "na", "")
        )
        if sun_day_avail and "DAY" not in weekly["Sunday"]:
            weekly["Sunday"].append("DAY")

        # Expand to concrete dates
        available = expand_availability(
            weekly, block_start, block_end, blackout_slots, blackout_dates
        )

        v = Volunteer(
            first_name=first,
            last_name=last,
            email=email,
            certification=cert,
            available=available,
            blackout_slots=blackout_slots,
            blackout_dates=blackout_dates,
        )
        volunteers.append(v)

    print(f"  Loaded {len(volunteers)} Ambulance EMT volunteers from form.")
    return volunteers