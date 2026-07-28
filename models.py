"""
models.py
=========
Shared data model for the Brown EMS scheduler: people, shift definitions,
crew sizes, and the rest rules. Everything date/shift-shaped lives here so
the parser, validators, solvers, and exporter all agree on the calendar.

Shift structure
---------------
Weekdays (Mon-Fri):  AM 0700-1300 (6h) | PM 1300-1900 (6h) | NIGHT 1900-0700 (12h)
Weekends (Sat-Sun):  DAY 0700-1900 (12h) | NIGHT 1900-0700 (12h)

Crew sizes
----------
Weekday AM/PM: 2 | NIGHT: 3 | "Big weekend" (Fri NIGHT through Sun DAY): 4
Sunday NIGHT is a normal night (3).

Campus Response
---------------
Weekday-only 3h blocks: A 0700-1000, B 1000-1300, C 1300-1600, D 1600-1900.

Rest rules (max 12h continuous work, then >=12h off)
----------------------------------------------------
AM+PM on the same day is allowed (one 12h stretch). A NIGHT shift conflicts
with every daytime shift (AM/PM/DAY) on the same day AND the next day.
Back-to-back NIGHTs are allowed (exactly 12h off between them).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

# ── Shift / block definitions ────────────────────────────────────────────────

SHIFT_HOURS = {"AM": 6, "PM": 6, "NIGHT": 12, "DAY": 12}

SHIFT_TIMES = {
    "AM":    ("0700", "1300"),
    "PM":    ("1300", "1900"),
    "NIGHT": ("1900", "0700+1"),
    "DAY":   ("0700", "1900"),
}

CAMPUS_BLOCKS = ("A", "B", "C", "D")
CAMPUS_BLOCK_HOURS = 3
CAMPUS_BLOCK_TIMES = {"A": "0700-1000", "B": "1000-1300", "C": "1300-1600", "D": "1600-1900"}

# A shift key is (date, shift_type); a campus key is (date, block).
ShiftKey = tuple[date, str]


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def shift_types_for(d: date) -> list[str]:
    return ["DAY", "NIGHT"] if is_weekend(d) else ["AM", "PM", "NIGHT"]


def is_big_weekend(d: date, shift_type: str) -> bool:
    """Fri NIGHT through Sun DAY run 4-person crews; Sun NIGHT is a normal night."""
    dow = d.weekday()
    if dow == 4 and shift_type == "NIGHT":
        return True
    if dow == 5:
        return True
    if dow == 6 and shift_type == "DAY":
        return True
    return False


def crew_cap(d: date, shift_type: str) -> int:
    if is_big_weekend(d, shift_type):
        return 4
    if shift_type == "NIGHT":
        return 3
    return 2  # weekday AM / PM


def block_dates(block_start: date, block_end: date) -> list[date]:
    return [block_start + timedelta(days=i) for i in range((block_end - block_start).days + 1)]


def all_shift_keys(schedule_dates: list[date]) -> list[ShiftKey]:
    return [(d, s) for d in schedule_dates for s in shift_types_for(d)]


def all_campus_keys(schedule_dates: list[date]) -> list[ShiftKey]:
    return [(d, b) for d in schedule_dates if not is_weekend(d) for b in CAMPUS_BLOCKS]


# ── Rest rules ───────────────────────────────────────────────────────────────

_DAYTIME = ("AM", "PM", "DAY")


def rest_conflict(k1: ShiftKey, k2: ShiftKey) -> bool:
    """
    True when one person may not work both shifts. Encodes "max 12h continuous,
    then at least 12h off": NIGHT excludes daytime shifts on its own calendar
    day and the following day. AM+PM same day and NIGHT->NIGHT are allowed.
    """
    (d1, s1), (d2, s2) = k1, k2
    if k1 == k2:
        return False
    if d2 < d1 or (d2 == d1 and s1 == "NIGHT" and s2 != "NIGHT"):
        (d1, s1), (d2, s2) = (d2, s2), (d1, s1)  # order so a NIGHT comes second/same day
    if d1 == d2:
        return "NIGHT" in (s1, s2) and (s1 in _DAYTIME or s2 in _DAYTIME)
    if s1 == "NIGHT" and d2 == d1 + timedelta(days=1):
        return s2 in _DAYTIME
    return False


def campus_ambulance_overlap(block: str, shift_type: str) -> bool:
    """Whether a campus block overlaps an ambulance shift on the same date."""
    if shift_type == "DAY":
        return True
    if shift_type == "AM":
        return block in ("A", "B")
    if shift_type == "PM":
        return block in ("C", "D")
    return False  # NIGHT


# ── People ───────────────────────────────────────────────────────────────────

@dataclass
class Volunteer:
    """Ambulance EMT (may also take campus response blocks)."""
    first_name: str
    last_name: str
    email: str
    certification: str                      # "EVDT" | "Auth" | "EMT"
    available: set = field(default_factory=set)          # {(date, shift_type)}
    campus_available: set = field(default_factory=set)   # {(date, block)}
    blackout_slots: set = field(default_factory=set)
    blackout_dates: set = field(default_factory=set)
    # Filled by the solvers:
    assigned: list = field(default_factory=list)         # [(date, shift_type)]
    campus_assigned: list = field(default_factory=list)  # [(date, block)]

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def is_evdt(self) -> bool:
        return self.certification == "EVDT"

    @property
    def is_driver(self) -> bool:
        return self.certification in ("EVDT", "Auth")

    @property
    def assigned_hours(self) -> int:
        return sum(SHIFT_HOURS[s] for (_, s) in self.assigned)

    @property
    def campus_assigned_hours(self) -> int:
        return CAMPUS_BLOCK_HOURS * len(self.campus_assigned)


@dataclass
class BertMember:
    """Campus-response-only member."""
    first_name: str
    last_name: str
    email: str
    certification: str = "BERT"
    campus_available: set = field(default_factory=set)   # {(date, block)}
    blackout_slots: set = field(default_factory=set)
    blackout_dates: set = field(default_factory=set)
    campus_assigned: list = field(default_factory=list)

    @property
    def full_name(self) -> str:
        return f"{self.first_name} {self.last_name}".strip()

    @property
    def campus_assigned_hours(self) -> int:
        return CAMPUS_BLOCK_HOURS * len(self.campus_assigned)


# ── Config expansion helpers ─────────────────────────────────────────────────

def expand_als_entries(entries: list, schedule_dates: list[date]) -> set[ShiftKey]:
    """
    Expand config "als_shifts" entries into shift keys. Each entry is
    "YYYY-MM-DD:DAY" / "YYYY-MM-DD:NIGHT" (or {"date":..., "shift":...}).
    Weekday DAY covers both AM and PM; weekend DAY is the single DAY shift.
    """
    in_block = set(schedule_dates)
    out: set[ShiftKey] = set()
    for raw in entries:
        item = raw
        if isinstance(raw, str):
            date_s, sep, kind_s = raw.strip().partition(":")
            if not sep:
                print(f"  [WARN] Skipping als_shifts entry {raw!r} (expected DATE:DAY or DATE:NIGHT)")
                continue
            item = {"date": date_s.strip(), "shift": kind_s.strip()}
        try:
            d = date.fromisoformat(item["date"])
            kind = str(item["shift"]).strip().upper()
        except (KeyError, ValueError, TypeError) as e:
            print(f"  [WARN] Skipping als_shifts entry {raw!r}: {e}")
            continue
        if d not in in_block:
            print(f"  [WARN] als_shifts date {d} outside scheduling block; skipped")
            continue
        if kind == "NIGHT":
            out.add((d, "NIGHT"))
        elif kind == "DAY":
            out.update({(d, "AM"), (d, "PM")} if not is_weekend(d) else {(d, "DAY")})
        else:
            print(f"  [WARN] als_shifts shift must be DAY or NIGHT, got {kind!r}")
    return out


_SHIFT_ORDER = {"AM": 0, "DAY": 0, "PM": 1, "NIGHT": 2}


def expand_blackout_period(start_date: date, start_shift: str, end_date: date, end_shift: str) -> set[ShiftKey]:
    """All shift keys from (start_date, start_shift) through (end_date, end_shift) inclusive."""
    out: set[ShiftKey] = set()
    d = start_date
    while d <= end_date:
        for s in shift_types_for(d):
            after_start = d > start_date or _SHIFT_ORDER[s] >= _SHIFT_ORDER[start_shift]
            before_end = d < end_date or _SHIFT_ORDER[s] <= _SHIFT_ORDER[end_shift]
            if after_start and before_end:
                out.add((d, s))
        d += timedelta(days=1)
    return out
