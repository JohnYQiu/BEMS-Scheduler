"""People, shift times and staffing rules shared by scheduling and reporting."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta

ShiftKey = tuple[date, str]
SHIFT_HOURS = {"AM": 6, "PM": 6, "NIGHT": 12, "DAY": 12}
SHIFT_TIMES = {"AM": ("0700", "1300"), "PM": ("1300", "1900"),
               "NIGHT": ("1900", "0700+1"), "DAY": ("0700", "1900")}
CAMPUS_BLOCKS = ("A", "B", "C", "D")
CAMPUS_BLOCK_HOURS = 3
CAMPUS_BLOCK_TIMES = {"A": "0700-1000", "B": "1000-1300",
                      "C": "1300-1600", "D": "1600-1900"}
START_HOURS = {"AM": 7, "PM": 13, "DAY": 7, "NIGHT": 19,
               "A": 7, "B": 10, "C": 13, "D": 16}


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def shift_types_for(d: date) -> list[str]:
    return ["DAY", "NIGHT"] if is_weekend(d) else ["AM", "PM", "NIGHT"]


def is_weekend_night(d: date, kind: str) -> bool:
    return kind == "NIGHT" and d.weekday() in (4, 5)


def is_weekend_day(d: date, kind: str) -> bool:
    return kind == "DAY" and is_weekend(d)


def is_big_weekend(d: date, kind: str) -> bool:
    return is_weekend_night(d, kind) or is_weekend_day(d, kind)


def crew_cap(d: date, kind: str) -> int:
    """Volunteer seats; the separately supplied supervisor is not included."""
    return 4 if is_big_weekend(d, kind) else 3 if kind == "NIGHT" else 2


def block_dates(start: date, end: date) -> list[date]:
    if end < start:
        raise ValueError("block_end must be on or after block_start")
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def all_shift_keys(dates: list[date]) -> list[ShiftKey]:
    return [(d, s) for d in dates for s in shift_types_for(d)]


def all_campus_keys(dates: list[date]) -> list[ShiftKey]:
    return [(d, b) for d in dates if not is_weekend(d) for b in CAMPUS_BLOCKS]


def interval(key: ShiftKey) -> tuple[datetime, datetime]:
    d, kind = key
    start = datetime.combine(d, time(START_HOURS[kind]))
    hours = SHIFT_HOURS[kind] if kind in SHIFT_HOURS else CAMPUS_BLOCK_HOURS
    return start, start + timedelta(hours=hours)


def campus_ambulance_overlap(block: str, kind: str) -> bool:
    return kind == "DAY" or (kind == "AM" and block in ("A", "B")) or (
        kind == "PM" and block in ("C", "D"))


@dataclass
class BertMember:
    first_name: str
    last_name: str
    email: str
    certification: str = "BERT"
    campus_available: set[ShiftKey] = field(default_factory=set)
    blackout_slots: set[ShiftKey] = field(default_factory=set)
    blackout_dates: set[date] = field(default_factory=set)
    campus_assigned: list[ShiftKey] = field(default_factory=list)

    def __post_init__(self):
        self.email = self.email.strip().lower()

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
    def campus_assigned_hours(self) -> int:
        return CAMPUS_BLOCK_HOURS * len(self.campus_assigned)


@dataclass
class Volunteer(BertMember):
    """Ambulance EMT, including EMT/ERT dual-role members."""
    certification: str = "EMT"
    available: set[ShiftKey] = field(default_factory=set)
    assigned: list[ShiftKey] = field(default_factory=list)

    @property
    def assigned_hours(self) -> int:
        return sum(SHIFT_HOURS[s] for _, s in self.assigned)


@dataclass(frozen=True)
class HourCaps:
    ambulance: int = 18
    campus_emt: int = 6
    campus_bert: int = 9

    def __post_init__(self):
        for value, unit in ((self.ambulance, 6), (self.campus_emt, 3), (self.campus_bert, 3)):
            if type(value) is not int or value < 0 or value % unit:
                raise ValueError(f"Hour caps must be nonnegative multiples of {unit}")

    def campus_for(self, person: BertMember) -> int:
        return self.campus_emt if isinstance(person, Volunteer) else self.campus_bert


@dataclass(frozen=True)
class LockedAssignment:
    email: str
    key: ShiftKey


@dataclass(frozen=True)
class SolveStage:
    name: str
    status: str
    value: int | None
    bound: float | None
    seconds: float


@dataclass
class Schedule:
    ambulance: dict[ShiftKey, list[Volunteer]]
    campus: dict[ShiftKey, list[BertMember]]
    stages: list[SolveStage] = field(default_factory=list)
