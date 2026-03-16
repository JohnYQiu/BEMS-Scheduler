"""
scheduler.py
============
Core scheduling engine for Brown EMS.

Slot structure (upgrades allowed, never downgrade):
  Weekday (3 slots):
    Slot 1: EVDT only       (EVDT preferred on ALS shifts — same rule)
    Slot 2: Auth or higher  (Auth or EVDT)
    Slot 3: any cert

  Weekend Fri NIGHT – Sun DAY (4 slots):
    Slot 1: EVDT only
    Slot 2: Auth or higher
    Slot 3: any cert
    Slot 4: any cert

Three-pass scheduling:
  Pass 1 — minimum crew:  fill Slot 1 + Slot 2 across all shifts
  Pass 2 — 18h fill-up:   fill remaining slots prioritising volunteers below 18h
  Pass 3 — top-off:       fill any remaining open slots

Rest rules:
  - Only one shift per calendar day
  - After a NIGHT shift, entire next calendar day is blocked

Flags (never downgrades):
  - Slot 1 empty → ALS/NO EVDT or NO EVDT warning
  - Slot 2 empty → NO AUTH DRIVER warning
"""

import random
from datetime import date, timedelta
from dataclasses import dataclass, field
from parse_form import Volunteer, SHIFT_HOURS

MAX_HOURS = 18


# ── Shift data class ──────────────────────────────────────────────────────────
@dataclass
class Shift:
    date:       date
    shift_type: str
    has_als:    bool = False
    volunteers: list = field(default_factory=list)

    @property
    def hours(self): return SHIFT_HOURS[self.shift_type]

    @property
    def is_weekend(self):
        dow = self.date.weekday()
        if dow == 4 and self.shift_type == "NIGHT": return True  # Fri NIGHT
        if dow == 5:                                return True  # All Sat
        if dow == 6 and self.shift_type == "DAY":   return True  # Sun DAY
        return False

    @property
    def max_slots(self):
        return 4 if self.is_weekend else 3

    @property
    def label(self):
        return f"{self.date.isoformat()} ({self.date.strftime('%a')}) {self.shift_type}"

    @property
    def has_evdt(self): return any(v.is_evdt for v in self.volunteers)

    @property
    def has_auth(self): return any(v.is_auth for v in self.volunteers)

    def slot_1_filled(self): return self.has_evdt
    def slot_2_filled(self): return self.has_auth


# ── Slot eligibility ──────────────────────────────────────────────────────────
def _eligible_for_slot(v: Volunteer, slot: int) -> bool:
    """
    slot 1 → EVDT only
    slot 2 → Auth or higher (Auth or EVDT)
    slot 3+ → any cert
    Upgrades are always allowed (e.g. EVDT can fill slot 2 or 3).
    """
    if slot == 1: return v.is_evdt
    if slot == 2: return v.is_auth
    return True


def _next_open_slot(shift: Shift) -> int:
    """
    Return the slot number (1-based) that the next volunteer should fill.
    Slot 1 must be EVDT, Slot 2 must be Auth+, Slots 3+ are open.
    We determine this by looking at what's already assigned.
    """
    evdt_count = sum(1 for v in shift.volunteers if v.is_evdt)
    auth_count = sum(1 for v in shift.volunteers if v.is_auth and not v.is_evdt)

    # Slot 1 needs an EVDT
    if evdt_count == 0:
        return 1
    # Slot 2 needs an Auth+ (Auth or EVDT not already counted in slot 1)
    # i.e. we need at least one Auth (non-EVDT) or a second EVDT
    auth_or_higher = sum(1 for v in shift.volunteers if v.is_auth)
    if auth_or_higher < 2 and not (evdt_count >= 2):
        # Slot 2 still open if we don't have a second auth-or-higher
        if auth_or_higher < 2:
            return 2
    return len(shift.volunteers) + 1  # next open general slot


# ── Rest rule helpers ─────────────────────────────────────────────────────────
def is_rest_blocked(v: Volunteer, candidate_date: date, candidate_shift: str) -> bool:
    for (sched_date, sched_shift) in v.scheduled_shifts:
        if sched_date == candidate_date:
            return True
        if sched_shift == "NIGHT" and candidate_date == sched_date + timedelta(days=1):
            return True
    return False


# ── Flexibility calculator ────────────────────────────────────────────────────
def compute_flexibility(volunteers: list) -> dict:
    return {v.email: len(v.available) for v in volunteers}


# ── Candidate filter ──────────────────────────────────────────────────────────
def _base_candidates(shift: Shift, vol_list: list, min_hours_only: bool) -> list:
    key = (shift.date, shift.shift_type)
    return [
        v for v in vol_list
        if key in v.available
        and v.scheduled_hours + shift.hours <= MAX_HOURS
        and v not in shift.volunteers
        and not is_rest_blocked(v, shift.date, shift.shift_type)
        and (not min_hours_only or v.scheduled_hours < MAX_HOURS)
    ]


def _sort_key(v: Volunteer, flexibility: dict, min_hours_only: bool, rng: random.Random):
    hours_gap = MAX_HOURS - v.scheduled_hours
    flex      = flexibility.get(v.email, 999)
    salt      = rng.random()
    if min_hours_only:
        return (-hours_gap, flex, salt)
    return (flex, salt)


# ── Slot filler ───────────────────────────────────────────────────────────────
def _fill_slot(shift: Shift, slot: int, vol_list: list,
               flexibility: dict, rng: random.Random,
               min_hours_only: bool = False):
    """Fill a specific slot number with an eligible volunteer."""
    candidates = [
        v for v in _base_candidates(shift, vol_list, min_hours_only)
        if _eligible_for_slot(v, slot)
    ]
    if not candidates:
        return
    candidates.sort(key=lambda v: _sort_key(v, flexibility, min_hours_only, rng))
    v = candidates[0]
    shift.volunteers.append(v)
    v.scheduled_hours += shift.hours
    v.scheduled_shifts.append((shift.date, shift.shift_type))


def _fill_open_slots(shift: Shift, vol_list: list,
                     flexibility: dict, rng: random.Random,
                     min_hours_only: bool = False):
    """
    Fill remaining open slots in strict slot order.
    Slot 1 (EVDT): if unavailable, leave empty, still fill slots 2+.
    Slot 2 (Auth+): if unavailable, leave empty, still fill slots 3+.
    Slots 3/4 (any): fill normally.
    Never places a volunteer into a slot they are not eligible for.
    """
    def best(cands):
        if not cands:
            return None
        cands.sort(key=lambda v: _sort_key(v, flexibility, min_hours_only, rng))
        return cands[0]

    def assign(v):
        shift.volunteers.append(v)
        v.scheduled_hours += shift.hours
        v.scheduled_shifts.append((shift.date, shift.shift_type))

    # Slot 1: EVDT only
    if not any(v.is_evdt for v in shift.volunteers):
        v = best([v for v in _base_candidates(shift, vol_list, min_hours_only) if v.is_evdt])
        if v:
            assign(v)

    # Slot 2: Auth+ — leave empty if unavailable, but continue to slots 3+
    auth_count = sum(1 for v in shift.volunteers if v.is_auth)
    if auth_count < 2:
        v = best([v for v in _base_candidates(shift, vol_list, min_hours_only) if v.is_auth])
        if v:
            assign(v)

    # Slots 3+ (and 4 for weekends): any cert
    while len(shift.volunteers) < shift.max_slots:
        v = best(_base_candidates(shift, vol_list, min_hours_only))
        if not v:
            break
        assign(v)


# ── Shift builder ─────────────────────────────────────────────────────────────
def build_shifts(schedule_dates: list, als_shifts: set) -> dict:
    all_shifts = {}
    for d in schedule_dates:
        shift_types = ["DAY", "NIGHT"] if d.weekday() >= 5 else ["AM", "PM", "NIGHT"]
        for s in shift_types:
            key = (d, s)
            all_shifts[key] = Shift(date=d, shift_type=s, has_als=(key in als_shifts))
    return all_shifts


def _build_blackout_slots(
    start_date: date, start_shift: str,
    end_date: date,   end_shift: str,
) -> set:
    """Return all (date, shift) tuples between start and end inclusive."""
    SHIFT_ORDER = {"AM": 0, "PM": 1, "NIGHT": 2, "DAY": 0}
    result = set()
    current = start_date
    while current <= end_date:
        for s in (["DAY", "NIGHT"] if current.weekday() >= 5 else ["AM", "PM", "NIGHT"]):
            after_start  = (current > start_date) or (SHIFT_ORDER[s] >= SHIFT_ORDER[start_shift])
            before_end   = (current < end_date)   or (SHIFT_ORDER[s] <= SHIFT_ORDER[end_shift])
            if after_start and before_end:
                result.add((current, s))
        current += timedelta(days=1)
    return result

# ── Orchestrator ──────────────────────────────────────────────────────────────
def run_schedule(volunteers: list, schedule_dates: list,
                 als_shifts: set, blackout_slots: set = None,
                 seed: int = 42) -> dict:
    """
    Three-pass scheduling:
      Pass 1 — minimum crew: fill Slot 1 (EVDT) and Slot 2 (Auth+) on all shifts
      Pass 2 — 18h fill-up:  fill remaining slots, prioritise volunteers under 18h
      Pass 3 — top-off:      fill any remaining open slots
    ALS shifts processed before non-ALS in every pass.
    """
    rng        = random.Random(seed)
    all_shifts = build_shifts(schedule_dates, als_shifts)
    if blackout_slots:
        for key in blackout_slots:
            all_shifts.pop(key, None)
    flexibility = compute_flexibility(volunteers)

    def sorted_keys():
        keys = list(all_shifts.keys())
        # ALS first, then by date/shift
        keys.sort(key=lambda k: (0 if all_shifts[k].has_als else 1, k))
        return keys

    # ── Pass 1: minimum crew (Slot 1 + Slot 2) ───────────────────────────────
    for key in sorted_keys():
        shift = all_shifts[key]
        # Slot 1: EVDT
        if not shift.slot_1_filled():
            _fill_slot(shift, 1, volunteers, flexibility, rng)
        # Slot 2: Auth+
        if not shift.slot_2_filled():
            _fill_slot(shift, 2, volunteers, flexibility, rng)

    # ── Pass 2: get everyone to 18h (remaining slots, min_hours_only) ────────
    for key in sorted_keys():
        shift = all_shifts[key]
        _fill_open_slots(shift, volunteers, flexibility, rng, min_hours_only=True)

    # ── Pass 3: top-off (fill any remaining open slots) ──────────────────────
    for key in sorted_keys():
        shift = all_shifts[key]
        _fill_open_slots(shift, volunteers, flexibility, rng, min_hours_only=False)

    return all_shifts