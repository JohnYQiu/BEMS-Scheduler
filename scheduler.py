"""
scheduler.py — Pure Python two-phase scheduler
"""

import heapq
from datetime import date, timedelta
from dataclasses import dataclass, field
from parse_form import Volunteer, SHIFT_HOURS

# Scheduling hour rules:
# - Everyone is capped at SOFT_MAX_HOURS (typically 18) to avoid "everyone goes to 24".
# - After coverage + min-fill, ONLY volunteers still under MIN_HOURS can be allowed
#   to go up to HARD_MAX_HOURS (typically 24) to catch up.
MIN_HOURS = 18
SOFT_MAX_HOURS = 18
HARD_MAX_HOURS = 24


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
        if dow == 4 and self.shift_type == "NIGHT": return True
        if dow == 5:                                return True
        if dow == 6 and self.shift_type == "DAY":   return True
        return False

    @property
    def is_weekday_daytime(self) -> bool:
        return self.date.weekday() < 5 and self.shift_type in ("AM", "PM")

    @property
    def max_slots(self):
        if self.is_weekday_daytime:
            return 2
        return 4 if self.is_weekend else 3

    @property
    def label(self):
        return f"{self.date.isoformat()} ({self.date.strftime('%a')}) {self.shift_type}"

    @property
    def has_evdt(self): return any(v.is_evdt for v in self.volunteers)

    @property
    def has_auth(self): return any(v.is_auth for v in self.volunteers)

    def slot_1_filled(self):
        return any(v.is_evdt for v in self.volunteers)

    def slot_2_filled(self):
        evdt_count = sum(1 for v in self.volunteers if v.is_evdt)
        auth_count = sum(1 for v in self.volunteers if v.is_auth)
        return auth_count > evdt_count

    def open_general_slots(self):
        """
        EMT-only slots:
          - Weekday: 1 EMT
          - Weekend: 2 EMTs
        These EMT positions should NOT increase just because the Auth slot is empty.
        """
        if self.is_weekday_daytime:
            # Weekday daytime shifts are 2-person crews and can be any cert mix.
            # Capacity is handled by max_slots / len(volunteers), not EMT-only slots.
            return max(0, self.max_slots - len(self.volunteers))
        emt_capacity = 2 if self.is_weekend else 1
        current_emts = sum(1 for v in self.volunteers if not v.is_auth)
        return max(0, emt_capacity - current_emts)


def is_rest_blocked(v: Volunteer, d: date, shift_type: str) -> bool:
    for (sd, ss) in v.scheduled_shifts:
        if sd == d:
            return True
        if ss == "NIGHT" and d == sd + timedelta(days=1):
            return True
    return False


def _eligible(
    v: Volunteer,
    shift: Shift,
    slot: int,
    min_hours_only: bool = False,
    max_hours_for=None,
) -> bool:
    key = (shift.date, shift.shift_type)
    if key not in v.available:                           return False
    if len(shift.volunteers) >= shift.max_slots:         return False
    if max_hours_for is None:
        max_hours = SOFT_MAX_HOURS
    else:
        max_hours = max_hours_for(v)
    if v.scheduled_hours + shift.hours > max_hours:      return False
    if min_hours_only and v.scheduled_hours >= MIN_HOURS: return False
    if v in shift.volunteers:                            return False
    if shift.is_weekday_daytime:
        # Weekday daytime is a 2-person crew (any cert mix).
        # We only use slot==1 to force EVDT placement for ALS daytime shifts.
        if slot == 1 and not v.is_evdt:
            return False
        return True
    if slot == 1 and not v.is_evdt:                      return False
    if slot == 2 and not v.is_auth:                      return False
    return True


def _evdt_available(shift: Shift, volunteers: list) -> bool:
    return any(_eligible(v, shift, 1) for v in volunteers)


def build_shifts(schedule_dates: list, als_shifts: set) -> dict:
    all_shifts = {}
    for d in schedule_dates:
        types = ["DAY", "NIGHT"] if d.weekday() >= 5 else ["AM", "PM", "NIGHT"]
        for s in types:
            key = (d, s)
            all_shifts[key] = Shift(date=d, shift_type=s, has_als=(key in als_shifts))
    return all_shifts


def _build_blackout_slots(start_date, start_shift, end_date, end_shift):
    SHIFT_ORDER = {"AM": 0, "PM": 1, "NIGHT": 2, "DAY": 0}
    result = set()
    current = start_date
    while current <= end_date:
        for s in (["DAY", "NIGHT"] if current.weekday() >= 5 else ["AM", "PM", "NIGHT"]):
            if ((current > start_date) or (SHIFT_ORDER[s] >= SHIFT_ORDER[start_shift])) and \
               ((current < end_date)   or (SHIFT_ORDER[s] <= SHIFT_ORDER[end_shift])):
                result.add((current, s))
        current += timedelta(days=1)
    return result


def _run_phase(
    volunteers,
    shift_keys,
    all_shifts,
    phase: str,
    min_hours_only: bool = False,
    max_hours_for=None,
):
    total_vols  = len(volunteers)
    flexibility = {v.email: sum(1 for k in shift_keys if k in v.available)
                   for v in volunteers}

    def want_shift(key):
        shift = all_shifts[key]
        # Phase 'evdt': focus on ALS shifts getting an EVDT in slot 1 if possible
        if phase == 'evdt':
            if not shift.has_als:
                return False
            has_evdt = _evdt_available(shift, volunteers) or shift.slot_1_filled()
            return has_evdt
        # Phase 'cover': only consider shifts that currently have nobody assigned
        if phase == 'cover':
            return len(shift.volunteers) == 0
        # Phase 'all': consider all shifts
        return True

    def score(v, shift, slot):
        if slot == 1: base = 10000 if shift.has_als else 5000
        elif slot == 2: base = 4500
        else: base = 1000
        if shift.is_weekday_daytime and v.is_evdt and slot != 1:
            # Soft preference: if we're filling a weekday daytime slot, prefer EVDT
            # when it doesn't harm fairness / hour caps.
            base += 200
        max_hours = SOFT_MAX_HOURS if max_hours_for is None else max_hours_for(v)
        hours_remaining = max_hours - v.scheduled_hours
        flex = flexibility.get(v.email, 999)
        return base + hours_remaining * total_vols - flex

    def rebuild():
        heap = []
        for key in shift_keys:
            if not want_shift(key):
                continue
            shift = all_shifts[key]

            # Weekday daytime: 2-person crew, any cert mix.
            # ALS weekday daytime still requires at least one EVDT (handled in 'evdt' phase).
            if shift.is_weekday_daytime:
                if phase == 'evdt':
                    if shift.has_als and not shift.slot_1_filled():
                        for i, v in enumerate(volunteers):
                            if _eligible(v, shift, 1, max_hours_for=max_hours_for):
                                heapq.heappush(heap, (-score(v, shift, 1), i, key, 1))
                    continue

                if len(shift.volunteers) < shift.max_slots:
                    for i, v in enumerate(volunteers):
                        if _eligible(v, shift, 3, min_hours_only=min_hours_only, max_hours_for=max_hours_for):
                            heapq.heappush(heap, (-score(v, shift, 3), i, key, 3))
                continue

            # Slot 1: EVDT only, used in EVDT-focused phase for ALS shifts
            if not shift.slot_1_filled() and phase == 'evdt':
                for i, v in enumerate(volunteers):
                    if _eligible(v, shift, 1, max_hours_for=max_hours_for):
                        heapq.heappush(heap, (-score(v, shift, 1), i, key, 1))

            # In the EVDT phase we ONLY place the ALS EVDT (slot 1).
            # Coverage and filling happen in later phases.
            if phase == 'evdt':
                continue

            # Slot 2: Auth/EVDT only
            if not shift.slot_2_filled():
                for i, v in enumerate(volunteers):
                    if _eligible(v, shift, 2, min_hours_only=min_hours_only, max_hours_for=max_hours_for):
                        heapq.heappush(heap, (-score(v, shift, 2), i, key, 2))

            # Slots 3+: any cert except Auth, fills freely
            # Slot 1 and slot 2 each reserve one position (filled or not)
            # So EMTs can fill up to max_slots - 2 positions
            if shift.open_general_slots() > 0:
                for i, v in enumerate(volunteers):
                    if _eligible(v, shift, 3, min_hours_only=min_hours_only, max_hours_for=max_hours_for) and not v.is_auth:
                        heapq.heappush(heap, (-score(v, shift, 3), i, key, 3))

        return heap

    heap = rebuild()
    count = 0
    while heap:
        neg_score, v_idx, key, slot = heapq.heappop(heap)
        shift = all_shifts[key]
        v     = volunteers[v_idx]

        # Revalidate
        if slot == 1 and (shift.slot_1_filled() or not want_shift(key)): continue
        if not shift.is_weekday_daytime:
            if slot == 2 and shift.slot_2_filled():                           continue
            if slot >= 3 and v.is_auth:                                       continue
            if slot >= 3 and shift.open_general_slots() <= 0:                 continue
        if not _eligible(v, shift, slot, min_hours_only=min_hours_only, max_hours_for=max_hours_for): continue
        if not want_shift(key):                                           continue

        shift.volunteers.append(v)
        v.scheduled_hours += shift.hours
        v.scheduled_shifts.append(key)
        count += 1
        heap = rebuild()

    return count


def run_schedule(volunteers, schedule_dates, als_shifts, blackout_slots=None):

    all_shifts = build_shifts(schedule_dates, als_shifts)
    if blackout_slots:
        for key in blackout_slots:
            all_shifts.pop(key, None)
        for v in volunteers:
            v.available -= blackout_slots

    shift_keys = sorted(all_shifts.keys())

    # Phase 1: prioritize ALS shifts getting an EVDT in slot 1 (only place slot 1)
    print("  Phase 1: ALS shifts — place EVDT (slot 1) where possible...")
    n1 = _run_phase(volunteers, shift_keys, all_shifts, 'evdt', min_hours_only=False)

    # Phase 2: ensure basic coverage — try to give every shift at least one person
    print("  Phase 2: basic coverage — ensuring every shift has at least one volunteer if possible...")
    n_cover = _run_phase(volunteers, shift_keys, all_shifts, 'cover', min_hours_only=False)

    # Phase 3: fill remaining slots up to MIN_HOURS, keeping everyone capped at SOFT_MAX_HOURS
    print(f"  Phase 3: remaining slots — getting everyone to {MIN_HOURS}h (soft cap {SOFT_MAX_HOURS}h)...")
    n3 = _run_phase(volunteers, shift_keys, all_shifts, 'all', min_hours_only=True)

    # Phase 4: if anyone is still under MIN_HOURS, allow ONLY those people to go up to HARD_MAX_HOURS
    unlock = {v.email for v in volunteers if v.scheduled_hours < MIN_HOURS}
    if unlock:
        print(f"  Phase 4: catch-up — unlocking up to {HARD_MAX_HOURS}h for {len(unlock)} under-{MIN_HOURS}h volunteer(s)...")

        def max_hours_for(v: Volunteer) -> int:
            return HARD_MAX_HOURS if v.email in unlock else SOFT_MAX_HOURS

        n4 = _run_phase(
            volunteers,
            shift_keys,
            all_shifts,
            'all',
            min_hours_only=False,
            max_hours_for=max_hours_for,
        )
    else:
        n4 = 0

    total = n1 + n_cover + n3 + n4
    print(f"  Total: {total} assignments.")
    return all_shifts