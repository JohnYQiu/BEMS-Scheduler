"""
scheduler.py — Pure Python multi-phase scheduler
"""

from __future__ import annotations

import heapq
from datetime import date, timedelta
from dataclasses import dataclass, field
from typing import Optional

from parse_form import Volunteer, SHIFT_HOURS


@dataclass
class AmbulanceHourLimits:
    """
    target_hours: goal used in phase 3 (min-fill) and phase 4 unlock threshold.
    soft_max_hours: normal cap before phase 4 catch-up.
    hard_max_hours: absolute max for volunteers still under target_hours in phase 4.
    """
    target_hours: int = 18
    soft_max_hours: int = 18
    hard_max_hours: int = 24


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
    def is_weekend_day(self) -> bool:
        return self.is_weekend and self.shift_type == "DAY"

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
        """ALS: EVDT in reserved slot 1. Non-ALS: any first assignee."""
        if self.has_als:
            return any(v.is_evdt for v in self.volunteers)
        return len(self.volunteers) >= 1

    def slot_2_filled(self):
        evdt_count = sum(1 for v in self.volunteers if v.is_evdt)
        auth_count = sum(1 for v in self.volunteers if v.is_auth)
        return auth_count > evdt_count

    def open_general_slots(self):
        """
        Non-auth EMT slots (weekend / night typed model).
        Weekday AM/PM uses max_slots and len(volunteers) only.
        """
        if self.is_weekday_daytime:
            return max(0, self.max_slots - len(self.volunteers))
        if not self.has_als:
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
    *,
    soft_max_hours: int = 18,
    target_hours: int = 18,
) -> bool:
    key = (shift.date, shift.shift_type)
    if key not in v.available:                           return False
    if len(shift.volunteers) >= shift.max_slots:         return False
    if max_hours_for is None:
        max_hours = soft_max_hours
    else:
        max_hours = max_hours_for(v)
    if v.scheduled_hours + shift.hours > max_hours:      return False
    if min_hours_only and v.scheduled_hours >= target_hours: return False
    if v in shift.volunteers:                            return False

    # Non-ALS: no EVDT-only slot; any cert in slots 1–2; higher slots are EMT-only (not Auth).
    if not shift.has_als:
        if shift.is_weekday_daytime:
            return True
        if slot in (1, 2):
            return True
        if slot >= 3:
            return not v.is_auth
        return True

    # ALS: slot 1 is EVDT-only or left empty (never Auth/EMT in slot 1).
    if shift.is_weekday_daytime:
        if slot == 1 and not v.is_evdt:
            return False
        return True

    if slot == 1 and not v.is_evdt:
        return False
    if slot == 2 and not v.is_auth:
        return False
    return True


def _evdt_available(shift: Shift, volunteers: list, soft_max_hours: int, target_hours: int) -> bool:
    return any(
        _eligible(v, shift, 1, soft_max_hours=soft_max_hours, target_hours=target_hours)
        for v in volunteers
    )


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


def _next_slot_for_non_als(shift: Shift) -> int:
    """Sequential slot index 1..max_slots for relaxed non-ALS night/weekend."""
    n = len(shift.volunteers)
    if n >= shift.max_slots:
        return 0
    return n + 1


def _run_phase(
    volunteers,
    shift_keys,
    all_shifts,
    phase: str,
    limits: AmbulanceHourLimits,
    min_hours_only: bool = False,
    max_hours_for=None,
):
    total_vols  = len(volunteers)
    flexibility = {v.email: sum(1 for k in shift_keys if k in v.available)
                   for v in volunteers}

    def want_shift(key):
        shift = all_shifts[key]
        if phase == 'evdt':
            if not shift.has_als:
                return False
            has_evdt = _evdt_available(
                shift, volunteers, limits.soft_max_hours, limits.target_hours
            ) or shift.slot_1_filled()
            return has_evdt
        if phase == 'cover':
            return len(shift.volunteers) == 0
        return True

    def score(v, shift, slot):
        if slot == 1: base = 10000 if shift.has_als else 5000
        elif slot == 2: base = 4500
        else: base = 1000
        if shift.is_weekday_daytime and not shift.has_als and v.is_evdt and slot != 1:
            base += 200
        max_hours = limits.soft_max_hours if max_hours_for is None else max_hours_for(v)
        hours_remaining = max_hours - v.scheduled_hours
        flex = flexibility.get(v.email, 999)
        return base + hours_remaining * total_vols - flex

    def rebuild():
        heap = []
        for key in shift_keys:
            if not want_shift(key):
                continue
            shift = all_shifts[key]

            # Weekday AM/PM
            if shift.is_weekday_daytime:
                if shift.has_als:
                    if phase == 'evdt' and not shift.slot_1_filled():
                        for i, v in enumerate(volunteers):
                            if _eligible(
                                v, shift, 1, max_hours_for=max_hours_for,
                                soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                            ):
                                heapq.heappush(heap, (-score(v, shift, 1), i, key, 1))
                    elif phase != 'evdt' and len(shift.volunteers) < shift.max_slots:
                        for i, v in enumerate(volunteers):
                            if _eligible(
                                v, shift, 3, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
                                soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                            ):
                                heapq.heappush(heap, (-score(v, shift, 3), i, key, 3))
                else:
                    if phase == 'evdt':
                        continue
                    if len(shift.volunteers) < shift.max_slots:
                        for i, v in enumerate(volunteers):
                            if _eligible(
                                v, shift, 3, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
                                soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                            ):
                                heapq.heappush(heap, (-score(v, shift, 3), i, key, 3))
                continue

            # Non-ALS night / weekend (typed slots relaxed: 1–2 any, 3+ EMT-only)
            if not shift.has_als:
                if phase == 'evdt':
                    continue
                slot = _next_slot_for_non_als(shift)
                if slot == 0:
                    continue
                for i, v in enumerate(volunteers):
                    if _eligible(
                        v, shift, slot, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
                        soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                    ):
                        heapq.heappush(heap, (-score(v, shift, slot), i, key, slot))
                continue

            # ALS weekend DAY / NIGHT — slot 1 EVDT only or empty; then Auth; then EMT slots
            if not shift.slot_1_filled() and phase == 'evdt':
                for i, v in enumerate(volunteers):
                    if _eligible(
                        v, shift, 1, max_hours_for=max_hours_for,
                        soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                    ):
                        heapq.heappush(heap, (-score(v, shift, 1), i, key, 1))

            if phase == 'evdt':
                continue

            if not shift.slot_2_filled():
                for i, v in enumerate(volunteers):
                    if _eligible(
                        v, shift, 2, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
                        soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                    ):
                        heapq.heappush(heap, (-score(v, shift, 2), i, key, 2))

            if shift.open_general_slots() > 0:
                for i, v in enumerate(volunteers):
                    if _eligible(
                        v, shift, 3, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
                        soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
                    ) and not v.is_auth:
                        heapq.heappush(heap, (-score(v, shift, 3), i, key, 3))

        return heap

    heap = rebuild()
    count = 0
    while heap:
        neg_score, v_idx, key, slot = heapq.heappop(heap)
        shift = all_shifts[key]
        v     = volunteers[v_idx]

        if shift.is_weekday_daytime:
            if slot == 1 and (shift.slot_1_filled() or not want_shift(key)):
                continue
            if slot == 3 and len(shift.volunteers) >= shift.max_slots:
                continue
        elif not shift.has_als:
            expected = _next_slot_for_non_als(shift)
            if slot != expected or expected == 0:
                continue
        else:
            if slot == 1 and (shift.slot_1_filled() or not want_shift(key)):
                continue
            if slot == 2 and shift.slot_2_filled():
                continue
            if slot >= 3 and v.is_auth:
                continue
            if slot >= 3 and shift.open_general_slots() <= 0:
                continue

        if not _eligible(
            v, shift, slot, min_hours_only=min_hours_only, max_hours_for=max_hours_for,
            soft_max_hours=limits.soft_max_hours, target_hours=limits.target_hours,
        ):
            continue
        if not want_shift(key):
            continue

        shift.volunteers.append(v)
        v.scheduled_hours += shift.hours
        v.scheduled_shifts.append(key)
        count += 1
        heap = rebuild()

    return count


def run_schedule(
    volunteers,
    schedule_dates,
    als_shifts,
    blackout_slots=None,
    hour_limits: Optional[AmbulanceHourLimits] = None,
):
    limits = hour_limits or AmbulanceHourLimits()
    all_shifts = build_shifts(schedule_dates, als_shifts)
    if blackout_slots:
        for key in blackout_slots:
            all_shifts.pop(key, None)
        for v in volunteers:
            v.available -= blackout_slots

    shift_keys = sorted(all_shifts.keys())

    print("  Phase 1: ALS shifts — place EVDT in slot 1 where possible (otherwise leave blank)...")
    n1 = _run_phase(volunteers, shift_keys, all_shifts, 'evdt', limits, min_hours_only=False)

    print("  Phase 2: basic coverage — ensuring every shift has at least one volunteer if possible...")
    n_cover = _run_phase(volunteers, shift_keys, all_shifts, 'cover', limits, min_hours_only=False)

    print(
        f"  Phase 3: remaining slots — getting everyone to {limits.target_hours}h "
        f"(soft cap {limits.soft_max_hours}h)..."
    )
    n3 = _run_phase(volunteers, shift_keys, all_shifts, 'all', limits, min_hours_only=True)

    unlock = {v.email for v in volunteers if v.scheduled_hours < limits.target_hours}
    if unlock:
        print(
            f"  Phase 4: catch-up — unlocking up to {limits.hard_max_hours}h for "
            f"{len(unlock)} under-{limits.target_hours}h volunteer(s)..."
        )

        def max_hours_for(v: Volunteer) -> int:
            return limits.hard_max_hours if v.email in unlock else limits.soft_max_hours

        n4 = _run_phase(
            volunteers,
            shift_keys,
            all_shifts,
            'all',
            limits,
            min_hours_only=False,
            max_hours_for=max_hours_for,
        )
    else:
        n4 = 0

    total = n1 + n_cover + n3 + n4
    print(f"  Total: {total} assignments.")
    reconcile_ambulance_hours(volunteers, all_shifts)
    return all_shifts


def reconcile_ambulance_hours(volunteers: list, all_shifts: dict) -> None:
    """
    Set each volunteer's scheduled_hours from scheduled_shifts (deduped by shift key).
    Keeps totals consistent with assignments and makes hour bugs obvious in exports.
    """
    for v in volunteers:
        seen = set()
        total = 0
        for key in v.scheduled_shifts:
            if key in seen:
                continue
            seen.add(key)
            if key in all_shifts:
                total += all_shifts[key].hours
            else:
                total += SHIFT_HOURS.get(key[1], 0)
        v.scheduled_hours = total
