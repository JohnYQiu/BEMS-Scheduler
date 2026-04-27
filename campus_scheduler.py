"""
campus_scheduler.py
===================
Schedules Campus Response shifts:
  - Weekdays only
  - Blocks: A (0700-1000), B (1000-1300), C (1300-1600), D (1600-1900)
  - Target staffing: 2 responders per block

Constraints:
  - Respect per-role campus hour caps (soft target + hard cap).
  - Do not overlap with assigned ambulance shifts for ambulance EMTs.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
import heapq
from typing import Optional

from parse_form import Volunteer, BertMember, CAMPUS_BLOCKS, CAMPUS_BLOCK_HOURS


@dataclass
class CampusHourCaps:
    """Max campus response hours per block role (3h per A/B/C/D block)."""
    ambulance_emt_max_hours: int = 6
    bert_max_hours: int = 9


@dataclass
class CampusShift:
    date: date
    block: str  # A/B/C/D
    responders: list = field(default_factory=list)  # list[Volunteer|BertMember]

    @property
    def key(self):
        return (self.date, self.block)


def _campus_overlaps_ambulance(block: str, ambulance_shift_type: str) -> bool:
    if ambulance_shift_type == "NIGHT":
        return False
    if ambulance_shift_type == "DAY":
        return True
    if ambulance_shift_type == "AM":
        return block in ("A", "B")
    if ambulance_shift_type == "PM":
        return block in ("C", "D")
    return False


def _ambulance_conflict(v: Volunteer, d: date, block: str) -> bool:
    for (sd, ss) in v.scheduled_shifts:
        if sd != d:
            continue
        if _campus_overlaps_ambulance(block, ss):
            return True
    return False


def _campus_cap_for(person, caps: CampusHourCaps) -> int:
    if isinstance(person, Volunteer):
        return caps.ambulance_emt_max_hours
    return caps.bert_max_hours


def _campus_hours(person) -> int:
    return getattr(person, "campus_scheduled_hours", 0)


def _campus_available(person, key) -> bool:
    return key in getattr(person, "campus_available", set())


def _eligible(person, key, shift: CampusShift, caps: CampusHourCaps) -> bool:
    d, block = key
    if not _campus_available(person, key):
        return False
    if person in shift.responders:
        return False
    if _campus_hours(person) + CAMPUS_BLOCK_HOURS > _campus_cap_for(person, caps):
        return False
    if isinstance(person, Volunteer) and _ambulance_conflict(person, d, block):
        return False
    return True


def run_campus_schedule(
    ambulance_volunteers: list[Volunteer],
    bert_members: list[BertMember],
    schedule_dates: list[date],
    responders_per_block: int = 2,
    hour_caps: Optional[CampusHourCaps] = None,
) -> dict:
    """
    Returns dict[(date, block)] -> CampusShift

    Strategy:
      - Fill every block to staffing target if possible.
      - Prefer assigning people who are below their cap and have low flexibility.
    """
    people = list(ambulance_volunteers) + list(bert_members)
    caps = hour_caps or CampusHourCaps()

    # Build shifts (weekdays only)
    shifts: dict = {}
    for d in schedule_dates:
        if d.weekday() >= 5:
            continue
        for b in CAMPUS_BLOCKS:
            shifts[(d, b)] = CampusShift(date=d, block=b)

    # Flexibility: number of campus blocks available
    flex = {getattr(p, "email", str(i)): len(getattr(p, "campus_available", set())) for i, p in enumerate(people)}

    def score(person, shift: CampusShift) -> int:
        cap = _campus_cap_for(person, caps)
        remaining = cap - _campus_hours(person)
        # Higher is better:
        # - prioritize those with remaining capacity
        # - prioritize less flexible people (smaller flex)
        return remaining * 1000 - flex.get(person.email, 999)

    def _fill_to_target(target_per_block: int) -> int:
        # Build heap of all possible assignments; rebuild after each assignment.
        def rebuild():
            heap = []
            for key, shift in shifts.items():
                if len(shift.responders) >= target_per_block:
                    continue
                for i, person in enumerate(people):
                    if _eligible(person, key, shift, caps):
                        heapq.heappush(heap, (-score(person, shift), i, key))
            return heap

        heap = rebuild()
        count = 0
        while heap:
            neg_s, p_idx, key = heapq.heappop(heap)
            shift = shifts[key]
            person = people[p_idx]

            if len(shift.responders) >= target_per_block:
                continue
            if not _eligible(person, key, shift, caps):
                continue

            shift.responders.append(person)
            person.campus_scheduled_hours = _campus_hours(person) + CAMPUS_BLOCK_HOURS
            person.campus_scheduled_shifts.append(key)
            count += 1
            heap = rebuild()
        return count

    # Ensure every block has 1 responder before doubling up.
    if responders_per_block >= 1:
        _fill_to_target(1)
    if responders_per_block >= 2:
        _fill_to_target(responders_per_block)

    return shifts

