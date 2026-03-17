"""
validate.py
===========
Pre-scheduling validation. Checks that each Ambulance EMT volunteer submitted
at least one availability slot for each required shift type:
  - At least 1 Weekday AM
  - At least 1 Weekday PM
  - At least 1 Weekday NIGHT
  - At least 1 Weekend DAY
  - At least 1 Weekend NIGHT

Prints a strike list of volunteers who did not meet requirements,
including exactly what they are missing.
Also prints a per-shift availability summary so personnel can spot thin days.
"""

from datetime import date
from parse_form import Volunteer, SHIFT_HOURS


# ── Helpers ───────────────────────────────────────────────────────────────────

def _is_weekday(d: date) -> bool:
    return d.weekday() < 5  # Mon–Fri

def _is_weekend(d: date) -> bool:
    return d.weekday() >= 5  # Sat–Sun


# ── Availability requirement check ───────────────────────────────────────────

REQUIRED_CATEGORIES = [
    ("WEEKDAY AM",    lambda d, s: _is_weekday(d) and s == "AM"),
    ("WEEKDAY PM",    lambda d, s: _is_weekday(d) and s == "PM"),
    ("WEEKDAY NIGHT", lambda d, s: _is_weekday(d) and s == "NIGHT"),
    ("WEEKEND DAY",   lambda d, s: _is_weekend(d) and s == "DAY"),
    ("WEEKEND NIGHT", lambda d, s: _is_weekend(d) and s == "NIGHT"),
]

def check_availability_requirements(volunteers: list[Volunteer]) -> list[dict]:
    """
    Returns a list of dicts for volunteers who are missing required categories:
      {"volunteer": v, "missing": ["WEEKDAY AM", "WEEKEND NIGHT", ...]}
    """
    violations = []
    for v in volunteers:
        missing = []
        for label, test in REQUIRED_CATEGORIES:
            has = any(test(d, s) for (d, s) in v.available)
            if not has:
                missing.append(label)
        if missing:
            violations.append({"volunteer": v, "missing": missing})
    return violations


def print_strike_list(violations: list[dict]):
    print("\n" + "=" * 60)
    print("STRIKE LIST — INSUFFICIENT AVAILABILITY SUBMITTED")
    print("=" * 60)
    if not violations:
        print("  ✓ All volunteers met the minimum availability requirements.")
        return
    print(f"  {len(violations)} volunteer(s) did not meet requirements:\n")
    for item in violations:
        v = item["volunteer"]
        missing_str = ", ".join(item["missing"])
        print(f"  • {v.full_name:<28} Missing: {missing_str}")
    print()


# ── Total hours check ─────────────────────────────────────────────────────────

def check_total_available_hours(volunteers: list[Volunteer]) -> list[dict]:
    """
    Flag volunteers whose total submitted availability is less than 18 hours.
    This is a secondary check — the primary is the category check above.
    """
    warnings = []
    for v in volunteers:
        total = sum(SHIFT_HOURS.get(s, 0) for (_, s) in v.available)
        if total < 18:
            warnings.append({"volunteer": v, "total_hours": total})
    return warnings


def print_hours_warnings(warnings: list[dict]):
    if not warnings:
        return
    print("\n" + "=" * 60)
    print("LOW AVAILABILITY WARNING — TOTAL HOURS < 18")
    print("=" * 60)
    for item in warnings:
        v = item["volunteer"]
        print(f"  • {v.full_name:<28} Only {item['total_hours']}h of availability submitted")
    print()


# ── Per-shift availability summary ───────────────────────────────────────────

def print_availability_summary(
    volunteers: list[Volunteer],
    schedule_dates: list[date],
    blackout_slots: set = None,
):
    """
    Print how many volunteers are available for each shift slot,
    so personnel can identify thin coverage days before scheduling.
    """
    from collections import defaultdict
    counts: dict = defaultdict(int)
    evdt_counts: dict = defaultdict(int)

    blackout_slots = blackout_slots or set()

    for v in volunteers:
        for (d, s) in v.available:
            if (d, s) in blackout_slots:
                continue
            counts[(d, s)] += 1
            if v.is_evdt:
                evdt_counts[(d, s)] += 1

    def shifts_for(d: date):
        if d.weekday() >= 5:
            return ["DAY", "NIGHT"]
        return ["AM", "PM", "NIGHT"]

    print("\n" + "=" * 60)
    print("AVAILABILITY SUMMARY (volunteers available per slot)")
    print("=" * 60)
    for d in schedule_dates:
        dow = d.strftime("%a")
        for s in shifts_for(d):
            n = counts.get((d, s), 0)
            e = evdt_counts.get((d, s), 0)
            flag = " ⚠ LOW" if n < 2 else ""
            print(f"  {d.isoformat()} ({dow}) {s:<6}  {n:>2} available  ({e} EVDT){flag}")
    print()