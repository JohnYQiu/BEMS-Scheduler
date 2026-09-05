"""
verify_schedule.py
==================
Independent rule checker: re-runs the pipeline, then asserts every hard
constraint against the produced assignments (not against the solver's own
bookkeeping). Run it after changing solver logic:

    python verify_schedule.py
"""

from __future__ import annotations

import json
from datetime import date

from ambulance_solver import solve_ambulance
from campus_solver import solve_campus
from models import (
    SHIFT_HOURS,
    block_dates,
    campus_ambulance_overlap,
    crew_cap,
    expand_als_entries,
    rest_conflict,
)
from parse_form import load_all_responses


def main() -> None:
    with open("config.json", encoding="utf-8") as f:
        cfg = json.load(f)
    block_start = date.fromisoformat(cfg["block_start"])
    block_end = date.fromisoformat(cfg["block_end"])
    schedule_dates = block_dates(block_start, block_end)
    als_shifts = expand_als_entries(cfg.get("als_shifts", []), schedule_dates)
    hours = cfg.get("hours", {})
    required = int(hours.get("ambulance_emt", 12))

    volunteers, bert_members = load_all_responses(cfg["form_csv"], block_start, block_end)
    assignments = solve_ambulance(volunteers, schedule_dates, als_shifts,
                                  required_hours=required)
    campus = solve_campus(volunteers, bert_members, schedule_dates,
                          responders_per_block=int(cfg.get("campus_responders_per_block", 2)),
                          emt_required_hours=int(hours.get("campus_emt", 3)),
                          bert_required_hours=int(hours.get("campus_bert", 6)),
                          require_driver=cfg.get("campus_driver_policy", "prefer") == "require")

    errors: list[str] = []

    # Availability + crew caps + ALS reservation
    for key, people in assignments.items():
        cap = crew_cap(*key)
        if len(people) > cap:
            errors.append(f"{key}: {len(people)} assigned > cap {cap}")
        for v in people:
            if key not in v.available:
                errors.append(f"{key}: {v.full_name} not available but assigned")
        if key in als_shifts and not any(v.is_evdt for v in people):
            if len(people) > cap - 1:
                errors.append(f"{key}: ALS with no EVDT but reserved slot filled")

    # Hours cap + rest rules, from the shift side
    for v in volunteers:
        keys = [k for k, people in assignments.items() if v in people]
        total = sum(SHIFT_HOURS[s] for (_, s) in keys)
        if total > required:
            errors.append(f"{v.full_name}: {total}h > required {required}h")
        if sorted(keys) != sorted(v.assigned):
            errors.append(f"{v.full_name}: .assigned out of sync with assignments")
        for i, k1 in enumerate(keys):
            for k2 in keys[i + 1:]:
                if rest_conflict(k1, k2):
                    errors.append(f"{v.full_name}: rest conflict {k1} / {k2}")

    # Campus: caps, availability, no ambulance overlap
    people = volunteers + bert_members
    for key, responders in campus.items():
        if len(responders) > int(cfg.get("campus_responders_per_block", 2)):
            errors.append(f"campus {key}: over responder cap")
        for p in responders:
            if key not in p.campus_available:
                errors.append(f"campus {key}: {p.full_name} not available but assigned")
    for p in people:
        target = int(hours.get("campus_emt", 3)) if p in volunteers else int(hours.get("campus_bert", 6))
        if p.campus_assigned_hours > target:
            errors.append(f"{p.full_name}: campus {p.campus_assigned_hours}h > {target}h")
    for v in volunteers:
        for (d, b) in v.campus_assigned:
            for (ad, s) in v.assigned:
                if ad == d and campus_ambulance_overlap(b, s):
                    errors.append(f"{v.full_name}: campus {d} {b} overlaps ambulance {s}")

    print("\n" + "=" * 55)
    if errors:
        print(f"✗ {len(errors)} RULE VIOLATION(S):")
        for e in errors:
            print(f"  - {e}")
        raise SystemExit(1)
    print("✓ All hard constraints verified: availability, crew caps, ALS "
          "reservation, hour caps, rest rules, campus caps, no overlaps.")


if __name__ == "__main__":
    main()
