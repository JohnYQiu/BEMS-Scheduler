"""Check the produced assignments independently of the solver's constraints."""

from collections import defaultdict
from datetime import timedelta

from models import HourCaps, Schedule, Volunteer, crew_cap, interval, SHIFT_HOURS


def validate_schedule(schedule: Schedule, people, providers, campus_keys,
                      caps: HourCaps, campus_capacity: int, locks=()) -> list[str]:
    errors = []
    roster = {p.email: p for p in people}
    if len(roster) != len(people):
        errors.append("Duplicate person identity in roster")
    if set(schedule.ambulance) != set(providers):
        errors.append("Ambulance schedule does not match configured active shifts")
    if set(schedule.campus) != set(campus_keys):
        errors.append("Campus schedule does not match configured active blocks")
    amb_person, cr_person = defaultdict(list), defaultdict(list)
    for ambulance, assignments, expected, index in (
        (True, schedule.ambulance, providers, amb_person),
        (False, schedule.campus, set(campus_keys), cr_person),
    ):
        for key, assigned in assignments.items():
            if key not in expected:
                continue
            if len({p.email for p in assigned}) != len(assigned):
                errors.append(f"{key}: same person occupies multiple seats")
            cap = crew_cap(*key) if ambulance else campus_capacity
            if len(assigned) > cap:
                errors.append(f"{key}: {len(assigned)} assigned exceeds capacity {cap}")
            if ambulance and providers[key] == "ALS" and sum(not p.is_evdt for p in assigned) > cap - 1:
                errors.append(f"{key}: ALS EVDT reservation filled by a non-EVDT")
            for p in assigned:
                if roster.get(p.email) is not p:
                    errors.append(f"{key}: unknown or duplicate identity {p.email}")
                if ambulance and not isinstance(p, Volunteer):
                    errors.append(f"{key}: campus-only member counted as ambulance EMT")
                available = getattr(p, "available", set()) if ambulance else p.campus_available
                if key not in available:
                    errors.append(f"{p.full_name}: unavailable assignment {key}")
                index[p.email].append(key)

    for p in people:
        ambulance, campus = amb_person[p.email], cr_person[p.email]
        ambulance_hours = sum(SHIFT_HOURS[k[1]] for k in ambulance)
        campus_hours = 3 * len(campus)
        if ambulance_hours > caps.ambulance:
            errors.append(f"{p.full_name}: ambulance hours exceed {caps.ambulance}")
        if campus_hours > caps.campus_for(p):
            errors.append(f"{p.full_name}: campus hours exceed {caps.campus_for(p)}")
        if sorted(ambulance) != sorted(getattr(p, "assigned", [])) or sorted(campus) != sorted(p.campus_assigned):
            errors.append(f"{p.full_name}: person totals do not match the exported assignments")
        work = sorted(interval(k) for k in ambulance + campus)
        if not work:
            continue
        run_start, run_end = work[0]
        for start, end in work[1:]:
            if start < run_end:
                errors.append(f"{p.full_name}: overlapping assignments at {start}")
                run_end = max(run_end, end)
            elif start == run_end:
                run_end = end
            else:
                if start.date() == run_start.date():
                    errors.append(f"{p.full_name}: gap within workday on {start.date()}")
                if start - run_end < timedelta(hours=12):
                    errors.append(f"{p.full_name}: less than 12 hours of rest before {start}")
                run_start, run_end = start, end
            if run_end - run_start > timedelta(hours=12):
                errors.append(f"{p.full_name}: more than 12 continuous hours ending {run_end}")
    for lock in locks:
        index = amb_person if lock.key[1] in SHIFT_HOURS else cr_person
        if lock.key not in index[lock.email.strip().lower()]:
            errors.append(f"Locked assignment missing: {lock.email} on {lock.key}")
    return errors
