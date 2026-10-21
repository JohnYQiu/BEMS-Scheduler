"""Check the produced assignments independently of the solver's constraints."""

from collections import defaultdict
from datetime import timedelta

from models import HourCaps, Schedule, Volunteer, ambulance_capacity, interval, SHIFT_HOURS
from policy import validate_hard_requirements, member_matches


def validate_schedule(schedule: Schedule, people, providers, campus_keys,
                      caps: HourCaps, campus_capacity: int, locks=(),
                      wellness_keys=(), wellness_capacity: int = 1,
                      wellness_shifts_per_week: int = 1,
                      max_evdt_als_shifts_per_week: int | None = None,
                      hard_requirements=()) -> list[str]:
    errors = []
    roster = {p.email: p for p in people}
    if len(roster) != len(people):
        errors.append("Duplicate person identity in roster")
    if set(schedule.ambulance) != set(providers):
        errors.append("Ambulance schedule does not match configured active shifts")
    if set(schedule.campus) != set(campus_keys):
        errors.append("Campus schedule does not match configured active blocks")
    if set(schedule.wellness) != set(wellness_keys):
        errors.append("Wellness Wagon schedule does not match configured active shifts")
    amb_person, cr_person, wellness_person = defaultdict(list), defaultdict(list), defaultdict(list)
    for ambulance, assignments, expected, index in (
        (True, schedule.ambulance, providers, amb_person),
        (False, schedule.campus, set(campus_keys), cr_person),
    ):
        for key, assigned in assignments.items():
            if key not in expected:
                continue
            if len({p.email for p in assigned}) != len(assigned):
                errors.append(f"{key}: same person occupies multiple seats")
            cap = ambulance_capacity(*key, providers[key]) if ambulance else campus_capacity
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

    for key, assigned in schedule.wellness.items():
        if len({p.email for p in assigned}) != len(assigned):
            errors.append(f"{key}: same person occupies multiple Wellness Wagon seats")
        if len(assigned) > wellness_capacity:
            errors.append(f"{key}: {len(assigned)} assigned exceeds Wellness Wagon capacity {wellness_capacity}")
        for p in assigned:
            if roster.get(p.email) is not p:
                errors.append(f"{key}: unknown or duplicate identity {p.email}")
            if key not in p.wellness_available:
                errors.append(f"{p.full_name}: unavailable Wellness Wagon assignment {key}")
            wellness_person[p.email].append(key)

    for p in people:
        ambulance, campus, wellness = amb_person[p.email], cr_person[p.email], wellness_person[p.email]
        ambulance_hours = sum(SHIFT_HOURS[k[1]] for k in ambulance)
        campus_hours = 3 * len(campus)
        if ambulance_hours > caps.ambulance:
            errors.append(f"{p.full_name}: ambulance hours exceed {caps.ambulance}")
        if campus_hours > caps.campus_for(p):
            errors.append(f"{p.full_name}: campus hours exceed {caps.campus_for(p)}")
        if (sorted(ambulance) != sorted(getattr(p, "assigned", [])) or
                sorted(campus) != sorted(p.campus_assigned) or
                sorted(wellness) != sorted(p.wellness_assigned)):
            errors.append(f"{p.full_name}: person totals do not match the exported assignments")
        weeks = defaultdict(int)
        for d, _ in wellness:
            weeks[d - timedelta(days=d.weekday())] += 1
        if any(count > wellness_shifts_per_week for count in weeks.values()):
            errors.append(f"{p.full_name}: more than {wellness_shifts_per_week} Wellness Wagon shift(s) in a week")
        if max_evdt_als_shifts_per_week is not None and isinstance(p, Volunteer) and p.is_evdt:
            als_weeks = defaultdict(int)
            for key in ambulance:
                if providers.get(key) == "ALS":
                    monday = key[0] - timedelta(days=key[0].weekday())
                    als_weeks[monday] += 1
            if any(count > max_evdt_als_shifts_per_week for count in als_weeks.values()):
                errors.append(f"{p.full_name}: more than {max_evdt_als_shifts_per_week} ALS shift(s) as EVDT in a week")
        work = sorted(interval(k) for k in ambulance + campus + wellness)
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
    for rule in validate_hard_requirements(hard_requirements):
        assignments = schedule.ambulance if rule["service"] == "ambulance" else schedule.campus
        for key, assigned in assignments.items():
            if rule["weekdays"] and key[0].weekday() not in rule["weekdays"]:
                continue
            if rule["shifts"] and key[1] not in rule["shifts"]:
                continue
            if rule["service"] == "ambulance" and rule["provider"] != "any" and providers[key] != rule["provider"]:
                continue
            count = sum(member_matches(person, rule) for person in assigned)
            if count < rule["min_people"]:
                errors.append(f"{key}: hard requirement '{rule['label']}' requires {rule['min_people']} matching people; got {count}")
    for lock in locks:
        index = (amb_person if lock.key in schedule.ambulance else
                 cr_person if lock.key in schedule.campus else wellness_person)
        if lock.key not in index[lock.email.strip().lower()]:
            errors.append(f"Locked assignment missing: {lock.email} on {lock.key}")
    return errors
