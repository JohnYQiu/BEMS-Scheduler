"""Joint ambulance/CR scheduling with hard work rules and staged objectives."""

from __future__ import annotations

from collections import defaultdict
from datetime import timedelta
from time import monotonic

from ortools.sat.python import cp_model

from models import (
    CAMPUS_BLOCK_HOURS, SHIFT_HOURS, START_HOURS, BertMember, HourCaps,
    LockedAssignment, Schedule, ShiftKey, SolveStage, Volunteer,
    WEEKEND_PRIORITY_LABELS, ambulance_capacity, weekend_priority,
)
from policy import (DEFAULT_OBJECTIVE_ORDER, OBJECTIVE_LABELS, validate_custom_priorities,
                    validate_hard_requirements, validate_objective_order)


def _present(model, variables, name):
    if not variables:
        return 0
    present = model.new_bool_var(name)
    model.add_max_equality(present, variables)
    return present


def _at_least(model, variables, minimum, name):
    if len(variables) < minimum:
        return 0
    enough = model.new_bool_var(name)
    model.add(sum(variables) >= minimum).only_enforce_if(enough)
    model.add(sum(variables) < minimum).only_enforce_if(enough.Not())
    return enough


def _matches_member(person, selected):
    """Match a form-derived role/driver status; never alter eligibility."""
    if selected == "any":
        return True
    if selected == "driver":
        return person.is_driver
    if selected == "evdt":
        return person.is_evdt
    if selected == "auth":
        return person.certification == "Auth"
    if selected == "non_driver":
        return not person.is_driver
    if selected == "emt":
        return isinstance(person, Volunteer)
    return not isinstance(person, Volunteer)  # validated value: bert


def _custom_objective(model, rule, people, providers, by_service):
    """Build a soft expression from supported, validated assignment filters."""
    terms = []
    for key, entries in by_service[rule["service"]].items():
        if rule["weekdays"] and key[0].weekday() not in rule["weekdays"]:
            continue
        if rule["shifts"] and key[1] not in rule["shifts"]:
            continue
        if rule["service"] == "ambulance" and rule["provider"] != "any" and providers[key] != rule["provider"]:
            continue
        selected = [var for pi, var in entries if _matches_member(people[pi], rule["member_filter"])]
        if rule["metric"] == "covered_shifts":
            terms.append(_present(model, selected, f"{rule['id']}_{key[0]}_{key[1]}"))
        elif rule["metric"] == "hours":
            hours = CAMPUS_BLOCK_HOURS if rule["service"] == "campus" else SHIFT_HOURS[key[1]]
            terms.extend(hours * var for var in selected)
        else:
            terms.extend(selected)
    expression = sum(terms)
    return -expression if rule["direction"] == "minimize" else expression


def _optimize(model, objectives, time_limit_s, workers):
    """Preserve each stage's best result; share one time budget across stages.

    FEASIBLE values may be preserved, but are never described as proven optima.
    If a later solve finds no incumbent, retain the last valid complete schedule.
    """
    deadline = monotonic() + time_limit_s
    stages, values = [], None
    active = [(stage[0], stage[1], stage[2] if len(stage) == 3 else "maximize")
              for stage in objectives if not isinstance(stage[1], int)]
    for i, (name, objective, direction) in enumerate(active):
        remaining = deadline - monotonic()
        if remaining <= 0:
            stages.append(SolveStage(name, "NOT_RUN", None, None, 0))
            break
        model.maximize(objective)
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = remaining / (len(active) - i)
        solver.parameters.num_workers = workers
        solver.parameters.random_seed = 0
        status = solver.solve(model)
        status_name = solver.status_name(status)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            if values is None or status in (cp_model.INFEASIBLE, cp_model.MODEL_INVALID):
                raise RuntimeError(f"Schedule solver failed at {name}: {status_name}. "
                                   "Check hard requirements and locked assignments, and increase the time budget if needed.")
            stages.append(SolveStage(name, status_name, None, None, solver.wall_time))
            break
        value = int(solver.value(objective))
        # The model maximizes -count for minimization. Report the real count and
        # a lower bound rather than a confusing negative score/bound.
        display_sign = -1 if direction == "minimize" else 1
        stages.append(SolveStage(name, status_name, display_sign * value,
                                 display_sign * solver.best_objective_bound, solver.wall_time))
        values = [solver.value(model.get_int_var_from_proto_index(j))
                  for j in range(len(model.proto.variables))]
        model.add(objective == value)
        model.clear_hints()
        for j, value in enumerate(values):
            model.add_hint(model.get_int_var_from_proto_index(j), value)
    if values is None:
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = max(0.001, deadline - monotonic())
        status = solver.solve(model)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            raise RuntimeError(f"Schedule solver failed: {solver.status_name(status)}")
        values = [solver.value(model.get_int_var_from_proto_index(j))
                  for j in range(len(model.proto.variables))]
    return values, stages


def solve_schedule(
    people: list[BertMember],
    providers: dict[ShiftKey, str],
    campus_keys: list[ShiftKey],
    caps: HourCaps = HourCaps(),
    campus_capacity: int = 2,
    locks: list[LockedAssignment] | None = None,
    time_limit_s: float = 30,
    workers: int = 8,
    wellness_keys: list[ShiftKey] | None = None,
    wellness_capacity: int = 1,
    objective_order: list[str] | tuple[str, ...] | None = None,
    wellness_shifts_per_week: int = 1,
    custom_priorities: list[dict] | tuple[dict, ...] | None = None,
    max_evdt_als_shifts_per_week: int | None = None,
    hard_requirements: list[dict] | tuple[dict, ...] | None = None,
) -> Schedule:
    custom_priorities = validate_custom_priorities(custom_priorities or ())
    hard_requirements = validate_hard_requirements(hard_requirements or ())
    objective_order = validate_objective_order(
        DEFAULT_OBJECTIVE_ORDER if objective_order is None else objective_order, custom_priorities)
    if time_limit_s <= 0 or type(workers) is not int or workers < 1:
        raise ValueError("Solver time budget and worker count must be positive")
    if type(campus_capacity) is not int or campus_capacity < 1:
        raise ValueError("campus_capacity must be a positive integer")
    if type(wellness_capacity) is not int or wellness_capacity < 1:
        raise ValueError("wellness_capacity must be a positive integer")
    if type(wellness_shifts_per_week) is not int or not 1 <= wellness_shifts_per_week <= 5:
        raise ValueError("wellness_shifts_per_week must be an integer from 1 to 5")
    if max_evdt_als_shifts_per_week is not None and (
            type(max_evdt_als_shifts_per_week) is not int or not 1 <= max_evdt_als_shifts_per_week <= 7):
        raise ValueError("max_evdt_als_shifts_per_week must be an integer from 1 to 7")
    if any(kind not in ("ALS", "BLS") for kind in providers.values()):
        raise ValueError("Every ambulance shift must explicitly specify ALS or BLS")
    people = sorted(people, key=lambda p: p.email)
    if any(not p.email for p in people) or len({p.email for p in people}) != len(people):
        raise ValueError("Every person must have one unique, nonempty email")

    model = cp_model.CpModel()
    x, y, z = {}, {}, {}
    amb_by_key, cr_by_key, wellness_by_key = defaultdict(list), defaultdict(list), defaultdict(list)
    by_person = defaultdict(list)
    campus_keys = sorted(set(campus_keys))
    wellness_keys = sorted(set(wellness_keys or ()))
    campus_set, wellness_set = set(campus_keys), set(wellness_keys)
    for pi, person in enumerate(people):
        services = [
            (y, person.campus_available & campus_set, cr_by_key),
            (z, person.wellness_available & wellness_set, wellness_by_key),
        ]
        if isinstance(person, Volunteer):
            services.append((x, person.available & providers.keys(), amb_by_key))
        for variables, available, index in services:
            for key in sorted(available):
                var = model.new_bool_var(f"p{pi}_{key[0]}_{key[1]}")
                variables[pi, key] = var
                index[key].append((pi, var))
                by_person[pi].append((key, var))

    for key, entries in amb_by_key.items():
        cap = ambulance_capacity(*key, providers[key])
        model.add(sum(v for _, v in entries) <= cap)
        if providers[key] == "ALS":
            # Preserve one EVDT Driver opportunity; other requests may still
            # fill the remaining crew capacity, but Auth cannot replace EVDT.
            model.add(sum(v for pi, v in entries if not people[pi].is_evdt) <= cap - 1)
    for entries in cr_by_key.values():
        model.add(sum(v for _, v in entries) <= campus_capacity)
    for entries in wellness_by_key.values():
        model.add(sum(v for _, v in entries) <= wellness_capacity)

    requirement_calendars = {"ambulance": amb_by_key, "campus": cr_by_key}
    for rule in hard_requirements:
        for key, entries in requirement_calendars[rule["service"]].items():
            if rule["weekdays"] and key[0].weekday() not in rule["weekdays"]:
                continue
            if rule["shifts"] and key[1] not in rule["shifts"]:
                continue
            if rule["service"] == "ambulance" and rule["provider"] != "any" and providers[key] != rule["provider"]:
                continue
            matching = [var for pi, var in entries if _matches_member(people[pi], rule["member_filter"])]
            model.add(sum(matching) >= rule["min_people"])

    # A/B/C/D are occupancy segments, not mandatory assignment bundles.
    patterns = [(0, 0, 0, 0)] + [
        tuple(int(start <= i < end) for i in range(4))
        for start in range(4) for end in range(start + 1, 5)
    ]
    amb_hours, cr_hours = [], []
    for pi, person in enumerate(people):
        assignments = by_person[pi]
        ambulance_hours = sum(SHIFT_HOURS[k[1]] * v for k, v in assignments if (pi, k) in x)
        campus_hours = sum(CAMPUS_BLOCK_HOURS * v for k, v in assignments if (pi, k) in y)
        model.add(ambulance_hours <= caps.ambulance)
        model.add(campus_hours <= caps.campus_for(person))
        amb_hours.append(ambulance_hours)
        cr_hours.append(campus_hours)
        days = defaultdict(list)
        for key, var in assignments:
            days[key[0]].append((key[1], var))
        weekly_wellness = defaultdict(list)
        for key, var in assignments:
            if (pi, key) in z:
                monday = key[0] - timedelta(days=key[0].weekday())
                weekly_wellness[monday].append(var)
        for entries in weekly_wellness.values():
            model.add(sum(entries) <= wellness_shifts_per_week)
        if max_evdt_als_shifts_per_week is not None and isinstance(person, Volunteer) and person.is_evdt:
            weekly_als = defaultdict(list)
            for key, var in assignments:
                if (pi, key) in x and providers[key] == "ALS":
                    monday = key[0] - timedelta(days=key[0].weekday())
                    weekly_als[monday].append(var)
            for entries in weekly_als.values():
                model.add(sum(entries) <= max_evdt_als_shifts_per_week)
        for d, entries in days.items():
            occupied = []
            for hour in (7, 10, 13, 16):
                active = [var for kind, var in entries if kind != "NIGHT" and
                          START_HOURS[kind] <= hour < START_HOURS[kind] + SHIFT_HOURS.get(kind, 3)]
                bit = model.new_bool_var(f"work_{pi}_{d}_{hour}")
                model.add(bit == sum(active))  # Also excludes simultaneous services.
                occupied.append(bit)
            model.add_allowed_assignments(occupied, patterns)
            # Nights take the full 12h allowance; the following day is rest.
            # Consecutive nights have exactly 12h off and remain possible.
            for night_date in (d, d - timedelta(days=1)):
                night = x.get((pi, (night_date, "NIGHT")))
                if night is not None:
                    model.add(sum(occupied) <= 4 * (1 - night))

    by_email = {p.email: pi for pi, p in enumerate(people)}
    for lock in locks or []:
        pi = by_email.get(lock.email.strip().lower())
        variables = x if lock.key[1] in SHIFT_HOURS else y
        var = variables.get((pi, lock.key))
        if var is None:
            raise ValueError(f"Locked assignment unavailable: {lock.email} on {lock.key}")
        model.add(var == 1)

    coverage, als_other, campus_coverage = [], [], []
    weekends = [{name: [] for name in ("als", "core", "seats")}
                for _ in WEEKEND_PRIORITY_LABELS]
    for key in sorted(providers):
        entries = amb_by_key[key]
        crew = [v for _, v in entries]
        evdts = [v for pi, v in entries if people[pi].is_evdt]
        coverage.append(_present(model, crew, f"covered_{key}"))
        rank = weekend_priority(*key)
        if providers[key] == "ALS":
            target = als_other if rank is None else weekends[rank]["als"]
            target.append(_present(model, evdts, f"evdt_{key}"))
        if rank is None or not entries:
            continue
        group = weekends[rank]
        group["seats"].extend(crew)
        core = model.new_int_var(0, 3, f"weekend_core_{key}")
        model.add_min_equality(core, [sum(crew), 3])
        group["core"].append(core)
    for key in campus_keys:
        campus_coverage.append(_present(model, [v for _, v in cr_by_key[key]], f"campus_{key}"))

    objectives = [("Ambulance shifts with an EMT", sum(coverage))]

    def add_weekend_objective(rank, category, description):
        objectives.append((f"{WEEKEND_PRIORITY_LABELS[rank]}: {description}",
                           sum(weekends[rank][category])))

    # Build Friday and Saturday nights toward three volunteers before weekend
    # day crews. This is ambulance staffing only: Utility vehicles are not modeled.
    for rank in (0, 1):
        add_weekend_objective(rank, "als", "ALS shifts with EVDT")
        add_weekend_objective(rank, "core", "crew seats toward three volunteers")
    # Keep weekend ALS driving ahead of fourth night volunteers.
    for rank in (2, 3):
        add_weekend_objective(rank, "als", "ALS shifts with EVDT")
    # Soft priority only: after coverage and higher-ranked weekend stages,
    # assign EVDTs to ALS rather than BLS shifts when the feasible options allow.
    objectives.append(("Other ALS shifts with EVDT", sum(als_other)))
    # Build Saturday/Sunday crews toward three volunteers before adding fourth
    # Friday/Saturday-night seats.
    for rank in (2, 3):
        add_weekend_objective(rank, "core", "crew seats toward three volunteers")
    for rank in (0, 1):
        add_weekend_objective(rank, "seats", "volunteer seats filled")
    for rank in (2, 3):
        add_weekend_objective(rank, "seats", "volunteer seats filled")
    objectives.extend([
        ("Ambulance hours within caps", sum(amb_hours)),
        ("Campus blocks with a responder", sum(campus_coverage)),
        ("Campus hours within caps", sum(cr_hours)),
        ("Wellness Wagon shifts filled", sum(v for entries in wellness_by_key.values() for _, v in entries)),
    ])
    # The list above remains the canonical/default ranking. Only its sequence is
    # configurable; the expressions and hard scheduling constraints are unchanged.
    objective_by_id = dict(zip(DEFAULT_OBJECTIVE_ORDER, objectives, strict=True))
    if any(name != OBJECTIVE_LABELS[identifier]
           for identifier, (name, _) in objective_by_id.items()):
        raise RuntimeError("Objective IDs and solver stages are out of sync")
    by_service = {"ambulance": amb_by_key, "campus": cr_by_key, "wellness": wellness_by_key}
    for rule in custom_priorities:
        objective_by_id[rule["id"]] = (
            f"Custom ({rule['direction']}): {rule['label']} [{rule['id']}]",
            _custom_objective(model, rule, people, providers, by_service),
            rule["direction"],
        )
    objectives = [objective_by_id[identifier] for identifier in objective_order]
    # Avoid sub-hour class conflicts after all higher-priority objectives tie.
    soft_conflict_count = sum(
        var for key, entries in cr_by_key.items()
        for pi, var in entries if key in people[pi].campus_soft_conflicts
    )
    if not isinstance(soft_conflict_count, int):
        objectives.append(("Avoid sub-hour class conflicts", -soft_conflict_count))
    values, stages = _optimize(model, objectives, time_limit_s, workers)
    result = Schedule(
        {k: [] for k in sorted(providers)},
        {k: [] for k in campus_keys},
        {k: [] for k in wellness_keys},
        stages,
    )
    for person in people:
        person.campus_assigned = []
        person.wellness_assigned = []
        if isinstance(person, Volunteer):
            person.assigned = []
    for variables, assignments, attribute in (
        (x, result.ambulance, "assigned"),
        (y, result.campus, "campus_assigned"),
        (z, result.wellness, "wellness_assigned"),
    ):
        for (pi, key), var in variables.items():
            if values[var.index]:
                person = people[pi]
                assignments[key].append(person)
                getattr(person, attribute).append(key)
    return result
