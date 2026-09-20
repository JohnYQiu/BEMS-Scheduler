"""Joint ambulance/CR scheduling with hard work rules and staged objectives."""

from __future__ import annotations

from collections import defaultdict
from datetime import timedelta
from time import monotonic

from ortools.sat.python import cp_model

from models import (
    CAMPUS_BLOCK_HOURS, SHIFT_HOURS, START_HOURS, BertMember, HourCaps,
    LockedAssignment, Schedule, ShiftKey, SolveStage, Volunteer,
    crew_cap, is_big_weekend, is_weekend_day, is_weekend_night,
)


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


def _optimize(model, objectives, time_limit_s, workers):
    """Preserve each stage's best result; share one time budget across stages.

    FEASIBLE values may be preserved, but are never described as proven optima.
    If a later solve finds no incumbent, retain the last valid complete schedule.
    """
    deadline = monotonic() + time_limit_s
    stages, values = [], None
    active = [(name, expr) for name, expr in objectives if not isinstance(expr, int)]
    for i, (name, objective) in enumerate(active):
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
                                   "Check locked assignments and increase the time budget if needed.")
            stages.append(SolveStage(name, status_name, None, None, solver.wall_time))
            break
        value = int(solver.value(objective))
        stages.append(SolveStage(name, status_name, value,
                                 solver.best_objective_bound, solver.wall_time))
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
) -> Schedule:
    if time_limit_s <= 0 or type(workers) is not int or workers < 1:
        raise ValueError("Solver time budget and worker count must be positive")
    if type(campus_capacity) is not int or campus_capacity < 1:
        raise ValueError("campus_capacity must be a positive integer")
    if any(kind not in ("ALS", "BLS") for kind in providers.values()):
        raise ValueError("Every ambulance shift must explicitly specify ALS or BLS")
    people = sorted(people, key=lambda p: p.email)
    if any(not p.email for p in people) or len({p.email for p in people}) != len(people):
        raise ValueError("Every person must have one unique, nonempty email")

    model = cp_model.CpModel()
    x, y = {}, {}
    amb_by_key, cr_by_key = defaultdict(list), defaultdict(list)
    by_person = defaultdict(list)
    campus_keys = sorted(set(campus_keys))
    campus_set = set(campus_keys)
    for pi, person in enumerate(people):
        services = [(y, person.campus_available & campus_set, cr_by_key)]
        if isinstance(person, Volunteer):
            services.append((x, person.available & providers.keys(), amb_by_key))
        for variables, available, index in services:
            for key in sorted(available):
                var = model.new_bool_var(f"p{pi}_{key[0]}_{key[1]}")
                variables[pi, key] = var
                index[key].append((pi, var))
                by_person[pi].append((key, var))

    for key, entries in amb_by_key.items():
        model.add(sum(v for _, v in entries) <= crew_cap(*key))
        if providers[key] == "ALS":
            # An unfilled EVDT seat stays open; Auth never satisfies ALS driving.
            model.add(sum(v for pi, v in entries if not people[pi].is_evdt) <= crew_cap(*key) - 1)
    for entries in cr_by_key.values():
        model.add(sum(v for _, v in entries) <= campus_capacity)

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

    coverage, als_nights, als_days, als_other, utility, campus_coverage = [], [], [], [], [], []
    weekend_ready, weekend_core, weekend_seats = [], [], []
    for key in sorted(providers):
        entries = amb_by_key[key]
        coverage.append(_present(model, [v for _, v in entries], f"covered_{key}"))
        if is_big_weekend(*key) and entries:
            crew = [v for _, v in entries]
            drivers = [v for pi, v in entries if people[pi].is_driver]
            evdts = [v for pi, v in entries if people[pi].is_evdt]
            weekend_seats.extend(crew)
            core = model.new_int_var(0, 3, f"weekend_core_{key}")
            model.add_min_equality(core, [sum(crew), 3])
            weekend_core.append(core)
            # BLS: supervisor + EMT, Utility driver + EMT (three volunteers).
            # ALS also needs a separate EVDT on the ambulance while the
            # supervisor treats; one volunteer cannot drive both vehicles.
            requirements = [_at_least(model, crew, 3, f"split_size_{key}"),
                            _at_least(model, drivers, 2 if providers[key] == "ALS" else 1,
                                      f"split_drivers_{key}")]
            if providers[key] == "ALS":
                requirements.append(_at_least(model, evdts, 1, f"split_evdt_{key}"))
            ready = model.new_bool_var(f"split_ready_{key}")
            model.add_min_equality(ready, requirements)
            weekend_ready.append(ready)
        if providers[key] == "ALS":
            evdt = _present(model, [v for pi, v in entries if people[pi].is_evdt], f"evdt_{key}")
            target = als_nights if is_weekend_night(*key) else als_days if is_weekend_day(*key) else als_other
            target.append(evdt)
        elif is_big_weekend(*key):
            # The BLS supervisor drives the ambulance. Auth and EVDT are equal
            # Utility drivers here; weekday BLS has no driver objective.
            utility.append(_present(model, [v for pi, v in entries if people[pi].is_driver], f"utility_{key}"))
    for key in campus_keys:
        campus_coverage.append(_present(model, [v for _, v in cr_by_key[key]], f"campus_{key}"))

    objectives = [
        ("Ambulance shifts with an EMT", sum(coverage)),
        ("ALS Friday/Saturday nights with EVDT", sum(als_nights)),
        ("ALS Saturday/Sunday days with EVDT", sum(als_days)),
        ("Weekend shifts ready for split crew", sum(weekend_ready)),
        ("Weekend crew seats toward three volunteers", sum(weekend_core)),
        ("Weekend BLS shifts with Utility driver", sum(utility)),
        ("Other ALS shifts with EVDT", sum(als_other)),
        ("Weekend volunteer seats filled", sum(weekend_seats)),
        ("Ambulance hours within caps", sum(amb_hours)),
        ("Campus blocks with a responder", sum(campus_coverage)),
        ("Campus hours within caps", sum(cr_hours)),
    ]
    values, stages = _optimize(model, objectives, time_limit_s, workers)
    result = Schedule({k: [] for k in sorted(providers)}, {k: [] for k in campus_keys}, stages)
    for person in people:
        person.campus_assigned = []
        if isinstance(person, Volunteer):
            person.assigned = []
    for variables, assignments, attribute in (
        (x, result.ambulance, "assigned"), (y, result.campus, "campus_assigned")
    ):
        for (pi, key), var in variables.items():
            if values[var.index]:
                person = people[pi]
                assignments[key].append(person)
                getattr(person, attribute).append(key)
    return result
