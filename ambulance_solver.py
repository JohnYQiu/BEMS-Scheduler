"""
ambulance_solver.py
===================
Ambulance shift scheduler built on Google OR-Tools CP-SAT.

Instead of greedily filling shifts one at a time, the whole block is solved as
one constraint-optimization problem: every possible (volunteer, shift)
assignment is a boolean variable, the rules below are constraints, and the
solver finds the assignment that maximizes a prioritized objective.

Hard constraints
----------------
  - Only assign people to shifts they marked available.
  - Crew caps per shift (see models.crew_cap).
  - Nobody exceeds their required hours ("the requirement is also the max").
  - Rest rules: max 12h continuous work, then >=12h off (models.rest_conflict).
  - ALS shifts always keep one slot that only an EVDT can occupy: if no EVDT
    is on the shift, at most cap-1 people may be assigned.

Objective (highest priority first — weights are separated by magnitude)
----------------------------------------------------------------------
  1. Cover every shift with at least one person.
  2. Get every volunteer to their required hours (minimize shortfall).
  3. Put an EVDT on every ALS shift.
  4. Put a driver (EVDT/Auth) on every night & weekend shift.
  5. Fill as many crew slots as possible — with hours fixed, this favors
     splitting people into 6h shifts over 12h shifts, maximizing coverage.
  6. Spread each person out (small penalties for same-day and adjacent-day
     pairs) and equalize any unavoidable shortfall across people.
"""

from __future__ import annotations

from datetime import date, timedelta

from ortools.sat.python import cp_model

from models import (
    SHIFT_HOURS,
    ShiftKey,
    Volunteer,
    all_shift_keys,
    crew_cap,
    is_big_weekend,
    rest_conflict,
)

# Objective weights, separated by orders of magnitude so higher tiers always win.
W_COVER = 1_000_000        # per shift with >= 1 person
W_SHORTFALL = -30_000      # per hour a volunteer falls short of required hours
W_EVDT_ON_ALS = 150_000    # per ALS shift with an EVDT
W_DRIVER = 40_000          # per night/weekend shift with any driver
W_SLOT = 500               # per filled crew slot
W_SAME_DAY = -200          # per same-person AM+PM pair on one day
W_ADJACENT = -100          # per same-person pair on consecutive days
W_MAX_SHORTFALL = -20_000  # times the largest individual shortfall


def _needs_driver(key: ShiftKey) -> bool:
    """Shifts where a crew without any driver is worth warning about."""
    d, s = key
    return s == "NIGHT" or is_big_weekend(d, s)


def solve_ambulance(
    volunteers: list[Volunteer],
    schedule_dates: list[date],
    als_shifts: set[ShiftKey],
    blackout_slots: set[ShiftKey] | None = None,
    locked_assignments: list[dict] | None = None,
    required_hours: int = 12,
    time_limit_s: float = 60.0,
) -> dict[ShiftKey, list[Volunteer]]:
    """
    Returns {shift_key: [assigned volunteers]} covering every non-blackout
    shift in the block, and fills each volunteer's .assigned list.
    """
    blackout_slots = blackout_slots or set()
    shift_keys = [k for k in all_shift_keys(schedule_dates) if k not in blackout_slots]

    model = cp_model.CpModel()

    # x[(email, key)] = 1 when that volunteer works that shift.
    x: dict[tuple[str, ShiftKey], cp_model.IntVar] = {}
    for v in volunteers:
        for key in v.available:
            if key in set(shift_keys):
                x[(v.email, key)] = model.new_bool_var(f"x_{v.email}_{key[0]}_{key[1]}")

    def vars_for_shift(key: ShiftKey, pred=None):
        return [
            x[(v.email, key)]
            for v in volunteers
            if (v.email, key) in x and (pred is None or pred(v))
        ]

    def vars_for_volunteer(v: Volunteer):
        return [(key, x[(v.email, key)]) for key in sorted(v.available) if (v.email, key) in x]

    # ── Hard constraints ─────────────────────────────────────────────────────
    for key in shift_keys:
        assigned = vars_for_shift(key)
        if not assigned:
            continue
        cap = crew_cap(*key)
        model.add(sum(assigned) <= cap)
        if key in als_shifts:
            # Reserve one slot for an EVDT: non-EVDTs can never take the last seat.
            non_evdt = vars_for_shift(key, lambda v: not v.is_evdt)
            if non_evdt:
                model.add(sum(non_evdt) <= cap - 1)

    for v in volunteers:
        pairs = vars_for_volunteer(v)
        if not pairs:
            continue
        model.add(sum(SHIFT_HOURS[k[1]] * var for k, var in pairs) <= required_hours)
        for i, (k1, var1) in enumerate(pairs):
            for k2, var2 in pairs[i + 1:]:
                if rest_conflict(k1, k2):
                    model.add(var1 + var2 <= 1)

    # Explicit personnel selections are hard constraints, validated before
    # solving so a typo or unavailable person never silently changes a schedule.
    by_email = {v.email.lower(): v for v in volunteers}
    for lock in locked_assignments or []:
        try:
            key = (date.fromisoformat(lock["date"]), str(lock["shift"]).upper())
            email = str(lock["email"]).lower()
        except (KeyError, TypeError, ValueError) as e:
            raise ValueError(f"Invalid locked ambulance assignment {lock}: {e}") from e
        volunteer = by_email.get(email)
        var = x.get((email, key))
        if volunteer is None or var is None:
            raise ValueError(f"Locked assignment unavailable: {email} on {key[0]} {key[1]}")
        model.add(var == 1)

    # ── Objective ────────────────────────────────────────────────────────────
    terms = []

    for key in shift_keys:
        assigned = vars_for_shift(key)
        if not assigned:
            continue
        covered = model.new_bool_var(f"cov_{key[0]}_{key[1]}")
        model.add(sum(assigned) >= 1).only_enforce_if(covered)
        terms.append(W_COVER * covered)

        if key in als_shifts:
            evdts = vars_for_shift(key, lambda v: v.is_evdt)
            if evdts:
                has_evdt = model.new_bool_var(f"evdt_{key[0]}_{key[1]}")
                model.add(sum(evdts) >= 1).only_enforce_if(has_evdt)
                terms.append(W_EVDT_ON_ALS * has_evdt)

        if _needs_driver(key):
            drivers = vars_for_shift(key, lambda v: v.is_driver)
            if drivers:
                has_driver = model.new_bool_var(f"drv_{key[0]}_{key[1]}")
                model.add(sum(drivers) >= 1).only_enforce_if(has_driver)
                terms.append(W_DRIVER * has_driver)

    shortfalls = []
    for v in volunteers:
        pairs = vars_for_volunteer(v)
        hours = sum(SHIFT_HOURS[k[1]] * var for k, var in pairs) if pairs else 0
        short = model.new_int_var(0, required_hours, f"short_{v.email}")
        model.add(short == required_hours - hours)
        shortfalls.append(short)
        terms.append(W_SHORTFALL * short)

        for _, var in pairs:
            terms.append(W_SLOT * var)
        for i, (k1, var1) in enumerate(pairs):
            for k2, var2 in pairs[i + 1:]:
                gap = abs((k2[0] - k1[0]).days)
                weight = W_SAME_DAY if gap == 0 else W_ADJACENT if gap == 1 else 0
                if weight and not rest_conflict(k1, k2):
                    both = model.new_bool_var(f"pair_{v.email}_{k1}_{k2}")
                    model.add(var1 + var2 - 1 <= both)
                    terms.append(weight * both)

    if shortfalls:
        max_short = model.new_int_var(0, required_hours, "max_shortfall")
        model.add_max_equality(max_short, shortfalls)
        terms.append(W_MAX_SHORTFALL * max_short)

    model.maximize(sum(terms))

    # ── Solve ────────────────────────────────────────────────────────────────
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_s
    solver.parameters.random_seed = 0
    solver.parameters.num_workers = 8
    status = solver.solve(model)
    status_name = solver.status_name(status)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(f"Ambulance solver failed: {status_name}")
    gap = solver.best_objective_bound - solver.objective_value
    print(f"  Solver status: {status_name} (objective {solver.objective_value:,.0f}, "
          f"gap ≤ {gap:,.0f}, {solver.wall_time:.1f}s)")

    # ── Extract ──────────────────────────────────────────────────────────────
    assignments: dict[ShiftKey, list[Volunteer]] = {key: [] for key in shift_keys}
    for v in volunteers:
        v.assigned = []
        for key in sorted(v.available):
            var = x.get((v.email, key))
            if var is not None and solver.value(var):
                assignments[key].append(v)
                v.assigned.append(key)
    return assignments
