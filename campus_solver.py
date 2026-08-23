"""
campus_solver.py
================
Campus Response scheduler (CP-SAT), run after the ambulance solve.

Staffs the weekday A/B/C/D blocks with BERT members and ambulance EMTs.
Campus hours are block requirements (e.g. EMT 3h = 1 block, BERT 6h = 2
blocks) that double as caps — nobody is scheduled above their requirement.

Hard constraints
----------------
  - Only assign people to blocks they marked / implied available.
  - At most `responders_per_block` people per block.
  - Nobody exceeds their campus hour requirement.
  - Ambulance EMTs never overlap their assigned ambulance shifts.
  - A staffed campus block always has a driver-eligible responder (EVDT or
    Authorized).  A block without one stays open rather than placing an EMT
    or BERT member in the S1 driver seat.

Objective (highest priority first)
----------------------------------
  1. Every block has at least one responder.
  2. Everyone reaches their campus hour requirement.
  3. Every block reaches the full responder target.
  4. Mild spread: prefer a person's blocks on different, non-adjacent days.
"""

from __future__ import annotations

from datetime import date

from ortools.sat.python import cp_model

from models import (
    CAMPUS_BLOCK_HOURS,
    BertMember,
    ShiftKey,
    Volunteer,
    all_campus_keys,
    campus_ambulance_overlap,
)

W_COVER = 1_000_000     # per block with >= 1 responder
W_SHORTFALL = -30_000   # per hour a person falls short of their campus requirement
W_FULL = 60_000         # per additional responder up to the block target
W_SAME_DAY = -50        # per same-person same-day block pair
W_ADJACENT = -25        # per same-person adjacent-day block pair


def solve_campus(
    volunteers: list[Volunteer],
    bert_members: list[BertMember],
    schedule_dates: list[date],
    responders_per_block: int = 2,
    emt_required_hours: int = 3,
    bert_required_hours: int = 6,
    time_limit_s: float = 60.0,
) -> dict[ShiftKey, list]:
    """
    Returns {(date, block): [assigned people]} for every weekday block and
    fills each person's .campus_assigned list.
    """
    block_keys = all_campus_keys(schedule_dates)
    people = list(volunteers) + list(bert_members)

    def required_for(p) -> int:
        return emt_required_hours if isinstance(p, Volunteer) else bert_required_hours

    model = cp_model.CpModel()

    # Keyed by person index, not email: a dual submission (EMT + BERT rows
    # sharing an email) must stay two separate people in the model.
    y: dict[tuple[int, ShiftKey], cp_model.IntVar] = {}
    for pi, p in enumerate(people):
        for key in p.campus_available:
            if key not in set(block_keys):
                continue
            if isinstance(p, Volunteer) and any(
                d == key[0] and campus_ambulance_overlap(key[1], s) for (d, s) in p.assigned
            ):
                continue  # overlaps an ambulance shift they were just given
            y[(pi, key)] = model.new_bool_var(f"y_{pi}_{key[0]}_{key[1]}")

    terms = []

    for key in block_keys:
        assigned = [y[(pi, key)] for pi, p in enumerate(people) if (pi, key) in y]
        if not assigned:
            continue
        model.add(sum(assigned) <= responders_per_block)

        covered = model.new_bool_var(f"cov_{key[0]}_{key[1]}")
        model.add(sum(assigned) >= 1).only_enforce_if(covered)
        # "covered" means genuinely coverable: an uncovered block must not
        # receive a non-driver merely to improve that person's hour total.
        model.add(sum(assigned) == 0).only_enforce_if(covered.negated())

        driver_vars = [
            y[(pi, key)]
            for pi, p in enumerate(people)
            if (pi, key) in y and getattr(p, "is_driver", False)
        ]
        if driver_vars:
            model.add(sum(driver_vars) >= 1).only_enforce_if(covered)
        else:
            # No person who could occupy S1 is available, so preserve the
            # open block for a manual dispatcher decision.
            model.add(covered == 0)
        terms.append(W_COVER * covered)

        # Reward each responder beyond the first, up to the target.
        extra = model.new_int_var(0, max(0, responders_per_block - 1), f"extra_{key[0]}_{key[1]}")
        model.add(extra <= sum(assigned) - 1).only_enforce_if(covered)
        model.add(extra == 0).only_enforce_if(covered.negated())
        terms.append(W_FULL * extra)

    for pi, p in enumerate(people):
        pairs = [(key, y[(pi, key)]) for key in sorted(p.campus_available) if (pi, key) in y]
        required = required_for(p)
        max_blocks = required // CAMPUS_BLOCK_HOURS
        hours = sum(CAMPUS_BLOCK_HOURS * var for _, var in pairs) if pairs else 0
        model.add(hours <= required)

        short = model.new_int_var(0, required, f"short_{pi}")
        model.add(short == required - hours)
        terms.append(W_SHORTFALL * short)

        if max_blocks > 1:
            for i, (k1, var1) in enumerate(pairs):
                for k2, var2 in pairs[i + 1:]:
                    gap = abs((k2[0] - k1[0]).days)
                    weight = W_SAME_DAY if gap == 0 else W_ADJACENT if gap == 1 else 0
                    if weight:
                        both = model.new_bool_var(f"pair_{pi}_{k1}_{k2}")
                        model.add(var1 + var2 - 1 <= both)
                        terms.append(weight * both)

    model.maximize(sum(terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_s
    solver.parameters.random_seed = 0
    solver.parameters.num_workers = 8
    status = solver.solve(model)
    status_name = solver.status_name(status)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(f"Campus solver failed: {status_name}")
    print(f"  Solver status: {status_name} (objective {solver.objective_value:,.0f}, "
          f"{solver.wall_time:.1f}s)")

    assignments: dict[ShiftKey, list] = {key: [] for key in block_keys}
    for pi, p in enumerate(people):
        p.campus_assigned = []
        for key in sorted(p.campus_available):
            var = y.get((pi, key))
            if var is not None and solver.value(var):
                assignments[key].append(p)
                p.campus_assigned.append(key)
    return assignments
