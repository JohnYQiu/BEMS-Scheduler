from __future__ import annotations

"""
main.py
=======
Entry point for the Brown EMS scheduling system.

Usage:
    python main.py

Reads config.json, then runs the full pipeline:
  1. Parse form responses
  2. Validate availability & print strike list
  3. Build and run the schedule
  4. Print and export results

ALS coverage: set "als_shifts" to a list of { "date", "shift" } where shift is
DAY or NIGHT. On weekdays, DAY marks both AM and PM as ALS. Omit "als_shifts"
to treat every shift in the block as ALS. Legacy "non_als_shifts" is still
honored when "als_shifts" is omitted.

Hour limits: optional "hour_limits" with "ambulance" (target_hours, soft_max_hours,
hard_max_hours) and "campus" (ambulance_emt_max_hours, bert_max_hours). See config.json.
"""

import json
import sys
from datetime import date, timedelta
from pathlib import Path

from parse_form  import load_all_responses
from validate    import (check_availability_requirements, print_strike_list,
                         check_total_available_hours,    print_hours_warnings,
                         print_availability_summary)
from scheduler   import run_schedule, _build_blackout_slots, AmbulanceHourLimits
from campus_scheduler import run_campus_schedule, CampusHourCaps
from output      import export_schedule_xlsx, print_summary, print_warnings


def load_config(path: str = "config.json") -> dict:
    if not Path(path).exists():
        print(f"\n  ERROR: config.json not found at '{path}'.")
        sys.exit(1)
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_shift_list(shift_list: list) -> set:
    result = set()
    for item in shift_list:
        try:
            d = date.fromisoformat(item["date"])
            s = item["shift"].upper()
            result.add((d, s))
        except (KeyError, ValueError) as e:
            print(f"  [WARN] Skipping malformed shift entry {item}: {e}")
    return result


def all_shift_keys_in_block(schedule_dates: list) -> set:
    s = set()
    for d in schedule_dates:
        for st in (["DAY", "NIGHT"] if d.weekday() >= 5 else ["AM", "PM", "NIGHT"]):
            s.add((d, st))
    return s


def expand_als_shifts_day_night(entries: list, schedule_dates: list) -> set:
    """
    Config entries use 12h blocks: DAY or NIGHT for a calendar date.
    - Weekday DAY  -> ALS on both AM and PM that day.
    - Weekend DAY -> ALS on the single DAY shift (0700–1900).
    - NIGHT -> ALS on NIGHT for that date (weekday or weekend).

    Each entry may be {"date": "YYYY-MM-DD", "shift": "DAY"|"NIGHT"} or the
    compact string "YYYY-MM-DD:DAY" / "YYYY-MM-DD:NIGHT".
    """
    schedule_set = set(schedule_dates)
    out = set()
    for raw in entries:
        item = raw
        if isinstance(raw, str):
            s = raw.strip()
            if ":" not in s:
                print(f"  [WARN] Skipping als_shifts entry {raw!r} (expected date:DAY or date:NIGHT)")
                continue
            date_s, _, kind_s = s.partition(":")
            item = {"date": date_s.strip(), "shift": kind_s.strip()}
        try:
            d = date.fromisoformat(item["date"])
            kind = str(item["shift"]).strip().upper()
        except (KeyError, ValueError, TypeError) as e:
            print(f"  [WARN] Skipping als_shifts entry {raw!r}: {e}")
            continue
        if d not in schedule_set:
            print(f"  [WARN] als_shifts date {d} outside scheduling block; skipped")
            continue
        dow = d.weekday()
        if kind == "NIGHT":
            out.add((d, "NIGHT"))
        elif kind == "DAY":
            if dow < 5:
                out.add((d, "AM"))
                out.add((d, "PM"))
            else:
                out.add((d, "DAY"))
        else:
            print(f"  [WARN] als_shifts shift must be DAY or NIGHT, got {kind!r}")
    return out


def build_als_shifts_set(cfg: dict, schedule_dates: list) -> set:
    """
    - If als_shifts is present (list): only those expanded keys are ALS.
    - If als_shifts is missing: every shift in the block is ALS (legacy default).
    - Legacy key non_als_shifts (if present and als_shifts missing): subtract from full set.
    """
    all_keys = all_shift_keys_in_block(schedule_dates)
    if "als_shifts" in cfg:
        raw = cfg.get("als_shifts")
        if raw is None:
            return all_keys
        if not isinstance(raw, list):
            print("  [WARN] als_shifts must be a list; using all shifts ALS")
            return all_keys
        return expand_als_shifts_day_night(raw, schedule_dates)
    if cfg.get("non_als_shifts"):
        return all_keys - parse_shift_list(cfg["non_als_shifts"])
    return all_keys


def load_hour_limits(cfg: dict) -> tuple[AmbulanceHourLimits, CampusHourCaps]:
    """
    hour_limits.ambulance: target_hours, soft_max_hours, hard_max_hours
    hour_limits.campus: ambulance_emt_max_hours, bert_max_hours
    """
    hl = cfg.get("hour_limits") or {}
    amb = hl.get("ambulance") or {}
    camp = hl.get("campus") or {}
    return (
        AmbulanceHourLimits(
            target_hours=int(amb.get("target_hours", 18)),
            soft_max_hours=int(amb.get("soft_max_hours", amb.get("target_hours", 18))),
            hard_max_hours=int(amb.get("hard_max_hours", 24)),
        ),
        CampusHourCaps(
            ambulance_emt_max_hours=int(camp.get("ambulance_emt_max_hours", 6)),
            bert_max_hours=int(camp.get("bert_max_hours", 9)),
        ),
    )


def main():
    print("\n▶  Loading configuration...")
    cfg = load_config()

    try:
        block_start = date.fromisoformat(cfg["block_start"])
        block_end   = date.fromisoformat(cfg["block_end"])
    except (KeyError, ValueError) as e:
        print(f"  ERROR: Invalid block_start/block_end in config.json: {e}")
        sys.exit(1)

    form_csv   = cfg.get("form_csv",   "form_responses.csv")
    output_csv = cfg.get("output_csv", "schedule_output.csv")

    # Build full schedule dates
    schedule_dates = [
        block_start + timedelta(days=i)
        for i in range((block_end - block_start).days + 1)
    ]

    als_shifts = build_als_shifts_set(cfg, schedule_dates)
    amb_limits, campus_caps = load_hour_limits(cfg)

    # Build blackout slots
    blackout_slots = set()
    for period in cfg.get("blackout_periods", []):
        try:
            s_date  = date.fromisoformat(period["start_date"])
            s_shift = period["start_shift"]
            e_date  = date.fromisoformat(period["end_date"])
            e_shift = period["end_shift"]
            blackout_slots |= _build_blackout_slots(s_date, s_shift, e_date, e_shift)
        except (KeyError, ValueError) as e:
            print(f"  [WARN] Skipping malformed blackout period {period}: {e}")

    print(f"  Block: {block_start} → {block_end}  ({len(schedule_dates)} days)")
    print(f"  ALS shifts: {len(als_shifts)}  |  Blackout slots: {len(blackout_slots)}")
    print(
        f"  Ambulance hours: target {amb_limits.target_hours} / soft max {amb_limits.soft_max_hours} / "
        f"hard max {amb_limits.hard_max_hours}"
    )
    print(
        f"  Campus caps: ambulance EMT ≤ {campus_caps.ambulance_emt_max_hours}h, "
        f"BERT ≤ {campus_caps.bert_max_hours}h"
    )

    # ── Step 1: Parse form ────────────────────────────────────────────────────
    print(f"\n▶  Parsing form responses from '{form_csv}'...")
    if not Path(form_csv).exists():
        print(f"  ERROR: Form CSV not found at '{form_csv}'.")
        sys.exit(1)

    volunteers, bert_members = load_all_responses(form_csv, block_start, block_end)

    if not volunteers:
        print("  ERROR: No Ambulance EMT volunteers found.")
        sys.exit(1)

    # ── Step 2: Validate ──────────────────────────────────────────────────────
    print("\n▶  Validating availability submissions...")
    violations     = check_availability_requirements(volunteers)
    hours_warnings = check_total_available_hours(volunteers, min_hours=amb_limits.target_hours)
    print_strike_list(violations)
    print_hours_warnings(hours_warnings, min_hours=amb_limits.target_hours)
    print_availability_summary(volunteers, schedule_dates, blackout_slots)

    # ── Step 3: Schedule ──────────────────────────────────────────────────────
    print("▶  Running scheduler...")
    all_shifts = run_schedule(
        volunteers, schedule_dates, als_shifts, blackout_slots, hour_limits=amb_limits,
    )

    print("▶  Running campus responder scheduler...")
    campus_shifts = run_campus_schedule(
        volunteers, bert_members, schedule_dates, responders_per_block=2, hour_caps=campus_caps,
    )

    # ── Step 4: Output ────────────────────────────────────────────────────────
    print_summary(all_shifts, volunteers, ambulance_target_hours=amb_limits.target_hours)
    print_warnings(all_shifts)
    print("▶  Exporting...")
    export_schedule_xlsx(
        all_shifts,
        volunteers + bert_members,
        output_csv,
        violations=violations,
        campus_shifts=campus_shifts,
        ambulance_target_hours=amb_limits.target_hours,
        campus_emt_max_hours=campus_caps.ambulance_emt_max_hours,
        campus_bert_max_hours=campus_caps.bert_max_hours,
    )
    print("✓  Done.\n")


if __name__ == "__main__":
    main()