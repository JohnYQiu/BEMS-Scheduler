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
"""

import json
import sys
from datetime import date, timedelta
from pathlib import Path

from parse_form  import load_responses
from validate    import (check_availability_requirements, print_strike_list,
                         check_total_available_hours,    print_hours_warnings,
                         print_availability_summary)
from scheduler   import run_schedule, _build_blackout_slots
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
    seed       = cfg.get("random_seed", 42)

    # Build full schedule dates
    schedule_dates = [
        block_start + timedelta(days=i)
        for i in range((block_end - block_start).days + 1)
    ]

    # Build ALS shifts (all shifts minus non-ALS exceptions)
    non_als_shifts = parse_shift_list(cfg.get("non_als_shifts", []))
    als_shifts = {
        (d, s)
        for d in schedule_dates
        for s in (["DAY", "NIGHT"] if d.weekday() >= 5 else ["AM", "PM", "NIGHT"])
    } - non_als_shifts

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

    # ── Step 1: Parse form ────────────────────────────────────────────────────
    print(f"\n▶  Parsing form responses from '{form_csv}'...")
    if not Path(form_csv).exists():
        print(f"  ERROR: Form CSV not found at '{form_csv}'.")
        sys.exit(1)

    volunteers = load_responses(form_csv, block_start, block_end)

    if not volunteers:
        print("  ERROR: No Ambulance EMT volunteers found.")
        sys.exit(1)

    # ── Step 2: Validate ──────────────────────────────────────────────────────
    print("\n▶  Validating availability submissions...")
    violations     = check_availability_requirements(volunteers)
    hours_warnings = check_total_available_hours(volunteers)
    print_strike_list(violations)
    print_hours_warnings(hours_warnings)
    print_availability_summary(volunteers, schedule_dates, blackout_slots)

    # ── Step 3: Schedule ──────────────────────────────────────────────────────
    print("▶  Running scheduler...")
    all_shifts = run_schedule(volunteers, schedule_dates, als_shifts,
                              blackout_slots, seed=seed)

    # ── Step 4: Output ────────────────────────────────────────────────────────
    print_summary(all_shifts, volunteers)
    print_warnings(all_shifts)
    print("▶  Exporting...")
    export_schedule_xlsx(all_shifts, volunteers, output_csv)
    print("✓  Done.\n")


if __name__ == "__main__":
    main()