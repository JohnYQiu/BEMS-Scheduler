"""
main.py
=======
Entry point for the Brown EMS scheduling system.

Usage:
    python main.py [config.json]

Pipeline:
  1. Parse the Google Form CSV export        (parse_form.py)
  2. Validate availability, print strike list (validate.py)
  3. Solve the ambulance schedule             (ambulance_solver.py, CP-SAT)
  4. Solve the campus response schedule       (campus_solver.py, CP-SAT)
  5. Export xlsx + print summary              (output.py)

See README.md for the config.json format and FORM_GUIDE.md for how the
Google Form must be structured for the parser.
"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

from ambulance_solver import solve_ambulance
from campus_solver import solve_campus
from models import block_dates, expand_als_entries, expand_blackout_period
from output import (
    collect_warnings,
    export_master_schedule_csv,
    export_schedule_xlsx,
    print_summary,
    print_warnings,
)
from parse_form import load_all_responses
from validate import (
    AvailabilityRequirements,
    check_ambulance_requirements,
    check_bert_requirements,
    check_total_available_hours,
    print_availability_summary,
    print_hours_warnings,
    print_strike_list,
)


def load_config(path: str) -> dict:
    if not Path(path).exists():
        sys.exit(f"ERROR: config not found at '{path}'.")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_blackout_periods(cfg: dict) -> set:
    slots = set()
    for period in cfg.get("blackout_periods", []):
        try:
            slots |= expand_blackout_period(
                date.fromisoformat(period["start_date"]), period["start_shift"],
                date.fromisoformat(period["end_date"]), period["end_shift"],
            )
        except (KeyError, ValueError) as e:
            print(f"  [WARN] Skipping malformed blackout period {period}: {e}")
    return slots


def main() -> None:
    config_path = sys.argv[1] if len(sys.argv) > 1 else "config.json"
    print("\n▶  Loading configuration...")
    cfg = load_config(config_path)

    try:
        block_start = date.fromisoformat(cfg["block_start"])
        block_end = date.fromisoformat(cfg["block_end"])
    except (KeyError, ValueError) as e:
        sys.exit(f"ERROR: invalid block_start/block_end in config: {e}")

    hours = cfg.get("hours", {})
    ambulance_required = int(hours.get("ambulance_emt", 12))
    campus_emt_required = int(hours.get("campus_emt", 3))
    campus_bert_required = int(hours.get("campus_bert", 6))
    responders_per_block = int(cfg.get("campus_responders_per_block", 2))
    require_campus_driver = cfg.get("campus_driver_policy", "prefer") == "require"
    time_limit_s = float(cfg.get("solver_time_limit_s", 30))
    reqs = AvailabilityRequirements.from_config(cfg)

    schedule_dates = block_dates(block_start, block_end)
    als_shifts = expand_als_entries(cfg.get("als_shifts", []), schedule_dates)
    blackout_slots = parse_blackout_periods(cfg)

    print(f"  Block: {block_start} → {block_end}  ({len(schedule_dates)} days)")
    print(f"  ALS shift slots: {len(als_shifts)}  |  Blackout slots: {len(blackout_slots)}")
    print(f"  Required hours — ambulance EMT: {ambulance_required}h, "
          f"campus EMT: {campus_emt_required}h, campus BERT: {campus_bert_required}h")

    # ── 1. Parse ─────────────────────────────────────────────────────────────
    form_csv = cfg.get("form_csv", "form_responses.csv")
    print(f"\n▶  Parsing form responses from '{form_csv}'...")
    if not Path(form_csv).exists():
        sys.exit(f"ERROR: form CSV not found at '{form_csv}'.")
    volunteers, bert_members = load_all_responses(form_csv, block_start, block_end)
    if not volunteers:
        sys.exit("ERROR: no Ambulance EMT volunteers found.")

    # ── 2. Validate ──────────────────────────────────────────────────────────
    print("\n▶  Validating availability submissions...")
    emt_violations = check_ambulance_requirements(volunteers, reqs)
    bert_violations = check_bert_requirements(bert_members, reqs)
    print_strike_list(emt_violations, "AMBULANCE EMT AVAILABILITY")
    print_strike_list(bert_violations, "BERT AVAILABILITY")
    print_hours_warnings(
        check_total_available_hours(volunteers, ambulance_required), ambulance_required
    )
    print_availability_summary(volunteers, schedule_dates, blackout_slots)

    # ── 3. Solve ambulance ───────────────────────────────────────────────────
    print("▶  Solving ambulance schedule (CP-SAT)...")
    assignments = solve_ambulance(
        volunteers, schedule_dates, als_shifts, blackout_slots,
        required_hours=ambulance_required, time_limit_s=time_limit_s,
    )

    # ── 4. Solve campus ──────────────────────────────────────────────────────
    print("▶  Solving campus response schedule (CP-SAT)...")
    campus_assignments = solve_campus(
        volunteers, bert_members, schedule_dates,
        responders_per_block=responders_per_block,
        emt_required_hours=campus_emt_required,
        bert_required_hours=campus_bert_required,
        require_driver=require_campus_driver,
        time_limit_s=time_limit_s,
    )

    # ── 5. Output ────────────────────────────────────────────────────────────
    print_summary(
        assignments, campus_assignments, volunteers, bert_members, als_shifts,
        ambulance_required, campus_emt_required, campus_bert_required, responders_per_block,
    )
    warnings = collect_warnings(
        assignments, als_shifts, volunteers, ambulance_required,
        campus_assignments, responders_per_block,
    )
    print_warnings(warnings)
    print("▶  Exporting...")
    export_schedule_xlsx(
        assignments, campus_assignments, volunteers + bert_members,
        cfg.get("output_xlsx", cfg.get("output_csv", "schedule_output.xlsx")),
        als_shifts=als_shifts,
        violations=emt_violations + bert_violations,
        ambulance_required=ambulance_required,
        campus_emt_required=campus_emt_required,
        campus_bert_required=campus_bert_required,
        responders_per_block=responders_per_block,
    )
    master_export = cfg.get("master_schedule_export", {})
    if master_export.get("enabled", True):
        export_master_schedule_csv(
            assignments,
            campus_assignments,
            block_start,
            output_path=master_export.get("path", "master_schedule.csv"),
            block=master_export.get("block", "F26B1"),
            daynum_start=int(master_export.get("daynum_start", 810)),
            vehicle=master_export.get("vehicle", "R1"),
        )
    print("✓  Done.\n")


if __name__ == "__main__":
    main()
