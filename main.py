"""Load responses, solve jointly, validate that result, then export."""

from __future__ import annotations

import argparse
from pathlib import Path

from configuration import load_config
from output import export_master_schedule_csv, export_schedule_xlsx, print_summary
from parse_form import load_all_responses
from solver import solve_schedule
from validation import validate_schedule


def run(config_path, check_only=False):
    config = load_config(config_path)
    volunteers, bert_members = load_all_responses(
        config.form_csv, config.dates[0], config.dates[-1], config.overrides)
    people = volunteers + bert_members
    if not people:
        raise ValueError("No scheduling members found in the response CSV")
    print(f"Block: {config.dates[0]} to {config.dates[-1]}")
    print("Solving ambulance and campus response together...")
    schedule = solve_schedule(
        people, config.providers, config.campus_keys, config.caps,
        config.campus_capacity, config.locks, config.time_limit_s, config.workers)
    errors = validate_schedule(schedule, people, config.providers, config.campus_keys,
                               config.caps, config.campus_capacity, config.locks)
    if errors:
        raise RuntimeError("Schedule failed validation; nothing exported:\n" + "\n".join(errors))
    print("All hard scheduling rules verified against the resulting assignments.")
    print_summary(schedule, people, config.providers, config.caps)
    if not check_only:
        path = export_schedule_xlsx(schedule, people, config.providers, config.caps,
                                    config.output_xlsx, config.campus_capacity)
        print(f"Workbook: {path}")
        if config.master_csv:
            path = export_master_schedule_csv(
                schedule, config.providers, config.dates[0], config.master_csv,
                config.block, config.daynum_start, config.campus_capacity)
            print(f"Master Schedule CSV: {path}")
    return schedule


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", nargs="?", default=Path(__file__).with_name("config.json"))
    parser.add_argument("--check", action="store_true", help="solve and validate without exporting")
    args = parser.parse_args()
    try:
        run(args.config, args.check)
    except (OSError, ValueError, KeyError, TypeError, RuntimeError) as error:
        parser.exit(1, f"ERROR: {error}\n")


if __name__ == "__main__":
    main()
