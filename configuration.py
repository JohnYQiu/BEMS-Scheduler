"""Load one explicit per-block configuration; paths are relative to this file."""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from models import (
    CAMPUS_BLOCKS, HourCaps, LockedAssignment, ShiftKey, all_campus_keys,
    all_shift_keys, block_dates, campus_ambulance_overlap, interval, is_weekend,
)


@dataclass
class Configuration:
    dates: list[date]
    providers: dict[ShiftKey, str]
    campus_keys: list[ShiftKey]
    caps: HourCaps
    campus_capacity: int
    time_limit_s: float
    workers: int
    form_csv: Path
    output_xlsx: Path
    master_csv: Path | None
    block: str
    daynum_start: int
    overrides: dict[str, str]
    locks: list[LockedAssignment]


def _json_file(path: Path):
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _expand_key(raw: str, calendar: set[ShiftKey]) -> list[ShiftKey]:
    date_s, kind = raw.rsplit(":", 1)
    d = date.fromisoformat(date_s)
    kind = kind.strip().upper()
    keys = [(d, "AM"), (d, "PM")] if kind == "DAY" and not is_weekend(d) else [(d, kind)]
    if any(k not in calendar for k in keys):
        raise ValueError(f"Shift is invalid or outside the block: {raw}")
    return keys


def load_config(path: str | Path) -> Configuration:
    path = Path(path).resolve()
    cfg = _json_file(path)
    if not isinstance(cfg, dict):
        raise ValueError("Configuration must be a JSON object")
    dates = block_dates(date.fromisoformat(cfg["block_start"]), date.fromisoformat(cfg["block_end"]))
    calendar = set(all_shift_keys(dates))

    def local_path(value):
        return (path.parent / value).resolve()

    blackouts = set()
    for period in cfg.get("blackout_periods", []):
        first = (date.fromisoformat(period["start_date"]), period["start_shift"].upper())
        last = (date.fromisoformat(period["end_date"]), period["end_shift"].upper())
        lo, hi = interval(first)[0], interval(last)[1]
        if hi <= lo:
            raise ValueError("Blackout period ends before it starts")
        blackouts.update(k for k in calendar if interval(k)[0] < hi and lo < interval(k)[1])
    active = calendar - blackouts
    configured = cfg.get("shift_providers")
    if not isinstance(configured, dict):
        raise ValueError("Set shift_providers in config.json: every active shift needs ALS or BLS")
    providers = {}
    for raw, provider in configured.items():
        for key in _expand_key(raw, calendar):
            if key not in active:
                continue
            if key in providers:
                raise ValueError(f"Duplicate provider setting for {key}; do not overlap DAY with AM/PM")
            providers[key] = provider.upper() if isinstance(provider, str) else provider
    unset = sorted(k for k in active if providers.get(k) not in ("ALS", "BLS"))
    if unset:
        examples = ", ".join(f"{d}:{s}" for d, s in unset[:6])
        raise ValueError(f"Choose ALS or BLS for {len(unset)} active shift(s): {examples}")
    campus_keys = [k for k in all_campus_keys(dates) if not any(
        k[0] == d and campus_ambulance_overlap(k[1], s) for d, s in blackouts)]

    hours = cfg.get("hours", {})
    caps = HourCaps(hours.get("ambulance_emt", 18), hours.get("campus_emt", 6), hours.get("campus_bert", 9))
    capacity = cfg.get("campus_capacity", 2)
    workers = cfg.get("solver_workers", 8)
    time_limit = cfg.get("solver_time_limit_s", 30)
    if type(capacity) is not int or capacity < 1 or type(workers) is not int or workers < 1:
        raise ValueError("campus_capacity and solver_workers must be positive integers")
    if isinstance(time_limit, bool) or not isinstance(time_limit, (int, float)) or not math.isfinite(time_limit) or time_limit <= 0:
        raise ValueError("solver_time_limit_s must be a positive finite number")

    overrides = dict(cfg.get("driver_status_overrides", {}))
    overrides_file = cfg.get("driver_status_overrides_file")
    if overrides_file and local_path(overrides_file).exists():
        local = _json_file(local_path(overrides_file))
        if not isinstance(local, dict):
            raise ValueError("Driver overrides must be a JSON object keyed by email")
        overrides.update(local)

    raw_locks = list(cfg.get("locked_assignments", []))
    locks_file = cfg.get("locked_assignments_file", cfg.get("locked_ambulance_assignments_file"))
    if locks_file and local_path(locks_file).exists():
        local = _json_file(local_path(locks_file))
        if not isinstance(local, list):
            raise ValueError("Locked assignments must be a JSON array")
        raw_locks.extend(local)
    locks = []
    for item in raw_locks:
        key = (date.fromisoformat(item["date"]), item["shift"].upper())
        if key[1] not in CAMPUS_BLOCKS and key not in active:
            raise ValueError(f"Locked ambulance shift is inactive: {key}")
        if key[1] in CAMPUS_BLOCKS and key not in campus_keys:
            raise ValueError(f"Locked campus block is inactive: {key}")
        locks.append(LockedAssignment(item["email"].strip().lower(), key))

    master = cfg.get("master_schedule_export", {})
    return Configuration(
        dates, providers, campus_keys, caps, capacity, float(time_limit), workers,
        local_path(cfg.get("form_csv", "inputs/responses.csv")),
        local_path(cfg.get("output_xlsx", "outputs/schedule.xlsx")),
        local_path(master.get("path", "outputs/master_schedule.csv")) if master.get("enabled", True) else None,
        master.get("block", ""), int(master.get("daynum_start", dates[0].month * 100 + dates[0].day)),
        overrides, locks,
    )
