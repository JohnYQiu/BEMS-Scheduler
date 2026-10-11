"""Validate required ambulance crews against the latest local form export."""

from __future__ import annotations

from collections import defaultdict
from datetime import date

from models import Volunteer, ambulance_capacity
from parse_form import load_all_responses


def load_volunteers(config):
    volunteers, _ = load_all_responses(
        config.form_csv, config.dates[0], config.dates[-1], config.overrides,
        require_driver_status=config.requires_driver_status)
    return volunteers


def validate_forced_crews(config, volunteers):
    """Raise a specific error before solving; solver still enforces all hard limits."""
    by_email = {person.email: person for person in volunteers}
    by_shift = defaultdict(list)
    for crew in config.forced_crews:
        key = (date.fromisoformat(crew["date"]), crew["shift"])
        for email in crew["members"]:
            person = by_email.get(email)
            if not isinstance(person, Volunteer):
                raise ValueError(f"Forced crew {key}: {email} is not an ambulance EMT in the current form responses")
            if key not in person.available:
                raise ValueError(f"Forced crew {key}: {email} is not available for this shift in the current form responses")
            by_shift[key].append(person)
    for key, people in by_shift.items():
        cap = ambulance_capacity(*key, config.providers[key])
        if len(people) > cap:
            raise ValueError(f"Forced crew {key}: {len(people)} people exceed {cap} available seats")
        if config.providers[key] == "ALS" and sum(not p.is_evdt for p in people) > cap - 1:
            raise ValueError(f"Forced crew {key}: leave an ALS EVDT seat available")
    return by_shift


def crew_options(config):
    """Information for the local editor; derived from saved CSV/ZIP, never online."""
    volunteers = load_volunteers(config)
    shifts = []
    for key, provider in sorted(config.providers.items()):
        candidates = sorted((p for p in volunteers if key in p.available),
                            key=lambda p: (p.last_name.lower(), p.first_name.lower(), p.email))
        shifts.append({
            "key": f"{key[0]}:{key[1]}",
            "label": f"{key[0].strftime('%a %b %d')} · {key[1]} · {provider}",
            "capacity": ambulance_capacity(*key, provider),
            "provider": provider,
            "candidates": [dict(email=p.email, name=p.full_name or p.email,
                                certification=p.certification) for p in candidates],
        })
    return shifts
