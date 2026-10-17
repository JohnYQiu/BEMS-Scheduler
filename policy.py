"""Validated, declarative soft priorities; hard scheduling rules live elsewhere.

The built-in default order exactly matches the historical solver sequence. Custom
priorities select an input-derived subset of assignments, then maximize or
minimize its count, coverage, or hours. No user-supplied code is executed.
"""

import re


CUSTOM_PRIORITY_CATALOG = {
    "services": {"ambulance": ["AM", "PM", "DAY", "NIGHT"],
                 "campus": ["A", "B", "C", "D"],
                 "wellness": ["WAM", "WPM"]},
    "metrics": ["assignments", "covered_shifts", "hours"],
    "member_filters": ["any", "driver", "evdt", "auth", "non_driver", "emt", "bert"],
    "directions": ["maximize", "minimize"],
}
CUSTOM_PRIORITY_KEYS = frozenset({"id", "label", "service", "metric", "member_filter",
                                  "provider", "weekdays", "shifts", "direction"})
MAX_CUSTOM_PRIORITIES = 24
HARD_REQUIREMENT_CATALOG = {
    "services": {"ambulance": ["AM", "PM", "DAY", "NIGHT"],
                 "campus": ["A", "B", "C", "D"]},
    "member_filters": ["any", "driver", "evdt", "auth", "non_driver", "emt", "bert"],
    "max_per_shift": 4,
}
HARD_REQUIREMENT_KEYS = frozenset({"id", "label", "service", "member_filter", "provider",
                                   "weekdays", "shifts", "min_people"})
MAX_HARD_REQUIREMENTS = 24
CUSTOM_ID = re.compile(r"custom:[a-z0-9][a-z0-9_-]{0,39}\Z")
HARD_ID = re.compile(r"hard:[a-z0-9][a-z0-9_-]{0,39}\Z")

OBJECTIVE_STAGES = (
    ("ambulance_coverage", "Ambulance shifts with an EMT"),
    ("friday_als_evdt", "Friday nights: ALS shifts with EVDT"),
    ("friday_core", "Friday nights: crew seats toward three volunteers"),
    ("saturday_night_als_evdt", "Saturday nights: ALS shifts with EVDT"),
    ("saturday_night_core", "Saturday nights: crew seats toward three volunteers"),
    ("saturday_day_als_evdt", "Saturday days: ALS shifts with EVDT"),
    ("sunday_day_als_evdt", "Sunday days: ALS shifts with EVDT"),
    ("other_als_evdt", "Other ALS shifts with EVDT"),
    ("saturday_day_core", "Saturday days: crew seats toward three volunteers"),
    ("sunday_day_core", "Sunday days: crew seats toward three volunteers"),
    ("friday_seats", "Friday nights: volunteer seats filled"),
    ("saturday_night_seats", "Saturday nights: volunteer seats filled"),
    ("saturday_day_seats", "Saturday days: volunteer seats filled"),
    ("sunday_day_seats", "Sunday days: volunteer seats filled"),
    ("ambulance_hours", "Ambulance hours within caps"),
    ("campus_coverage", "Campus blocks with a responder"),
    ("campus_hours", "Campus hours within caps"),
    ("wellness_coverage", "Wellness Wagon shifts filled"),
)
DEFAULT_OBJECTIVE_ORDER = tuple(identifier for identifier, _ in OBJECTIVE_STAGES)
OBJECTIVE_LABELS = dict(OBJECTIVE_STAGES)


def validate_custom_priorities(value):
    """Accept only supported filter combinations, not free-form expressions."""
    if not isinstance(value, (list, tuple)) or len(value) > MAX_CUSTOM_PRIORITIES:
        raise ValueError(f"custom_priorities must be a list of at most {MAX_CUSTOM_PRIORITIES} items")
    seen = set()
    normalized = []
    for item in value:
        if not isinstance(item, dict) or set(item) != CUSTOM_PRIORITY_KEYS:
            raise ValueError("Each custom priority needs id, label, service, metric, member_filter, provider, weekdays, shifts, and direction")
        identifier, label = item["id"], item["label"]
        if type(identifier) is not str or not CUSTOM_ID.fullmatch(identifier) or identifier in seen:
            raise ValueError("Custom priority IDs must be unique custom: identifiers")
        seen.add(identifier)
        if type(label) is not str or not 1 <= len(label.strip()) <= 80 or "\x00" in label:
            raise ValueError("Custom priority labels must contain 1–80 characters")
        service = item["service"]
        if type(service) is not str or service not in CUSTOM_PRIORITY_CATALOG["services"]:
            raise ValueError(f"Unknown custom priority service: {service}")
        for field in ("metric", "member_filter", "direction"):
            if item[field] not in CUSTOM_PRIORITY_CATALOG[field + "s"]:
                raise ValueError(f"Unsupported custom priority {field}: {item[field]}")
        provider = item["provider"]
        if provider not in ("any", "ALS", "BLS") or (service != "ambulance" and provider != "any"):
            raise ValueError("ALS/BLS provider filters apply only to ambulance priorities")
        if service == "ambulance" and item["member_filter"] == "bert":
            raise ValueError("Campus-only BERT members cannot staff ambulance shifts")
        weekdays = item["weekdays"]
        if not isinstance(weekdays, list) or any(type(d) is not int or d not in range(7) for d in weekdays) or len(set(weekdays)) != len(weekdays):
            raise ValueError("weekdays must be unique numbers from 0 (Monday) to 6 (Sunday), or [] for any day")
        shifts = item["shifts"]
        if not isinstance(shifts, list) or any(type(s) is not str or s not in CUSTOM_PRIORITY_CATALOG["services"][service] for s in shifts) or len(set(shifts)) != len(shifts):
            raise ValueError("shifts must be unique supported shift types for the chosen service, or [] for any shift")
        normalized.append({**item, "label": label.strip(), "weekdays": list(weekdays), "shifts": list(shifts)})
    return tuple(normalized)


def validate_hard_requirements(value):
    """Validate supported per-shift minimum-staff rules."""
    if not isinstance(value, (list, tuple)) or len(value) > MAX_HARD_REQUIREMENTS:
        raise ValueError(f"hard_requirements must be a list of at most {MAX_HARD_REQUIREMENTS} items")
    seen, normalized = set(), []
    for item in value:
        if not isinstance(item, dict) or set(item) != HARD_REQUIREMENT_KEYS:
            raise ValueError("Each hard requirement needs id, label, service, member_filter, provider, weekdays, shifts, and min_people")
        identifier, label = item["id"], item["label"]
        if type(identifier) is not str or not HARD_ID.fullmatch(identifier) or identifier in seen:
            raise ValueError("Hard requirement IDs must be unique hard: identifiers")
        seen.add(identifier)
        if type(label) is not str or not 1 <= len(label.strip()) <= 80 or "\\x00" in label:
            raise ValueError("Hard requirement labels must contain 1–80 characters")
        service = item["service"]
        if type(service) is not str or service not in HARD_REQUIREMENT_CATALOG["services"]:
            raise ValueError(f"Unknown hard requirement service: {service}")
        if item["member_filter"] not in HARD_REQUIREMENT_CATALOG["member_filters"]:
            raise ValueError(f"Unsupported hard requirement member filter: {item['member_filter']}")
        provider = item["provider"]
        if provider not in ("any", "ALS", "BLS") or (service != "ambulance" and provider != "any"):
            raise ValueError("ALS/BLS provider filters apply only to ambulance hard requirements")
        if service == "ambulance" and item["member_filter"] == "bert":
            raise ValueError("Campus-only BERT members cannot staff ambulance shifts")
        if type(item["min_people"]) is not int or not 1 <= item["min_people"] <= HARD_REQUIREMENT_CATALOG["max_per_shift"]:
            raise ValueError("min_people must be an integer from 1 to 4")
        weekdays = item["weekdays"]
        if not isinstance(weekdays, list) or any(type(d) is not int or d not in range(7) for d in weekdays) or len(set(weekdays)) != len(weekdays):
            raise ValueError("weekdays must be unique numbers from 0 (Monday) to 6 (Sunday), or [] for any day")
        shifts = item["shifts"]
        if not isinstance(shifts, list) or any(type(s) is not str or s not in HARD_REQUIREMENT_CATALOG["services"][service] for s in shifts) or len(set(shifts)) != len(shifts):
            raise ValueError("shifts must be unique supported shift types for the chosen service, or [] for any shift")
        normalized.append({**item, "label": label.strip(), "weekdays": list(weekdays), "shifts": list(shifts)})
    return tuple(normalized)


def _hard_requirement_member_match(person, selected):
    if selected == "any": return True
    if selected == "driver": return person.is_driver
    if selected == "evdt": return person.is_evdt
    if selected == "auth": return person.certification == "Auth"
    if selected == "non_driver": return not person.is_driver
    if selected == "emt":
        from models import Volunteer
        return isinstance(person, Volunteer)
    from models import Volunteer
    return not isinstance(person, Volunteer)


def hard_requirement_warnings(requirements, people, providers, campus_keys, campus_capacity=2):
    """Heuristic preflight warnings based on per-shift availability only."""
    warnings = []
    calendars = {"ambulance": set(providers), "campus": set(campus_keys)}
    for rule in validate_hard_requirements(requirements):
        matched_keys = [key for key in calendars[rule["service"]]
                        if (not rule["weekdays"] or key[0].weekday() in rule["weekdays"])
                        and (not rule["shifts"] or key[1] in rule["shifts"])
                        and (rule["service"] != "ambulance" or rule["provider"] == "any" or providers[key] == rule["provider"])]
        if not matched_keys:
            warnings.append(f"Hard requirement '{rule['label']}' matches no active shifts; check its service/provider/day/shift filters.")
            continue
        for key in sorted(matched_keys):
            if rule["service"] == "ambulance":
                from models import ambulance_capacity
                capacity = ambulance_capacity(*key, providers[key])
            else:
                capacity = campus_capacity
            if rule["min_people"] > capacity:
                warnings.append(f"Hard requirement '{rule['label']}' cannot fit on {key[0]} {key[1]}: requires {rule['min_people']} people but shift capacity is {capacity}.")
            available = sum(
                key in (person.available if rule["service"] == "ambulance" else person.campus_available)
                and _hard_requirement_member_match(person, rule["member_filter"])
                for person in people
            )
            if available < rule["min_people"]:
                warnings.append(f"Hard requirement '{rule['label']}' is not individually feasible on {key[0]} {key[1]}: only {available} matching available candidate(s), requires {rule['min_people']}. The solver will fail if the full rule set is infeasible.")
            elif available == rule["min_people"]:
                warnings.append(f"Hard requirement '{rule['label']}' is fragile on {key[0]} {key[1]}: exactly {available} matching available candidate(s), no backup. Rest, hour caps, or other rules may make it infeasible.")
    return warnings


def validate_objective_order(value, custom_priorities=()):
    """Require every built-in and configured custom stage exactly once."""
    if not isinstance(value, (list, tuple)) or any(type(item) is not str for item in value):
        raise ValueError("objective_order must be a list of priority identifiers")
    expected = set(DEFAULT_OBJECTIVE_ORDER) | {item["id"] for item in custom_priorities}
    if len(value) != len(expected) or set(value) != expected:
        raise ValueError("objective_order must include each built-in and custom priority exactly once; unknown, missing, or duplicate identifier")
    return tuple(value)
