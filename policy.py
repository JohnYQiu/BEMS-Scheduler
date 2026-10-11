"""Named optimization stages officers may reorder without editing solver code.

The default order exactly matches the historical sequence in solver.py. These are
soft objectives, not safety/credential/availability constraints.
"""

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


def validate_objective_order(value):
    """Require all supported stages once; no silent dropping of an objective."""
    if not isinstance(value, (list, tuple)) or any(type(item) is not str for item in value):
        raise ValueError("objective_order must be a list of priority identifiers")
    if len(value) != len(DEFAULT_OBJECTIVE_ORDER) or set(value) != set(DEFAULT_OBJECTIVE_ORDER):
        raise ValueError("objective_order must include each supported priority exactly once; unknown, missing, or duplicate identifier")
    return tuple(value)
