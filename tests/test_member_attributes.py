import csv
import tempfile
import unittest
from pathlib import Path

from member_attributes import apply_member_attributes, load_attribute_catalog
from models import HourCaps, Volunteer
from policy import DEFAULT_OBJECTIVE_ORDER, hard_requirement_warnings
from solver import solve_schedule
from validation import validate_schedule
from datetime import date
from configuration import load_config
import json


class MemberAttributeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "roster.csv"
        with self.path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=[
                "BrownEmail", "ClassYear", "PrimaryStatus", "RoleStatus", "DriverRank"
            ])
            writer.writeheader()
            writer.writerow({
                "BrownEmail": "senior@example.edu", "ClassYear": "2027",
                "PrimaryStatus": "Active", "RoleStatus": "Ambulance EMT", "DriverRank": "1"
            })
            writer.writerow({
                "BrownEmail": "junior@example.edu", "ClassYear": "2028",
                "PrimaryStatus": "Active", "RoleStatus": "EVDT", "DriverRank": "3"
            })

    def test_roster_fields_load_by_email_and_provide_editor_values(self):
        senior = Volunteer("Senior", "Member", "SENIOR@example.edu")
        junior = Volunteer("Junior", "Member", "junior@example.edu")
        self.assertEqual(apply_member_attributes(self.path, [senior, junior]), [])
        self.assertEqual(senior.attributes["ClassYear"], "2027")
        self.assertEqual(junior.attributes["RoleStatus"], "EVDT")
        self.assertEqual(load_attribute_catalog(self.path)["DriverRank"], ["1", "3"])

    def test_attribute_hard_rule_is_enforced_and_validated(self):
        shift = (date(2026, 9, 21), "AM")
        senior = Volunteer("Senior", "Member", "senior@example.edu", available={shift})
        junior = Volunteer("Junior", "Member", "junior@example.edu", available={shift})
        apply_member_attributes(self.path, [senior, junior])
        rule = {
            "id": "hard:senior", "label": "Senior coverage", "service": "ambulance",
            "member_filter": "attribute", "member_attribute_field": "ClassYear",
            "member_attribute_value": "2027", "provider": "BLS", "weekdays": [0],
            "shifts": ["AM"], "min_people": 1,
        }
        providers = {shift: "BLS"}
        self.assertTrue(any('fragile' in warning for warning in
                            hard_requirement_warnings([rule], [senior, junior], providers, [])))
        result = solve_schedule(
            [senior, junior], providers, [], HourCaps(6, 0, 9), time_limit_s=5, workers=1,
            objective_order=DEFAULT_OBJECTIVE_ORDER, hard_requirements=[rule],
        )
        self.assertIn(senior, result.ambulance[shift])
        self.assertEqual(validate_schedule(
            result, [senior, junior], providers, [], HourCaps(6, 0, 9), 2,
            hard_requirements=[rule],
        ), [])

    def test_configuration_rejects_missing_roster_value_for_rule(self):
        cfg = json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text())
        cfg["member_attributes_csv"] = str(self.path)
        cfg["hard_requirements"] = [{
            "id": "hard:class-year", "label": "Require class year", "service": "ambulance",
            "member_filter": "attribute", "member_attribute_field": "ClassYear",
            "member_attribute_value": "2099", "provider": "any", "weekdays": [],
            "shifts": [], "min_people": 1,
        }]
        config_path = Path(self.temp.name) / "config.json"
        config_path.write_text(json.dumps(cfg))
        with self.assertRaisesRegex(ValueError, "unavailable roster value"):
            load_config(config_path)

    def test_attribute_custom_priority_changes_who_is_assigned(self):
        shift = (date(2026, 9, 21), "AM")
        senior = Volunteer("Senior", "Member", "senior@example.edu", available={shift})
        junior = Volunteer("Junior", "Member", "junior@example.edu", available={shift})
        apply_member_attributes(self.path, [senior, junior])
        priority = {
            "id": "custom:class-year", "label": "Prioritize class year",
            "service": "ambulance", "metric": "assignments",
            "member_filter": "attribute", "member_attribute_field": "ClassYear",
            "member_attribute_value": "2027", "provider": "BLS", "weekdays": [],
            "shifts": ["AM"], "direction": "maximize",
        }
        result = solve_schedule(
            [junior, senior], {shift: "BLS"}, [], HourCaps(6, 0, 9), time_limit_s=5,
            workers=1, objective_order=(priority["id"], *DEFAULT_OBJECTIVE_ORDER),
            custom_priorities=[priority],
        )
        self.assertEqual(result.ambulance[shift], [senior])


if __name__ == "__main__":
    unittest.main()
