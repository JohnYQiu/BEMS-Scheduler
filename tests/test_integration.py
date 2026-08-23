import csv
import tempfile
import unittest
from datetime import date
from pathlib import Path

from campus_solver import solve_campus
from models import BertMember, Volunteer
from output import export_master_schedule_csv
from parse_form import _build_column_maps


class MasterScheduleExportTests(unittest.TestCase):
    def test_export_uses_f26b1_layout_and_keeps_all_people(self):
        start = date(2026, 8, 10)
        driver = Volunteer("Driver", "One", "driver@example.com", "EVDT")
        second_driver = Volunteer("Driver", "Two", "driver2@example.com", "Auth")
        crew = Volunteer("Crew", "One", "crew@example.com", "EMT")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "schedule.csv"
            export_master_schedule_csv(
                {(start, "AM"): [driver, second_driver, crew]},
                {(start, "A"): [crew, second_driver]},
                start,
                str(path),
                block="F26B1",
                daynum_start=810,
            )
            with path.open(newline="") as csv_file:
                rows = list(csv.reader(csv_file))

        self.assertEqual(rows[0], ["Block", "ShiftID", "Date", "Shift", "Vehicle", "Seat", "Requires", "Assigned/Name"])
        self.assertEqual(rows[1][1], "F26B1-0810-AM-R1-EVDT")
        self.assertEqual(rows[1][2], "8/10/26")
        self.assertEqual([row[7] for row in rows[1:4]], ["Driver One", "Driver Two", "Crew One"])
        self.assertEqual(rows[4][5:8], ["S1", "AUTH", "Driver Two"])


class CampusDriverConstraintTests(unittest.TestCase):
    def test_block_with_no_driver_eligible_person_stays_open(self):
        d = date(2026, 8, 10)
        emt = Volunteer("EMT", "Only", "emt@example.com", "EMT", campus_available={(d, "A")})
        bert = BertMember("BERT", "Only", "bert@example.com", campus_available={(d, "A")})
        assignments = solve_campus([emt], [bert], [d], time_limit_s=5)
        self.assertEqual(assignments[(d, "A")], [])

    def test_driver_eligible_person_is_first_in_campus_assignment(self):
        d = date(2026, 8, 10)
        driver = Volunteer("Auth", "Driver", "auth@example.com", "Auth", campus_available={(d, "A")})
        bert = BertMember("BERT", "Member", "bert@example.com", campus_available={(d, "A")})
        assignments = solve_campus([driver], [bert], [d], time_limit_s=5)
        self.assertTrue(assignments[(d, "A")])
        self.assertTrue(assignments[(d, "A")][0].is_driver)


class ParserCompatibilityTests(unittest.TestCase):
    def test_legacy_and_current_form_headers_are_both_mapped(self):
        headers = [
            "Day Shifts [Mon 8/10]", "Night Shifts [Mon 8/10]",
            "Week 1 — Ambulance Availability [Tue 8/11]",
            "Week 1 — Campus Response Availability [Tue 8/11]",
        ]
        maps = _build_column_maps(headers, date(2026, 8, 10), date(2026, 9, 3))
        self.assertEqual(maps["emt_day"][0], date(2026, 8, 10))
        self.assertEqual(maps["emt_night"][1], date(2026, 8, 10))
        self.assertEqual(maps["emt_week"][2], date(2026, 8, 11))
        self.assertEqual(maps["bert"][3], date(2026, 8, 11))


if __name__ == "__main__":
    unittest.main()
