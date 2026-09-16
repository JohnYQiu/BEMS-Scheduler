import csv
import tempfile
import unittest
from datetime import date
from pathlib import Path

from campus_solver import solve_campus
from ambulance_solver import solve_ambulance
from models import BertMember, Volunteer
from output import export_master_schedule_csv
from parse_form import _build_column_maps
from validate import AvailabilityRequirements, check_ambulance_requirements


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

    def test_authorized_driver_is_not_exported_as_evdt(self):
        start = date(2026, 8, 10)
        authorized = Volunteer("Authorized", "Only", "auth@example.com", "Auth")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "schedule.csv"
            export_master_schedule_csv(
                {(start, "AM"): [authorized]}, {}, start, str(path),
            )
            with path.open(newline="") as csv_file:
                rows = list(csv.reader(csv_file))
        self.assertEqual(rows[1][1], "F26B1-0810-AM-R1-AUTH")
        self.assertEqual(rows[1][5:7], ["Driver", "AUTH"])

    def test_open_seats_are_exported_as_blank_rows(self):
        start = date(2026, 8, 10)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "schedule.csv"
            export_master_schedule_csv(
                {(start, "AM"): []}, {(start, "A"): []}, start, str(path),
            )
            with path.open(newline="") as csv_file:
                rows = list(csv.reader(csv_file))
        self.assertEqual([row[5] for row in rows[1:]], ["Driver", "C2", "S1", "S2"])
        self.assertTrue(all(row[7] == "" for row in rows[1:]))

    def test_saturday_night_preserves_evdt_opening_when_auth_can_ride_crew(self):
        block_start = date(2026, 9, 8)
        shift_date = date(2026, 9, 12)
        auth = Volunteer("Timothy", "Ro", "timothy@example.com", "Auth")
        crew_one = Volunteer("Samuel", "Salter", "samuel@example.com", "EMT")
        crew_two = Volunteer("Roma", "Shah", "roma@example.com", "EMT")

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "schedule.csv"
            export_master_schedule_csv(
                {(shift_date, "NIGHT"): [auth, crew_one, crew_two]},
                {},
                block_start,
                str(path),
                block="F26B1",
                daynum_start=908,
            )
            with path.open(newline="") as csv_file:
                rows = list(csv.reader(csv_file))

        ambulance_rows = rows[1:5]
        self.assertEqual(ambulance_rows[0][1], "F26B1-0912-NIGHT-R1-EVDT")
        self.assertEqual(ambulance_rows[0][5:8], ["Driver", "EVDT", ""])
        self.assertEqual(ambulance_rows[1][5:8], ["C2", "CREW", "Timothy Ro"])
        self.assertEqual(
            {row[7] for row in ambulance_rows[1:]},
            {"Timothy Ro", "Samuel Salter", "Roma Shah"},
        )


class CampusDriverConstraintTests(unittest.TestCase):
    def test_block_with_no_driver_eligible_person_is_still_covered(self):
        d = date(2026, 8, 10)
        emt = Volunteer("EMT", "Only", "emt@example.com", "EMT", campus_available={(d, "A")})
        bert = BertMember("BERT", "Only", "bert@example.com", campus_available={(d, "A")})
        assignments = solve_campus([emt], [bert], [d], time_limit_s=5)
        self.assertEqual(len(assignments[(d, "A")]), 2)

    def test_driver_can_still_be_required_when_requested(self):
        d = date(2026, 8, 10)
        emt = Volunteer("EMT", "Only", "emt@example.com", "EMT", campus_available={(d, "A")})
        assignments = solve_campus([emt], [], [d], require_driver=True, time_limit_s=5)
        self.assertEqual(assignments[(d, "A")], [])

    def test_driver_eligible_person_is_first_in_campus_assignment(self):
        d = date(2026, 8, 10)
        driver = Volunteer("Auth", "Driver", "auth@example.com", "Auth", campus_available={(d, "A")})
        bert = BertMember("BERT", "Member", "bert@example.com", campus_available={(d, "A")})
        assignments = solve_campus([driver], [bert], [d], time_limit_s=5)
        self.assertTrue(assignments[(d, "A")])
        self.assertTrue(assignments[(d, "A")][0].is_driver)


class LockedAssignmentTests(unittest.TestCase):
    def test_requested_assignment_is_preserved(self):
        d = date(2026, 8, 10)
        volunteer = Volunteer("Chosen", "Person", "chosen@example.com", "EMT", available={(d, "AM")})
        assignments = solve_ambulance(
            [volunteer], [d], set(), locked_assignments=[
                {"date": "2026-08-10", "shift": "AM", "email": "chosen@example.com"}
            ], time_limit_s=5,
        )
        self.assertEqual(assignments[(d, "AM")], [volunteer])


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

    def test_shopping_period_form_headers_are_mapped(self):
        headers = [
            "Last Name", "First Name",
            "Please indicate your availability for the below dates and shifts. (Weekdays) [Tue  9/8]",
            "Please indicate your availability for the below dates and shifts. (Weekend Nights) [Sat 9/12]",
            "Last Name", "First Name",
            "Please indicate your availability for the below dates and shifts. [Tue 9/8]",
        ]
        maps = _build_column_maps(headers, date(2026, 9, 8), date(2026, 9, 20))
        self.assertEqual(maps["emt_week"][2], date(2026, 9, 8))
        self.assertEqual(maps["emt_week"][3], date(2026, 9, 12))
        self.assertEqual(maps["bert"][6], date(2026, 9, 8))


class ShoppingPeriodValidationTests(unittest.TestCase):
    def test_strict_form_categories_require_each_kind_of_availability(self):
        reqs = AvailabilityRequirements(
            emt_min_weekday_am_shifts=1, emt_min_weekday_pm_shifts=1,
            emt_min_weekday_night_shifts=1, emt_min_weekend_day_shifts=1,
            emt_min_weekend_night_shifts=1,
        )
        only_am = Volunteer("Test", "EMT", "test@example.com", "EMT", available={
            (date(2026, 9, 8), "AM"), (date(2026, 9, 9), "AM"),
        })
        missing = check_ambulance_requirements([only_am], reqs)[0]["missing"]
        self.assertIn("weekday PM shifts (0/1)", missing)
        self.assertIn("weekend night shifts (0/1)", missing)

    def test_friday_night_does_not_count_as_weekday_night(self):
        reqs = AvailabilityRequirements(
            emt_min_weekday_night_shifts=1,
            emt_min_weekend_night_shifts=1,
        )
        friday_only = Volunteer("Test", "EMT", "test@example.com", "EMT", available={
            (date(2026, 9, 11), "NIGHT"),
        })
        missing = check_ambulance_requirements([friday_only], reqs)[0]["missing"]
        self.assertIn("weekday night shifts (0/1)", missing)


if __name__ == "__main__":
    unittest.main()
