import csv
import contextlib
import io
import json
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path

from configuration import load_config
from main import run
from models import Volunteer
from parse_form import _build_column_maps, infer_campus_availability, load_all_responses

D = date(2026, 9, 21)
HEADERS = ['Timestamp', 'Username', 'Are you an ambulance EMT or BERT member?',
           'First Name', 'Last Name', 'Driver Status',
           'Ambulance Availability: Weekday AM/PM [Mon 9/21]',
           'First Name', 'Last Name', 'Week 1 — Campus Response Availability [Mon 9/21]',
           'Wellness Wagon Availability: Weekday AM/PM [Mon 9/21]']


class InputTests(unittest.TestCase):
    def config(self, **updates):
        cfg = {'block_start': str(D), 'block_end': str(D),
               'shift_providers': {f'{D}:DAY': 'BLS', f'{D}:NIGHT': 'ALS'},
               'form_csv': 'inputs/responses.csv'}
        cfg.update(updates)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'config.json'
            path.write_text(json.dumps(cfg))
            result = load_config(path)
            self.assertEqual(result.form_csv, (Path(tmp) / 'inputs/responses.csv').resolve())
            return result

    def test_manual_provider_config_and_day_expansion(self):
        cfg = self.config()
        self.assertEqual(cfg.providers, {(D, 'AM'): 'BLS', (D, 'PM'): 'BLS', (D, 'NIGHT'): 'ALS'})
        self.assertEqual((cfg.caps.ambulance, cfg.caps.campus_emt, cfg.caps.campus_bert), (18, 6, 9))

    def test_provider_choices_are_required(self):
        for values in ({}, {f'{D}:DAY': None, f'{D}:NIGHT': 'BLS'},
                       {f'{D}:DAY': 'typo', f'{D}:NIGHT': 'BLS'}):
            with self.subTest(values=values), self.assertRaisesRegex(ValueError, 'Choose ALS or BLS'):
                self.config(shift_providers=values)

    def test_previous_block_provider_dates_are_rejected(self):
        with self.assertRaisesRegex(ValueError, 'outside the block'):
            self.config(shift_providers={'2026-09-08:DAY': 'ALS'})

    def test_duplicate_day_and_am_provider_settings_are_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate'):
            self.config(shift_providers={f'{D}:DAY': 'BLS', f'{D}:AM': 'ALS', f'{D}:NIGHT': 'BLS'})

    def test_blackout_removes_both_services(self):
        cfg = self.config(blackout_periods=[{'start_date': str(D), 'end_date': str(D),
                                             'start_shift': 'AM', 'end_shift': 'AM'}])
        self.assertNotIn((D, 'AM'), cfg.providers)
        self.assertEqual(cfg.campus_keys, [(D, 'C'), (D, 'D')])

    def test_invalid_caps_and_dates_are_rejected(self):
        for updates in ({'hours': {'ambulance_emt': 19}}, {'hours': {'campus_emt': -3}},
                        {'campus_capacity': 0}, {'solver_time_limit_s': -1},
                        {'block_end': '2026-09-20'}):
            with self.subTest(updates=updates), self.assertRaises(ValueError):
                self.config(**updates)

    def load(self, rows, zipped=False, **kwargs):
        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow(HEADERS)
        writer.writerows(rows)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ('responses.zip' if zipped else 'responses.csv')
            if zipped:
                with zipfile.ZipFile(path, 'w') as archive:
                    archive.writestr('responses.csv', buffer.getvalue())
            else:
                path.write_text(buffer.getvalue())
            return load_all_responses(path, D, D, **kwargs)

    def test_latest_submission_wins_across_roles(self):
        old = ['9/18/2026 10:00:00', 'SAME@example.com', 'Ambulance EMT', 'Old', 'Name', 'EVDT', 'AM;PM', '', '', '', '']
        new = ['9/18/2026 11:00:00', 'same@example.com', 'BERT Member Only', '', '', '', '', 'New', 'Name', 'Block A', '']
        emts, berts = self.load([old, new], zipped=True)
        self.assertEqual(emts, [])
        self.assertEqual(len(berts), 1)
        self.assertEqual(berts[0].full_name, 'New Name')

    def test_emt_cr_is_inferred_even_if_campus_section_is_filled(self):
        row = ['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT (EMT only & EMT/ERT dual-role)',
               'Test', 'Person', 'Authorized - Utility 1', 'AM', 'Test', 'Person', 'Block D', 'PM']
        emts, berts = self.load([row])
        self.assertEqual(berts, [])
        self.assertEqual(emts[0].campus_available, {(D, 'A'), (D, 'B')})
        self.assertEqual(emts[0].certification, 'Auth')

    def test_wellness_columns_never_enter_campus_map(self):
        maps = _build_column_maps(HEADERS, D, D)
        self.assertNotIn(10, maps['bert'])

    def test_legacy_form_columns_still_parse(self):
        maps = _build_column_maps(['Day Shifts [Mon 9/21]', 'Night Shifts [Mon 9/21]',
                                   'Weekend Day [Sat 9/26]'], D, date(2026, 9, 27))
        self.assertEqual(maps['emt_day'], {0: D})
        self.assertEqual(maps['emt_night'], {1: D})
        self.assertEqual(maps['emt_weekend'], {2: date(2026, 9, 26)})

    def test_inference_respects_explicit_block_blackout(self):
        p = Volunteer('Test', 'Person', 'test@example.com', available={(D, 'AM'), (D, 'PM')},
                      blackout_slots={(D, 'B')})
        self.assertEqual(infer_campus_availability(p), {(D, 'A'), (D, 'C'), (D, 'D')})

    def test_driver_overrides_are_normalized_and_validated(self):
        row = ['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT', 'Test', 'Person', '', 'AM', '', '', '', '']
        emts, _ = self.load([row], driver_status_overrides={'TEST@example.com': 'AUTH'})
        self.assertEqual(emts[0].certification, 'Auth')
        with self.assertRaisesRegex(ValueError, 'Driver overrides'):
            self.load([row], driver_status_overrides={'test@example.com': 'invalid'})

    def test_pipeline_reads_config_relative_inputs_validates_and_exports(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with (root / 'responses.csv').open('w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(HEADERS)
                writer.writerow(['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT',
                                 'Test', 'Person', 'Authorized - Utility 1', 'AM;PM', '', '', '', ''])
            config = {'block_start': str(D), 'block_end': str(D),
                      'shift_providers': {f'{D}:DAY': 'BLS', f'{D}:NIGHT': 'BLS'},
                      'form_csv': 'responses.csv', 'output_xlsx': 'out/schedule.xlsx',
                      'hours': {'ambulance_emt': 6, 'campus_emt': 6, 'campus_bert': 9},
                      'master_schedule_export': {'enabled': True, 'block': 'TEST', 'path': 'out/master.csv'},
                      'solver_workers': 1, 'solver_time_limit_s': 5}
            path = root / 'config.json'
            path.write_text(json.dumps(config))
            with contextlib.redirect_stdout(io.StringIO()):
                schedule = run(path, check_only=True)
                self.assertFalse((root / 'out').exists())
                run(path)
            self.assertTrue((root / 'out/schedule.xlsx').exists())
            self.assertTrue((root / 'out/master.csv').exists())
            self.assertEqual(sum(map(len, schedule.ambulance.values())), 1)
            self.assertEqual(sum(map(len, schedule.campus.values())), 2)


if __name__ == '__main__':
    unittest.main()
