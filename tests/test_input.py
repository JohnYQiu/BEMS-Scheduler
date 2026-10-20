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
from forces import crew_options
from main import run
from models import Volunteer
from parse_form import (_build_column_maps, infer_campus_availability,
                        load_all_responses, parse_class_conflict_blocks, parse_class_conflicts)

D = date(2026, 9, 21)
HEADERS = ['Timestamp', 'Username', 'Are you an ambulance EMT or BERT member?',
           'First Name', 'Last Name', 'Driver Status',
           'Ambulance Availability: Weekday AM/PM [Mon 9/21]',
           'First Name', 'Last Name', 'Week 1 — Campus Response Availability [Mon 9/21]',
           'Wellness Wagon Availability: Weekday AM/PM [Mon 9/21]',
           'Class/lab/conference conflicts (Ambulance EMTs)']


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

    def test_class_conflicts_use_current_block3_header_and_threshold(self):
        maps = _build_column_maps(["Class/lab/conference conflicts"],
                                  date(2026, 10, 19), date(2026, 11, 15))
        self.assertEqual(maps["idx_class_conflicts"], 0)
        hard, soft = parse_class_conflict_blocks(
            "WEEKLY TUE 13:00-14:20\nWEEKLY TUE 18:30-19:30",
            date(2026, 10, 19), date(2026, 11, 15))
        self.assertIn((date(2026, 11, 10), "C"), hard)
        self.assertNotIn((date(2026, 11, 10), "D"), hard)
        self.assertIn((date(2026, 11, 10), "D"), soft)

    def test_class_conflicts_under_hour_are_soft_not_hard(self):
        raw = '2026-09-21 09:30-10:30\n2026-09-21 15:30-16:30'
        hard, soft = parse_class_conflict_blocks(raw, D, D)
        self.assertEqual(hard, set())
        self.assertEqual(soft, {(D, 'A'), (D, 'B'), (D, 'C'), (D, 'D')})
        self.assertEqual(parse_class_conflicts(raw, D, D), set())

    def test_weekly_soft_class_conflicts_expand_through_the_block(self):
        hard, soft = parse_class_conflict_blocks(
            'WEEKLY MON 09:30-10:30', D, date(2026, 9, 28))
        self.assertEqual(hard, set())
        self.assertEqual(soft, {(D, 'A'), (D, 'B'),
                                (date(2026, 9, 28), 'A'), (date(2026, 9, 28), 'B')})

    def test_sixty_minutes_is_hard_and_59_minutes_is_soft(self):
        hard, soft = parse_class_conflict_blocks(
            '2026-09-21 10:00-11:00; 2026-09-21 13:00-13:59', D, D)
        self.assertEqual(hard, {(D, 'B')})
        self.assertEqual(soft, {(D, 'C')})

    def test_block3_ninety_minute_class_blocks_three_hour_cr_slot(self):
        start, end = date(2026, 10, 19), date(2026, 11, 15)
        conflicts = parse_class_conflicts('WEEKLY TUE 10:30-12:00', start, end)
        self.assertEqual(conflicts, {
            (date(2026, 10, 20), 'B'), (date(2026, 10, 27), 'B'),
            (date(2026, 11, 3), 'B'), (date(2026, 11, 10), 'B'),
        })

    def test_class_conflict_removes_only_cr_availability_not_ambulance(self):
        row = ['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT',
               'Test', 'Person', 'Not a driver', 'AM', '', '', '', '',
               '2026-09-21 10:00-11:00']
        emts, _ = self.load([row])
        self.assertEqual(emts[0].available, {(D, 'AM')})
        self.assertEqual(emts[0].campus_available, {(D, 'A')})

    def test_malformed_class_conflict_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Could not parse class-conflict line'):
            parse_class_conflicts('Tuesday 10:30 to 11:50', D, D)

    def test_driver_overrides_are_normalized_and_validated(self):
        row = ['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT', 'Test', 'Person', '', 'AM', '', '', '', '']
        emts, _ = self.load([row], driver_status_overrides={'TEST@example.com': 'AUTH'})
        self.assertEqual(emts[0].certification, 'Auth')
        with self.assertRaisesRegex(ValueError, 'Driver overrides'):
            self.load([row], driver_status_overrides={'test@example.com': 'invalid'})

    def test_driver_dependent_policy_rejects_missing_or_ambiguous_status(self):
        row = ['9/18/2026 10:00:00', 'test@example.com', 'Ambulance EMT',
               'Test', 'Person', 'Not authorized', 'AM', '', '', '', '']
        with self.assertRaisesRegex(ValueError, 'Unrecognized Driver Status'):
            self.load([row], require_driver_status=True)
        emts, _ = self.load([row], require_driver_status=True,
                            driver_status_overrides={'test@example.com': 'EMT'})
        self.assertEqual(emts[0].certification, 'EMT')
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'responses.csv'
            with path.open('w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(HEADERS[:5] + HEADERS[6:])
                writer.writerow(row[:5] + row[6:])
            with self.assertRaisesRegex(ValueError, 'Driver Status column'):
                load_all_responses(path, D, D, require_driver_status=True)

    def test_f26b3_config_preserves_minimum_hours_for_same_length_block(self):
        root = Path(__file__).resolve().parents[1]
        prior = load_config(root / 'config.json')
        next_block = load_config(root / 'config_f26b3.json')
        self.assertEqual(len(prior.dates), 28)
        self.assertEqual(len(next_block.dates), 28)
        self.assertEqual(next_block.caps, prior.caps)
        self.assertEqual(next_block.caps.ambulance, 18)
        self.assertEqual(next_block.caps.campus_emt, 6)
        self.assertEqual(next_block.caps.campus_bert, 9)
        self.assertEqual(len(next_block.providers), 76)
        self.assertEqual(next_block.providers[(date(2026, 10, 19), 'AM')], 'ALS')
        self.assertEqual(next_block.providers[(date(2026, 10, 25), 'NIGHT')], 'BLS')
        self.assertEqual(next_block.providers[(date(2026, 11, 14), 'DAY')], 'BLS')

    def test_form_driver_status_flows_into_custom_priority(self):
        from policy import DEFAULT_OBJECTIVE_ORDER
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with (root / 'responses.csv').open('w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(HEADERS)
                writer.writerow(['9/18/2026 10:00:00', 'driver@example.com', 'Ambulance EMT',
                                 'Driver', 'Person', 'Authorized - Utility 1', 'AM', '', '', '', ''])
                writer.writerow(['9/18/2026 10:00:00', 'other@example.com', 'Ambulance EMT',
                                 'Other', 'Person', 'Not a driver', 'AM', '', '', '', ''])
            custom = dict(id='custom:bls-driver', label='BLS driver coverage',
                          service='ambulance', metric='covered_shifts', member_filter='driver',
                          provider='BLS', weekdays=[0], shifts=['AM'], direction='maximize')
            cfg = dict(block_start=str(D), block_end=str(D),
                       shift_providers={f'{D}:DAY': 'BLS', f'{D}:NIGHT': 'BLS'},
                       form_csv='responses.csv', hours={'ambulance_emt': 6},
                       solver_workers=1, solver_time_limit_s=5,
                       custom_priorities=[custom],
                       objective_order=[custom['id'], *DEFAULT_OBJECTIVE_ORDER])
            path = root / 'config.json'
            path.write_text(json.dumps(cfg))
            with contextlib.redirect_stdout(io.StringIO()):
                schedule = run(path, check_only=True)
            self.assertEqual([p.email for p in schedule.ambulance[(D, 'AM')]],
                             ['driver@example.com'])
            self.assertEqual(schedule.stages[0].value, 1)

    def test_forced_crew_is_required_and_respects_hours_and_availability(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with (root / 'responses.csv').open('w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(HEADERS)
                writer.writerow(['9/18/2026 10:00:00', 'driver@example.com', 'Ambulance EMT',
                                 'Driver', 'Person', 'EVDT', 'AM;PM', '', '', '', ''])
                writer.writerow(['9/18/2026 10:00:00', 'crew@example.com', 'Ambulance EMT',
                                 'Crew', 'Person', 'EMT', 'AM;PM', '', '', '', ''])
            config = {'block_start': str(D), 'block_end': str(D),
                      'shift_providers': {f'{D}:DAY': 'ALS', f'{D}:NIGHT': 'BLS'},
                      'form_csv': 'responses.csv', 'hours': {'ambulance_emt': 6},
                      'solver_workers': 1, 'solver_time_limit_s': 5,
                      'forced_crews': [{'date': str(D), 'shift': 'AM',
                                        'members': ['driver@example.com', 'crew@example.com'],
                                        'note': 'Busy morning'}]}
            path = root / 'config.json'
            path.write_text(json.dumps(config))
            with contextlib.redirect_stdout(io.StringIO()):
                schedule = run(path, check_only=True)
            self.assertEqual({p.email for p in schedule.ambulance[(D, 'AM')]},
                             {'driver@example.com', 'crew@example.com'})
            self.assertEqual(schedule.ambulance[(D, 'PM')], [])
            self.assertEqual(len(load_config(path).locks), 2)
            with contextlib.redirect_stdout(io.StringIO()):
                choices = crew_options(load_config(path))
            morning = next(shift for shift in choices if shift['key'] == f'{D}:AM')
            self.assertEqual({p['email'] for p in morning['candidates']},
                             {'driver@example.com', 'crew@example.com'})
            self.assertEqual(morning['provider'], 'ALS')
            config['forced_crews'][0]['members'] = ['missing@example.com']
            path.write_text(json.dumps(config))
            with self.assertRaisesRegex(ValueError, 'not an ambulance EMT'):
                with contextlib.redirect_stdout(io.StringIO()):
                    run(path, check_only=True)
            config['forced_crews'][0]['members'] = ['driver@example.com']
            config['forced_crews'][0]['shift'] = 'NIGHT'
            path.write_text(json.dumps(config))
            with self.assertRaisesRegex(ValueError, 'not available'):
                with contextlib.redirect_stdout(io.StringIO()):
                    run(path, check_only=True)
            config['forced_crews'] = [
                {'date': str(D), 'shift': 'AM', 'members': ['driver@example.com']},
                {'date': str(D), 'shift': 'PM', 'members': ['driver@example.com']},
            ]
            path.write_text(json.dumps(config))
            with self.assertRaisesRegex(RuntimeError, 'failed'):
                with contextlib.redirect_stdout(io.StringIO()):
                    run(path, check_only=True)

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
