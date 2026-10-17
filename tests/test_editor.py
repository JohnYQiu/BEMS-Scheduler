"""Focused tests for the standalone local policy editor."""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import editor


class EditorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'config.json'
        self.cfg = json.loads((editor.ROOT / 'config.json').read_text())
        self.cfg['unknown'] = {'keep': ['unchanged'], 'scalar': 'old'}
        self.path.write_text(json.dumps(self.cfg))

    def submission(self):
        cfg, digest = editor.read_config(self.path)
        expanded = {}
        for key, value in cfg['shift_providers'].items():
            if key.endswith(':DAY') and __import__('datetime').date.fromisoformat(key[:10]).weekday() < 5:
                expanded[key[:10]+':AM'] = value
                expanded[key[:10]+':PM'] = value
            else:
                expanded[key] = value
        return digest, {'block_start': cfg['block_start'], 'block_end': cfg['block_end'],
                        'shift_providers': expanded,
                        'scalars': {p: self.get_pointer(cfg,p) for p in editor.scalar_paths(cfg)},
                        'objective_order': list(editor.DEFAULT_PRIORITIES),
                        'optional_rules': {}}

    @staticmethod
    def get_pointer(cfg, pointer):
        node = cfg
        for p in editor._parts(pointer):
            node = node[p]
        return node

    def test_save_backup_unknown_keys_and_weekday_expansion(self):
        version, submitted = self.submission()
        submitted['scalars']['/campus_capacity'] = 3
        backup = editor.save_config(self.path, version, submitted)
        old = json.loads(backup.read_text())
        new = json.loads(self.path.read_text())
        self.assertEqual(old, self.cfg)
        self.assertEqual(new['unknown'], self.cfg['unknown'])
        self.assertNotIn('/unknown/scalar', editor.scalar_paths(self.cfg))
        self.assertEqual(new['campus_capacity'], 3)
        self.assertEqual(new['shift_providers']['2026-09-21:AM'], 'ALS')
        self.assertNotIn('2026-09-21:DAY', new['shift_providers'])
        self.assertEqual(new['objective_order'], list(editor.DEFAULT_PRIORITIES))

    def test_backup_preserves_private_permissions_and_fractional_time(self):
        os.chmod(self.path, 0o600)
        version, submitted = self.submission()
        submitted['scalars']['/solver_time_limit_s'] = 30.5
        backup = editor.save_config(self.path, version, submitted)
        self.assertEqual(self.path.stat().st_mode & 0o777, 0o600)
        self.assertEqual(backup.stat().st_mode & 0o777, 0o600)
        self.assertEqual(json.loads(self.path.read_text())['solver_time_limit_s'], 30.5)

    def test_hard_requirement_round_trips_through_local_editor(self):
        version, submitted = self.submission()
        rule = {'id': 'hard:driver', 'label': 'BLS weekend driver', 'service': 'ambulance',
                'member_filter': 'driver', 'provider': 'BLS', 'weekdays': [4, 5],
                'shifts': ['NIGHT'], 'min_people': 1}
        submitted['hard_requirements'] = [rule]
        editor.save_config(self.path, version, submitted)
        from configuration import load_config
        self.assertEqual(load_config(self.path).hard_requirements, (rule,))
        version, submitted = self.submission()
        submitted['hard_requirements'] = [{**rule, 'min_people': 5}]
        with self.assertRaisesRegex(editor.EditorError, 'min_people must'):
            editor.save_config(self.path, version, submitted)

    def test_supported_rule_can_be_added_and_removed(self):
        version, submitted = self.submission()
        submitted['optional_rules'] = {'/rules/wellness_shifts_per_week': 2}
        editor.save_config(self.path, version, submitted)
        self.assertEqual(json.loads(self.path.read_text())['rules']['wellness_shifts_per_week'], 2)
        from configuration import load_config
        self.assertEqual(load_config(self.path).wellness_shifts_per_week, 2)
        version, submitted = self.submission()
        editor.save_config(self.path, version, submitted)
        self.assertEqual(load_config(self.path).wellness_shifts_per_week, 1)
        self.assertNotIn('wellness_shifts_per_week', json.loads(self.path.read_text()).get('rules', {}))

    def test_evdt_weekly_als_rule_is_supported(self):
        version, submitted = self.submission()
        submitted['optional_rules'] = {'/rules/max_evdt_als_shifts_per_week': 2}
        editor.save_config(self.path, version, submitted)
        from configuration import load_config
        self.assertEqual(load_config(self.path).max_evdt_als_shifts_per_week, 2)
        version, submitted = self.submission()
        editor.save_config(self.path, version, submitted)
        self.assertIsNone(load_config(self.path).max_evdt_als_shifts_per_week)

    def test_custom_priority_saved_with_order_and_unknown_data_preserved(self):
        version, submitted = self.submission()
        custom = dict(id='custom:driver-first', label='Driver coverage on BLS nights',
                      service='ambulance', metric='covered_shifts', member_filter='driver',
                      provider='BLS', weekdays=[4, 5], shifts=['NIGHT'], direction='maximize')
        submitted['custom_priorities'] = [custom]
        submitted['objective_order'].insert(0, custom['id'])
        editor.save_config(self.path, version, submitted)
        new = json.loads(self.path.read_text())
        self.assertEqual(new['custom_priorities'], [custom])
        self.assertEqual(new['objective_order'][0], custom['id'])
        self.assertEqual(new['unknown'], self.cfg['unknown'])
        from configuration import load_config
        self.assertEqual(load_config(self.path).custom_priorities[0], custom)

    def test_custom_priority_rejects_unknown_fields_and_missing_order_without_write(self):
        version, submitted = self.submission()
        custom = dict(id='custom:driver-first', label='Driver', service='ambulance',
                      metric='assignments', member_filter='driver', provider='any',
                      weekdays=[], shifts=[], direction='maximize')
        submitted['custom_priorities'] = [custom]
        with self.assertRaisesRegex(editor.EditorError, 'exactly once'):
            editor.save_config(self.path, version, submitted)
        self.assertEqual(json.loads(self.path.read_text()), self.cfg)
        submitted['objective_order'].append(custom['id'])
        custom['extra'] = 'unsupported code'
        with self.assertRaisesRegex(editor.EditorError, 'Each custom priority needs'):
            editor.save_config(self.path, version, submitted)

    def test_incompatible_saved_custom_priority_cannot_be_silently_dropped(self):
        self.cfg['custom_priorities'] = [{'id': 'custom:future', 'label': 'Future rule',
                                          'service': 'ambulance', 'metric': 'unrecognized'}]
        self.path.write_text(json.dumps(self.cfg))
        version, submitted = self.submission()
        submitted['custom_priorities'] = []
        with self.assertRaisesRegex(editor.EditorError, 'blocked to prevent data loss'):
            editor.save_config(self.path, version, submitted)
        self.assertEqual(json.loads(self.path.read_text()), self.cfg)

    def test_unknown_rule_is_rejected(self):
        _, submitted = self.submission()
        submitted['optional_rules'] = {'/rules/fictional_rule': 2}
        with self.assertRaisesRegex(editor.EditorError, 'implemented rules'):
            editor.build_candidate(self.cfg, submitted)

    def test_missing_shift_is_rejected_without_write(self):
        version, submitted = self.submission()
        submitted['shift_providers'].pop('2026-09-21:AM')
        before = self.path.read_bytes()
        with self.assertRaisesRegex(editor.EditorError, 'cover exactly'):
            editor.save_config(self.path, version, submitted)
        self.assertEqual(self.path.read_bytes(), before)

    def test_invalid_cap_rejected_without_write(self):
        version, submitted = self.submission()
        submitted['scalars']['/hours/ambulance_emt'] = 17
        with self.assertRaisesRegex(ValueError, 'multiples'):
            editor.save_config(self.path, version, submitted)
        self.assertEqual(json.loads(self.path.read_text()), self.cfg)

    def test_stale_version_rejected(self):
        version, submitted = self.submission()
        self.path.write_text(self.path.read_text() + '\n')
        with self.assertRaisesRegex(editor.EditorError, 'changed on disk'):
            editor.save_config(self.path, version, submitted)

    def test_duplicate_priority_rejected(self):
        _, submitted = self.submission()
        submitted['objective_order'] = ['one', 'one']
        with self.assertRaisesRegex(editor.EditorError, 'exactly once'):
            editor.build_candidate(self.cfg, submitted)


if __name__ == '__main__':
    unittest.main()
