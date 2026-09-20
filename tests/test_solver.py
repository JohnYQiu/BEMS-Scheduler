import unittest
from datetime import date, timedelta
from unittest.mock import Mock, patch

from ortools.sat.python import cp_model

from models import BertMember, HourCaps, LockedAssignment, Volunteer, is_big_weekend, weekend_priority
from solver import _optimize, solve_schedule
from validation import validate_schedule

D = date(2026, 9, 21)


def emt(name="Person", cert="EMT", ambulance=(), campus=()):
    return Volunteer(name, "Test", f"{name.lower()}@example.com", cert,
                     available=set(ambulance), campus_available=set(campus))


class SchedulingTests(unittest.TestCase):
    def solve(self, people, providers=None, campus=None, caps=HourCaps(), locks=()):
        if providers is None:
            providers = {k: "BLS" for p in people for k in getattr(p, "available", ())}
        if campus is None:
            campus = sorted({k for p in people for k in p.campus_available})
        result = solve_schedule(people, providers, campus, caps, 2, list(locks), 5, 1)
        self.assertEqual(validate_schedule(result, people, providers, campus, caps, 2, locks), [])
        self.assertTrue(all(s.status == "OPTIMAL" for s in result.stages))
        return result

    def test_approved_contiguous_combinations(self):
        for ambulance, campus in [(('AM', 'PM'), ()), (('AM',), ('C',)),
                                  (('AM',), ('C', 'D')), (('PM',), ('A', 'B')),
                                  (('PM',), ('B',)), ((), ('B', 'C', 'D'))]:
            with self.subTest(ambulance=ambulance, campus=campus):
                p = emt(ambulance=[(D, s) for s in ambulance], campus=[(D, s) for s in campus])
                locks = [LockedAssignment(p.email, (D, s)) for s in ambulance + campus]
                self.solve([p], caps=HourCaps(18, 9, 9), locks=locks)
                self.assertEqual(len(p.assigned) + len(p.campus_assigned), len(locks))

    def test_gap_combinations_are_rejected_when_ambulance_is_locked(self):
        for shift, block in [('AM', 'D'), ('PM', 'A')]:
            with self.subTest(shift=shift, block=block):
                p = emt(ambulance={(D, shift)}, campus={(D, block)})
                self.solve([p], locks=[LockedAssignment(p.email, (D, shift))])
                self.assertEqual(p.campus_assigned, [])

    def test_bert_cannot_work_disconnected_blocks(self):
        p = BertMember('Campus', 'Only', 'campus@example.com', campus_available={(D, 'A'), (D, 'D')})
        self.solve([p])
        self.assertEqual(p.campus_assigned_hours, 3)

    def test_cr_blocks_are_individual_options(self):
        p = emt(campus={(D, 'A'), (D, 'B')})
        self.solve([p], caps=HourCaps(0, 3, 9))
        self.assertEqual(len(p.campus_assigned), 1)

    def test_joint_solver_can_choose_ambulance_shift_to_allow_cr(self):
        p = emt(ambulance={(D, 'AM'), (D, 'PM')}, campus={(D, 'B')})
        self.solve([p], caps=HourCaps(6, 3, 9))
        self.assertEqual(p.assigned, [(D, 'PM')])
        self.assertEqual(p.campus_assigned, [(D, 'B')])

    def test_night_excludes_campus_on_same_and_following_day(self):
        for day in (D, D + timedelta(days=1)):
            for block in ('A', 'B', 'C', 'D'):
                with self.subTest(day=day, block=block):
                    p = emt(ambulance={(D, 'NIGHT')}, campus={(day, block)})
                    self.solve([p], locks=[LockedAssignment(p.email, (D, 'NIGHT'))])
                    self.assertEqual(p.campus_assigned, [])

    def test_consecutive_nights_have_twelve_hours_rest(self):
        keys = {(D, 'NIGHT'), (D + timedelta(days=1), 'NIGHT')}
        p = emt(ambulance=keys)
        self.solve([p], caps=HourCaps(24, 0, 9))
        self.assertEqual(set(p.assigned), keys)

    def test_evening_to_next_morning_has_twelve_hours_rest(self):
        keys = {(D, 'D'), (D + timedelta(days=1), 'A')}
        p = emt(campus=keys)
        self.solve([p])
        self.assertEqual(set(p.campus_assigned), keys)

    def test_ambulance_and_campus_have_separate_caps(self):
        night = (D + timedelta(days=3), 'NIGHT')
        p = emt(ambulance={(D, 'AM'), (D, 'PM'), night},
                campus={(D, b) for b in ('A', 'B', 'C', 'D')})
        self.solve([p], locks=[LockedAssignment(p.email, night)])
        self.assertEqual(p.assigned_hours, 18)
        self.assertEqual(p.campus_assigned_hours, 6)

    def test_inadequate_availability_leaves_hours_short(self):
        p = emt(ambulance={(D, 'PM'), (D + timedelta(days=3), 'NIGHT')},
                campus={(D, 'C'), (D, 'D')})
        self.solve([p])
        self.assertEqual(p.assigned_hours, 18)
        self.assertEqual(p.campus_assigned_hours, 0)

    def test_conflicting_locks_fail_instead_of_overriding_rules(self):
        p = emt(ambulance={(D, 'AM')}, campus={(D, 'D')})
        with self.assertRaisesRegex(RuntimeError, 'INFEASIBLE'):
            self.solve([p], locks=[LockedAssignment(p.email, k) for k in [(D, 'AM'), (D, 'D')]])

    def test_unavailable_lock_fails(self):
        with self.assertRaisesRegex(ValueError, 'unavailable'):
            self.solve([emt()], locks=[LockedAssignment('person@example.com', (D, 'AM'))])

    def test_duplicate_identity_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'unique'):
            self.solve([emt(), BertMember('Person', 'Again', 'PERSON@example.com')])

    def test_campus_only_person_does_not_cover_ambulance(self):
        p = BertMember('Campus', 'Only', 'campus@example.com', campus_available={(D, 'A')})
        result = self.solve([p], providers={(D, 'AM'): 'BLS'})
        self.assertEqual(result.ambulance[(D, 'AM')], [])
        self.assertEqual(result.campus[(D, 'A')], [p])

    def test_weekend_definition_and_priority_excludes_friday_day_and_sunday_night(self):
        expected = {(4, 'NIGHT'): 0, (5, 'NIGHT'): 1, (5, 'DAY'): 2, (6, 'DAY'): 3}
        for offset in range(7):
            for kind in ('AM', 'PM', 'DAY', 'NIGHT'):
                with self.subTest(offset=offset, kind=kind):
                    day = D + timedelta(days=offset)
                    self.assertEqual(weekend_priority(day, kind), expected.get((offset, kind)))
                    self.assertEqual(is_big_weekend(day, kind), (offset, kind) in expected)

    def test_split_crews_follow_friday_saturday_nights_then_saturday_sunday_days(self):
        groups = [(4, 'NIGHT'), (5, 'NIGHT'), (5, 'DAY'), (6, 'DAY')]
        for higher in range(len(groups)):
            for lower in range(higher + 1, len(groups)):
                with self.subTest(higher=higher, lower=lower):
                    first = (D + timedelta(days=groups[higher][0]), groups[higher][1])
                    second = (D + timedelta(days=groups[lower][0]), groups[lower][1])
                    flexible = emt('Flexible', ambulance={first, second})
                    people = [flexible, emt('FirstAuth', 'Auth', {first}), emt('FirstEMT', ambulance={first}),
                              emt('SecondAuth', 'Auth', {second}), emt('SecondEMT', ambulance={second})]
                    result = self.solve(people, {first: 'BLS', second: 'BLS'}, caps=HourCaps(12, 0, 9))
                    self.assertEqual(flexible.assigned, [first])
                    self.assertEqual(len(result.ambulance[first]), 3)
                    self.assertEqual(len(result.ambulance[second]), 2)

    def test_one_night_split_crew_outranks_two_weekend_day_split_crews(self):
        night = (D + timedelta(days=4), 'NIGHT')
        saturday, sunday = (D + timedelta(days=5), 'DAY'), (D + timedelta(days=6), 'DAY')
        first, second = emt('First', ambulance={night, saturday}), emt('Second', ambulance={night, sunday})
        people = [first, second, emt('NightAuth', 'Auth', {night}),
                  emt('SaturdayAuth', 'Auth', {saturday}), emt('SaturdayEMT', ambulance={saturday}),
                  emt('SundayAuth', 'Auth', {sunday}), emt('SundayEMT', ambulance={sunday})]
        result = self.solve(people, {night: 'BLS', saturday: 'BLS', sunday: 'BLS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(first.assigned, [night])
        self.assertEqual(second.assigned, [night])
        self.assertEqual([len(result.ambulance[k]) for k in (night, saturday, sunday)], [3, 2, 2])

    def test_saturday_night_split_crew_outranks_fourth_friday_volunteer(self):
        friday, saturday = (D + timedelta(days=4), 'NIGHT'), (D + timedelta(days=5), 'NIGHT')
        flexible = emt('Flexible', ambulance={friday, saturday})
        people = [flexible, emt('FridayAuth', 'Auth', {friday}), emt('Friday1', ambulance={friday}),
                  emt('Friday2', ambulance={friday}), emt('SaturdayAuth', 'Auth', {saturday}),
                  emt('SaturdayEMT', ambulance={saturday})]
        result = self.solve(people, {friday: 'BLS', saturday: 'BLS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(flexible.assigned, [saturday])
        self.assertEqual([len(result.ambulance[k]) for k in (friday, saturday)], [3, 3])

    def test_fourth_night_volunteer_outranks_completing_weekend_day_split_crew(self):
        for offset in (4, 5):
            with self.subTest(night_offset=offset):
                night, day = (D + timedelta(days=offset), 'NIGHT'), (D + timedelta(days=6), 'DAY')
                flexible = emt('Flexible', ambulance={night, day})
                people = [flexible, emt('NightAuth', 'Auth', {night}), emt('Night1', ambulance={night}),
                          emt('Night2', ambulance={night}), emt('DayAuth', 'Auth', {day}),
                          emt('DayEMT', ambulance={day})]
                result = self.solve(people, {night: 'BLS', day: 'BLS'}, caps=HourCaps(12, 0, 9))
                self.assertEqual(flexible.assigned, [night])
                self.assertEqual([len(result.ambulance[k]) for k in (night, day)], [4, 2])

    def test_bls_night_split_crew_outranks_weekend_day_als_driver(self):
        night, day = (D + timedelta(days=4), 'NIGHT'), (D + timedelta(days=5), 'DAY')
        driver = emt('Driver', 'EVDT', {night, day})
        people = [driver, emt('NightAuth', 'Auth', {night}), emt('NightEMT', ambulance={night}),
                  emt('DayEMT', ambulance={day})]
        result = self.solve(people, {night: 'BLS', day: 'ALS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(driver.assigned, [night])
        self.assertTrue(all(result.ambulance.values()))

    def test_coverage_outranks_evdt_placement(self):
        friday = D + timedelta(days=4)
        keys = {(D, 'AM'): 'BLS', (D, 'PM'): 'BLS', (friday, 'NIGHT'): 'ALS'}
        p = emt(cert='EVDT', ambulance=set(keys))
        result = self.solve([p], keys, caps=HourCaps(12, 0, 9))
        self.assertEqual(set(p.assigned), {(D, 'AM'), (D, 'PM')})
        self.assertFalse(result.ambulance[(friday, 'NIGHT')])

    def test_als_weekend_night_has_priority_over_weekend_day(self):
        friday, saturday = D + timedelta(days=4), D + timedelta(days=5)
        night, day = (friday, 'NIGHT'), (saturday, 'DAY')
        driver = emt('Driver', 'EVDT', {night, day})
        n, d = emt('Night', ambulance={night}), emt('Day', ambulance={day})
        self.solve([driver, n, d], {night: 'ALS', day: 'ALS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(driver.assigned, [night])

    def test_als_weekend_day_has_priority_over_weekday_als(self):
        weekday, weekend = (D, 'NIGHT'), (D + timedelta(days=5), 'DAY')
        driver = emt('Driver', 'EVDT', {weekday, weekend})
        people = [driver, emt('Weekday', ambulance={weekday}), emt('Weekend', ambulance={weekend})]
        self.solve(people, {weekday: 'ALS', weekend: 'ALS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(driver.assigned, [weekend])

    def test_bls_weekend_needs_three_with_a_driver_before_weekday_als(self):
        weekday, weekend = (D, 'NIGHT'), (D + timedelta(days=4), 'NIGHT')
        for existing_crew in (1, 2, 3):
            with self.subTest(existing_crew=existing_crew):
                driver = emt('Driver', 'EVDT', {weekday, weekend})
                people = [driver, emt('Auth', 'Auth', {weekend}), emt('Weekday', ambulance={weekday})]
                people.extend(emt(f'Crew{i}', ambulance={weekend}) for i in range(existing_crew - 1))
                result = self.solve(people, {weekday: 'ALS', weekend: 'BLS'}, caps=HourCaps(12, 0, 9))
                self.assertTrue(all(result.ambulance.values()))
                self.assertEqual(driver.assigned, [weekend if existing_crew < 3 else weekday])
                self.assertEqual(len(result.ambulance[weekend]), min(existing_crew + 1, 3))

    def test_three_weekend_emts_without_a_driver_still_need_the_evdt(self):
        weekday, weekend = (D, 'NIGHT'), (D + timedelta(days=4), 'NIGHT')
        driver = emt('Driver', 'EVDT', {weekday, weekend})
        people = [driver, emt('Weekday', ambulance={weekday})]
        people.extend(emt(f'Crew{i}', ambulance={weekend}) for i in range(3))
        result = self.solve(people, {weekday: 'ALS', weekend: 'BLS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(driver.assigned, [weekend])
        self.assertGreaterEqual(len(result.ambulance[weekend]), 3)

    def test_als_split_crew_needs_two_distinct_qualified_drivers(self):
        key = (D + timedelta(days=5), 'DAY')
        for truck_cert, second_cert, expected_ready in [('EVDT', 'EMT', 0), ('EVDT', 'Auth', 1),
                                                        ('EVDT', 'EVDT', 1), ('Auth', 'Auth', 0)]:
            with self.subTest(truck_cert=truck_cert, second_cert=second_cert):
                people = [emt('Truck', truck_cert, {key}), emt('Utility', second_cert, {key}),
                          emt('Crew', ambulance={key})]
                result = self.solve(people, {key: 'ALS'}, caps=HourCaps(12, 0, 9))
                ready = next(s for s in result.stages if s.name == 'Saturday days: shifts ready for split crew')
                self.assertEqual(ready.value, expected_ready)
                self.assertEqual(len(result.ambulance[key]), 3)

    def test_three_person_split_crews_outrank_a_fourth_weekend_volunteer(self):
        first, second = (D + timedelta(days=4), 'NIGHT'), (D + timedelta(days=11), 'NIGHT')
        flexible = emt('Flexible', ambulance={first, second})
        people = [flexible, emt('FirstAuth', 'Auth', {first}), emt('First1', ambulance={first}),
                  emt('First2', ambulance={first}), emt('SecondAuth', 'Auth', {second}),
                  emt('Second1', ambulance={second})]
        result = self.solve(people, {first: 'BLS', second: 'BLS'}, caps=HourCaps(12, 0, 9))
        self.assertEqual(flexible.assigned, [second])
        self.assertEqual([len(result.ambulance[k]) for k in (first, second)], [3, 3])

    def test_ordinary_nights_do_not_get_the_weekend_staffing_priority(self):
        weekend = (D + timedelta(days=4), 'NIGHT')
        for offset in (3, 6):  # Thursday and Sunday nights are ordinary nights.
            with self.subTest(offset=offset):
                ordinary = (D + timedelta(days=offset), 'NIGHT')
                flexible = emt('Flexible', ambulance={ordinary, weekend})
                people = [flexible, emt('Ordinary', 'Auth', {ordinary}), emt('Weekend', 'Auth', {weekend})]
                result = self.solve(people, {ordinary: 'BLS', weekend: 'BLS'}, caps=HourCaps(12, 0, 9))
                self.assertEqual(flexible.assigned, [weekend])
                self.assertEqual(len(result.ambulance[ordinary]), 1)

    def test_weekend_crews_fill_before_adding_weekday_seats(self):
        weekdays = {(D, 'AM'), (D, 'PM')}
        weekend_types = [(4, 'NIGHT'), (5, 'NIGHT'), (5, 'DAY'), (6, 'DAY')]
        for offset, kind in weekend_types:
            for provider in ('ALS', 'BLS'):
                with self.subTest(offset=offset, kind=kind, provider=provider):
                    weekend = (D + timedelta(days=offset), kind)
                    flexible = emt('Flexible', ambulance=weekdays | {weekend})
                    people = [flexible, emt('Day', ambulance=weekdays),
                              emt('Driver', 'EVDT' if provider == 'ALS' else 'Auth', {weekend}),
                              emt('Crew1', ambulance={weekend}), emt('Crew2', ambulance={weekend})]
                    providers = {k: 'BLS' for k in weekdays}
                    providers[weekend] = provider
                    result = self.solve(people, providers, caps=HourCaps(12, 0, 9))
                    self.assertEqual(flexible.assigned, [weekend])
                    self.assertEqual(len(result.ambulance[weekend]), 4)
                    self.assertTrue(all(len(result.ambulance[k]) == 1 for k in weekdays))

    def test_weekend_staffing_outranks_total_hours_without_losing_coverage(self):
        weekday_night = (D + timedelta(days=2), 'NIGHT')
        friday_am, friday_night = (D + timedelta(days=4), 'AM'), (D + timedelta(days=4), 'NIGHT')
        keys = {weekday_night, friday_am, friday_night}
        flexible = emt('Flexible', ambulance=keys)
        people = [flexible, emt('Night', ambulance={weekday_night}),
                  emt('Morning', ambulance={friday_am}), emt('Driver', 'Auth', {friday_night})]
        result = self.solve(people, {k: 'BLS' for k in keys})
        self.assertTrue(all(result.ambulance.values()))
        self.assertEqual(flexible.assigned, [friday_night])
        self.assertEqual(flexible.assigned_hours, 12)
        self.assertEqual(len(result.ambulance[friday_night]), 2)

    def test_extra_weekday_seats_do_not_outrank_campus_coverage(self):
        weekdays = {(D, 'AM'), (D, 'PM')}
        night = (D + timedelta(days=3), 'NIGHT')
        flexible = emt('Flexible', ambulance=weekdays | {night}, campus={(D, 'C'), (D, 'D')})
        people = [flexible, emt('Day', ambulance=weekdays), emt('Night', ambulance={night})]
        result = self.solve(people, caps=HourCaps(12, 6, 9))
        self.assertTrue(all(result.ambulance.values()))
        self.assertEqual(flexible.assigned, [night])
        self.assertEqual(flexible.campus_assigned_hours, 6)

    def test_weekday_bls_has_no_driver_objective(self):
        p = emt('Driver', 'EVDT', {(D, 'AM')})
        result = self.solve([p])
        self.assertFalse(any('EVDT' in s.name or 'driver' in s.name for s in result.stages))

    def test_als_seat_cannot_be_filled_by_auth(self):
        key = (D, 'AM')
        people = [emt(str(i), 'Auth', {key}) for i in range(3)]
        result = self.solve(people, {key: 'ALS'}, caps=HourCaps(6, 0, 9))
        self.assertEqual(len(result.ambulance[key]), 1)

    def test_crew_and_campus_capacity(self):
        people = [emt(str(i), ambulance={(D, 'AM')}, campus={(D, 'C')}) for i in range(6)]
        result = self.solve(people)
        self.assertEqual(len(result.ambulance[D, 'AM']), 2)
        self.assertEqual(len(result.campus[D, 'C']), 2)

    def test_rerun_does_not_accumulate_assignments(self):
        p = emt(ambulance={(D, 'AM')}, campus={(D, 'C')})
        self.solve([p])
        self.solve([p])
        self.assertEqual(p.assigned_hours, 6)
        self.assertEqual(p.campus_assigned_hours, 3)

    def test_empty_roster_and_unused_shifts(self):
        result = self.solve([], {(D, 'AM'): 'BLS'}, [(D, 'A')])
        self.assertEqual(result.ambulance[D, 'AM'], [])
        self.assertEqual(result.campus[D, 'A'], [])

    def test_later_timeout_preserves_previous_feasible_solution(self):
        model = cp_model.CpModel()
        assigned = model.new_bool_var('assigned')
        first = cp_model.CpSolver()
        timeout = Mock()
        timeout.solve.return_value = cp_model.UNKNOWN
        timeout.status_name.return_value = 'UNKNOWN'
        timeout.wall_time = 0.01
        with patch('solver.cp_model.CpSolver', side_effect=[first, timeout]):
            values, stages = _optimize(model, [('coverage', assigned), ('hours', assigned)], 5, 1)
        self.assertEqual(values[assigned.index], 1)
        self.assertEqual([s.status for s in stages], ['OPTIMAL', 'UNKNOWN'])

    def test_feasible_stage_is_not_reported_as_proven_optimal(self):
        model = cp_model.CpModel()
        assigned = model.new_bool_var('assigned')
        real = cp_model.CpSolver()
        original_solve = real.solve
        def feasible_status(model):
            original_solve(model)
            return cp_model.FEASIBLE
        with patch.object(real, 'solve', side_effect=feasible_status):
            with patch('solver.cp_model.CpSolver', return_value=real):
                values, stages = _optimize(model, [('coverage', assigned)], 5, 1)
        self.assertEqual(values[assigned.index], 1)
        self.assertEqual(stages[0].status, 'FEASIBLE')


if __name__ == '__main__':
    unittest.main()
