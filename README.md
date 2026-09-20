# BEMS Scheduler

Schedules ambulance and Campus Response together from a Google Form CSV or
single-CSV ZIP. Wellness Wagon scheduling and availability strikes are outside
this program.

## Run

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python main.py
```

Before each block, edit `config.json`. Set the dates, response file, hour caps,
and **every active shift's supervisor type** in `shift_providers`. The current
Block 2 configuration covers September 21–October 18, 2026, with the supplied
ALS roster and BLS for all remaining shifts. Review these choices for each new
block; every active shift must explicitly be `"ALS"` or `"BLS"`.

```json
"shift_providers": {
  "2026-09-21:DAY": "ALS",
  "2026-09-21:NIGHT": "ALS"
}
```

On weekdays, `DAY` sets both AM and PM. To configure different providers,
replace that DAY entry with separate AM and PM entries. Weekends use DAY
and NIGHT. Missing, invalid, overlapping, and out-of-block entries stop the
run. Supplied blackouts remove slots from both services where they overlap.

Paths resolve relative to the configuration file, so this also works from
another directory:

```sh
.venv/bin/python main.py /path/to/config.json
.venv/bin/python main.py config.json --check
```

`--check` runs the same solve and validation without exporting. A normal run
validates the exact resulting assignments before writing:

- `outputs/schedule.xlsx`: Schedule, Campus Response, Hour Summary, Warnings,
  and Solver sheets.
- `outputs/master_schedule.csv`: flat Master Schedule rows. Export only;
  review before importing into another system.

Inputs, outputs, and personnel overrides are Git-ignored. The original response
archive can stay in Downloads; the local Block 2 input is `inputs/block2.csv.zip`.

## Work rules

- Only assign submitted availability. For EMTs, AM availability makes A and B
  individually eligible for CR; PM makes C and D eligible. Eligibility does not
  require assigning both blocks. EMT/ERT dual-role members use the EMT quota.
- Campus-only ERT/BERT members use their explicit A–D selections. They never
  count as ambulance EMT coverage. Extra campus sections filled by an EMT do
  not create another person or override the inferred availability.
- One person per normalized email; the latest complete submission wins across
  roles. Unreadable timestamps or unknown roles stop parsing rather than silently
  selecting an older response.
- Maximum **12 continuous hours across both services**, with **12 hours off**
  between work periods. Multiple assignments on a day must be contiguous.
  AM+C, AM+C+D, A+B+PM, and AM+PM are valid. AM+D and A+PM are not.
- NIGHT excludes daytime ambulance and CR on that date and the next date.
  Consecutive nights have exactly 12 hours off and are allowed.
- There is **no preference for spreading or clustering assignments**.
- The hour requirements are also hard caps: EMT 18 ambulance plus 6 CR hours;
  campus-only members 9 CR hours. They are independently counted. Insufficient
  availability leaves a reported shortfall; caps are never exceeded.
- Volunteer crew caps remain 2 for AM/PM, 3 for normal nights, and 4 for Friday
  NIGHT, Saturday DAY/NIGHT, and Sunday DAY. Supervisors are supplied separately.
- `campus_capacity` is a maximum, not required staffing. There is no warning
  just because a CR block has fewer than two responders.

## Supervisors and driving

Every shift already has a supervisor who can drive the ambulance. An EVDT on
an ALS shift lets the ALS provider treat the patient in the back during transport.
An ALS shift therefore reserves one volunteer seat for an EVDT; it stays open
otherwise, and the missing EVDT is reported. Other EMTs may still be assigned.

On BLS weekends, the supervisor can drive the ambulance and an Auth can drive
Utility for split crew. EVDT and Auth are equally eligible for that Utility role.
There is no EVDT bonus just for being a weekend, and no driver preference on
weekday BLS shifts. A missing Utility volunteer is not a missing ambulance driver.
EVDTs remain eligible for CR under the same availability and work rules as other EMTs.

## Optimization

The whole block is one CP-SAT model. The stages below preserve the value reached
by each preceding stage; they do not freeze individual ambulance assignments.
Campus objectives can still rearrange ambulance assignments while preserving
the earlier results.

1. Maximize ambulance shifts with at least one EMT alongside the supervisor.
2. Maximize ambulance hours within the individual caps.
3. Maximize ALS Friday/Saturday nights with an EVDT.
4. Maximize ALS Saturday/Sunday days with an EVDT.
5. Maximize other ALS shifts with an EVDT.
6. Maximize BLS weekend shifts with a Utility-qualified volunteer (Auth or EVDT).
7. Maximize filled ambulance volunteer seats.
8. Maximize campus blocks with at least one responder.
9. Maximize campus hours within the individual caps.

`solver_time_limit_s` is one total search budget shared across stages. Stages
with no relevant variables are skipped. Each stage reports its status, attained
value, upper bound, and time. `OPTIMAL` means that stage was proved optimal
given earlier attained values. A `FEASIBLE` result can be retained at the time
limit, but is not a proof of optimality. If a later stage finds no solution in
its allocation, the last complete feasible schedule is retained and flagged.
Individual open shifts and shortfalls are never described as inherently unavoidable.

For reproducible runs with the same inputs and pinned dependencies, set
`solver_workers` to 1 and allow enough time to finish. Multiple workers improve
search speed but may choose different equally good schedules.

## Master Schedule export

The column order remains `Block, ShiftID, Date, Shift, Vehicle, Seat, Requires,
Assigned/Name`. The supervisor is not duplicated in volunteer rows.

- ALS ambulance Driver rows use `R1` and require `EVDT`.
- Weekend Utility Driver rows use `U1` and require `AUTH` (EVDT also qualifies).
- Other volunteer seats are `CREW`. A crew-only EMT is never labelled a driver.
- BLS shifts do not reserve a fictitious EVDT ambulance-driver seat. Weekday
  BLS rows are crew seats. Weekend open capacity may include a Utility opening.
- CR rows use `CR`, S1 through the configured capacity, and `CREW`; they do
  not imply that campus-only members are ambulance EMTs or require a driver.

Vehicle and seat IDs now reflect those roles. Regenerate/review the whole export
when importing; do not assume the previous R1-only Driver IDs still apply.
`daynum_start` preserves the existing sequential day-number convention, rather
than switching formats at the start of a month.

## Inputs and local corrections

The parser supports the existing legacy, weekly, shopping-period, and Block 2
headers. Use bracketed date labels such as `[Mon 9/21]`, an email column
(`Email Address` or `Username`), the role question, and the usual First/Last Name
fields. Ambulance cells contain AM/PM/DAY/NIGHT; campus cells contain A/B/C/D.
Wellness Wagon columns are explicitly ignored.

Existing structured blackout syntax in the difficulties field remains supported:
`9/25`, `9/25-9/27`, `9/25 NIGHT; 9/28 AM`, or campus block tokens. Separate
independent entries with semicolons. Arbitrary prose is not a structured rule;
personnel should review it separately.

Optional `driver_status_overrides.local.json` contains email-to-credential
corrections, with values EVDT, Auth, or EMT. Optional
`locked_assignments.local.json` contains assignments to preserve:

```json
[{"date": "2026-09-21", "shift": "AM", "email": "person@example.com"}]
```

Locks can use ambulance shift names or individual campus blocks. Unavailable or
conflicting locks fail clearly; they never override a hard work rule.

## Development

```sh
.venv/bin/python -m unittest discover -s tests -v
```

`main.py` orchestrates; `configuration.py` loads settings; `models.py` owns
people and times; `parse_form.py` reads responses; `solver.py` builds the joint
model; `validation.py` independently checks the resulting intervals and staffing;
`output.py` formats those validated assignments. Availability strikes are handled
elsewhere, so the former strike workflow and duplicate checker were removed.
