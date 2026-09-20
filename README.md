# BEMS Scheduler

Creates an ambulance and Campus Response schedule from the CSV exported from a
Google Form's **Form Responses** sheet. A ZIP containing that CSV also works.
The schedule includes volunteer assignments; supervisors are supplied separately.
Wellness Wagon scheduling and availability strikes are handled elsewhere.

**Workflow: export responses → edit `config.json` → run → review the output.**
The program reads a downloaded file. It does not connect to or update Google Sheets.

## 1. Set up the project once

These commands are for macOS/Linux and have been tested with Python 3.11.
If you already have this project, open Terminal in the `BEMS-Scheduler` folder
(the folder containing `main.py` and `config.json`) and skip the download commands.

To download a new copy:

```sh
git clone https://github.com/JohnYQiu/BEMS-Scheduler.git
cd BEMS-Scheduler
```

Then create the project's Python environment and install its dependencies:

```sh
python3 --version
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

If `python3` is missing, install Python 3.11 before continuing. If `.venv` already
exists and the dependencies are installed, you can skip setup. All commands
below run from the project folder; there is no need to activate the environment.

## 2. Save the responses for this block

Export the **Form Responses** tab as a CSV. Keep the original question headers
and response rows; no manual reformatting is needed. The parser supports the
existing BEMS legacy, shopping-period, and Block 2 forms. A redesigned form may
need parser changes.

Create an input folder if it does not exist:

```sh
mkdir -p inputs
```

Put the exported CSV inside `inputs`, rename it `form_responses.csv`, and set
this entry in `config.json`:

```json
"form_csv": "inputs/form_responses.csv"
```

**Already have a ZIP?** You can use it without extracting it. The current Block 2
configuration points to `inputs/form_responses.csv.zip`. The ZIP must contain exactly one
CSV. Choose either a plain CSV or a ZIP and make `form_csv` match the file you use.
Any filename is fine; paths are relative to the folder containing `config.json`.

Response files are private local inputs and are not included in a Git checkout.
For new or corrected responses, download a fresh export and replace the input
before running again. An older block's CSV has the right kind of format but the
wrong availability dates.

## 3. Configure the block

Open `config.json` in a text editor. Update these settings for every new block:

| Setting | What to enter |
| --- | --- |
| `block_start`, `block_end` | First and last schedule dates, inclusive, as `YYYY-MM-DD`. |
| `form_csv` | Path to this block's CSV or ZIP from step 2. |
| `hours.ambulance_emt` | Ambulance hours per EMT; currently `18`. |
| `hours.campus_emt` | Separate Campus Response hours per EMT; currently `6`. |
| `hours.campus_bert` | Campus Response hours per campus-only ERT/BERT member; currently `9`. |
| `shift_providers` | Every date's manually supplied ALS/BLS supervisor type, as explained below. |
| `master_schedule_export.block` | Block label used in exported shift IDs; currently `F26B2`. |
| `master_schedule_export.daynum_start` | First day's sequential ID number; currently `921`. See export details below. |

The hour settings are both targets and hard maximums. A person may finish below
target and will be reported; the solver never exceeds the maximum. Ambulance
hours must be a multiple of 6 and campus hours a multiple of 3.

The checked-in configuration is for **September 21–October 18, 2026**, with that
block's supplied ALS roster and BLS on all remaining shifts. For another block,
replace the dates and provider entries rather than reusing that roster.

### Enter ALS/BLS for every day and night

Inside `shift_providers`, include `DAY` and `NIGHT` entries for **every date** in
the block. For example, the entries for one date look like this:

```json
"2026-09-21:DAY": "ALS",
"2026-09-21:NIGHT": "ALS"
```

This is an example for one date, not a complete block configuration.

- On Monday–Friday, `DAY` applies to both AM (07:00–13:00) and PM (13:00–19:00).
- On Saturday–Sunday, `DAY` is the full 07:00–19:00 ambulance shift.
- `NIGHT` starts at 19:00 on the named date and ends at 07:00 the next morning.
- Use only `"ALS"` or `"BLS"`. Missing entries stop the run; BLS is not assumed.

If weekday AM and PM have different providers, replace that date's `DAY` entry
with separate `AM` and `PM` entries:

```json
"2026-09-21:AM": "ALS",
"2026-09-21:PM": "BLS",
"2026-09-21:NIGHT": "ALS"
```

Do not keep `DAY` alongside `AM`/`PM` for the same date. Remove entries outside the
new block. JSON needs double quotes and commas between entries, with no comma
after the final entry and no comments.

Leave the other settings at their existing values unless you need an optional
correction, blackout, different output location, or longer solve time.

## 4. Generate and review the schedule

Run:

```sh
.venv/bin/python main.py
```

The program reads responses, solves ambulance and campus assignments together,
and independently checks the resulting assignments before exporting. A successful
run prints `All hard scheduling rules verified against the resulting assignments.`
and the paths to the output files.

| Output | What to review |
| --- | --- |
| `outputs/schedule.xlsx` | Open in Excel or import into Google Sheets. **Schedule** shows ambulance assignments; **Campus Response** shows campus assignments; **Hour Summary** lists each person's hours and shortfalls; **Warnings** lists staffing/driver/hour issues; **Solver** records how thoroughly each objective was solved. |
| `outputs/master_schedule.csv` | Volunteer seat rows for the Master Schedule. Review before importing into another system. |

The `outputs` folder is created automatically. Rerunning overwrites files at the
configured output paths, so save a copy first if you want to keep an earlier run.
Nothing is uploaded or published automatically.

**Passing validation does not mean every staffing target was met.** Review
Warnings and Hour Summary before using the schedule:

- `NO AMBULANCE EMT`: no volunteer EMT was assigned alongside the supervisor.
- `ALS — NO EVDT`: no assigned EVDT to drive while the ALS provider treats during transport.
- `NO UTILITY DRIVER`: a BLS weekend shift lacks a qualified volunteer for split crew; the supervisor can still drive the ambulance.
- `AMBULANCE UNDER HOURS` / `CAMPUS UNDER HOURS`: a person is below that service's target.
- `SOLVER LIMIT`: the solver returned a valid schedule but did not prove all priority stages optimal. Try increasing `solver_time_limit_s` from `30` to `60`.

To solve and validate **without writing output files**, use this instead:

```sh
.venv/bin/python main.py --check
```

## If something goes wrong

| Message or problem | What to do |
| --- | --- |
| File not found | Check `form_csv` and the actual filename/location. Private inputs must be supplied separately after cloning. |
| No scheduling columns match the configured block dates | Check that the input is this block's response export and that the configured dates match it. |
| Choose ALS or BLS / invalid or outside the block / duplicate provider setting | Complete the provider list, remove old dates, and use either `DAY` or `AM` + `PM` for each weekday. |
| JSON error | Check quotes and commas in `config.json` and any optional local JSON files. |
| Missing Python module | Run `.venv/bin/python -m pip install -r requirements.txt`, then use `.venv/bin/python main.py`. |
| Unknown member role / response timestamp cannot be parsed | Check the original response export. Changed form wording or timestamp formats may require a parser update. |
| Solver fails with `INFEASIBLE` | Check optional locked assignments for conflicts with availability, hour caps, or work/rest rules. |
| Solver fails with `UNKNOWN` | Increase `solver_time_limit_s` and retry. |
| Cannot write the workbook | Close it in Excel, check the output folder is writable, and rerun. |

## Scheduling rules

- Availability is required. For EMTs, AM makes campus A and B individually
  eligible; PM makes C and D eligible. This does not require assigning both
  campus blocks. Campus-only members use their submitted A–D availability.
- Campus blocks run on weekdays: A 07:00–10:00, B 10:00–13:00,
  C 13:00–16:00, D 16:00–19:00. Campus-only members never count as ambulance EMTs.
- Multiple assignments on the same day must be contiguous, with no overlaps,
  at most 12 continuous hours, and at least 12 hours off between work periods.
  AM+C, AM+C+D, A+B+PM, and AM+PM are allowed; AM+D and A+PM are not.
- A NIGHT assignment excludes daytime work on its starting date and the next
  date. Consecutive nights have 12 hours off and are allowed when hour caps permit.
- Ambulance and campus hours are counted separately. There is no preference for
  spreading or clustering attendance days.
- Volunteer ambulance capacity is 2 for AM/PM, 3 for ordinary nights, and 4 for
  Friday/Saturday nights and Saturday/Sunday days. Supervisors are additional.
- `campus_capacity` is a maximum per block, currently 2, not a required minimum.
  There is no warning just because a campus block has fewer than two responders.

### Driving and priorities

Every shift has a supervisor who can drive the ambulance. On ALS shifts, an
EVDT lets the ALS provider treat in the back during transport; one volunteer
seat is reserved for an EVDT and stays open if none is assigned. An Auth cannot
fill that ALS driving role.

On Friday/Saturday nights and Saturday/Sunday days with a BLS supervisor, an
Auth or EVDT can drive Utility for split crew while the supervisor drives the
ambulance. They are equally eligible for Utility. Weekday BLS shifts have no
driver preference. EVDTs remain eligible for campus work under the usual rules.

Weekend staffing uses exactly these four shift categories:

| Priority | Staffing tier | Shift |
| --- | --- | --- |
| 1 | Nights | Friday NIGHT |
| 2 | Nights | Saturday NIGHT |
| 3 | Days | Saturday DAY |
| 4 | Days | Sunday DAY |

Friday daytime and Sunday night are ordinary shifts for staffing purposes.
A NIGHT shift belongs to the date when it starts at 19:00, even though it ends
at 07:00 the next morning.

A BLS shift is ready for split crew with **three volunteers, including at least
one Auth or EVDT**: supervisor + EMT on the ambulance, driver + EMT on Utility.
An ALS split crew needs three volunteers including an EVDT for the ambulance
and a **second, distinct Auth or EVDT** for Utility.

**Nights are a higher staffing tier than days.** The solver builds Friday-night
split crews first, then Saturday-night split crews. A fourth Friday volunteer
cannot take priority over the third person needed for Saturday-night split crew.
Within each night category, completing a three-person split crew takes priority
over adding fourth volunteers to other shifts in that category.

Once night split crews are protected, ALS driving takes priority over fourth
night volunteers. This preserves the rule that an EVDT can cover weekday ALS
once a BLS night already has three people and a qualified driver. Remaining
fourth night seats are then filled before building larger weekend day crews.
Day crews cannot displace higher-priority night staffing.

The solver applies these priorities in order, preserving each earlier attained
result while choosing later assignments:

1. Ambulance shifts with at least one EMT alongside the supervisor.
2. Friday nights: ALS EVDT coverage, qualified split crews, staffing toward three volunteers, and BLS Utility driving.
3. Saturday nights: the same goals, after Friday-night results are protected.
4. Remaining ALS driving: Saturday days, Sunday days, then other ALS shifts.
5. Extra night volunteers up to the four-person capacity: Friday nights, then Saturday nights.
6. Weekend day split crews and staffing toward three: Saturday days, then Sunday days. Drivers must qualify for each vehicle.
7. Extra day volunteers up to capacity: Saturday days, then Sunday days.
8. Ambulance hours within individual caps.
9. Campus blocks with at least one responder, then campus hours within individual caps.

These are separate optimization stages, not small score bonuses: extra day
staffing cannot compensate for a worse attained night result. The Solver sheet
reports each category separately. Weekday shifts receive basic coverage, but
filling all weekday seats is not an objective. A person may stay below their
hour target when higher priorities require it; all hour, availability, overlap,
and rest limits remain hard constraints.

`solver_time_limit_s` is a total search budget shared across stages. `OPTIMAL`
means a stage was proved optimal given earlier attained results; `FEASIBLE`
means a valid result was found without that proof. A later timeout retains the
last complete valid schedule and flags the limit. Open shifts and shortfalls
are not necessarily unavoidable.

For repeatable runs with the same input and dependency versions, set
`solver_workers` to `1` and allow enough solve time. The default of `8` workers
can choose different schedules with equally good objective values.

## Optional corrections and export settings

These files are optional; you do not need to create empty copies to run:

- `driver_status_overrides.local.json`: email-to-credential corrections, for example
  `{"person@example.com": "EVDT"}`. Values may be `EVDT`, `Auth`, or `EMT`.
- `locked_assignments.local.json`: assignments to preserve, for example
  `[{"date": "2026-09-21", "shift": "AM", "email": "person@example.com"}]`.
  Use ambulance shift names or individual campus blocks A–D. Locks must satisfy
  availability and all work rules. On weekdays, lock AM/PM individually.

`blackout_periods` in `config.json` excludes service periods for everyone.
Leave it as `[]` unless needed. For example, this removes September 25 daytime
and night shifts, including overlapping campus blocks:

```json
"blackout_periods": [
  {
    "start_date": "2026-09-25",
    "start_shift": "AM",
    "end_date": "2026-09-25",
    "end_shift": "NIGHT"
  }
]
```

The form's difficulties answers also support structured unavailability such as
`9/25`, `9/25-9/27`, or `9/25 NIGHT; 9/28 AM`. Separate entries with semicolons.
Review free-text notes yourself; arbitrary prose is not reliably interpreted.
Duplicate responses use the latest timestamp for each email, including role changes.

`output_xlsx` sets the workbook path. `master_schedule_export.path` sets the CSV
path; set `master_schedule_export.enabled` to `false` if you only want the workbook.
The CSV columns are `Block, ShiftID, Date, Shift, Vehicle, Seat, Requires, Assigned/Name`.
ALS ambulance Driver rows use `R1`/`EVDT`; Utility Driver rows use `U1`/`AUTH`
(EVDT also qualifies); other volunteer seats use `CREW`. Campus rows use `CR`.
Supervisors are not duplicated in these volunteer rows.

`daynum_start` is a sequential identifier, not a changing calendar date: starting
at `921`, September 30 is `930` and October 1 is `931`. Coordinate this with the
Master Schedule's ID convention when setting up a new block. Vehicle/seat IDs
reflect the current driving roles; review the whole export before replacing
rows from the older R1-only export format.

## What each file is for

| File or folder | Purpose |
| --- | --- |
| `main.py` | Command-line entry point: load, solve, validate, export. |
| `configuration.py` | Read settings and expand the manual ALS/BLS roster. |
| `parse_form.py` | Read CSV/ZIP responses, availability, identities, and credentials. |
| `models.py` | Shared people, shift times, hour caps, and schedule data. |
| `solver.py` | Choose ambulance and campus assignments together. |
| `validation.py` | Independently check the produced assignments before export. |
| `output.py` | Write the workbook and Master Schedule CSV. |
| `config.json` | Settings and provider roster for the current block. |
| `requirements.txt` | Pinned solver and Excel-export dependencies. |
| `tests/` | Checks for parsing/configuration, scheduling rules, validation, and exports. |
| `README.md`, `.gitignore` | Instructions and rules excluding local/generated files from Git. |
| `inputs/`, `outputs/` | Private local responses and generated schedules. Not committed. |
| `.venv/` | Installed Python environment. Local and required for the commands above. |
| `.git/` | Repository history and Git configuration. |

AI-tool settings such as `.claude/`, Python caches, and macOS metadata are not
needed to schedule. They are ignored by Git if tools recreate them. Historical
workbooks can be kept under `outputs/archive/` to distinguish them from new runs.

Run the existing checks after changing code:

```sh
.venv/bin/python -m unittest discover -s tests -v
```
