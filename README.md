# BEMS Scheduler

Creates an ambulance and Campus Response schedule from the CSV exported from a
Google Form's **Form Responses** sheet. A ZIP containing that CSV also works.
The schedule includes volunteer assignments; supervisors are supplied separately.
Wellness Wagon can be enabled as a separate, manual-entry workbook. It never changes the Master Schedule CSV or writes to Google Sheets.

**Workflow: export responses → edit `config.json` → run → review the output.**
The program reads a downloaded file. It does not connect to or update Google Sheets.

This `feature/custom-priorities` branch adds input-driven custom preferences and
an optional weekly EVDT limit to the local browser editor and macOS installer.
The stable Terminal scheduler without the editor is on [`main`](https://github.com/noahlangbo/BEMS-Scheduler),
and the pre-forced-crew version remains on
[`legacy-before-forces`](https://github.com/noahlangbo/BEMS-Scheduler/tree/legacy-before-forces).

## 1. Set up the project once

These commands are for macOS/Linux and have been tested with Python 3.11.
If you already have this project, open Terminal in the `BEMS-Scheduler` folder
(the folder containing `main.py` and `config.json`) and skip the download commands.

To download a new copy:

```sh
git clone --branch feature/custom-priorities --single-branch https://github.com/noahlangbo/BEMS-Scheduler.git
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
existing BEMS legacy, shopping-period, Block 2, and Block 3 forms, including the
optional EMT class-conflict question. If a later form changes question wording,
check the header mappings and parser before scheduling.

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
| `shift_providers` | Every date's ALS/BLS classification from the supervisor calendar, as explained below. |
| `master_schedule_export.block` | Block label used in exported shift IDs; the F26B3 config uses `F26B3`. |

The hour settings are both targets and hard maximums. A person may finish below
target and will be reported; the solver never exceeds the maximum. Ambulance
hours must be a multiple of 6 and campus hours a multiple of 3. Both F26B2 and
F26B3 are 28-day blocks, so F26B3 keeps the same per-block targets: 18 ambulance
hours per EMT, 6 Campus Response hours per EMT, and 9 Campus Response hours per
campus-only ERT/BERT member.

The default configuration is for **September 21–October 18, 2026**. The separate
`config_f26b3.json` is for **October 19–November 15, 2026** and has its own
provider map and output paths. For F26B3, only shifts supervised by Chief
Sanderson, Emily Binns, Jeff Magill, or Andrew Tetreault are ALS; all other
supervisor-calendar entries, including OPEN and Utility Ops, are BLS.

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
| `outputs/wellness_wagon.xlsx` | Separate weekday AM/PM Wellness Wagon grid for manual entry into the V3 monthly view. It is not part of the Master Schedule CSV. |

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
- Ambulance EMTs may optionally report intended class, lab, or conference
  sessions as `WEEKLY TUE 10:30-11:50` or `2026-10-27 13:00-14:30`. Any time
  overlap hard-blocks the affected Campus Response block, without changing
  Ambulance availability. Invalid nonempty entries stop parsing rather than
  silently allowing a conflicting assignment.
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

The supervisor calendar determines shift type. For F26B3, only Chief Sanderson,
Emily Binns, Jeff Magill, and Andrew Tetreault make a shift ALS; all other
supervisors and OPEN/Utility Ops entries are BLS in that block's supplied
classification. An ALS shift without an assigned EVDT remains classified ALS;
the output flags `ALS — NO EVDT` for officer review. It is not automatically
reclassified or treated as an operational status decision.

The R1 `Driver` seat is credentialed, not a separate vehicle. Put an EVDT there
whenever assigned, on either ALS or BLS. An AUTH driver may take that seat on a
BLS shift if no EVDT is assigned. On ALS shifts, AUTH stays in C2 or later and
the EVDT-required Driver seat remains open, preserving the chance for an EVDT to
pick it up. Remaining volunteers occupy C2, C3, and so on with `CREW` requirements.
The solver prefers EVDTs on ALS shifts when feasible, without overriding the
other hard constraints or staged scheduling priorities. Availability requests
are not filtered by driver credentials: non-drivers can still be assigned as
crew on remaining seats, including extra hands on ALS shifts while the EVDT
Driver seat remains open. A person in a crew seat is not treated as a driver.

EVDTs remain eligible for campus work under the usual rules.

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

The current solver **does not model Utility vehicle assignments or drivers**.
It optimizes ambulance volunteer seats and exports R1/C2+ rows. A three-person
weekend crew is a staffing target within the ambulance model, not a claim that
Utility is staffed or ready. Do not use the workbook as a Utility readiness check.

**Nights are a higher staffing tier than days in the default priority order.**
The solver builds Friday-night crews toward three volunteers, then
Saturday-night crews, before prioritizing Saturday/Sunday daytime crews.
An extra fourth volunteer comes after the higher-priority three-person goals.
The exact sequence of all 18 soft goals is shown in the browser editor and in
the workbook's Solver sheet. Officers may reorder those goals for a block
without changing any hard scheduling constraints.

The default solver applies these priority groups in order, preserving each
earlier attained result while choosing later assignments:

1. Ambulance shifts with at least one EMT alongside the supervisor.
2. Friday and Saturday nights: ALS EVDT coverage, then crew seats toward three.
3. Saturday and Sunday days: ALS EVDT coverage; then other ALS shifts with EVDT.
4. Saturday and Sunday day crew seats toward three.
5. Extra weekend volunteer seats: Friday night, Saturday night, Saturday day, Sunday day.
6. Ambulance hours within individual caps.
7. Campus blocks with at least one responder, then campus hours within individual caps.
8. Wellness Wagon shifts filled, if enabled.

These are separate optimization stages, not small score bonuses: extra day
staffing cannot compensate for a worse attained night result. The Solver sheet
reports each category separately. Weekday shifts receive basic coverage, but
filling all weekday seats is not an objective. A person may stay below their
hour target when higher priorities require it; all hour, availability, overlap,
and rest limits remain hard constraints.

### Wellness Wagon

Set `wellness_wagon.enabled` to `true` in `config.json` to create a separate
`outputs/wellness_wagon.xlsx` file. The scheduler reads the form's **Are you working
for the Wellness Wagon?** response and its weekday AM/PM Wellness Wagon availability
columns. It assigns at most the configured `capacity` people per Wellness shift and at
most one Wellness shift per person in each Monday–Friday week by default (the supported
`rules.wellness_shifts_per_week` option can change this from 1 to 5). Wellness is a final,
lower-priority objective: it cannot reduce the attained ambulance or Campus Response
results, does not count toward ambulance/CR requirements, and follows the same overlap
and 12-hour rest rules. It is intentionally excluded from `master_schedule.csv`.

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
The CSV columns match the live V3 Master Schedule: `ShiftID, Date, Shift, Vehicle, Seat, Requires, Assigned/Name`.
Shift IDs use the calendar date as MMDD, for example `F26B3-1019-AM-R1-EVDT`.
Driver rows use `R1`/`Driver`, with `Requires` set to `EVDT` or `AUTH` as
applicable; crew rows use `C2`, `C3`, and so on with `CREW`. Campus rows use `CR`.
Supervisors are not duplicated in these volunteer rows.

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

## Local browser policy editor (optional)

For a first-time macOS setup, open the `feature/custom-priorities` branch on GitHub,
choose **Code → Download ZIP**, extract the archive, then double-click
**Install on Desktop.command** in its top-level folder. Do not download `main` for
the editor; that branch contains only the Terminal scheduler. It creates a separate
`~/Desktop/BEMS Scheduler` folder, installs pinned dependencies in `.venv`,
and places **Open BEMS Editor.command** and **Run Scheduler.command** launchers
on the Desktop. It never overwrites an existing installation and does not copy
private inputs or generated outputs. Python 3.11–3.13 and internet access for
pinned package downloads are required; Windows needs manual setup. The installer
has been tested on a Mac with Python 3.11 via Terminal, including dependency
installation and a repeat-run no-overwrite check. Finder/Gatekeeper behavior on a
fresh Mac still needs a separate trial.
See [OFFICER_GUIDE.md](OFFICER_GUIDE.md) for a nontechnical explanation of
sources, constraints, priorities, forced crews, and review steps.

You can configure a block without editing Python or JSON by running this from the
project folder:

```sh
.venv/bin/python editor.py config.json
```

For another block, first make a separate config file (for example,
`cp config.json config_next_block.json`), then run
`.venv/bin/python editor.py config_next_block.json`. The editor opens only on
`127.0.0.1` on your computer. It does **not** run the solver, upload responses,
or publish the schedule. Keep the Terminal window open while using it; press
Ctrl-C when done.

In the browser, set the block dates, review settings, mark each ambulance shift
ALS or BLS, and reorder built-in or custom objective priorities if policy calls for it.
For a new date range, use **Set unassigned shifts to BLS**, then mark the ALS
shifts from the supervisor calendar. Weekday AM and PM may be classified
separately. Every active shift must be classified before saving. The editor
validates the full config and makes a unique `.bak-...` copy of the prior file
before saving. Review the displayed backup path if you need to restore it.

The **default priority order is exactly the previous solver order**. Without
custom priorities, behavior is unchanged. The editor can create a new soft
priority from supported ingredients: ambulance, Campus Response, or Wellness
assignments; count seats, shifts with at least one matching person, or hours;
filter by day of week, shift type, ALS/BLS (ambulance only), and member role or
Driver Status (any, EVDT, Auth, driver, non-driver, EMT, or campus-only BERT);
then maximize or minimize the result. Driver Status comes from the saved Form
Responses export, with configured overrides taking precedence. For example,
create **BLS weekend nights with a driver** using ambulance / covered shifts /
driver / BLS / Friday and Saturday / NIGHT / maximize. Move the new item above
or below the 18 existing objectives. All configured priorities must appear once;
a stage with no eligible matching assignments is skipped. Driver-dependent
priorities and the EVDT cap require a Driver Status column and recognized values
(EVDT, Auth/Authorized, or non-driver EMT); an unknown or blank answer stops the
run unless an officer supplies a validated per-person override. They remain
sequential lexicographic goals, **not numerical weights or an arbitrary-code
scripting language**. Minimize stages show the actual nonnegative count in the
Solver tab, with a lower bound rather than an upper bound. A high-priority minimize goal can lower coverage, so inspect the
Solver and Warnings tabs. More stages share the same solve time budget.

The **Hard requirements** builder can require a minimum number of matching, available people on every selected ambulance or Campus Response shift. Choose service, member/driver filter, ALS/BLS, weekday, shift kind, and minimum 1–4. These are enforced by the solver and independent validator. Before solving, the scheduler warns when a matching shift has too few candidates, exactly enough candidates (no backup), or the rule matches no shifts. This is an availability-only screening check; hour caps, rest, forced crews, or interactions can still make a rule set infeasible. The solver is authoritative and stops without exporting if the full hard-rule set cannot be satisfied. Avoid turning a desirable goal into a hard requirement unless failing the whole schedule is acceptable.

The **Supported rule library** includes the per-person weekly Wellness maximum
(default one) and an optional weekly maximum on ALS assignments for each EVDT.
The latter uses form-derived Driver Status and, if enabled, may leave ALS shifts
without an EVDT. Both are hard limits enforced by the solver and independent
validator. Neither overrides availability, capacity, hours, rest, or credential
rules. A new hard rule outside this catalogue still needs implementation and
tests; typing an invented key does not enforce it.

**Forced crews:** after saving the dates and form-response path, use the editor's
Forced crews section to choose a shift and members who appear available in the
latest local form export. Each listed member is required on that shift. The
solver may fill remaining seats, but cannot exceed capacity, hours, or rest
limits. If a required member is unavailable or forces conflict, the check fails
rather than silently dropping the force. The config stores these in
`forced_crews`; pre-existing individual `locked_assignments` remain supported.

After saving, use the normal commands with that same config file, for example:

```sh
.venv/bin/python main.py config_next_block.json --check
.venv/bin/python main.py config_next_block.json
```

The editor exposes existing hour/capacity settings, forced crews, supported
optional rules, and declarative custom soft priorities. It does not execute
free-form code or invent hard rules. A new hard rule, shift geometry, credential
policy, or scoring metric outside the supported vocabulary still requires a
reviewed code change and matching validation tests. Do not assume that moving a priority guarantees a
better schedule: compare the Solver, Warnings, and Hour Summary tabs, especially
when the solve reports `FEASIBLE` or a time limit.
