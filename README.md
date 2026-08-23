# BEMS Scheduler

Shift scheduler for Brown EMS. Takes the availability Google Form's CSV
export and produces a formatted `.xlsx` schedule for a block, covering:

- **Ambulance shifts** — weekday AM (0700–1300) / PM (1300–1900) / NIGHT
  (1900–0700), weekend DAY (0700–1900) / NIGHT. Driver certifications:
  EVDT (Rescue 1) > Authorized (Utility 1) > EMT.
- **Campus Response** — weekday 3h blocks A–D (0700–1900), staffed by BERT
  members and ambulance EMTs.

## Setup

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

> Anaconda note: OR-Tools crashes on import under Anaconda Python on macOS
> (protobuf symbol clash). Use a plain python.org/Homebrew Python for the
> venv, e.g. `/opt/homebrew/bin/python3.11 -m venv .venv`.

## Usage

1. Export the form responses sheet as CSV (see [FORM_GUIDE.md](FORM_GUIDE.md)
   for how the form must be structured).
2. Edit `config.json` (block dates, hour requirements, ALS shifts).
3. Run:

```bash
.venv/bin/python main.py
```

Output: `schedule_output.xlsx` with five sheets — Schedule, Campus Response,
Hour Summary, Warnings, Strike List — plus a console report.

To independently re-check every hard rule against a fresh solve:

```bash
.venv/bin/python verify_schedule.py
```

## How scheduling works

Both schedules are solved with [CP-SAT](https://developers.google.com/optimization/cp)
(Google OR-Tools), a constraint solver: every possible (person, shift)
assignment is a variable, the rules are constraints, and the solver searches
for the assignment that maximizes a prioritized objective — considering the
whole block at once instead of filling shifts greedily one at a time.

**Hard rules (never violated):**

- People are only assigned to slots they marked available.
- Crew caps: weekday AM/PM = 2, NIGHT = 3, big weekend (Fri NIGHT → Sun DAY) = 4.
- Hour requirement is also a hard maximum — nobody is scheduled above it.
- Rest: max 12h continuous (AM+PM is allowed), then at least 12h off — a
  NIGHT excludes all daytime shifts that day and the next; NIGHT→NIGHT is fine.
- ALS shifts always hold one seat only an EVDT can fill (left open otherwise).
- Campus blocks never overlap the same person's ambulance shifts.

**Priorities (highest first):**

1. Every shift has at least one person.
2. Everyone reaches their required hours.
3. An EVDT on every ALS shift; a driver on every night/weekend shift.
4. Maximize total filled crew slots — with hours capped, this automatically
   prefers giving people 6h shifts over 12h shifts, spreading coverage.
5. Spread people out (avoid same-day/back-to-back days) and equalize any
   unavoidable hour shortfalls.

Anything that still falls short lands on the **Warnings** sheet (unfilled
shift, ALS without EVDT, crew without a driver, member under hours), so the
gaps that remain are provably unavoidable given submitted availability —
they need a human conversation, not a better algorithm.

## Configuration (`config.json`)

| Key | Meaning |
| --- | --- |
| `block_start` / `block_end` | Block dates, inclusive (ISO `YYYY-MM-DD`). |
| `form_csv` | Path to the form's CSV export. |
| `output_xlsx` | Output workbook path. |
| `master_schedule_export` | Optional flat CSV export for the Master Schedule. Set its `block`, `daynum_start`, and `vehicle` from the target block; F26B1 starts at `0810`. |
| `hours.ambulance_emt` | Required (= max) ambulance hours per EMT this block. |
| `hours.campus_emt` / `hours.campus_bert` | Required (= max) campus hours per role. |
| `campus_responders_per_block` | Staffing target per campus block. |
| `solver_time_limit_s` | Solver budget per stage (30s is plenty; raise for huge blocks). |
| `availability_requirements` | Minimum shifts/blocks people must *submit* (drives the Strike List). |
| `als_shifts` | `"YYYY-MM-DD:DAY"` / `"YYYY-MM-DD:NIGHT"` entries; weekday `DAY` covers AM+PM. |
| `blackout_periods` | `{start_date, start_shift, end_date, end_shift}` ranges removed from the block. |

`master_schedule.csv` has the exact column order used by the Master Schedule:
`Block, ShiftID, Date, Shift, Vehicle, Seat, Requires, Assigned/Name`.
It is deliberately a local export only; review it, then paste the rows into the
Sheet. The scheduler never calls the Google Sheets or Google Calendar APIs.

For F26B1, `config.json` is set to August 10–September 3, 2026 and a day
number start of `810`. Populate `als_shifts` from the approved F26B1 ALS list
before each live run—the historical Spring list was removed rather than reused.
The parser accepts both the legacy separate-shift grids and the current
weekly-grid form format, so the availability-format decision does not require
code changes as long as the documented header phrases are retained.

## Files

| File | Role |
| --- | --- |
| `main.py` | Pipeline: parse → validate → solve → export. |
| `models.py` | Shift/block calendar, crew caps, rest rules, people dataclasses. |
| `parse_form.py` | Google Form CSV → people (handles both form formats, dedup, blackouts). |
| `validate.py` | Strike list + availability summaries (thresholds from config). |
| `ambulance_solver.py` | CP-SAT ambulance scheduler. |
| `campus_solver.py` | CP-SAT campus response scheduler. |
| `output.py` | xlsx export + console summary. |
| `verify_schedule.py` | Independent checker for every hard rule. |
| `FORM_GUIDE.md` | How to build the Google Form so it parses cleanly. |
