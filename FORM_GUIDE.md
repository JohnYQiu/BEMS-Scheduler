# Google Form Guide

How to build the availability form so that (a) students can fill it out day by
day against their Google Calendar, and (b) the CSV export parses with zero
manual cleanup. The parser ([parse_form.py](parse_form.py)) reads this format
*and* the legacy per-shift-type format, so old response sheets still work.

## Design principles

- **One checkbox grid per week, rows = days, columns = shifts.** Students plan
  out of Google Calendar's week view; the form should mirror it. They open
  their calendar next to the form, walk Monday→Sunday once per week, and are
  done — no revisiting the same date across three different questions (the old
  form's day/night/weekend split).
- **Every row requires an answer.** Turn on "Require a response in each row"
  and keep a `NOT AVAILABLE` column. An untouched row is indistinguishable
  from a forgotten one; forcing a choice catches half-finished submissions.
  If someone checks a shift *and* `NOT AVAILABLE`, the parser trusts the shift.
- **Keep dates in row labels, times in column labels.** The parser keys on the
  bracketed row label (`[Mon 4/27]`) that Google appends to each exported
  column header, and on the shift token (`AM`, `PM`, `NIGHT`, `DAY`, `A`–`D`)
  at the start of the column label. Use 24h times like `0700-1300` in labels —
  never `7a`/`1p` style, which can confuse token parsing.

## Form structure

**Settings:** collect email addresses (verified). This produces the
`Email Address` column the parser uses to deduplicate (latest submission per
email wins — tell students they can just resubmit to make corrections).

### Section 1 — Everyone

1. The SOG/requirement acknowledgements (unchanged; parser ignores them).
2. **"Are you an ambulance EMT or BERT member?"** — multiple choice, exact
   options, with section branching:
   - `Ambulance EMT (EMT only & EMT/ERT dual-role)` → Section 2
   - `BERT Member Only` → Section 3

### Section 2 — Ambulance EMT

1. `Last Name` (short answer), `First Name` (short answer)
2. **`Driver Status`** — multiple choice:
   - `EVDT - Rescue 1/Utility 1`
   - `Authorized - Utility 1`
   - `Not a driver`
3. **One checkbox grid per week**, question title:

   > `Week 1 — Ambulance Availability (Apr 27 – May 3)`

   The words **"Ambulance Availability"** must appear in the title — that is
   what the parser keys on. Number the weeks for students; the parser doesn't
   care.

   - **Rows:** one per date, labeled `Mon 4/27`, `Tue 4/28`, … `Sun 5/3`
     (day-name + M/D — the parser reads the M/D).
   - **Columns** (exact labels):
     - `AM (0700-1300) — weekdays`
     - `PM (1300-1900) — weekdays`
     - `DAY (0700-1900) — weekends`
     - `NIGHT (1900-0700)`
     - `NOT AVAILABLE`
   - Turn **on** "Require a response in each row".

   A weekend-only column checked on a weekday row (or vice versa) is ignored
   by the parser, so stray clicks can't corrupt the schedule.

4. **`Do you foresee any difficulties that Personnel should be made aware of
   in completing your block requirements?`** — paragraph. Instruct students to
   list blackout dates as `5/3`, ranges as `5/3-5/6`, and specific shifts as
   `5/3 NIGHT`, separated by semicolons. The parser reads exactly that syntax
   and removes those slots from their availability.

### Section 3 — BERT

1. `Last Name`, `First Name` (their own copies — the parser expects the
   second occurrence of each).
2. **One checkbox grid per week**, question title:

   > `Week 1 — Campus Response Availability (Select Minimum of 1 A/B AND 1 C/D)`

   The words **"Campus Response Availability"** must appear in the title.

   - **Rows:** weekdays only (`Mon 4/27` … `Fri 5/1`).
   - **Columns** (exact labels):
     - `Block A (0700-1000)`
     - `Block B (1000-1300)`
     - `Block C (1300-1600)`
     - `Block D (1600-1900)`
     - `NOT AVAILABLE`
   - Turn **on** "Require a response in each row".

3. The same difficulties question (blocks work in the syntax too: `5/4 A B`).

## Exporting for the scheduler

Google Sheets (linked responses) → File → Download → CSV. Point `form_csv` in
[config.json](config.json) at the file. Nothing else to clean: duplicate
submissions, `NOT AVAILABLE` noise, and out-of-block dates are all handled.

## Why not something fancier?

Two alternatives considered and rejected for now:

- **Calendar (ICS) import / free-busy via Apps Script** — automatic, but
  requires every student to grant calendar access, and "free" on a calendar is
  not the same as "willing to take a 1900–0700 ambulance shift". Availability
  is a decision, not a fact about the calendar.
- **One multi-select question listing every date+shift** — parses trivially
  but is a 50-checkbox wall that students will underfill; per-week grids keep
  the day-by-day flow the corps prefers.

If the form's wording must change, keep the key phrases ("Ambulance
Availability", "Campus Response Availability", `Driver Status`, the bracketed
`[Day M/D]` row labels, and shift tokens at the start of column labels) and
the parser will keep up.
