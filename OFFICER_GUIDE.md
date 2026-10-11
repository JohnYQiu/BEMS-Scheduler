# BEMS Scheduler: an officer's guide

This guide is for officers who do not write code. The scheduler turns members' availability into a proposed schedule. It does **not** decide who the supervisors are, read Google Sheets live, or publish shifts for you.

## Where each piece of information comes from

1. **Member availability and qualifications:** download the Form Responses sheet as CSV (or a ZIP containing one CSV). The program reads the saved file. If the form question wording changes, the parser may need maintenance.
2. **Supervisor shift type:** an officer checks the supervisor calendar and marks each ambulance shift **ALS** or **BLS** in the local editor. This is an input, not something inferred from the member form. Weekdays can have separate AM and PM choices.
3. **Scheduling policy and forced crews:** the selected block's config file stores hour limits, capacities, priority order, optional supported rules, dates, required crews, and file locations. The browser editor changes that file, not Python source code.
4. **The scheduling engine:** the Python solver uses those inputs and the rules implemented in the code. It checks availability, qualifications, hour caps, no overlaps, and rest before weighing preferences. A separate validator checks the resulting assignments.
5. **Outputs:** local Excel workbooks and, if enabled, a Master Schedule CSV. Review these before using them anywhere else; no automatic Google Sheet update occurs.

## Hard rules versus priorities

A **hard rule** is a boundary the schedule cannot cross, such as assigning someone who was unavailable or exceeding their configured hours. Some limits are already editable: hours and Campus/Wellness capacity. The supported rule library includes a weekly Wellness limit and an optional weekly cap on ALS shifts for each EVDT. The EVDT cap uses Driver Status from the saved responses or an officer's configured override and may reduce ALS coverage. Other safety/credential/rest rules are deliberately fixed until the solver and validator are both updated and tested.

A **priority** is a preference among valid schedules. The solver tries the top goal first, preserves its attained result, then moves to the next goal. For example, if ambulance coverage ranks above Campus coverage, it will not give up an attained ambulance-coverage result just to improve Campus coverage. Reordering changes this tradeoff. You can also create a **custom soft priority** by choosing a service, a count (assignments, shifts covered by a matching member, or hours), filters for provider, weekday, shift type, and member role/Driver Status, and whether to maximize or minimize. For example, prioritize an EVDT on ALS nights or drivers on BLS weekend shifts. Driver Status is normalized from the local form export as EVDT, Auth, or non-driver EMT, with configured overrides taking precedence. The solver only considers eligible assignments; custom preferences never waive hard rules. If a driver-dependent priority or EVDT cap is enabled, the export must have a Driver Status column and recognized answers; blank or ambiguous answers require a corrected export or a verified per-person override. These are **not numerical weights**, free-form formulas, or arbitrary code. All 18 built-in goals remain in the order alongside your custom ones. A goal with no matching eligible assignments may be skipped. A high-priority minimize goal can reduce coverage, and extra stages share the time budget, so review Warnings and Solver.

**Default behavior:** without `custom_priorities`, `objective_order`, or optional `rules`, the solver keeps its historical priority order and one Wellness shift per person per week. The EVDT weekly cap is off unless added. Simply opening the editor changes nothing. Saving adds explicit choices to the config file and creates a backup.

## Typical block workflow

1. Obtain the latest Form Responses CSV/ZIP. Keep it local and set its path in the editor.
2. Use a separate config file for the block. Set its start/end dates and output filenames.
3. Check the supervisor calendar. In the editor, set any new unassigned shifts to BLS, then change the ALS shifts. Review every day. Do not guess missing classifications.
4. For a busy ambulance shift, use **Forced crews** to pick required members from the eligible list derived from the saved form export. Save dates and form path first if you changed them. A forced member must be available in the current responses; other seats can still be filled by the solver. Do not force someone just because they may become available later.
5. Review hour limits, capacities, and priorities with the officer responsible for scheduling policy. Add a custom priority only when you can explain its filters and likely tradeoff. Add only supported hard rules that you intend to change. The editor validates the config and keeps a backup on save.
6. Run the normal check command, then generate outputs using the **same config file**:

   ```sh
   .venv/bin/python main.py config_next_block.json --check
   .venv/bin/python main.py config_next_block.json
   ```

7. Read the workbook's Schedule, Campus Response, Hour Summary, Warnings, and Solver tabs. A valid schedule can still have open shifts or shortfalls. `FEASIBLE` means a valid result was found but not proven best; `OPTIMAL` means the goal was proven best given earlier attained results. A time limit can leave uncertainty. Review, correct inputs if needed, and rerun before publishing.

## What the editor can and cannot do

- It runs only on your computer at `127.0.0.1`. It saves a local config file and a unique `.bak-...` backup. It never runs the solver itself.
- A forced crew is a set of **required** ambulance assignments, not a frozen entire shift. The scheduler fills any remaining seats if useful. All forced members still count toward hour caps and must satisfy availability, credential, overlap, and rest constraints. If forces conflict, the run fails instead of quietly dropping them.
- It offers a catalogue of **implemented** fields and hard rules, plus a bounded builder for custom soft priorities. You can combine supported filters and metrics, not write arbitrary code or a new scoring formula. A genuinely new hard rule needs code in the config loader, solver, and validator, plus tests. Never place a proposed rule in the config and assume it is enforced.
- The current solver does not model Utility vehicle assignments. Do not interpret a plain-language reference to Utility as proof that it created those assignments.
- Keep response exports and generated schedules private. They are excluded from Git, but avoid placing them in shared folders or emailing them casually.

## Installation

On macOS, after obtaining the project folder from a trusted source, double-click **Install on Desktop.command** in that folder. It copies the scheduler to a new Desktop folder, installs pinned Python dependencies in a local virtual environment, and creates desktop launchers. It does not overwrite an existing installation or fetch the scheduler from the internet. Installing dependencies requires an internet connection and a compatible Python installation. Follow the installer's messages if a prerequisite is missing.

For manual setup or troubleshooting, see the main [README](README.md). Windows support is not included in this installer yet.
