# BEMS Scheduler

A local scheduling tool for BEMS officers. Download the latest Google Forms response export, set the block calendar and policy in the browser editor, run the scheduler, then review its proposed assignments. It does not connect to Google Forms, Sheets, or Drive, and it never publishes a schedule for you.

## Install the editor

Use the editor-enabled [`feature/custom-priorities` branch](https://github.com/noahlangbo/BEMS-Scheduler/tree/feature/custom-priorities), not `main`.

1. [Download the editor version as a ZIP](https://github.com/noahlangbo/BEMS-Scheduler/archive/refs/heads/feature/custom-priorities.zip) and extract it.
2. On macOS, double-click **Install on Desktop.command** in the extracted folder and confirm the prompt. The installer creates a separate `~/Desktop/BEMS Scheduler` folder and two Desktop launchers.
3. If Python 3.11, 3.12, or 3.13 is not already usable, setup downloads Python 3.13.16 for Apple Silicon or Intel, verifies the pinned SHA-256 checksum, and installs it for your user under `~/Library/Application Support/BEMS Scheduler/`. It does not use `sudo`, change system Python, or edit shell settings.
4. Keep an internet connection available for the Python runtime (when needed) and pinned dependency downloads. Windows is not supported by this installer.

The installer will not overwrite an existing `~/Desktop/BEMS Scheduler` folder. To test a fresh version when one is already installed, first make a backup and move the old folder aside yourself. The installer does not migrate local inputs, outputs, custom configs, or backups; keep and restore those deliberately. Existing Desktop shortcuts are not replaced.

## Use the browser editor

Double-click **Open BEMS Editor.command**. Press Return to use `config.json`, or enter another configuration filename in the project folder. A browser tab opens to a local-only editor. Leave the Terminal window open while editing; press Ctrl-C there to stop the editor.

The page is organized into these sections:

| Editor section | What you do |
| --- | --- |
| **Block dates** | Choose the first and last day. The coverage grid updates to that window. |
| **Supervisor ALS/BLS calendar** | Enter the supervisor calendar's ALS/BLS classification. Use **Fill blank shifts with BLS** as a starting point, then change each ALS shift. Weekdays have AM, PM, and NIGHT; weekends have DAY and NIGHT. This calendar is manual, not synced. |
| **Policy settings** | Choose the local response CSV or ZIP, output paths, hour targets and hard caps, campus capacity, Wellness settings, block label, and solver budget. Help text explains each supported setting. |
| **Forced crews** | Require selected available people on particular shifts. Other seats may still be filled by the solver. Availability, credentials, hour caps, rest, and capacity still apply. |
| **Supported rules** | Add only rules implemented by the scheduler. A setting written elsewhere is not automatically enforced. |
| **Scheduling priorities** | Reorder soft preferences, add supported custom priorities, or define hard requirements. Priorities are resolved in order, not treated as numeric weights. |
| **Save configuration** | Validate and save the selected local config. A unique backup is created first. Saving does not run the scheduler. |

### Hard requirements and priorities

A **hard requirement** is a must-have: for every matching shift, require a minimum of 1 to 4 matching people. Choose Ambulance or Campus Response, the member/driver filter, ALS/BLS when relevant, and optional weekdays or shift kinds. The editor warns when no shifts match, available candidates are too few, or a shift has no backup candidates. These are screening warnings only. Hour caps, rest, forced crews, and interactions can still make a requirement set infeasible. If the full schedule cannot satisfy all hard rules, the solver stops without exporting.

A **priority** is a preference among schedules that satisfy hard rules. Move an item higher to protect its result before lower items are considered. Custom priorities can count assignments, covered shifts, or hours using supported filters. A high-priority minimize preference can reduce coverage, so review the resulting warnings and Solver tab.

### How the screen works

The editor opens with an **Officer workspace** heading, a short path from form export to reviewed schedule, a local-save status, and a section navigation list. The date controls are followed by the calendar grid. The lower sections are cards for policy settings, forced crews, supported rules, and priorities. Hard requirements and custom priorities are built from labeled dropdowns and checkboxes, not code. At the bottom, **Save configuration** and **Discard unsaved edits** make the handoff explicit. A **How it works** section explains the data sources and what saving does.

## Prepare the availability file

Export the Google Form's **Form Responses** sheet as a CSV, or use a ZIP containing exactly one CSV. Keep the file on your computer. In **Policy settings**, select its path in **Form responses file**. The editor reads the selected local file when it offers people for forced crews. It does not fetch responses or sync edits to Google.

For a new block, use the current response export and a separate config file when appropriate. Check that the block dates, output paths, provider classifications, hour limits, and capacities match the block before saving. Do not guess ALS/BLS classifications or force someone who is not available in the saved responses.

## Run and review

Double-click **Run Scheduler.command**, press Return to use `config.json` or enter another config filename, then choose:

1. **Check only** to solve and validate without exporting files. This is the recommended first run after changing settings.
2. **Full run** to write the configured output files. Type `EXPORT` to confirm. A full run may replace files at those output paths.

Review the results before sharing or publishing:

- `outputs/schedule.xlsx`: check **Schedule**, **Campus Response**, **Hour Summary**, **Warnings**, and **Solver**.
- `outputs/master_schedule.csv`: volunteer seat rows for the Master Schedule, if enabled.
- `outputs/wellness_wagon.xlsx`: separate manual-entry Wellness grid, if enabled.

A valid result can still have unfilled shifts or hour shortfalls. `OPTIMAL` means the solver proved that stage best given earlier results; `FEASIBLE` means it found a valid schedule without proving the best result, often because the time budget ran out. Read **Warnings** and **Hour Summary** before deciding the schedule is ready. Outputs remain local unless you choose to share them.

## What this tool does not do

- It does not read a live Google Form, Sheet, or Drive file.
- It does not automatically choose ALS/BLS classifications from the supervisor calendar.
- It does not approve, publish, or upload a schedule.
- It does not model Utility vehicle assignments or confirm Utility readiness.
- It will not enforce a proposed rule unless that rule appears in the editor's supported controls and is implemented in the solver and validator.

For the scheduling rules and officer review guidance, see [OFFICER_GUIDE.md](OFFICER_GUIDE.md). For maintainers, run the test suite with `.venv/bin/python -m unittest discover -s tests -v`.
