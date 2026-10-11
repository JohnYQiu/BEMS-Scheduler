"""Editor-facing descriptions of supported configuration controls.

A catalogue entry alone never creates a scheduling rule. Each optional rule
must also be implemented in the config loader, solver, and validator with tests.
The browser editor reads these descriptors so a new supported field does not
require rewriting the user interface.
"""

FIELD_CATALOG = [
    dict(path="/form_csv", label="Form responses file", group="Inputs", type="text",
         description="Local CSV or ZIP exported from the Google Form responses sheet. No live connection.", source="Google Form export, saved on this computer"),
    dict(path="/hours/ambulance_emt", label="Ambulance hours per EMT", group="Staffing limits", type="number",
         description="Target and hard maximum per EMT for this block; use a multiple of 6.", source="This block's policy configuration"),
    dict(path="/hours/campus_emt", label="Campus hours per EMT", group="Staffing limits", type="number",
         description="Separate Campus Response target and hard maximum per EMT; use a multiple of 3.", source="This block's policy configuration"),
    dict(path="/hours/campus_bert", label="Campus hours per campus-only member", group="Staffing limits", type="number",
         description="Campus Response target and hard maximum; use a multiple of 3.", source="This block's policy configuration"),
    dict(path="/campus_capacity", label="Campus responders per block", group="Staffing limits", type="number",
         description="Maximum assigned people in each Campus Response block.", source="This block's policy configuration"),
    dict(path="/wellness_wagon/enabled", label="Schedule Wellness Wagon", group="Wellness Wagon", type="boolean",
         description="Include Wellness shifts using the form's Wellness availability.", source="This block's policy configuration and Form responses"),
    dict(path="/wellness_wagon/capacity", label="Wellness people per shift", group="Wellness Wagon", type="number",
         description="Maximum assigned people on each Wellness shift.", source="This block's policy configuration"),
    dict(path="/solver_time_limit_s", label="Solver time budget (seconds)", group="Run settings", type="number", step="any",
         description="Total time shared across priority stages. A short run may be valid but not proven best.", source="This block's policy configuration"),
    dict(path="/solver_workers", label="Solver workers", group="Run settings", type="number",
         description="Use 1 for more repeatable runs; multiple workers can choose different equally good schedules.", source="This block's policy configuration"),
    dict(path="/output_xlsx", label="Schedule workbook path", group="Output files", type="text",
         description="Local destination for the schedule workbook.", source="This block's policy configuration"),
    dict(path="/wellness_wagon/output_xlsx", label="Wellness workbook path", group="Output files", type="text",
         description="Local destination if Wellness is enabled.", source="This block's policy configuration"),
    dict(path="/master_schedule_export/enabled", label="Export Master Schedule CSV", group="Output files", type="boolean",
         description="Create a local CSV for the Master Schedule; does not upload it.", source="This block's policy configuration"),
    dict(path="/master_schedule_export/path", label="Master CSV path", group="Output files", type="text",
         description="Local destination for the Master Schedule CSV.", source="This block's policy configuration"),
    dict(path="/master_schedule_export/block", label="Block label", group="Output files", type="text",
         description="Identifier included in exported shift IDs, such as F26B3.", source="This block's policy configuration"),
]

OPTIONAL_RULES = [
    dict(path="/rules/wellness_shifts_per_week", label="Wellness shifts per person per week",
         group="Supported rule library", type="number", default=1, min=1, max=5,
         description="Hard maximum Monday–Friday Wellness shifts for one person. The default is one.",
         source="This block's policy configuration; enforced by the solver and independent validator",
         optional=True),
]

FIELD_BY_PATH = {entry["path"]: entry for entry in FIELD_CATALOG}
RULE_BY_PATH = {entry["path"]: entry for entry in OPTIONAL_RULES}
