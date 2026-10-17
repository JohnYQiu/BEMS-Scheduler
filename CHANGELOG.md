# Changelog

## Unreleased

- Fixed BLS ambulance capacity so the driver position is counted alongside crew seats. Weekday AM/PM shifts now allow Driver + C2; regular nights allow Driver + C2/C3; high-capacity weekend shifts allow Driver + C2/C3/C4.
- Kept every required BLS seat in the generated Master Schedule export, including an open Driver row when no qualified volunteer driver is assigned and open crew rows when seats remain unfilled. Existing assigned volunteers remain in their seat rows.
- Kept driver assignment as a preference, not a hard BLS eligibility constraint. When the shift would otherwise remain short, an assigned crew member may use the Driver row as a CREW override; that does not label the person as driver-qualified, and the supervisor remains the fallback driver.
- Added regression coverage for open weekday C2 seats with an assigned driver, open BLS driver seats, last-resort crew overrides, capacity-based scheduling, and the Campus Response labels/next-week view.

### Cause

The BLS daytime capacity was set to one volunteer, while the exporter used that value as the total number of positions and then reserved a driver position when a volunteer driver was assigned. The result was zero crew positions for a staffed weekday BLS ambulance, so no C2 row was generated and the solver could not assign a second volunteer. Separately, the exporter omitted a BLS Driver row when the volunteer driver position was empty. Both behaviors made required seats disappear from generated output.

This project generates local schedule files; the change does not edit the live Google Sheet or alter existing assignments in that sheet. The weekly view formula and giveaway release/cancellation automation are outside this Python repository and need their own source/workflow review.
